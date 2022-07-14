
$converter = @{
    33 = 'PGUP'
    34 = 'PGDN'
    45 = 'INS'
    48 = '0'
    49 = '1'
    50 = '2'
    51 = '3'
    52 = '4'
    53 = '5'
    54 = '6'
    55 = '7'
    56 = '8'
    57 = '9'
    106 = 'MULTIPLY'
    107 = 'ADD'
    109 = 'MINUS'
    110 = 'DECIMAL'
    111 = 'DIVIDE'
    186 = ';'
    187 = '='
    188 = ','
    189 = '-'
    190 = '.'
    191 = '/'
    192 = "'"
    219 = '['
    220 = '\'
    221 = ']'
    222 = '#'
    223 = '`'
}
$scanCodesModifiers = @{
    'Shift'   = '0x10'
    'Control' = '0x11'
    'Alt'     = '0x12'
}
$ignoreKeys = @(9,16,17,18,20,91,144,145)
function scancodes($element) {
    if (!($ignoreKeys).Contains($element[2])) {
        $scanCodes = @()
        $mods = ([string]$_.modifiers).Replace(', ', ' + ')
        if($converter.Contains($element[2])) {
            $key = $converter[$element[2]]
        }
        else {
            $key = $element[1]
        }
        if ($mods -ne 'None') {
            if([bool]$element[0].PSObject.Properties["Text"]) {
                $element[0].Text = "$mods + $key"
            }
            else {
                $element[0].Value = "$mods + $key"
            }
            foreach ($mod in $mods.split(' + ')) {
                $scanCodes += $scanCodesModifiers[$mod]
            }
        }
        else {
            if([bool]$element[0].PSObject.Properties["Text"]) {
                $element[0].Text = $key
            }
            else {
                $element[0].Value = $key
            }
        }
        $scanCodes += '0x' + ('{0:x}' -f $element[2]).ToUpper()
        return $scanCodes -join ','
    }
}
function expand {
    if ($expandCollapse.Text -eq 'Collapse') {
        $syncHash.Window.Height = 66
        $syncHash.Window.Width = 140
        $syncHash.Window.FormBorderStyle = 'None'
        $expandCollapse.Text = 'Expand'
        $expandCollapsePanel.Left = -80
        $FormTabControl.Visible = $false
    }
    else {
        $syncHash.Window.Height = 470
        $syncHash.Window.Width = 570
        $syncHash.Window.FormBorderStyle = 'Fixed3D'
        $expandCollapse.Text = 'Collapse'
        $expandCollapsePanel.Left = 370
        $FormTabControl.Visible = $true
    }
}
function loadKeybinds {
    $i = 0
    $json = Get-Content "$PSScriptRoot\..\keybinds.json" | ConvertFrom-Json
    $json.Keybinds | Foreach-Object {
        $skillsGrid.Rows[$i].Cells[2].Value = $_.Keybind
        $skillsGrid.Rows[$i].Cells[9].Value = $_.Scancode
        $i++
    }
    $json = Get-Content "$PSScriptRoot\..\controls.json" | ConvertFrom-Json
    $confirmKeyTxt.Text = $json.ConfirmKey.Keybind -join ''
    $confirmKeyTxt.Tag = $json.ConfirmKey.Scancode -join ''
    $foodBuffTxT.Text = $json.FoodBuffKey.Keybind -join ''
    $foodBuffTxT.Tag = $json.FoodBuffKey.Scancode -join ''
    $medicineTxt.Text = $json.MedicineKey.Keybind -join ''
    $medicineTxt.Tag = $json.MedicineKey.Scancode -join ''
    $craftLogTxT.Text = $json.CraftingLogKey.Keybind -join ''
    $craftLogTxT.Tag = $json.CraftingLogKey.Scancode -join ''
}
function craft {
    $convertDelay = @{
        Step = 2600
        Buff = 1600
    }
    $macros = @()
    $numSteps = $craftingGrid.Rows.Count - 1
    for ($i = 0; $i -lt $numSteps; $i++) {
        $id = $craftingGrid.Rows[$i].Cells[0].Value
        $skill = $skillsGrid.Rows[$id].Cells
        $delay = $convertDelay[$skill[6].Value]
        $macros += @{
            SendKeys = $skill[9].Value
            Delay    = $delay
        }
    }
    $confirmKey = $confirmKeyTxt.Tag -join ''
    $medicineKey = $medicineTxt.Tag -join ''
    $foodbuffKey = $foodBuffTxt.Tag -join ''
    $craftingLog = $craftLogTxT.Tag -join ''
    $ffxiv = Get-Process | Where-Object { $_.MainWindowTitle -like 'Final Fantasy XIV' }
    if ($ffxiv) {
        $syncHash.Crafts = $craftNumeric.Value
        $syncHash.Macros = $macros
        $syncHash.ConfirmKey = $confirmKey
        $syncHash.MedicineKey = $medicineKey
        $syncHash.MedicineCheck = $useMedicine.Checked
        $syncHash.FoodbuffKey = $foodbuffKey
        $syncHash.FoodbuffCheck = $useFoodbuff.Checked
        $syncHash.CraftingLog = $craftingLog
        $syncHash.Window.Text = 'FFXIV Macro Crafter - Running'
        $syncHash.FFXIVHandle = 0
        $syncHash.HDC = 0
        $NativeMethods = Add-Type -name user32 -passThru -MemberDefinition '
        [DllImport("user32.dll")]
        public static extern bool PostMessageA(int hWnd, int hMsg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern bool BlockInput(bool fBlockIt);
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern Int32 ReleaseDC(IntPtr hwnd, IntPtr hdc);
        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetPixel(IntPtr dc, int x, int y);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
        public static extern void SetThreadExecutionState(uint esFlags);
'
        $AssemblyEntry = New-Object System.Management.Automation.Runspaces.SessionStateAssemblyEntry -ArgumentList System.Windows.Forms
    
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        $initialSessionState.Assemblies.Add($AssemblyEntry)
        $initialSessionState.ExecutionPolicy = 4
    
        $Runspace = [runspacefactory]::CreateRunspace()
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("NativeMethods",$NativeMethods)
        $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    
        $Powershell = [powershell]::Create($initialSessionState)
        $Powershell.Runspace = $Runspace
        
        $Powershell.AddScript({        
            $ES_CONTINUOUS = [uint32]"0x80000000"
            $ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
            $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED)
            $ffxivHandle = $NativeMethods::FindWindow(0, 'FINAL FANTASY XIV')
            $hdc = $NativeMethods::GetDC($ffxivHandle)
            $syncHash.FFXIVHandle = $ffxivHandle
            $syncHash.HDC = $hdc
            $confirmDelay = 1500
            $loopDelay = 1800
            $currenTime = Get-Date
            $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 29)
            $medicineTimestamp = $currenTime + (New-Timespan -Minutes 14)
            $craftingbuffs = $syncHash.FoodbuffCheck -or $syncHash.MedicineCheck
            $steps = $($syncHash.Macros).Count
            $i = 0
            while (($i -lt $syncHash.Crafts) -and !($syncHash.Stop -or $syncHash.Abort)) {
                $syncHash.Window.Text = "FFXIV Macro Crafter - Running  Crafted: $($i)  Remaining: $($syncHash.Crafts-$i)"
                do{
                    if($syncHash.Abort){break}
                    Start-Sleep  -m 250
                } while($syncHash.Pause)
                $currenTime = Get-Date
                $buffTimestamps = ($currenTime -gt $foodBuffTimestamp) -or ($currenTime -gt $medicineTimestamp)
                if ($craftingbuffs -and ($i -eq 0 -or $buffTimestamps)) {
                    if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.CraftingLog, 0)}
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.CraftingLog, 0)
                    Start-Sleep 2
                    if ($syncHash.FoodbuffCheck) {
                        if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.FoodbuffKey, 0)} # Use Foodbuff
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.FoodbuffKey, 0)
                        $foodBuffTimestamp = (Get-Date) + (New-Timespan -Minutes 29)
                        Start-Sleep 4
                    }
                    if ($syncHash.MedicineCheck) {
                        if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.MedicineKey, 0)} # Use Medicine
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.MedicineKey, 0)
                        $medicineTimestamp = (Get-Date) + (New-Timespan -Minutes 14)
                        Start-Sleep 4
                    }
                    if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.CraftingLog, 0)} # Open Crafting Log
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.CraftingLog, 0)
                    Start-Sleep 2
                }
                do{
                   $mouseButtons = [System.Windows.Forms.UserControl]::MouseButtons
                   Start-Sleep 1
                } while ($mouseButtons -ne 'None')
                $NativeMethods::BlockInput(1)
                Start-Sleep -m 100
                for($k = 1; $k -le 4; $k++) {
                    if($k -ne 4) {$delay = $confirmDelay - 500} else {$delay = $confirmDelay}
                    if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.ConfirmKey, 0)} # Press Confirm Key
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.ConfirmKey, 0)
                    Start-Sleep -m $delay
                }
                $NativeMethods::BlockInput(0)
                $j = 1
                foreach ($step in $syncHash.Macros) {
                    $scancodes = ($step.Sendkeys).Split(',')
                    $collectable = $NativeMethods::GetPixel($hdc, 283, 281)
                    $normal = $NativeMethods::GetPixel($hdc, 225, 317)
                    #$condition = $NativeMethods::GetPixel($hdc, 46, 326)
                    $quality = [bool](($normal -eq 5232008) -or ($collectable -eq 11927477))
                    if ($quality -and $j -lt $steps) {continue}
                    if ($scancodes.Length -gt 1) {
                        #Double key keybind
                        if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 0)} #Press keys
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[1], 0) 
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[1], 0) #Release keys
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0)
                    }
                    else {
                        #Single key keybind
                        if ($syncHash.Abort) {break} else {[void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 0)} # Key press down
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0) # Key press release
                    }
                    if($syncHash.Pause) {
                        do{
                            if($syncHash.Abort){break}
                            Start-Sleep -m 250
                        } while($syncHash.Pause)
                    } else {
                        Start-Sleep -m $step.Delay
                    }
                    $j += 1
                }
                Start-Sleep -m $loopDelay
                $i++
            }
            $NativeMethods::BlockInput(0)
            $NativeMethods::ReleaseDC($ffxivHandle, $hdc)
            $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS)
        })
        
        $BackgroundJob = $Powershell.BeginInvoke()
        
        #Wait for code to complete and keep UI responsive
         do {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 1
        } while (!$BackgroundJob.IsCompleted)
        
        $Result = $Powershell.EndInvoke($BackgroundJob)
        
        #Clean up
        $syncHash.Stop = $false
        $syncHash.Pause = $false
        $syncHash.Abort = $false
        $syncHash.Window.Text = 'FFXIV Macro Crafter'
        $syncHash.CraftBtn.Text = 'Craft'
        $syncHash.PauseBtn.Text = 'Pause'
        $syncHash.PauseBtn.Enabled = $false
        $Powershell.Dispose()
        $Runspace.Close()
    }
    else {
        [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
    }
}