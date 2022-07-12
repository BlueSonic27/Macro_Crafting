
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
        $main.Height = 66
        $main.Width = 140
        $main.FormBorderStyle = 'None'
        $expandCollapse.Text = 'Expand'
        $expandCollapsePanel.Left = -80
        $FormTabControl.Visible = $false
    }
    else {
        $main.Height = 470
        $main.Width = 570
        $main.FormBorderStyle = 'Fixed3D'
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
        expand
        $timer.Enabled = $True
        $main.Text = 'FFXIV Macro Crafter - Running'
        Start-Job -Name Craft -ScriptBlock {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type @'
            using System;
            using System.Runtime.InteropServices;
            public static class NativeMethods {
                [DllImport("user32.dll")] public static extern bool PostMessageA(int hWnd, int hMsg, int wParam, int lParam);
                [DllImport("user32.dll")] public static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);
                [DllImport("user32.dll")] public static extern bool BlockInput(bool fBlockIt);
                [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
                [DllImport("user32.dll", SetLastError = true)]
                public static extern Int32 ReleaseDC(IntPtr hwnd, IntPtr hdc);
                [DllImport("gdi32.dll", SetLastError = true)]
                public static extern uint GetPixel(IntPtr dc, int x, int y);
                [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
                public static extern void SetThreadExecutionState(uint esFlags);
            }
'@
        function keyPress() {
            param(
                [Parameter(Mandatory=$true)]
                $key
            )
            [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $key, 0) | Out-Null # Key press down
            Start-Sleep -m 50
            [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $key, 0) | Out-Null # Key press release
        }
            $ES_CONTINUOUS = [uint32]"0x80000000"
            $ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
            [NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED)
            $ffxivHandle = [NativeMethods]::FindWindow(0, 'FINAL FANTASY XIV')
            $hdc = [NativeMethods]::GetDC($ffxivHandle)
            $confirmDelay = 1500
            $loopDelay = 1800
            $currenTime = Get-Date
            $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 29)
            $medicineTimestamp = $currenTime + (New-Timespan -Minutes 14)
            $craftingbuffs = ($args[6] -eq $true) -or ($args[4] -eq $true)
            $steps = $args[1].Count
            for ($i = 0; $i -lt $args[0]; $i++) {
                $currenTime = Get-Date
                $buffTimestamps = ($currenTime -gt $foodBuffTimestamp) -or ($currenTime -gt $medicineTimestamp)
                if ($craftingbuffs -eq $true -and ($i -eq 0 -or $buffTimestamps -eq $true)) {
                    keyPress($args[7]) # Close crafting Log
                    Start-Sleep 2
                    if ($args[6] -eq $true) {
                        keyPress($args[5]) # Use Foodbuff
                        $foodBuffTimestamp = (Get-Date) + (New-Timespan -Minutes 29)
                        Start-Sleep 4
                    }
                    if ($args[4] -eq $true) {
                        keyPress($args[3]) # Use Medicine
                        $medicineTimestamp = (Get-Date) + (New-Timespan -Minutes 14)
                        Start-Sleep 4
                    }
                    keyPress($args[7]) # Open Crafting Log
                    Start-Sleep 2
                }
                do{
                   $mouseButtons = [System.Windows.Forms.UserControl]::MouseButtons
                   Start-Sleep 1
                } while ($mouseButtons -ne 'None')
                [NativeMethods]::BlockInput(1) | Out-Null
                Start-Sleep -m 100
                for($k = 1; $k -le 4; $k++) {
                    $delay = (($k -ne 4) ? ($confirmDelay - 500) : $confirmDelay)
                    keyPress($args[2]) # Press Confirm Key
                    Start-Sleep -m $delay
                }
                [NativeMethods]::BlockInput(0) | Out-Null
                $j = 1
                foreach ($step in $args[1]) {
                    $scancodes = ($step.Sendkeys).Split(',')
                    $collectable = [NativeMethods]::GetPixel($hdc, 283, 281)
                    $normal = [NativeMethods]::GetPixel($hdc, 225, 317)
                    #$condition = [NativeMethods]::GetPixel($hdc, 46, 326)
                    $quality = [bool](($normal -eq 5232008) -or ($collectable -eq 11927477))
                    if (($quality -eq $false) -or ($quality -eq $true -and $j -eq $steps)) {
                        if ($scancodes.Length -gt 1) {
                            #Double key keybind
                            [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 0) | Out-Null #Press keys
                            Start-Sleep -m 50
                            [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $scancodes[1], 0) | Out-Null 
                            Start-Sleep -m 50
                            [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $scancodes[1], 0) | Out-Null #Release keys
                            Start-Sleep -m 50
                            [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0) | Out-Null
                        }
                        else {
                            #Single key keybind
                            keyPress($scancodes[0])
                        }
                        Start-Sleep -m $step.Delay
                    }
                    $j += 1
                }
                Start-Sleep -m $loopDelay
                "Crafted: $($i+1)  Remaining: $($args[0]-($i+1))"
            }
            [NativeMethods]::ReleaseDC($ffxivHandle, $hdc) | Out-Null
            [NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS)
        } -ArgumentList $craftNumeric.Value, $macros, $confirmKey, $medicineKey, $useMedicine.Checked, $foodbuffKey, $useFoodbuff.Checked, $craftingLog
    }
    else {
        [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
    }
}