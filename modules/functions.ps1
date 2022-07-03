
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
        $mods = ([string]$_.modifiers).Replace(', ', ' + ')
        if($converter.Contains($element[2])) {
            $key = $converter[$element[2]]
        }
        else {
            $key = $element[1]
        }
        $scanCodes = @()
        if ($mods -ne 'None') {
            foreach ($mod in $mods.split(' + ')) {
                $scanCodes += $scanCodesModifiers[$mod]
            }
        }
        $scanCodes += '0x' + ('{0:x}' -f $element[2]).ToUpper()
        if($mods -ne 'None') {
            $element[0].Text = "$mods + $key"
        }
        else {
            $element[0].Text = $key
        }
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
    for ($i = 0; $i -lt ($craftingGrid.Rows.Count - 1); $i++) {
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
                [DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
                public static extern void SetThreadExecutionState(uint esFlags);
            }
'@
            $ES_CONTINUOUS = [uint32]"0x80000000"
            $ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
            [NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED)
            [int]$ffxivHandle = [NativeMethods]::FindWindow(0, 'FINAL FANTASY XIV')
            $confirmDelay = 1500
            $loopDelay = 1800
            $currenTime = Get-Date
            $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 29)
            $medicineTimestamp = $currenTime + (New-Timespan -Minutes 14)
            $craftingbuffs = ($args[6] -eq $true) -or ($args[4] -eq $true)
            for ($i = 0; $i -lt $args[0]; $i++) {
                $currenTime = Get-Date
                $buffTimestamps = ($currenTime -gt $foodBuffTimestamp) -or ($currenTime -gt $medicineTimestamp)
                if ($craftingbuffs -eq $true -and ($i -eq 0 -or $buffTimestamps -eq $true)) {
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[7], 0) | Out-Null #Close crafting Log
                    Start-Sleep -m 50
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[7], 0) | Out-Null
                    Start-Sleep 2
                    if ($args[6] -eq $true) {
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[5], 0) | Out-Null #Use Foodbuff
                        Start-Sleep -m 50
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[5], 0) | Out-Null
                        $foodBuffTimestamp = (Get-Date) + (New-Timespan -Minutes 29)
                        Start-Sleep 4
                    }
                    if ($args[4] -eq $true) {
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[3], 0) | Out-Null #Use Medicine
                        Start-Sleep -m 50
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[3], 0) | Out-Null
                        $medicineTimestamp = (Get-Date) + (New-Timespan -Minutes 14)
                        Start-Sleep 4
                    }
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[7], 0) | Out-Null #Open Crafting Log
                    Start-Sleep -m 50
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[7], 0) | Out-Null
                    Start-Sleep 2
                }
                do{
                   $mouseButtons = [System.Windows.Forms.UserControl]::MouseButtons
                   Start-Sleep 1
                } while ($mouseButtons -ne 'None')
                [NativeMethods]::BlockInput(1) | Out-Null
                Start-Sleep -m 100
                if ($craftingbuffs -eq $true) {
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                    Start-Sleep -m 50
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                    Start-Sleep -m ($confirmDelay - 500)
                }
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                Start-Sleep -m 50
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                Start-Sleep -m ($confirmDelay - 500)
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                Start-Sleep -m 50
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                Start-Sleep -m ($confirmDelay - 500)
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                Start-Sleep -m 50
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                Start-Sleep -m $confirmDelay
                [NativeMethods]::BlockInput(0) | Out-Null
                foreach ($step in $args[1]) {
                    $scancodes = ($step.Sendkeys).Split(',')
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
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 0) | Out-Null #Press key
                        Start-Sleep -m 50
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0) | Out-Null #Release key
                    }
                    Start-Sleep -m $step.Delay
                }
                Start-Sleep -m $loopDelay
                "Crafted: $($i+1)  Remaining: $($args[0]-($i+1))"
            }
            [NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS)
        } -ArgumentList $craftNumeric.Value, $macros, $confirmKey, $medicineKey, $useMedicine.Checked, $foodbuffKey, $useFoodbuff.Checked, $craftingLog
    }
    else {
        [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
    }
}