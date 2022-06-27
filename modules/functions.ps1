
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
        Step = 2500
        Buff = 1300
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
            Add-Type @'
            using System;
            using System.Runtime.InteropServices;
            public static class NativeMethods {
                [DllImport("user32.dll")] public static extern bool PostMessageA(int hWnd, int hMsg, int wParam, int lParam);
                [DllImport("user32.dll")] public static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);
                [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
                [DllImport("user32.dll")] public static extern bool BlockInput(bool fBlockIt);
                [DllImport("user32.dll")] public static extern void mouse_event(int flags, int dx, int dy, int cButtons, int info);
            }
'@
            [int]$ffxivHandle = [NativeMethods]::FindWindow(0, 'FINAL FANTASY XIV')
            $confirmDelay = 1300
            $loopDelay = 1700
            $currenTime = Get-Date
            $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 30 -Seconds 10)
            $medicineTimestamp = $currenTime + (New-Timespan -Minutes 15 -Seconds 10)
            for ($i = 0; $i -lt $args[0]; $i++) {
                $currenTime = Get-Date
                if ($args[4] -eq $true -or $args[6] -eq $true) {
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[7], 0) | Out-Null #Open crafting Log
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[7], 0) | Out-Null
                    Start-Sleep 3
                    if ($currenTime -gt $foodBuffTimestamp -or $i -eq 0) {
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[5], 0) | Out-Null #Use Foodbuff
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[5], 0) | Out-Null
                        $foodBuffTimestamp = (Get-Date) + (New-Timespan -Minutes 30)
                        Start-Sleep 3
                    }
                    if ($currenTime -gt $medicineTimestamp -or $i -eq 0) {
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[3], 0) | Out-Null #Use Medicine
                        [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[3], 0) | Out-Null
                        $medicineTimestamp = (Get-Date) + (New-Timespan -Minutes 15)
                        Start-Sleep 3
                    }
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[7], 0) | Out-Null #Close Crafting Log
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[7], 0) | Out-Null
                    Start-Sleep 2
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                    [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                    Start-Sleep -m $confirmDelay
                }
                [NativeMethods]::BlockInput(1) | Out-Null
                Start-Sleep -m 100
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                Start-Sleep -m ($confirmDelay - 500)
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0101, $args[2], 0) | Out-Null #Release Confirm Key
                Start-Sleep -m ($confirmDelay - 500)
                [NativeMethods]::PostMessageA($ffxivHandle, 0x0100, $args[2], 0) | Out-Null #Press Confirm Key
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
        } -ArgumentList $craftNumeric.Value, $macros, $confirmKey, $medicineKey, $useMedicine.Checked, $foodbuffKey, $useFoodbuff.Checked, $craftingLog
    }
    else {
        [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
    }
}