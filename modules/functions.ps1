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
            $mods.split(' + ') | % {
                $scanCodes += $scanCodesModifiers[$_]
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
    $json.Keybinds | % {
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

function loadGearsets{
    $json = Get-Content "$PSScriptRoot\..\gearset.json" | ConvertFrom-Json
    if($json){
        $gearsetGroupBox.Controls.Controls.Controls | Where-Object{$_.Tag -eq 'Numeric'} | % {
            $name = $_.Name
            $value = $json | Where-Object{$_.Name -eq $name} | Select-Object -ExpandProperty Value
            $_.Value = $value
        }
    }
}

function craft {
    $convertDelay = @{
        Step = 2600
        Buff = 1600
    }
    $macros = @()
    $craftingGrid.Rows | % {
        if(!$_.Cells[0].Value) {return}
        $id = $_.Cells[0].Value
        $skill = $skillsGrid.Rows[$id].Cells
        $delay = $convertDelay[$skill[6].Value]
        $macros += @{
            SendKeys = $skill[9].Value
            Delay    = $delay
        }
    }
    $syncHash.Macros = $macros
    $syncHash.Window.Text = 'FFXIV Macro Crafter - Running'
    
    $Powershell.AddScript({        
        $ES_CONTINUOUS = [uint32]"0x80000000"
        $ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
        $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED)
        $ffxivHandle = $syncHash.FFXIVHandle
        $hdc = $syncHash.HDC
        $confirmDelay = 1500
        $loopDelay = 1800
        $currenTime = Get-Date
        $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 29)
        $medicineTimestamp = $currenTime + (New-Timespan -Minutes 14)
        $craftingbuffs = $syncHash.FoodbuffCheck -or $syncHash.MedicineCheck
        $i = 0
        while (($i -lt $syncHash.Crafts) -and !$syncHash.Stop) {
            while($syncHash.Pause) {
                $syncHash.Window.Text = "FFXIV Macro Crafter - Paused"
                Start-Sleep -m 250
            }
            $syncHash.Window.Text = "FFXIV Macro Crafter - Running  Crafted: $($i)  Remaining: $($syncHash.Crafts-$i)"
            $currenTime = Get-Date
            $buffTimestamps = ($currenTime -gt $foodBuffTimestamp) -or ($currenTime -gt $medicineTimestamp)
            if ($craftingbuffs -and ($i -eq 0 -or $buffTimestamps)) {
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.CraftingLog, 1)
                Start-Sleep -m 50
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.CraftingLog, 0xC0000001)
                Start-Sleep 2
                if ($syncHash.FoodbuffCheck) {
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.FoodbuffKey, 1) # Use Foodbuff
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.FoodbuffKey, 0xC0000001)
                    $foodBuffTimestamp = (Get-Date) + (New-Timespan -Minutes 29)
                    Start-Sleep 4
                }
                if ($syncHash.MedicineCheck) {
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.MedicineKey, 1) # Use Medicine
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.MedicineKey, 0xC0000001)
                    $medicineTimestamp = (Get-Date) + (New-Timespan -Minutes 14)
                    Start-Sleep 4
                }
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.CraftingLog, 1) # Open Crafting Log
                Start-Sleep -m 50
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.CraftingLog, 0xC0000001)
                Start-Sleep 2
            }
            do{
                $mouseButtons = [System.Windows.Forms.UserControl]::MouseButtons
                Start-Sleep 1
            } while ($mouseButtons -ne 'None')
            $NativeMethods::BlockInput(1)
            Start-Sleep -m 100
            1..4 | % {
                if($_ -ne 4) {$delay = $confirmDelay - 500} else {$delay = $confirmDelay}
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $syncHash.ConfirmKey, 1) # Press Confirm Key
                Start-Sleep -m 50
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $syncHash.ConfirmKey, 0xC0000001)
                Start-Sleep -m $delay
            }
            $NativeMethods::BlockInput(0)
            0..$syncHash.Macros.Count | % {
                $scancodes = ($syncHash.Macros[$_].Sendkeys).Split(',')
                $collectable = $NativeMethods::GetPixel($hdc, 283, 281)
                $normal = $NativeMethods::GetPixel($hdc, 225, 317)
                #$condition = $NativeMethods::GetPixel($hdc, 46, 326)
                $quality = [bool](($normal -eq 5232008) -or ($collectable -eq 11927477))
                if ($quality -and $_ -lt $($syncHash.Macros).Count-1) {return}
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 1)
                if ($scancodes.Length -gt 1) {
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[1], 0x80000001) 
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[1], 0xC0000001)
                }
                Start-Sleep -m 50
                [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0xC0000001)
                if($syncHash.Pause) {
                    while($syncHash.Pause) {
                        $syncHash.Window.Text = "FFXIV Macro Crafter - Paused"
                        Start-Sleep -m 250
                    }
                    $syncHash.Window.Text = "FFXIV Macro Crafter - Running  Crafted: $($i)  Remaining: $($syncHash.Queue[$q].Crafts-$i)"
                } else {
                    Start-Sleep -m $syncHash.Macros[$_].Delay
                }
            }
            Start-Sleep -m $loopDelay
            $i++
        }
    })
    
    $BackgroundJob = $Powershell.BeginInvoke()
    
    #Wait for code to complete and keep UI responsive
    do {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 1
    } while (!$BackgroundJob.IsCompleted)
    
    if($Powershell.InvocationStateInfo.State -eq 'Completed') {$Result = $Powershell.EndInvoke($BackgroundJob)}
    
    #Clean up
    $NativeMethods::BlockInput(0)
    $NativeMethods::ReleaseDC($syncHash.FFXIVHandle, $syncHash.HDC)
    $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS)
    $syncHash.Stop = $false
    $syncHash.Pause = $false
    $syncHash.Window.Text = 'FFXIV Macro Crafter'
    $syncHash.CraftBtn.Text = 'Craft'
    $syncHash.PauseBtn.Text = 'Pause'
    $syncHash.PauseBtn.Enabled = $false
}
function craftQueue {
    $convertDelay = @{
        Step = 2600
        Buff = 1600
    }
    [System.Collections.ArrayList]$syncHash.Queue = @()
    [System.Collections.ArrayList]$syncHash.Start = @()
    [System.Collections.ArrayList]$syncHash.Begin = @()

    :queue foreach($row in $queueGrid.Rows) {
        [System.Collections.ArrayList]$columns = @()
        $macro = @()
        foreach($cell in $row.Cells) {
            if(!$cell.Value) {break queue}
            if($cell.OwningColumn.Name -eq 'Rotation') {
                foreach($id in $(Get-Content "$PSScriptRoot\..\Rotations\$($cell.Value).json" | ConvertFrom-Json)) {
                    $skill = $skillsGrid.Rows[$id].Cells
                    $delay = $convertDelay[$skill[6].Value]
                    $macro += @{
                        SendKeys = $skill[9].Value
                        Delay    = $delay
                    }
                }
                $columns.Add(@{$cell.OwningColumn.Name = $macro})
                continue
            }
            $columns.Add(@{$cell.OwningColumn.Name = $cell.Value})
        }
        $syncHash.Queue.Add($columns)
    }
    if ($syncHash.MedicineCheck) {
        $syncHash.Begin.Add(@{Keybind = $syncHash.MedicineKey; Delay = 4000; Repeat = 1}) # Use Medicine
    }
    if ($syncHash.FoodbuffCheck) {
        $syncHash.Begin.Add(@{Keybind = $syncHash.FoodBuffKey; Delay = 4000; Repeat = 1}) # Use Food
    }
    $syncHash.Begin.Add(@{Keybind =     $syncHash.CraftingLog; Delay = 750;  Repeat = 1}) # Crafting Log
    $syncHash.Begin.Add(@{Keybind =     '0x22';                Delay = 100;  Repeat = 4}) # PgDn
    $syncHash.Begin.Add(@{Keybind =     '0x62';                Delay = 100;  Repeat = 4}) # NumPad 2
    $syncHash.Begin.Add(@{Keybind =     '0x22';                Delay = 100;  Repeat = 4}) # PgDn
    $syncHash.Begin.Add(@{Keybind =     '0x62';                Delay = 100;  Repeat = 1}) # NumPad 2
    $syncHash.Begin.Add(@{Keybind =     $syncHash.ConfirmKey;  Delay = 100;  Repeat = 1}) # Confirm
    $syncHash.Begin.Add(@{Keybind =     '0x0D';                Delay = 5000; Repeat = 1}) # Enter
    $syncHash.Begin.Add(@{Keybind =     $syncHash.ConfirmKey;  Delay = 100;  Repeat = 1}) # Confirm
    $syncHash.Begin.Add(@{Keybind =     '0x68';                Delay = 100;  Repeat = 1}) # NumPad 8
    $syncHash.Begin.Add(@{Keybind =     '0x66';                Delay = 100;  Repeat = 2}) # NumPad 6

    $syncHash.Window.Text = 'FFXIV Macro Crafter - Running'
    
    $Powershell.AddScript({
        $WindowsForms = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");     
        $ES_CONTINUOUS = [uint32]"0x80000000"
        $ES_DISPLAY_REQUIRED = [uint32]"0x00000002"
        $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_DISPLAY_REQUIRED)
        $ffxivHandle = $syncHash.FFXIVHandle
        $hdc = $syncHash.HDC
        $confirmDelay = 1500
        $loopDelay = 1800
        $pasteText = {
            param(
                [Parameter(Mandatory)]
                $text
            )
            Set-Clipboard -Value $text
            [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, 0x11, 1)
            Start-Sleep -m 50
            & $singleKeyPress -keybind 0x56 -lParam1 0x80000001
            Start-Sleep -m 50
            [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, 0x11, 0xC0000001)
        }
        $selectRecipe = {
            $syncHash.Start | % {
                & $pause
                $syncHash.Window.Text = "FFXIV Macro Crafter - Running"
                $keybind = $_.Keybind
                $delay = $_.Delay
                switch ($keybind) {
                    $syncHash.FoodBuffKey {
                        if ($syncHash.FoodbuffCheck) {
                            $script:foodBuffTimestamp = (Get-Date + (New-Timespan -Minutes 29))
                        }
                    }
                    $syncHash.MedicineKey {
                        if ($syncHash.MedicineCheck) {
                            $script:medicineTimestamp = (Get-Date + (New-Timespan -Minutes 14))
                        }
                    }
                    '0x0D' {
                        Set-Clipboard -Value $syncHash.Queue[$q].Recipe
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, 0x11, 1)
                        Start-Sleep -m 50
                        & $singleKeyPress -keybind 0x56 -lParam1 0x80000001
                        Start-Sleep -m 50
                        [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, 0x11, 0xC0000001)
                        Start-Sleep -m 250
                    }
                }
                1..$_.Repeat | % {
                    & $singleKeyPress -keybind $keybind
                    Start-Sleep -m $delay
                }
            }
        }
        for($q = 0; $q -lt $syncHash.Queue.Count; $q++) {
            if($syncHash.Stop) {break}
            $syncHash.Start = $syncHash.Begin.Clone()
            1..$syncHash.Queue[$q].Materials | % {
                $syncHash.Start.Add(@{Keybind = $syncHash.ConfirmKey; Delay = 100; Repeat = 6}) # Confirm
                $syncHash.Start.Add(@{Keybind = '0x68'; Delay = 100; Repeat = 1})               # NumPad 8
            }
            $syncHash.Start.Add(@{Keybind = '0x68'; Delay = 100; Repeat = 2})                   # NumPad 8
            $syncHash.Start.Add(@{Keybind = $syncHash.ConfirmKey; Delay = 1500; Repeat = 1})    # Confirm
            if($q -eq 0) {[System.Windows.MessageBox]::Show('Close all in-game Windows before pressing OK','Information','OK','Information')}
            $NativeMethods::BlockInput(1)
            Start-Sleep -m 100
            #Change Crafter Job
            $gearsetNumbers = $gearsetGroupBox.Controls.Controls.Controls | Where-Object{$_.Tag -eq 'Numeric'} | Select-Object Value
            $duplicateGearsets = ($gearsetNumbers | Select-Object -ExpandProperty Value | Group-Object | ?{$_.Count -gt 1}).Values
            if($duplicateGearsets){
                [System.Windows.MessageBox]::Show('Gearset Numbers need to be unique', 'Error', 'OK', 'Error')
                break
            }
            $gearsetNumber = $syncHash.Queue[$q].Crafter
            & $singleKeyPress -Keybind 0xBF
            & $pasteText -Text "gs change $gearsetNumber"
            & $singleKeyPress -Keybind 0x0D

            Start-Sleep 1
            $currenTime = Get-Date
            $foodBuffTimestamp = $currenTime + (New-Timespan -Minutes 29)
            $medicineTimestamp = $currenTime + (New-Timespan -Minutes 14)
            $craftingbuffs = $syncHash.Begin.Count -gt 10
            & $selectRecipe
            $NativeMethods::BlockInput(0)
            for($i = 0; $i -lt $syncHash.Queue[$q].Crafts; $i++) {
                if($syncHash.Stop){break}
                & $pause
                $syncHash.Window.Text = "FFXIV Macro Crafter - Running  Crafted: $($i)  Remaining: $($syncHash.Queue[$q].Crafts-$i)"
                $currenTime = Get-Date
                $buffTimestamps = [bool]($currenTime -gt $foodBuffTimestamp -or $currenTime -gt $medicineTimestamp)
                if (($i -gt 0 -or $q -gt 0) -and ($craftingbuffs -and $buffTimestamps)) {
                    & $singleKeyPress -keybind $syncHash.CraftingLog
                    Start-Sleep 2
                    & $selectRecipe
                }
                while ([System.Windows.Forms.UserControl]::MouseButtons -ne 'None') {
                    Start-Sleep 1
                }
                $NativeMethods::BlockInput(1)
                Start-Sleep -m 100
                if($i -gt 0) {
                    1..4 | % {
                        if($i -eq 0) {return}
                        if($_ -lt 3) {$delay = $confirmDelay - 500} else {$delay = $confirmDelay}
                        & $singleKeyPress -keybind $syncHash.ConfirmKey
                        Start-Sleep -m $delay
                    }
                }
                $NativeMethods::BlockInput(0)
                for($r = 0;$r -lt $syncHash.Queue[$q].Rotation.Count;$r++){
                    $collectable = $NativeMethods::GetPixel($hdc, 283, 281)
                    $normal      = $NativeMethods::GetPixel($hdc, 225, 317)
                  # $condition   = $NativeMethods::GetPixel($hdc, 46, 326)
                    $quality     = [bool]($normal -eq 5232008 -or $collectable -eq 11927477)
                    if ($quality -and ($r -lt ($syncHash.Queue[$q].Rotation.Count - 1))) {continue}
                    $scancodes = ($syncHash.Queue[$q].Rotation[$r].SendKeys).Split(',')
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0100, $scancodes[0], 1)
                    if ($scancodes.Length -gt 1) {
                        Start-Sleep -m 50
                        & $singleKeyPress -keybind $scancodes[1]
                    }
                    Start-Sleep -m 50
                    [void]$NativeMethods::PostMessageA($ffxivHandle, 0x0101, $scancodes[0], 0xC0000001)
                    if($syncHash.Pause) {
                        & $pause
                        $syncHash.Window.Text = "FFXIV Macro Crafter - Running  Crafted: $($i)  Remaining: $($syncHash.Queue[$q].Crafts-$i)"
                    } else {
                        Start-Sleep -m $syncHash.Queue[$q].Rotation[$r].Delay
                    }
                }
                Start-Sleep -m $loopDelay
            }
            Start-Sleep -m 1700
            & $singleKeyPress -keybind $syncHash.CraftingLog
            Start-Sleep 4
        }
    })
    
    $BackgroundJob = $Powershell.BeginInvoke()
    
    #Wait for code to complete and keep UI responsive
    do {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 1
    } while (!$BackgroundJob.IsCompleted)
    
    if($Powershell.InvocationStateInfo.State -eq 'Completed') {$Result = $Powershell.EndInvoke($BackgroundJob)}
    
    #Clean up
    $NativeMethods::BlockInput(0)
    $NativeMethods::ReleaseDC($syncHash.FFXIVHandle, $syncHash.HDC)
    $NativeMethods::SetThreadExecutionState($ES_CONTINUOUS)
    $syncHash.Stop = $false
    $syncHash.Pause = $false
    $syncHash.Window.Text = 'FFXIV Macro Crafter'
    $syncHash.CraftQueueBtn.Text = 'Craft'
    $syncHash.PauseQueueBtn.Text = 'Pause'
    $syncHash.PauseQueueBtn.Enabled = $false
}