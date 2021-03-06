$importTextCheck = $false
$numericKeys = @(8,37,39,46,48,49,50,51,52,53,54,55,56,57,96,97,98,99,100,101,102,103,104,105)
$keybindsTab.Add_Enter({
    $skillsGrid.Rows[0].Cells[2].Selected = $true
})
$expandCollapse.Add_Click({
    expand
})
$importMacros.Add_Click({
    $importForm.ShowDialog()
})
$importText.Add_TextChanged({
    if ($importTextCheck -eq $false) {
        $importTextCheck = $true
        $replace = @()
        $lines = $this.Text -split "\r?\n|\r"
        foreach ($line in $lines) {
            $replace += $line
        }
        $this.Lines = $replace
    }
})
$importBtn.Add_Click({
    $text = $importText.Text
    $craftingGrid.Rows.Clear()
    foreach ($line in ($text -split "\r?\n|\r")) {
        $line = $line -replace '(^/ac )|( ?<[\S]+>$)|"', ''
        $lastRow = $craftingGrid.NewRowIndex
        $craftingGrid.Rows.Add();
        $craftingGrid.Rows[$lastRow].Cells[0].Value = $ds.Tables[0] | Where-Object { $_.Name -eq $line } | Select-Object -ExpandProperty ID
    }
    $clearRotation.Visible = $true
    $importForm.Close()
    $importText.Clear()
})
$classDropdown.Add_SelectionChangeCommitted({
    $class = $this.SelectedItem
    $classButtons = @('advancedTouch', 'basicSynth', 'basicTouch', 'delicateSynth', 'focusedSynth', 'focusedTouch', 'groundwork', 'intensiveSynth', 'preciseTouch', 'preparatoryTouch', 'prudentSynth', 'prudentTouch', 'standardTouch')
    if ($this.SelectedItem -ne $null) {
        foreach ($button in $classButtons) {
            $(Get-Variable -Name $button -ValueOnly).Image = [System.Drawing.Image]::FromFile("$scriptDir\icons\$class\$button.png")
        }
    }
})
$saveQueueBtn.Add_Click({
    $rows = $queueGrid.Rows | Select-Object -SkipLast 1
    $queuelist = $rows | Select-Object @{
        Label='Recipe'   ;Expression={$_.Cells[0].Value}
        },@{
        Label='Materials';Expression={$_.Cells[1].Value}
        },@{
        Label='Rotation' ;Expression={$_.Cells[2].Value}
        },@{
        Label='Crafter'  ;Expression={$_.Cells[3].Value}
        },@{
        Label='Crafts'   ;Expression={$_.Cells[4].Value}
        }
    $mod5 = ((($queueGrid.Rows | Select-Object -SkipLast 1).Cells | Select-Object -ExpandProperty Value).Count % 5) -ne 0
    if($queueGrid.RowCount -eq 1 -or $mod5){return}
    $saveFileBrowser.InitialDirectory = "$scriptDir\Queues"
    [void]$saveFileBrowser.ShowDialog()
    if(!$saveFileBrowser.Filename){return}
    $queueList | ConvertTo-Json | Out-File -FilePath $saveFileBrowser.Filename -NoNewline
})
$loadQueueBtn.Add_Click({
    $openFileBrowser.InitialDirectory = "$scriptDir\Queues"
    [void]$openFileBrowser.ShowDialog()
    if(!$openFileBrowser.Filename){return}
    $queue = Get-Content $openFileBrowser.Filename | ConvertFrom-Json
    $queueGrid.Rows.Clear()
    $clearRotation.Visible = $true
    for ($i = 0; $i -lt $queue.Length; $i++) {
        $lastRow = $queueGrid.NewRowIndex
        $column = $queue[$i]
        $queueGrid.Rows.Add()
        $queueGrid.Rows[$lastRow].Cells[0].Value = $column.Recipe
        $queueGrid.Rows[$lastRow].Cells[1].Value = $column.Materials
        $queueGrid.Rows[$lastRow].Cells[2].Value = $column.Rotation
        $queueGrid.Rows[$lastRow].Cells[3].Value = $column.Crafter
        $queueGrid.Rows[$lastRow].Cells[4].Value = $column.Crafts
    }
})
$syncHash.CraftQueueBtn.Add_Click({
    if($queueGrid.RowCount -eq 1){return}
    switch($syncHash.CraftQueueBtn.Text){
        'Stop' {
            $syncHash.Stop = $true
            $syncHash.PauseQueueBtn.Enabled = $false
            $syncHash.Window.Text = 'FFXIV Macro Crafter - Stopping'
            $syncHash.CraftQueueBtn.Text = 'Abort'
        }
        'Abort' {
            $Powershell.Stop()
            $syncHash.PauseQueueBtn.Enabled = $false
            $syncHash.Window.Text = 'FFXIV Macro Crafter'
            $syncHash.CraftQueueBtn.Text = 'Craft'
        }
        default {
            $syncHash.FFXIVHandle = $NativeMethods::FindWindow(0, 'FINAL FANTASY XIV')
            if ($syncHash.FFXIVHandle -ne 0) {
                $syncHash.HDC = $NativeMethods::GetDC($syncHash.FFXIVHandle)
                $syncHash.PauseQueueBtn.Enabled = $true
                $syncHash.Window.Text = 'FFXIV Macro Crafter'
                $syncHash.CraftQueueBtn.Text = 'Stop'
                $syncHash.ConfirmKey = $confirmKeyTxt.Tag -join ''
                $syncHash.MedicineKey = $medicineTxt.Tag -join ''
                $syncHash.MedicineCheck = $useMedicine.Checked
                $syncHash.FoodbuffKey = $foodBuffTxt.Tag -join ''
                $syncHash.FoodbuffCheck = $useFoodbuff.Checked
                $syncHash.CraftingLog = $craftLogTxT.Tag -join ''
                craftQueue
            }
            else {
                [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
            }
        }
    }
})
$syncHash.PauseQueueBtn.Add_click({
    if($syncHash.Pause) {
        $syncHash.Pause = $false
        $syncHash.PauseQueueBtn.Text = 'Pause'
        $syncHash.CraftQueueBtn.Text = 'Stop'
    } else {
        $syncHash.Pause = $true
        $syncHash.PauseQueueBtn.Text = 'Resume'
        $syncHash.CraftQueueBtn.Text = 'Abort'
    }
})
$syncHash.CraftBtn.Add_click({
    if($craftingGrid.RowCount -eq 1){return}
    switch($syncHash.CraftBtn.Text) {
        'Stop' {
            $syncHash.Stop = $true
            $syncHash.PauseBtn.Enabled = $false
            $syncHash.Window.Text = 'FFXIV Macro Crafter - Stopping'
            $syncHash.CraftBtn.Text = 'Abort'
        }
        'Abort' {
            $Powershell.Stop()
            $syncHash.PauseBtn.Enabled = $false
            $syncHash.Window.Text = 'FFXIV Macro Crafter'
            $syncHash.CraftBtn.Text = 'Craft'
        }
        default {
            $syncHash.FFXIVHandle = $NativeMethods::FindWindow(0, 'FINAL FANTASY XIV')
            if ($syncHash.FFXIVHandle -ne 0) {
                $syncHash.HDC = $NativeMethods::GetDC($syncHash.FFXIVHandle)
                $syncHash.PauseBtn.Enabled = $true
                $syncHash.Window.Text = 'FFXIV Macro Crafter'
                $syncHash.CraftBtn.Text = 'Stop'
                $syncHash.ConfirmKey = $confirmKeyTxt.Tag -join ''
                $syncHash.MedicineKey = $medicineTxt.Tag -join ''
                $syncHash.MedicineCheck = $useMedicine.Checked
                $syncHash.FoodbuffKey = $foodBuffTxt.Tag -join ''
                $syncHash.FoodbuffCheck = $useFoodbuff.Checked
                $syncHash.CraftingLog = $craftLogTxT.Tag -join ''
                $syncHash.Crafts = $craftNumeric.Value
                craft
            }
            else {
                [System.Windows.MessageBox]::Show('Final Fantasy XIV Client cannot be found', 'Error', 'OK', 'Error')
            }
        }
    }
})
$saveGearset.Add_Click({
    $gearsetNumbers = $gearsetGroupBox.Controls.Controls.Controls | Where-Object{$_.Tag -eq 'Numeric'} | Select-Object Value
    $duplicates = ($gearsetNumbers | Select-Object -ExpandProperty Value | Group-Object | ?{$_.Count -gt 1}).Values
    if($duplicates){
        [System.Windows.MessageBox]::Show('Gearset Numbers need to be unique', 'Error', 'OK', 'Error')
    } else {
        ConvertTo-Json -InputObject $gearsetNumbers -Compress | Out-File -FilePath "$scriptDir\gearset.json" -NoNewline
    }
})
$loadGearset.Add_Click({
    loadGearsets
})
$syncHash.PauseBtn.Add_click({
    if($syncHash.Pause){
        $syncHash.Pause = $false
        $syncHash.PauseBtn.Text = 'Pause'
        $syncHash.CraftBtn.Text = 'Stop'
    } else {
        $syncHash.Pause = $true
        $syncHash.PauseBtn.Text = 'Resume'
        $syncHash.CraftBtn.Text = 'Abort'
    }
})
$clearRotation.Add_Click({
    switch($FormTabControl.SelectedTab.Text){
        'Crafting' {
            $craftingGrid.Rows.Clear()
            $reflect.Enabled = $true
            $trainedEye.Enabled = $true
            $muscleMemory.Enabled = $true
        }
        'Queue' {
            $queueGrid.Rows.Clear()
        }
    }
    $clearRotation.Visible = $false
})
$loadRecipeBtn.Add_Click({
    $openFileBrowser.InitialDirectory = "$scriptDir\Rotations"
    [void]$openFileBrowser.ShowDialog()
    if(!$openFileBrowser.Filename){return}
    $rotation = Get-Content $openFileBrowser.Filename | ConvertFrom-Json
    $craftingGrid.Rows.Clear()
    $clearRotation.Visible = $true
    for ($i = 0; $i -lt $rotation.Length; $i++) {
        $lastRow = $craftingGrid.NewRowIndex
        $step = [string]$rotation[$i]
        $craftingGrid.Rows.Add();
        $craftingGrid.Rows[$lastRow].Cells[0].Value = $step
    }
    if ($craftingGrid.NewRowIndex -gt 0) {
        $reflect.Enabled = $false
        $trainedEye.Enabled = $false
        $muscleMemory.Enabled = $false
    }
})
$saveRecipeBtn.Add_Click({
    if($craftingGrid.RowCount -eq 1){return}
    $saveFileBrowser.InitialDirectory = "$scriptDir\Rotations"
    [void]$saveFileBrowser.ShowDialog()
    if(!$saveFileBrowser.Filename){return}
    $rotation = $craftingGrid.Rows.Cells | Select-Object -SkipLast 1 -ExpandProperty Value | ConvertTo-Json -Compress
    $rotation | Out-File -FilePath $saveFileBrowser.Filename -NoNewline
})
$confirmKeyTxt.Add_KeyUp({
    $this.Tag = scancodes($this, $_.keyCode, $_.KeyValue)
})
$foodBuffTxt.Add_KeyUp({
    $this.Tag = scancodes($this, $_.keyCode, $_.KeyValue)
})
$medicineTxt.Add_KeyUp({
    $this.Tag = scancodes($this, $_.keyCode, $_.KeyValue)
})
$craftLogTxt.Add_KeyUp({
    $this.Tag = scancodes($this, $_.keyCode, $_.KeyValue)
})
$saveKeybindBtn.Add_Click({
    $keybinds = $skillsGrid.DataSource | Select-Object Keybind, Scancode
    $keybindsJson = @{'Keybinds' = $keybinds } | ConvertTo-Json -Compress
    $keybindsJson | Out-File -FilePath "$scriptDir\keybinds.json" -NoNewline
    $controlJson = @"
{
    `"ConfirmKey`" : [
        {`"Keybind`" : `"$($confirmKeyTxt.Text)`"},
        {`"Scancode`" : `"$($confirmKeyTxt.Tag)`"}
    ],
    `"FoodBuffKey`" : [
        {`"Keybind`" : `"$($foodBuffTxt.Text)`"},
        {`"Scancode`" : `"$($foodbuffTxt.Tag)`"}
    ],
    `"MedicineKey`" : [
        {`"Keybind`" : `"$($medicineTxt.Text)`"},
        {`"Scancode`" : `"$($medicineTxt.Tag)`"}
    ],
    `"CraftingLogKey`" : [
        {`"Keybind`" : `"$($craftLogTxt.Text)`"},
        {`"Scancode`" : `"$($craftLogTxt.Tag)`"}
    ]
}
"@
    $controlJson | Out-File -FilePath "$scriptDir\controls.json" -NoNewline
})
$loadKeybindBtn.Add_Click({
    loadKeybinds
})
$levelNumeric.Add_TextChanged({
    foreach ($skill in ($ds.Tables[0].Rows)) {
        $id = $skill.Key
        $level = [int]$skill.Level
        if ($level -gt $this.Value) {
            $(Get-Variable -Name $id -ValueOnly).Enabled = $false
        }
        else {
            $(Get-Variable -Name $id -ValueOnly).Enabled = $true
        }
    }
})
$checkNumeric = {
    $col = $this.EditingControlDataGridView.CurrentCell.OwningColumn.Name
    $materialsAllowedKeys = @(8,46,49,50,51,52,53,54,97,98,99,100,101,102)
    if(
        !($numericKeys).Contains($_.keyValue) -or 
        ($col -eq 'Materials' -and !$materialsAllowedKeys.Contains($_.keyValue))
    ) {
        $_.SuppressKeyPress = $true
    }
}
$queueGrid.Add_EditingControlShowing({
    $clearRotation.Visible = $true
    $col = $queueGrid.CurrentCell.OwningColumn.Name
    $_.Control.Remove_KeyDown($checkNumeric)
    if(($col -eq 'Crafts') -or ($col -eq 'Materials')) {
        $_.Control.Add_KeyDown($checkNumeric)
    }
})
$queueGrid.Add_CellMouseUp({
    if ($_.Button -eq 'Right') {
        $rowIndex = $_.RowIndex
        if ($rowIndex -ne $this.NewRowIndex) {
            $this.Rows.RemoveAt($rowIndex)
            if ($this.NewRowIndex -gt 0) {
                $clearRotation.Visible = $true
            }
            else {
                $clearRotation.Visible = $false
            }
        }
    }
})
$skillsGrid.Add_KeyDown({
    $_.SuppressKeyPress = $true
})
$skillsGrid.Add_KeyUp({
    $this.SelectedCells[9].Value = scancodes($this.SelectedCells[2], $_.keyCode, $_.KeyValue)
})
$skillsGrid.Add_CellValidating({
    if ($_.ColumnIndex -eq 1) { return }
    $uniqueCount = ($skillsGrid.Rows.Cells | Where-Object { $_.ColumnIndex -eq 2 -and $_.Value -ne '' } | Select-Object Value -unique).Count
    $count = ($skillsGrid.Rows.Cells | Where-Object { $_.ColumnIndex -eq 2 -and $_.Value -ne '' } | Select-Object Value).Count
    if ($uniqueCount -lt $count) {
        $_.Cancel = $true
        [System.Windows.MessageBox]::Show('Keybind is not unique. Please press another keybind', 'Error', 'OK', 'Error')
    }
})
$skillsGrid.Add_CellValidated({
    $this.Rows[$_.RowIndex].ErrorText = $null
})
$skillsGrid.Add_CellMouseUp({
    if ($_.Button -eq 'Right') {
        $this.Rows[$args[1].RowIndex].Cells[2].Value = ''
    }
    else {
        $this.Rows[$args[1].RowIndex].Cells[2].Selected = $true
    }
})
$craftingGrid.Add_CellMouseUp({
    if ($_.Button -eq 'Right') {
        $rowIndex = $_.RowIndex
        $firstStep = $this.Rows[0].Cells[0].Value
        if ($rowIndex -ne $this.NewRowIndex) {
            $this.Rows.RemoveAt($rowIndex)
            if ($this.NewRowIndex -gt 0) {
                $clearRotation.Visible = $true
            }
            else {
                $clearRotation.Visible = $false
                $queueBtn.Visible = $false
            }
        }
        else {
            $this.Rows[$rowIndex].Cells[0].Value = ''
        }
        if ($this.NewRowIndex -eq 0 -or ($firstStep -eq 28 -and $this.NewRowIndex -le 1)) {
            $reflect.Enabled = $true
            $trainedEye.Enabled = $true
            $muscleMemory.Enabled = $true
        }
        elseif ($firstStep -ne 28 -or $this.NewRowIndex -gt 1) {
            $reflect.Enabled = $false
            $trainedEye.Enabled = $false
            $muscleMemory.Enabled = $false
        }
    }
})
$syncHash.Window.Add_Closing({
    $NativeMethods::SetThreadExecutionState([uint32]"0x80000000")
    $NativeMethods::BlockInput(0)
    if($syncHash.FFXIVHandle -and $syncHash.HDC) {$NativeMethods::ReleaseDC($syncHash.FFXIVHandle, $syncHash.HDC)}
    $Powershell.Dispose()
    $Runspace.Close()
})
$syncHash.Window.Add_Load({
    $tooltip1 = New-Object System.Windows.Forms.ToolTip
    $skillsGrid.AutoSizeColumnsMode = 'AllCells'
    $skillsGrid.ColumnHeadersHeightSizeMode = 1
    $skillsGrid.AllowUserToAddRows = $false
    $skillsGrid.AllowUserToResizeRows = $false
    $skillsGrid.AutoGenerateColumns = $true
    $skillsGrid.RowHeadersVisible = $false
    $skillsGrid.MultiSelect = $false
    $skillsGrid.SelectionMode = 1
    $skillsGrid.ReadOnly = $true
    for ($local:i = 0; $i -lt $ds.Tables[0].Columns.Count; $i++) {
        if ($i -ne 1 -and $i -ne 2) {
            $skillsGrid.Columns[$i].Visible = $false
        }
        $skillsGrid.Columns[$i].SortMode = 0
    }
    $skillsGrid.Rows[0].Cells[2].Selected = $true

    $craftingGrid.AutoSizeColumnsMode = 'Fill'
    $craftingGrid.ColumnHeadersHeightSizeMode = 1
    $craftingGrid.AllowUserToResizeRows = $false
    $craftingGrid.AutoGenerateColumns = $true
    $craftingGrid.ColumnHeadersVisible = $false
    $craftingGrid.RowHeadersVisible = $false
    $craftingGrid.MultiSelect = $false

    $queueGrid.AutoGenerateColumns = $true
    $queueGrid.ColumnHeadersHeightSizeMode = 1
    $queueGrid.AllowUserToAddRows = $true
    $queueGrid.AllowUserToResizeColumns = $false
    $queueGrid.AllowUserToResizeRows = $false
    $queueGrid.ColumnHeadersVisible = $true
    $queueGrid.RowHeadersVisible = $false
    $queueGrid.MultiSelect = $false

    loadKeybinds
    foreach ($row in $ds.Tables[0].DefaultView) {
        if(!$row.Key) { continue }
            $tooltip = @"
$($row.Name) - CP $($row.CP)
$($row.Description)
"@
        $tooltip1.SetToolTip($(Get-Variable $row.Key -ValueOnly), $tooltip)
        $(Get-Variable $row.Key -ValueOnly).Add_Click({
            $lastRow = $craftingGrid.NewRowIndex
            $clearRotation.Visible = $true
            $craftingGrid.Rows.Add()
            $craftingGrid.Rows[$lastRow].Cells[0].Value = "$($this.Tag)"
            if ($craftingGrid.FirstDisplayedCell.Value -eq 28 -and $lastRow -lt 1) {
                $reflect.Enabled = $true
                $trainedEye.Enabled = $true
                $muscleMemory.Enabled = $true
            }
            elseif ($craftingGrid.FirstDisplayedCell.Value -ne 28 -or $lastRow -gt 0) {
                $reflect.Enabled = $false
                $trainedEye.Enabled = $false
                $muscleMemory.Enabled = $false
            }
        })
    }
})