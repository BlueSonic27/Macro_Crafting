$importTextCheck = $false
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
                $(Get-Variable -Name $button -ValueOnly).Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\..\icons\$class\$button.png")
            }
        }
    })
$syncHash.CraftBtn.Add_click({
        if ($syncHash.CraftBtn.Text -eq 'Stop') {
            $syncHash.Stop = $true
            $syncHash.Window.Text = 'FFXIV Macro Crafter - Stopping'
            $syncHash.CraftBtn.Text = 'Abort'
        }
        elseif ($syncHash.CraftBtn.Text -eq 'Abort') {
            $syncHash.Abort = $true
            $syncHash.Pause = $false
            $syncHash.Window.Text = 'FFXIV Macro Crafter'
            $syncHash.CraftBtn.Text = 'Craft'
        } else {
            $syncHash.CraftBtn.Text = 'Stop'
            $syncHash.PauseBtn.Enabled = $true
            craft
        }
    })
$syncHash.PauseBtn.Add_click({
    if($syncHash.Pause){
        $syncHash.PauseBtn.Text = 'Pause'
        $syncHash.CraftBtn.Text = 'Stop'
        $syncHash.Pause = $false
    } else {
        $syncHash.PauseBtn.Text = 'Resume'
        $syncHash.CraftBtn.Text = 'Abort'
        $syncHash.Pause = $true
    }
    })
$loadRecipeDropdown.Add_TextChanged({
        [string]$text = $this.Text
        $escape = '[<>:"\\/|?*]'
        if ($text -match $escape) {
            $this.Text = $text.Substring(0, $text.Length - 1)
            $errorText = @'
A Filename can't contain any of the following characters:
        <>:"/\|?*
'@
            [System.Windows.MessageBox]::Show($errorText, 'Error', 'OK', 'Error')
        }
    })
$loadRecipeDropdown.Add_SelectionChangeCommitted({
        $filename = $args[0].SelectedItem
        $rotation = Get-Content "$PSScriptRoot\..\Rotations\$filename.json" | ConvertFrom-Json
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
$clearRotation.Add_Click({
        $craftingGrid.Rows.Clear()
        $clearRotation.Visible = $false
        $reflect.Enabled = $true
        $trainedEye.Enabled = $true
        $muscleMemory.Enabled = $true    
    })
$saveRecipeBtn.Add_Click({
        if ($null -eq $loadRecipeDropdown.SelectedItem) {
            $filename = $loadRecipeDropdown.Text
        }
        else {
            $filename = $loadRecipeDropdown.SelectedItem
        }
        $rotation = @()
        foreach ($row in $craftingGrid.Rows) {
            $cell = $row.Cells[0].Value
            if ($cell -ne $null) { $rotation += $cell }
        }
        if ($null -eq $filename -or $filename -eq '') {
            [System.Windows.MessageBox]::Show('Please enter a filename', 'Error', 'OK', 'Error')
        }
        else {
            New-Item -Path "$PSScriptRoot\..\Rotations" -ItemType Directory -Force
            ConvertTo-Json -InputObject $rotation -Compress | Out-File -FilePath "$PSScriptRoot\..\Rotations\$filename.json" -NoNewline
            $loadRecipeDropdown.Items.Clear()
            $loadRecipeDropdown.Items.AddRange((Get-ChildItem "$PSScriptRoot\..\Rotations\*.json").BaseName)
            [System.Windows.MessageBox]::Show('Rotation saved', 'Information', 'OK', 'Information')
        }
    })
$deleteRecipeBtn.Add_Click({
        if ($loadRecipeDropdown.SelectedItem) {
            $filename = $loadRecipeDropdown.SelectedItem
        }
        else {
            $filename = $false
        }
        if ($filename) {
            if ([System.Windows.MessageBox]::Show(@"
Are you sure you want to delete this rotation?

        [ $filename ]
"@, 'Warning', 'YesNo', 'Warning') -eq 'Yes') {
                $loadRecipeDropdown.Items.Clear()
                $loadRecipeDropdown.Text = ''
                $loadRecipeDropdown.Items.AddRange((Get-ChildItem "$PSScriptRoot\..\Rotations\*.json").BaseName)
                Remove-Item "$PSScriptRoot\..\Rotations\$filename.json"
            }
        }
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
        $keybindsJson | Out-File -FilePath "$PSScriptRoot\..\keybinds.json" -NoNewline
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
        $controlJson | Out-File -FilePath "$PSScriptRoot\..\controls.json" -NoNewline
    })
$loadKeybindBtn.Add_Click({
        loadKeybinds
    })
$levelNumeric.Add_TextChanged({
        foreach ($skill in ($ds.Tables[0].Rows)) {
            $id = $skill.Key
            $level = [int]$skill.Level
            if ($id -ne 'manipulation') {
                if ($level -gt $this.Value) {
                    $(Get-Variable -Name $id -ValueOnly).Enabled = $false
                }
                else {
                    $(Get-Variable -Name $id -ValueOnly).Enabled = $true
                }
            }
            else {
                if(($level -gt $this.Value) -or ($manipulationCheck.Checked -ne $true)) {
                    $manipulation.Enabled = $false
                }
                else {
                    $manipulation.Enabled = $true
                }
            }
        }
    })
$manipulationCheck.Add_Click({
        if ($levelNumeric.Value -ge 65 ) {
            $manipulation.Enabled = $this.Checked}
        else {
            $manipulation.Enabled = $false
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
    [NativeMethods]::SetThreadExecutionState([uint32]"0x80000000")
    [NativeMethods]::BlockInput(0)
    [NativeMethods]::ReleaseDC($syncHash.FFXIVHandle, $syncHash.HDC)
    $syncHash.Abort = $true
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
        $craftingGrid.RowHeadersVisible = $false
        $craftingGrid.MultiSelect = $false
        $craftingGrid.ColumnHeadersVisible = $false

        loadKeybinds
        foreach ($row in $ds.Tables[0].DefaultView) {
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