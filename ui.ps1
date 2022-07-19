$craftingJobs = @('Carpenter','Blacksmith','Armorer','Goldsmith','Leatherworker','Weaver','Alchemist','Culinarian')
<#
----------------------------------------------------------------
                            Import Form
#>
$importText = New-Object System.Windows.Forms.TextBox
$importText.Dock = 'Bottom'
$importText.Height = 300
$importText.Multiline = $true
$importText.Scrollbars = 'Vertical'

$importBtn = New-Object System.Windows.Forms.Button
$importBtn.Dock = 'Bottom'
$importBtn.Text = 'Import Macro'

$panel1 = New-Object System.Windows.Forms.Panel
$panel1.Dock = 'Bottom'
$panel1.Height = 40
$panel1.Controls.Add($importBtn)

$importForm = New-Object System.Windows.Forms.Form
$importForm.Text = 'Import in-game macros'
$importForm.Height = 400
$importForm.Width = 400
$importForm.Padding = 10
$importForm.FormBorderStyle = 'Fixed3D'
$importForm.StartPosition = 'CenterScreen'
$importForm.Topmost = $true
$importForm.MaximizeBox = $false
$importForm.Controls.AddRange(@($importText,$panel1))
<#                          Import Form
----------------------------------------------------------------
                            Main Form
#>

$main = [System.Windows.Forms.Form] @{
Text            = 'FFXIV Macro Crafter'
Height          = 510
Width           = 570
FormBorderStyle = 'Fixed3D'
StartPosition   = 'CenterScreen'
Topmost         = $true
MaximizeBox     = $false}
$syncHash.Window = $main

                            <# Collapse/Expand #>

$expandCollapse = [System.Windows.Forms.Button] @{
Text    = 'Collapse'
Dock    = 'Right'
Height  = 23
Width   = 60}

$clearRotation = [System.Windows.Forms.Button] @{
Text    = 'Clear All'
Dock    = 'Right'
Height  = 23
Width   = 60
Visible = $false}

                            <# Tabpages #>

$craftingTab = [System.Windows.Forms.Tabpage] @{
Text = 'Crafting'
Padding = New-Object System.Windows.Forms.Padding(10,0,10,0)}

$queueTab = [System.Windows.Forms.Tabpage] @{
Text = 'Queue'
Padding = New-Object System.Windows.Forms.Padding(10,0,10,0)}

$keybindsTab = [System.Windows.Forms.Tabpage] @{
Text = 'Keybinds'
Padding = New-Object System.Windows.Forms.Padding(75,0,75,0)}

$FormTabControl = [System.Windows.Forms.TabControl] @{Dock = 5}
$FormTabControl.Controls.AddRange(@($craftingTab,$queueTab,$keybindsTab))

<#                          Main Form
---------------------------------------------------------------------------------------
                            Queue Tab
#>

$recipeColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$recipeColumn.Name = 'Recipe'
$recipeColumn.HeaderText = 'In-Game Recipe Name'
$recipeColumn.Width = 240

$rotationDataTable = New-Object System.Data.DataTable
[void]$rotationDataTable.Columns.Add('Rotations')
$(Get-ChildItem "$PSScriptRoot\Rotations\*.json").BaseName | % {
    [void]$rotationDataTable.Rows.Add($_)
}

$materialsColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$materialsColumn.Name = 'Materials'
$materialsColumn.HeaderText = 'Material(s)'
$materialsColumn.MaxInputLength = 1
$materialsColumn.Width = 65

$rotationColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$rotationColumn.Name = 'Rotation'
$rotationColumn.HeaderText = 'Rotation'
$rotationColumn.DataSource = $rotationDataTable
$rotationColumn.DisplayMember = 'Rotations'
$rotationColumn.ValueMember = 'Rotations'
$rotationColumn.Width = 95
$rotationColumn.DropDownWidth = 160

$jobDataTable = New-Object System.Data.DataTable
[void]$jobDataTable.Columns.AddRange(@('ID','Crafter'))
$gearsetJson = Get-Content "$PSScriptRoot\gearset.json" | ConvertFrom-Json
0..($craftingJobs.Count-1) | % {
    $crafter = $craftingJobs[$_]
    $gearsetNumber = $gearsetJson | Where-Object{$_.Name -eq $crafter} | Select-Object -ExpandProperty Value
    [void]$jobDataTable.Rows.Add(@($gearsetNumber,$crafter))
}

$jobColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$jobColumn.Name = 'Crafter'
$jobColumn.HeaderText = 'Crafter'
$jobColumn.DataSource = $jobDataTable
$jobColumn.DisplayMember = 'Crafter'
$jobColumn.ValueMember = 'ID'
$jobColumn.Width = 70
$jobColumn.DropDownWidth  = 90

$craftsColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$craftsColumn.Name = 'Crafts'
$craftsColumn.HeaderText = 'Craft(s)'
$craftsColumn.MaxInputLength = 3
$craftsColumn.Width = 49

$queueGrid = New-Object System.Windows.Forms.DataGridView
$queueGrid.Dock = 1
$queueGrid.Height = 250
$queueGrid.EditMode = 'EditOnEnter'
[void]$queueGrid.Columns.AddRange($recipeColumn,$materialsColumn,$rotationColumn,$jobColumn,$craftsColumn)

                                    <# Controls #>

$useFoodbuffQueue = New-Object System.Windows.Forms.CheckBox
$useFoodbuffQueue.AutoSize = $true
$useFoodbuffQueue.Dock = 'Right'

$useFoodbuffLblQueue = New-Object System.Windows.Forms.Label
$useFoodbuffLblQueue.AutoSize = $true
$useFoodbuffLblQueue.Text = 'Food Buff'
$useFoodbuffLblQueue.Dock = 'Right'

$useMedicineQueue = New-Object System.Windows.Forms.CheckBox
$useMedicineQueue.AutoSize = $true
$useMedicineQueue.Dock = 'Right'

$useMedicineLblQueue = New-Object System.Windows.Forms.Label
$useMedicineLblQueue.AutoSize = $true
$useMedicineLblQueue.Text = 'Medicine'
$useMedicineLblQueue.Dock = 'Right'

$craftBtnQueue = New-Object System.Windows.Forms.Button
$craftBtnQueue.Text = 'Craft'
$craftBtnQueue.Dock = 'Right'
$syncHash.CraftQueueBtn = $craftBtnQueue

$pauseBtnQueue = New-Object System.Windows.Forms.Button
$pauseBtnQueue.Text = 'Pause'
$pauseBtnQueue.Dock = 'Right'
$pauseBtnQueue.Enabled = $false
$syncHash.PauseQueueBtn = $pauseBtnQueue

$gearsetLabelDefault =  @{
    Dock = 'Left'
    Height = 40
}

$gearsetNumericDefault = @{
    Dock = 'Right'
    Width = 50
    Minimum = 1
    Maximum = 47
    Tag = 'Numeric'
}

$gearsetPanelDefault = @{
    Dock = 'Top'
    Height = 24
}

$panelColumnDefault = @{
    Dock = 'Left'
    Width = 258
}

$carpenterLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$carpenterLbl.Text = 'Carpenter'
$carpenterNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$carpenterNumeric.Name = 'Carpenter'
$carpenterGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$carpenterGroup.Controls.AddRange(@($carpenterLbl,$carpenterNumeric))

$blacksmithLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$blacksmithLbl.Text = 'Blacksmith'
$blacksmithNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$blacksmithNumeric.Name = 'Blacksmith'
$blacksmithGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$blacksmithGroup.Controls.AddRange(@($blacksmithLbl,$blacksmithNumeric))

$armorerLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$armorerLbl.Text = 'Armorer'
$armorerNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$armorerNumeric.Name = 'Armorer'
$armorerGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$armorerGroup.Controls.AddRange(@($armorerLbl,$armorerNumeric))

$goldsmithLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$goldsmithLbl.Text = 'Goldsmith'
$goldsmithNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$goldsmithNumeric.Name = 'Goldsmith'
$goldsmithGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$goldsmithGroup.Controls.AddRange(@($goldsmithLbl,$goldsmithNumeric))

$leatherworkerLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$leatherworkerLbl.Text = 'Leatherworker'
$leatherworkerNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$leatherworkerNumeric.Name= 'Leatherworker'
$leatherworkerGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$leatherworkerGroup.Controls.AddRange(@($leatherworkerLbl,$leatherworkerNumeric))

$weaverLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$weaverLbl.Text = 'Weaver'
$weaverNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$weaverNumeric.Name = 'Weaver'
$weaverGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$weaverGroup.Controls.AddRange(@($weaverLbl,$weaverNumeric))

$alchemistLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$alchemistLbl.Text = 'Alchemist'
$alchemistNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$alchemistNumeric.Name = 'Alchemist'
$alchemistGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$alchemistGroup.Controls.AddRange(@($alchemistLbl,$alchemistNumeric))

$culinarianLbl = New-Object System.Windows.Forms.Label -Property $gearsetLabelDefault
$culinarianLbl.Text = 'Culinarian'
$culinarianNumeric = New-Object System.Windows.Forms.NumericUpDown -Property $gearsetNumericDefault
$culinarianNumeric.Name = 'Culinarian'
$culinarianGroup = New-Object System.Windows.Forms.Panel -Property $gearsetPanelDefault
$culinarianGroup.Controls.AddRange(@($culinarianLbl,$culinarianNumeric))

$saveGearset = [System.Windows.Forms.Button] @{Text = 'Save';Dock = 'Right';MaximumSize = New-Object System.Drawing.Size(100,24)}
$loadGearset = [System.Windows.Forms.Button] @{Text = 'Load';Dock = 'Left' ;MaximumSize = New-Object System.Drawing.Size(100,24)}

                                <# Panels #>
$gearsetPanel1 = New-Object System.Windows.Forms.Panel -Property $panelColumnDefault
$gearsetPanel2 = New-Object System.Windows.Forms.Panel -Property $panelColumnDefault
$gearsetPanel1.Controls.AddRange(@($loadGearset,$culinarianGroup,$weaverGroup,$goldsmithGroup,$blacksmithGroup))
$gearsetPanel2.Controls.AddRange(@($saveGearset,$alchemistGroup,$leatherworkerGroup,$armorerGroup,$carpenterGroup))

$queueGroup = New-Object System.Windows.Forms.Panel
$queueGroup.Dock = 'Top'
$queueGroup.Height = 23
$queueGroup.Controls.AddRange(@($useFoodbuffLblQueue,$useFoodbuffQueue,$useMedicineLblQueue,$useMedicineQueue,$syncHash.CraftQueueBtn,$syncHash.PauseQueueBtn))

$gearsetGroupBox = [System.Windows.Forms.GroupBox] @{Dock = 'Top';Height = 150;Text = 'Gearset Number'}
$gearsetGroupBox.Controls.AddRange(@($gearsetPanel1,$gearsetPanel2))

loadGearsets

<#                              Queue Tab
---------------------------------------------------------------------------------------
                                Keybinds Tab
#>

$skillsGrid = New-Object System.Windows.Forms.DataGridView
$skillsGrid.Dock = 'Top'
$skillsGrid.Height = 200
$skillsGrid.DataSource = $ds.Tables[0].DefaultView

$saveKeybindBtn = New-Object System.Windows.Forms.Button
$saveKeybindBtn.Text = 'Save'
$saveKeybindBtn.Dock = 'Top'

$loadKeybindBtn = New-Object System.Windows.Forms.Button
$loadKeybindBtn.Text = 'Load'
$loadKeybindBtn.Dock = 'Top'

$confirmKeyLbl = New-Object System.Windows.Forms.Label
$confirmKeyLbl.Width = 100
$confirmKeyLbl.Text = 'Confirm'
$confirmKeyLbl.Dock = 'Left'

$confirmKeyTxt = New-Object System.Windows.Forms.TextBox
$confirmKeyTxt.Dock = 'Left'
$confirmKeyTxt.Width = 100
$confirmKeyTxt.ReadOnly = $true

$foodBuffLbl = New-Object System.Windows.Forms.Label
$foodBuffLbl.Width = 100
$foodBuffLbl.Text = 'Food Buff'
$foodBuffLbl.Dock = 'Left'

$foodBuffTxt = New-Object System.Windows.Forms.TextBox
$foodBuffTxt.Dock = 'Left'
$foodBuffTxT.Width = 100
$foodBuffTxT.ReadOnly = $true

$medicineLbl = New-Object System.Windows.Forms.Label
$medicineLbl.Width = 100
$medicineLbl.Text = 'Medicine'
$medicineLbl.Dock = 'Left'

$medicineTxt = New-Object System.Windows.Forms.TextBox
$medicineTxt.Dock = 'Left'
$medicineTxT.Width = 100
$medicineTxT.ReadOnly = $true

$craftLogLbl = New-Object System.Windows.Forms.Label
$craftLogLbl.Width = 100
$craftLogLbl.Text = 'Crafting Log'
$craftLogLbl.Dock = 'Left'

$craftLogTxt = New-Object System.Windows.Forms.TextBox
$craftLogTxt.Dock = 'Left'
$craftLogTxT.Width = 100
$craftLogTxT.ReadOnly = $true

$foodBuffPanel = New-Object System.Windows.Forms.Panel
$foodBuffPanel.Dock = 'Top'
$foodBuffPanel.Height = 24
$foodBuffPanel.Controls.AddRange(@($foodBuffTxt,$foodBuffLbl))

$confirmPanel = New-Object System.Windows.Forms.Panel
$confirmPanel.Dock = 'Top'
$confirmPanel.Height = 24
$confirmPanel.Controls.AddRange(@($confirmKeyTxt,$confirmKeyLbl))

$medicinePanel = New-Object System.Windows.Forms.Panel
$medicinePanel.Dock = 'Top'
$medicinePanel.Height = 24
$medicinePanel.Controls.AddRange(@($medicineTxt,$medicineLbl))

$craftLogPanel = New-Object System.Windows.Forms.Panel
$craftLogPanel.Dock = 'Top'
$craftLogPanel.Height = 24
$craftLogPanel.Controls.AddRange(@($craftLogTxt,$craftLogLbl))

<#                              Keybinds Tab
-------------------------------------------------------------------------------------
                                Crafting Tab
#>

$marcoColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$marcoColumn.Name = 'Macro'
$marcoColumn.DataSource = $ds.Tables[0]
$marcoColumn.DisplayMember = 'Name'
$marcoColumn.ValueMember = 'ID'
$marcoColumn.DisplayStyle = 'Nothing'

$craftingGrid = New-Object System.Windows.Forms.DataGridView
$craftingGrid.Dock = 1
$craftingGrid.Height = 200
$craftingGrid.ReadOnly = $true
[void]$craftingGrid.Columns.Add($marcoColumn)

                                <# Recipe Controls #>

$levelLbl = New-Object System.Windows.Forms.Label
$levelLbl.Text = 'Level'
$levelLbl.Dock = 'Left'
$levelLbl.AutoSize = $true

$levelNumeric = New-Object System.Windows.Forms.NumericUpDown
$levelNumeric.Dock = 'Left'
$levelNumeric.Minimum = 1
$levelNumeric.Maximum = 90
$levelNumeric.Value = 90
$levelNumeric.Width = 50

$classLbl = New-Object System.Windows.Forms.Label
$classLbl.Text = 'Class'
$classLbl.Dock = 'Left'
$classLbl.AutoSize = $true

$classDropdown = New-Object System.Windows.Forms.ComboBox
$classDropdown.Dock = 'Left'
$classDropdown.Width = 60
$classDropdown.DropDownStyle = 'DropDownList'
$classDropdown.Items.AddRange(@('ALC','ARM','BSM','CRP','CUL','GSM','LTW','WVR'))
$classDropdown.SelectedItem = 'CUL'

$loadRecipeDropdown = New-Object System.Windows.Forms.ComboBox
$loadRecipeDropdown.Dock = 'Right'
$loadRecipeDropdown.Height = 25
$listRotations = $(Get-ChildItem "$PSScriptRoot\Rotations\*.json").BaseName
if(-not($null -eq $listRotations)){
    $loadRecipeDropdown.Items.AddRange($listRotations)
}

$saveRecipeBtn = New-Object System.Windows.Forms.Button
$saveRecipeBtn.Dock = 'Right'
$saveRecipeBtn.Width = 50
$saveRecipeBtn.Text = 'Save'

$deleteRecipeBtn = New-Object System.Windows.Forms.Button
$deleteRecipeBtn.Dock = 'Right'
$deleteRecipeBtn.Width = 50
$deleteRecipeBtn.Text = 'Delete'

$importMacros = New-Object System.Windows.Forms.Button
$importMacros.Dock = 'Right'
$importMacros.AutoSize = $true
$importMacros.Text = 'Import Macro(s)'

                                            <# Skill Buttons #>

$basicSynth = New-Object System.Windows.Forms.Button
$basicSynth.Tag = 0
$basicSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\basicSynth.png")

$rapidSynth = New-Object System.Windows.Forms.Button
$rapidSynth.Tag = 1
$rapidSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\rapidSynth.png")

$muscleMemory = New-Object System.Windows.Forms.Button
$muscleMemory.Tag = 2
$muscleMemory.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\muscleMemory.png")

$carefulSynth = New-Object System.Windows.Forms.Button
$carefulSynth.Tag = 3
$carefulSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\carefulSynth.png")

$focusedSynth = New-Object System.Windows.Forms.Button
$focusedSynth.Tag = 4
$focusedSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\focusedSynth.png")

$groundwork = New-Object System.Windows.Forms.Button
$groundwork.Tag = 5
$groundwork.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\groundwork.png")

$delicateSynth = New-Object System.Windows.Forms.Button
$delicateSynth.Tag = 6
$delicateSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\delicateSynth.png")

$observe = New-Object System.Windows.Forms.Button
$observe.Tag = 7
$observe.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\observe.png")

$intensiveSynth = New-Object System.Windows.Forms.Button
$intensiveSynth.Tag = 8
$intensiveSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\intensiveSynth.png")

$prudentSynth = New-Object System.Windows.Forms.Button
$prudentSynth.Tag = 9
$prudentSynth.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\prudentSynth.png")

$basicTouch = New-Object System.Windows.Forms.Button
$basicTouch.Tag = 10
$basicTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\basicTouch.png")

$standardTouch = New-Object System.Windows.Forms.Button
$standardTouch.Tag = 11
$standardTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\standardTouch.png")

$byregotsBlessing = New-Object System.Windows.Forms.Button
$byregotsBlessing.Tag = 12
$byregotsBlessing.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\byregotsBlessing.png")

$preciseTouch = New-Object System.Windows.Forms.Button
$preciseTouch.Tag = 13
$preciseTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\preciseTouch.png")

$prudentTouch = New-Object System.Windows.Forms.Button
$prudentTouch.Tag = 14
$prudentTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\prudentTouch.png")

$focusedTouch = New-Object System.Windows.Forms.Button
$focusedTouch.Tag  = 15
$focusedTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\focusedTouch.png")

$reflect = New-Object System.Windows.Forms.Button
$reflect.Tag = 16
$reflect.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\reflect.png")

$preparatoryTouch = New-Object System.Windows.Forms.Button
$preparatoryTouch.Tag = 17
$preparatoryTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\preparatoryTouch.png")

$trainedEye = New-Object System.Windows.Forms.Button
$trainedEye.Tag = 18
$trainedEye.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\trainedEye.png")

$advancedTouch = New-Object System.Windows.Forms.Button
$advancedTouch.Tag = 19
$advancedTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\CUL\advancedTouch.png")

$trainedFinesse = New-Object System.Windows.Forms.Button
$trainedFinesse.Tag = 20
$trainedFinesse.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\trainedFinesse.png")

$mastersMind = New-Object System.Windows.Forms.Button
$mastersMind.Tag = 21
$mastersMind.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\mastersMind.png")

$wasteNot = New-Object System.Windows.Forms.Button
$wasteNot.Tag = 22
$wasteNot.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\wasteNot.png")

$wasteNot2 = New-Object System.Windows.Forms.Button
$wasteNot2.Tag = 23
$wasteNot2.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\wasteNot2.png")

$manipulation = New-Object System.Windows.Forms.Button
$manipulation.Tag = 24
$manipulation.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\manipulation.png")

$veneration = New-Object System.Windows.Forms.Button
$veneration.Tag = 25
$veneration.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\veneration.png")

$greatStrides = New-Object System.Windows.Forms.Button
$greatStrides.Tag = 26
$greatStrides.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\greatStrides.png")

$innovation = New-Object System.Windows.Forms.Button
$innovation.Tag = 27
$innovation.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\innovation.png")

$finalAppraisal = New-Object System.Windows.Forms.Button
$finalAppraisal.Tag = 28
$finalAppraisal.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\finalAppraisal.png")

$hastyTouch = New-Object System.Windows.Forms.Button
$hastyTouch.Tag = 29
$hastyTouch.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\hastyTouch.png")

$heartSoul = New-Object System.Windows.Forms.Button
$heartSoul.Tag =  30
$heartSoul.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\heartSoul.png")

foreach($button in ($ds.Tables[0].Rows | Select-Object -ExpandProperty Key)){
    if(!$button){ continue }
    $(Get-Variable -Name $button -ValueOnly).Dock = 'Left'
    $(Get-Variable -Name $button -ValueOnly).FlatStyle = 'Flat'
    $(Get-Variable -Name $button -ValueOnly).FlatAppearance.BorderSize = 0
    $(Get-Variable -Name $button -ValueOnly).MaximumSize = New-Object System.Drawing.Size(43,41)
}

                                            <# Skill Labels #>

$progressLbl = New-Object System.Windows.Forms.Label
$progressLbl.Text = 'Progression'
$progressLbl.Dock = 1
$progressLbl.Height = 17

$qualityLbl = New-Object System.Windows.Forms.Label
$qualityLbl.Text = 'Quality'
$qualityLbl.Dock = 1
$qualityLbl.Height = 17

$otherLbl = New-Object System.Windows.Forms.Label
$otherLbl.Text = 'Buff/Repair/Other'
$otherLbl.Dock = 1
$otherLbl.Height = 17

                                            <# Craft Controls #>

$useFoodbuff = New-Object System.Windows.Forms.CheckBox
$useFoodbuff.AutoSize = $true
$useFoodbuff.Dock = 'Right'

$useFoodbuffLbl = New-Object System.Windows.Forms.Label
$useFoodbuffLbl.AutoSize = $true
$useFoodbuffLbl.Text = 'Food Buff'
$useFoodbuffLbl.Dock = 'Right'

$useMedicine = New-Object System.Windows.Forms.CheckBox
$useMedicine.AutoSize = $true
$useMedicine.Dock = 'Right'

$useMedicineLbl = New-Object System.Windows.Forms.Label
$useMedicineLbl.AutoSize = $true
$useMedicineLbl.Text = 'Medicine'
$useMedicineLbl.Dock = 'Right'

$craftLog = New-Object System.Windows.Forms.CheckBox
$craftLog.AutoSize = $true
$craftLog.Dock = 'Right'

$craftLogLbl = New-Object System.Windows.Forms.Label
$craftLogLbl.AutoSize = $true
$craftLogLbl.Text = 'Crafting Log'
$craftLogLbl.Dock = 'Right'

$craftLbl = New-Object System.Windows.Forms.Label
$craftLbl.AutoSize = $true
$craftLbl.Text = 'Craft(s)'
$craftLbl.Dock = 'Right'

$craftNumeric = New-Object System.Windows.Forms.NumericUpDown
$craftNumeric.Dock = 'Right'
$craftNumeric.Maximum = 999
$craftNumeric.Minimum = 1
$craftNumeric.Width = 40

$craftBtn = New-Object System.Windows.Forms.Button
$craftBtn.Text = 'Craft'
$craftBtn.Dock = 'Right'
$syncHash.CraftBtn = $craftBtn

$pauseBtn = New-Object System.Windows.Forms.Button
$pauseBtn.Text = 'Pause'
$pauseBtn.Dock = 'Right'
$pauseBtn.Enabled = $false
$syncHash.PauseBtn = $pauseBtn

<# Panels #>

$expandCollapsePanel = New-Object System.Windows.Forms.Panel
$expandCollapsePanel.Height = 22
$expandCollapsePanel.Top = 0
$expandCollapsePanel.Left = 350
$expandCollapsePanel.Margin = 0
$expandCollapsePanel.Controls.AddRange(@($clearRotation,$expandCollapse))

$levelPanel = New-Object System.Windows.Forms.Panel
$levelPanel.Dock = 'Top'
$levelPanel.Height = 23
$levelPanel.Controls.AddRange(@($levelNumeric,$levelLbl,$classDropdown,$classLbl,$loadRecipeDropdown,$deleteRecipeBtn,$saveRecipeBtn,$importMacros))

$progress = New-Object System.Windows.Forms.Panel
$progress.Height = 41
$progress.Dock = 1
$progress.Controls.AddRange(@($muscleMemory,$intensiveSynth,$focusedSynth,$groundwork,$rapidSynth,$prudentSynth,$carefulSynth,$basicSynth))

$quality = New-Object System.Windows.Forms.Panel
$quality.Height = 41
$quality.Dock = 1
$quality.Controls.AddRange(@($reflect,$trainedEye,$trainedFinesse,$byregotsBlessing,$preparatoryTouch,$prudentTouch,$focusedTouch,$preciseTouch,$hastyTouch,$advancedTouch,$standardTouch,$basicTouch))

$buff = New-Object System.Windows.Forms.Panel
$buff.AutoSize = $true
$buff.Height = 41
$buff.Dock = 'Left'
$buff.Controls.AddRange(@($heartSoul,$finalAppraisal,$veneration,$innovation,$greatStrides,$wasteNot2,$wasteNot))

$repair = New-Object System.Windows.Forms.Panel
$repair.AutoSize = $true
$repair.Dock = 'Left'
$repair.Height = 60
$repair.Padding = New-Object System.Windows.Forms.Padding(10,0,10,0)
$repair.Controls.AddRange(@($manipulation,$mastersMind))

$other = New-Object System.Windows.Forms.Panel
$other.AutoSize = $true
$other.Dock = 'Left'
$other.Controls.AddRange(@($delicateSynth,$observe))

$craftOther = New-Object System.Windows.Forms.Panel
$craftOther.Dock = 1
$craftOther.Controls.AddRange(@($buff, $repair, $other))

$craftGroup = New-Object System.Windows.Forms.Panel
$craftGroup.Dock = 1
$craftGroup.Height = 210
$craftGroup.Padding = New-Object System.Windows.Forms.Padding(0,10,0,0)
$craftGroup.Controls.AddRange(@($craftOther,$otherLbl,$quality,$qualityLbl,$progress,$progressLbl,$levelPanel))

$craftGroup2 = New-Object System.Windows.Forms.Panel
$craftGroup2.Dock = 'Top'
$craftGroup2.Height = 23
$craftGroup2.Controls.AddRange(@($useFoodbuffLbl,$useFoodbuff,$useMedicineLbl,$useMedicine,$craftLbl,$craftNumeric,$syncHash.CraftBtn,$syncHash.PauseBtn))

$queueTab.Controls.AddRange(@($gearsetGroupBox,$queueGroup,$queueGrid))
$keybindsTab.Controls.AddRange(@($saveKeybindBtn,$loadKeybindBtn,$medicinePanel,$foodBuffPanel,$craftLogPanel,$confirmPanel,$skillsGrid))
$craftingTab.Controls.AddRange(@($craftGroup2,$craftGroup,$craftingGrid))
$syncHash.Window.Controls.AddRange(@($expandCollapsePanel,$FormTabControl))