if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(IntPtr ZeroOnly, string lpWindowName);


    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);


    [DllImport("user32.dll")]
    public static extern bool EnableWindow(IntPtr hWnd, int bEnable);
}
'@
[NativeMethods]::ShowWindowAsync((Get-Process -id $pid).MainWindowHandle, 2) | Out-Null
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
$ds = New-Object System.Data.Dataset
$null = $ds.ReadXml("$PSScriptRoot\skills.xml")
$listRotations = $(Get-ChildItem "$PSScriptRoot\Rotations\*.json").BaseName
$importTextCheck = $false
if(!(Test-Path -Path "$PSScriptRoot\keybinds.json" -PathType Leaf) -or !(Test-Path -Path "$PSScriptRoot\controls.json" -PathType Leaf)) {
    [System.Windows.MessageBox]::Show('Please configure your keybinds before using this tool','Information','OK','Information')
}

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

$main = New-Object System.Windows.Forms.Form
$main.Text ='FFXIV Macro Crafter'
$main.FormBorderStyle  = 0
$main.Height = 510
$main.Width = 570
$main.FormBorderStyle = 'Fixed3D'
$main.StartPosition = 'CenterScreen'
$main.Topmost = $true
$main.MaximizeBox = $false

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000

$craftingTab = New-object System.Windows.Forms.Tabpage
$craftingTab.DataBindings.DefaultDataSourceUpdateMode = 0
$craftingTab.Name = 'Tab1'
$craftingTab.Text = 'Crafting Macro'
$craftingTab.Padding = New-Object System.Windows.Forms.Padding(10,00,10,0)

$keybindsTab = New-object System.Windows.Forms.Tabpage
$keybindsTab.DataBindings.DefaultDataSourceUpdateMode = 0
$keybindsTab.Name = 'Tab2'
$keybindsTab.Text = 'Keybinds'
$keybindsTab.Padding = New-Object System.Windows.Forms.Padding(75,0,75,0)

$FormTabControl = New-object System.Windows.Forms.TabControl
$FormTabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$FormTabControl.Dock = 5
$FormTabControl.Controls.AddRange(@($craftingTab,$keybindsTab))

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
$confirmKeyLbl.Text = 'Confirm Key'
$confirmKeyLbl.Dock = 'Left'

$confirmKeyTxt = New-Object System.Windows.Forms.TextBox
$confirmKeyTxt.Dock = 'Left'
$confirmKeyTxt.Width = 100
$confirmKeyTxt.ReadOnly = $true

$foodBuffLbl = New-Object System.Windows.Forms.Label
$foodBuffLbl.Width = 100
$foodBuffLbl.Text = 'Food Buff Key'
$foodBuffLbl.Dock = 'Left'

$foodBuffTxt = New-Object System.Windows.Forms.TextBox
$foodBuffTxt.Dock = 'Left'
$foodBuffTxT.Width = 100
$foodBuffTxT.ReadOnly = $true

$medicineLbl = New-Object System.Windows.Forms.Label
$medicineLbl.Width = 100
$medicineLbl.Text = 'Medicine Key'
$medicineLbl.Dock = 'Left'

$medicineTxt = New-Object System.Windows.Forms.TextBox
$medicineTxt.Dock = 'Left'
$medicineTxT.Width = 100
$medicineTxT.ReadOnly = $true

$craftLogLbl = New-Object System.Windows.Forms.Label
$craftLogLbl.Width = 100
$craftLogLbl.Text = 'Crafting Log Key'
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

$marcoColumn = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$marcoColumn.Name = 'Macro'
$marcoColumn.HeaderText = 'Name'
$marcoColumn.DataSource = $ds.Tables[0]
$marcoColumn.DisplayMember = 'Name'
$marcoColumn.ValueMember = 'ID'
$marcoColumn.DisplayStyle = 'Nothing'

$craftingGrid = New-Object System.Windows.Forms.DataGridView
$craftingGrid.Dock = 1
$craftingGrid.Height = 200
$craftingGrid.ReadOnly = $true
$craftingGrid.Columns.Add($marcoColumn)

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
if(-not($null -eq $listRotations)){
    $loadRecipeDropdown.Items.AddRange($listRotations)
}

$expandCollapse = New-Object System.Windows.Forms.Button
$expandCollapse.Text = 'Collapse'
$expandCollapse.Dock = 'Right'
$expandCollapse.Height = 23
$expandCollapse.Width = 60

$clearRotation = New-Object System.Windows.Forms.Button
$clearRotation.Text = 'Clear All'
$clearRotation.Dock = 'Right'
$clearRotation.Height = 23
$clearRotation.Width = 60
$clearRotation.Visible = $false

$expandCollapsePanel = New-Object System.Windows.Forms.Panel
$expandCollapsePanel.Height = 23
$expandCollapsePanel.Top = 0
$expandCollapsePanel.Left = 350
$expandCollapsePanel.Margin = 0
$expandCollapsePanel.Controls.AddRange(@($clearRotation,$expandCollapse))

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

$levelPanel = New-Object System.Windows.Forms.Panel
$levelPanel.Dock = 'Top'
$levelPanel.Height = 23
$levelPanel.Controls.AddRange(@($levelNumeric,$levelLbl,$classDropdown,$classLbl,$loadRecipeDropdown,$deleteRecipeBtn,$saveRecipeBtn,$importMacros))

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

$manipulationCheck = New-Object System.Windows.Forms.CheckBox
$manipulationCheck.BackColor = [System.Drawing.Color]::FromName("Transparent")
$manipulationCheck.Location = New-Object System.Drawing.Point(83,0)
$manipulationCheck.AutoSize = $true

$manipulation = New-Object System.Windows.Forms.Button
$manipulation.Tag = 24
$manipulation.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\icons\manipulation.png")
$manipulation.Enabled = $false

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

$tooltip1 = New-Object System.Windows.Forms.ToolTip

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
$buff.Controls.AddRange(@($finalAppraisal,$veneration,$innovation,$greatStrides,$wasteNot2,$wasteNot))

$repair = New-Object System.Windows.Forms.Panel
$repair.AutoSize = $true
$repair.Dock = 'Left'
$repair.Height = 60
$repair.Padding = New-Object System.Windows.Forms.Padding(10,0,10,0)
$repair.Controls.AddRange(@($manipulationCheck,$manipulation,$mastersMind))

$other = New-Object System.Windows.Forms.Panel
$other.AutoSize = $true
$other.Dock = 'Left'
$other.Controls.AddRange(@($delicateSynth,$observe))

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

$craftOther = New-Object System.Windows.Forms.Panel
$craftOther.Dock = 1
$craftOther.Controls.AddRange(@($buff, $repair, $other))

$craftGroup = New-Object System.Windows.Forms.Panel
$craftGroup.Dock = 1
$craftGroup.Height = 210
$craftGroup.Padding = New-Object System.Windows.Forms.Padding(0,10,0,0)
$craftGroup.Controls.AddRange(@($craftOther,$otherLbl,$quality,$qualityLbl,$progress,$progressLbl,$levelPanel))

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
$craftLogLbl.Text = 'Crafting Log Key'
$craftLogLbl.Dock = 'Right'

$craftLbl = New-Object System.Windows.Forms.Label
$craftLbl.Text = 'Num of Craft(s)'
$craftLbl.Dock = 'Right'

$craftNumeric = New-Object System.Windows.Forms.NumericUpDown
$craftNumeric.Dock = 'Right'
$craftNumeric.Maximum = 999
$craftNumeric.Minimum = 1
$craftNumeric.Width = 70

$craftBtn = New-Object System.Windows.Forms.Button
$craftBtn.Text = 'Craft'
$craftBtn.Dock = 'Right'

$craftGroup2 = New-Object System.Windows.Forms.Panel
$craftGroup2.Dock = 'Top'
$craftGroup2.Height = 23
$craftGroup2.Controls.AddRange(@($useFoodbuffLbl,$useFoodbuff,$useMedicineLbl,$useMedicine,$craftLbl,$craftNumeric,$craftBtn))

foreach($button in ($ds.Tables[0].Rows | Select-Object -ExpandProperty Key)){
    $(Get-Variable -Name $button -ValueOnly).Dock = 'Left'
    $(Get-Variable -Name $button -ValueOnly).FlatStyle = 'Flat'
    $(Get-Variable -Name $button -ValueOnly).FlatAppearance.BorderSize = 0
    $(Get-Variable -Name $button -ValueOnly).MaximumSize = New-Object System.Drawing.Size(43,41)
}

$keybindsTab.Controls.AddRange(@($saveKeybindBtn,$loadKeybindBtn,$medicinePanel,$foodBuffPanel,$craftLogPanel,$confirmPanel,$skillsGrid))
$craftingTab.Controls.AddRange(@($craftGroup2,$craftGroup,$craftingGrid))
$main.Controls.AddRange(@($expandCollapsePanel,$FormTabControl))

Get-ChildItem -Path "$PSScriptRoot\modules" | ForEach-Object {. $($_.FullName)}

$main.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
$main.Dispose()