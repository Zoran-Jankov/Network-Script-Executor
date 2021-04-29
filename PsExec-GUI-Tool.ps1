#---------------------------------------------------------[Initialisations]--------------------------------------------------------

[hashtable]$ComputerList

$ApplicationTitle = "PsExec-GUI-Tool"
$ItemBackgroundColor = "#003F9A"
$FormBackgroundColor = "#468FEA"
$FormForegroundColor = "#FFFFFF"

$MarginSize = 25
$Separator = 10
$LabelWidth = 100
$BoxWidth = 300
$ItemWidth = $LabelWidth + $BoxWidth
$ItemHeight = 25
$ListBoxHeight = 400
$ProgressBarHeight = 25
$ButtonWidth = 120
$ButtonHeight = 35

$MainFormWidth = $ItemWidth + ($MarginSize * 2)
$MainFormHeight = ($ItemHeight * 3) + $ListBoxHeight + ($Separator * 4) + $ProgressBarHeight + $ButtonHeight + ($MarginSize * 3)
$BoxPosition = $MarginSize + $LabelWidth
$RunItemPosition = $MarginSize
$ScriptItemPosition = $RunItemPosition + $Separator + $ItemHeight
$ComputerListItemPosition = $ScriptItemPosition + $Separator + $ItemHeight
$ListBoxPosition = $ComputerListItemPosition + $Separator + $ItemHeight
$ProgressBarPosition = $ComputerListBox + $Separator + $ListBoxHeight
$ButtonPosition = $ProgressBarPosition + $MarginSize + $ProgressBarHeight

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm = New-Object system.Windows.Forms.Form
$MainForm.ClientSize = New-Object System.Drawing.Point($MainFormWidth, $MainFormHeight)
$MainForm.Text = $ApplicationTitle
$MainForm.TopMost = $true
$MainForm.FormBorderStyle = 'Fixed3D'
$MainForm.MaximizeBox = $false
$MainForm.ShowIcon = $false
$MainForm.BackColor = $FormBackgroundColor
$MainForm.ForeColor = $FormForegroundColor

$RunLabel = New-Object system.Windows.Forms.Label
$RunLabel.Text = "  Run"
$RunLabel.Width = $LabelWidth
$RunLabel.Height = $ItemHeight
$RunLabel.Location = New-Object System.Drawing.Point($MarginSize, $RunItemPosition)
$RunLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$RunLabel.TextAlign = "MiddleLeft"
$RunLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($RunLabel)

$RunComboBox = New-Object system.Windows.Forms.ComboBox
$RunComboBox.Width = $BoxWidth
$RunComboBox.Height = $ItemHeight
$RunComboBox.Location = New-Object System.Drawing.Point($BoxPosition, $RunItemPosition)
$RunComboBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$RunComboBox.AutoCompleteMode = 'SuggestAppend'
$RunComboBox.AutoCompleteSource = 'ListItems'
$MainForm.Controls.Add($RunComboBox)

$ScriptLabel = New-Object system.Windows.Forms.Label
$ScriptLabel.Text = "  Script"
$ScriptLabel.Width = $LabelWidth
$ScriptLabel.Height = $ItemHeight
$ScriptLabel.Location = New-Object System.Drawing.Point($MarginSize, $ScriptItemPosition)
$ScriptLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ScriptLabel.TextAlign = "MiddleLeft"
$ScriptLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($ScriptLabel)

$ScriptTextBox = New-Object system.Windows.Forms.TextBox
$ScriptTextBox.Multiline = $false
$ScriptTextBox.Width = $BoxWidth
$ScriptTextBox.Height = $ItemHeight
$ScriptTextBox.Location = New-Object System.Drawing.Point($BoxPosition, $ScriptItemPosition)
$ScriptTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$MainForm.Controls.Add($ScriptTextBox)

$ComputerListLabel = New-Object system.Windows.Forms.Label
$ComputerListLabel.Text = "  Computer List"
$ComputerListLabel.Width = $LabelWidth
$ComputerListLabel.Height = $ItemHeight
$ComputerListLabel.Location = New-Object System.Drawing.Point($MarginSize, $ComputerListItemPosition)
$ComputerListLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$ComputerListLabel.TextAlign = "MiddleLeft"
$ComputerListLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($ComputerListLabel)

$ComputerListTextBox = New-Object system.Windows.Forms.TextBox
$ComputerListTextBox.Multiline = $false
$ComputerListTextBox.Width = $BoxWidth
$ComputerListTextBox.Height = $ItemHeight
$ComputerListTextBox.Location = New-Object System.Drawing.Point($BoxPosition, $ComputerListItemPosition)
$ComputerListTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$MainForm.Controls.Add($ComputerListTextBox)

$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Width = $ItemWidth
$ListBox.Height = $ListBoxHeight
$ListBox.Location = New-Object System.Drawing.Point($MarginSize, $ListBoxPosition)
$MainForm.Controls.Add($ListBox)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point($MarginSize, $ProgressBarPosition)
$ProgressBar.Size = New-Object System.Drawing.Size($ItemWidth, $ProgressBarHeight)
$ProgressBar.Style = "Continuous"
$ProgressBar.MarqueeAnimationSpeed
$MainForm.Controls.Add($ProgressBar)

$InstallButton = New-Object system.Windows.Forms.Button
$InstallButton.Text = "Run"
$InstallButton.Width = $ButtonWidth
$InstallButton.Height = $ButtonHeight
$InstallButton.Location = New-Object System.Drawing.Point((($MainFormWidth - $ButtonWidth) / 2), $ButtonPosition)
$InstallButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$InstallButton.BackColor = $ItemBackgroundColor
$MainForm.controls.Add($InstallButton)

[void]$MainForm.ShowDialog()