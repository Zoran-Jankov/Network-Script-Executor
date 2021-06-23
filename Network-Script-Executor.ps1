#---------------------------------------------------------[Initialisations]--------------------------------------------------------
$global:Computers = $null
$global:Script = $null
$global:ProgressBarIncrement = 0

$ApplicationTitle = "Network Script Executor"
$ItemBackgroundColor = "#004589"
$FormBackgroundColor = "#0080FF"
$FormForegroundColor = "#D8EBFF"

$MarginSize = 25
$Separator = 10
$LabelWidth = 100
$BoxWidth = 500
$ItemWidth = $LabelWidth + $BoxWidth
$ItemHeight = 25
$ListBoxHeight = 400
$ProgressBarHeight = 25
$ButtonHeight = 35

$MainFormWidth = $ItemWidth + ($MarginSize * 2)
$MainFormHeight = ($ItemHeight * 4) + $ListBoxHeight + ($Separator * 5) + $ProgressBarHeight + $ButtonHeight + ($MarginSize * 3)
$BoxPosition = $MarginSize + $LabelWidth
$UsingItemPosition = $MarginSize
$RunItemPosition = $UsingItemPosition + $Separator + $ItemHeight
$ScriptItemPosition = $RunItemPosition + $Separator + $ItemHeight
$ComputerListItemPosition = $ScriptItemPosition + $Separator + $ItemHeight
$ListBoxPosition = $ComputerListItemPosition + $Separator + $ItemHeight
$ProgressBarPosition = $ListBoxPosition + $Separator + $ListBoxHeight
$ButtonPosition = $ProgressBarPosition + $MarginSize + $ProgressBarHeight
$ButtonWidth = ($ItemWidth - ($MarginSize * 2)) / 3

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm = New-Object system.Windows.Forms.Form
$MainForm.ClientSize = New-Object System.Drawing.Point($MainFormWidth, $MainFormHeight)
$MainForm.Text = $ApplicationTitle
$MainForm.FormBorderStyle = 'Fixed3D'
$MainForm.MaximizeBox = $false
$MainForm.ShowIcon = $false
$MainForm.BackColor = $FormBackgroundColor
$MainForm.ForeColor = $FormForegroundColor
$MainForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

$UsingLabel = New-Object system.Windows.Forms.Label
$UsingLabel.Text = "  Using"
$UsingLabel.Width = $LabelWidth
$UsingLabel.Height = $ItemHeight
$UsingLabel.Location = New-Object System.Drawing.Point($MarginSize, $UsingItemPosition)
$UsingLabel.TextAlign = "MiddleLeft"
$UsingLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($UsingLabel)

$UsingComboBox = New-Object system.Windows.Forms.ComboBox
$UsingComboBox.Width = $BoxWidth
$UsingComboBox.Height = $ItemHeight
$UsingComboBox.Location = New-Object System.Drawing.Point($BoxPosition, $UsingItemPosition)
$UsingComboBox.AutoCompleteMode = 'SuggestAppend'
$UsingComboBox.AutoCompleteSource = 'ListItems'
$MainForm.Controls.Add($UsingComboBox)

$UsingComboBox.Items.AddRange(@("PsExec", "Windows PowerShell"))

$UsingComboBox.Add_TextChanged({
    switch ($UsingComboBox.Text) {
        "Windows PowerShell" {
            $RunComboBox.Enabled = $false
            $RunComboBox.Text = "Windows PowerShell"
        }
        Default {
            $RunComboBox.Enabled = $true
            $RunComboBox.Text = ""
        }
    }
})

$RunLabel = New-Object system.Windows.Forms.Label
$RunLabel.Text = "  Run"
$RunLabel.Width = $LabelWidth
$RunLabel.Height = $ItemHeight
$RunLabel.Location = New-Object System.Drawing.Point($MarginSize, $RunItemPosition)
$RunLabel.TextAlign = "MiddleLeft"
$RunLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($RunLabel)

$RunComboBox = New-Object system.Windows.Forms.ComboBox
$RunComboBox.Width = $BoxWidth
$RunComboBox.Height = $ItemHeight
$RunComboBox.Location = New-Object System.Drawing.Point($BoxPosition, $RunItemPosition)
$RunComboBox.AutoCompleteMode = 'SuggestAppend'
$RunComboBox.AutoCompleteSource = 'ListItems'
$MainForm.Controls.Add($RunComboBox)

$RunComboBox.Items.AddRange(@("Windows PowerShell Script", "Windows Batch Script", "gpupdate"))

$RunComboBox.Add_TextChanged({
    switch ($RunComboBox.Text) {
        "gpupdate" {
            $ScriptTextBox.Enabled = $false
            $ScriptTextBox.Text = ""
        }
        Default {
            $ScriptTextBox.Enabled = $true
        }
    }
})

$ScriptLabel = New-Object system.Windows.Forms.Label
$ScriptLabel.Text = "  Script"
$ScriptLabel.Width = $LabelWidth
$ScriptLabel.Height = $ItemHeight
$ScriptLabel.Location = New-Object System.Drawing.Point($MarginSize, $ScriptItemPosition)
$ScriptLabel.TextAlign = "MiddleLeft"
$ScriptLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($ScriptLabel)

$ScriptTextBox = New-Object system.Windows.Forms.TextBox
$ScriptTextBox.Multiline = $false
$ScriptTextBox.ReadOnly = $true
$ScriptTextBox.Width = $BoxWidth
$ScriptTextBox.Height = $ItemHeight
$ScriptTextBox.Location = New-Object System.Drawing.Point($BoxPosition, $ScriptItemPosition)
$MainForm.Controls.Add($ScriptTextBox)

$ScriptTextBox.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $OpenFileDialog.Filter = "PS1 Files (*.ps1) | *.ps1"
    $OpenFileDialog.ShowDialog() | Out-Null
    $ScriptTextBox.Text =  $OpenFileDialog.Filename
    if (($ScriptTextBox.Text -like "*ps1") -or ($ScriptTextBox.Text -like "*bat")) {
        $OpenScriptButton.Enabled = $true
    }
    else {
        $OpenScriptButton.Enabled = $false
    }
    $global:Script = Get-Content -Path ($ScriptTextBox.Text)
})

$ComputerListLabel = New-Object system.Windows.Forms.Label
$ComputerListLabel.Text = "  Computer List"
$ComputerListLabel.Width = $LabelWidth
$ComputerListLabel.Height = $ItemHeight
$ComputerListLabel.Location = New-Object System.Drawing.Point($MarginSize, $ComputerListItemPosition)
$ComputerListLabel.TextAlign = "MiddleLeft"
$ComputerListLabel.BackColor = $ItemBackgroundColor
$MainForm.Controls.Add($ComputerListLabel)

$ComputerListTextBox = New-Object system.Windows.Forms.TextBox
$ComputerListTextBox.Multiline = $false
$ComputerListTextBox.ReadOnly = $true
$ComputerListTextBox.Width = $BoxWidth
$ComputerListTextBox.Height = $ItemHeight
$ComputerListTextBox.Location = New-Object System.Drawing.Point($BoxPosition, $ComputerListItemPosition)
$MainForm.Controls.Add($ComputerListTextBox)

$ComputerListTextBox.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $OpenFileDialog.Filter = "Text Files (*.txt) | *.txt"
    $OpenFileDialog.ShowDialog() | Out-Null
    $ComputerListTextBox.Text = $OpenFileDialog.Filename
    $global:Computers = Get-Content -Path ($ComputerListTextBox.Text) -Encoding UTF8
    if ($global:Computers -ne $null) {
        $RunButton.Enabled = $true
    }
    else {
        $RunButtonButton.Enabled = $false
    }
    $global:ProgressBarIncrement = 10000 / $Computers.Count
})

$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Width = $ItemWidth
$ListBox.Height = $ListBoxHeight
$ListBox.Location = New-Object System.Drawing.Point($MarginSize, $ListBoxPosition)
$ListBox.Text = "jkj"
$ListBox.ForeColor = "#000000"
$MainForm.Controls.Add($ListBox)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point($MarginSize, $ProgressBarPosition)
$ProgressBar.Size = New-Object System.Drawing.Size($ItemWidth, $ProgressBarHeight)
$ProgressBar.Style = "Continuous"
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = 10000
$MainForm.Controls.Add($ProgressBar)

$OpenScriptButton = New-Object System.Windows.Forms.Button
$OpenScriptButton.Enabled = $false
$OpenScriptButton.Text = "Open Script"
$OpenScriptButton.Width = $ButtonWidth
$OpenScriptButton.Height = $ButtonHeight
$OpenScriptButton.Location = New-Object System.Drawing.Point(($MarginSize * 2 + $ButtonWidth), $ButtonPosition)
$OpenScriptButton.BackColor = $ItemBackgroundColor
$MainForm.controls.Add($OpenScriptButton)

$OpenScriptButton.Add_Click({
    if ((Test-Path -Path "C:\Program Files\Notepad++") -or (Test-Path -Path "C:\Program Files (x86)\Notepad++")) {
        Start-Process notepad++ $ScriptTextBox.Text
    }
    else {
        Start-Process notepad $ScriptTextBox.Text
    }
})

$RunButton = New-Object System.Windows.Forms.Button
$RunButton.Enabled = $true
$RunButton.Text = "Run"
$RunButton.Width = $ButtonWidth
$RunButton.Height = $ButtonHeight
$RunButton.Location = New-Object System.Drawing.Point(($MarginSize * 3  + ($ButtonWidth * 2)), $ButtonPosition)
$RunButton.BackColor = $ItemBackgroundColor
$MainForm.controls.Add($RunButton)

$RunButton.Add_Click({
    foreach ($Computer in $global:Computers) {
        if (Test-Connection $Computer -Quiet -Count 1) {  
            $ListBox.Text += "Executing script on $Computer computer . . .`r`n"
            switch ($UsingComboBox.Text) {
                "Windows PowerShell" {
                    Invoke-Command -ComputerName $Computer -ScriptBlock $global:Script
                }
                Default {
                    $TargetComputer = "\\$Computer"
                    Start-Process -FilePath PsExec.exe -ArgumentList $TargetComputer $global:Script -h
                }
            }
            $ListBox.Text += "Script executed on $Computer computer . . .`r`n"
        }
        else {
            $ListBox.Text += "ERROR: $Computer computer is offline`r`n"
        }
        $ProgressBar.Value += $global:ProgressBarIncrement
    }
})

[void]$MainForm.ShowDialog()