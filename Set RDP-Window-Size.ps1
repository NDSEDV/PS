<#
.SYNOPSIS
  Set RDP-Window-Size
.DESCRIPTION
  Das Tool soll dir bei deiner täglichen Arbeit helfen.
.PARAMETER Sprache
    Das Tool hat eine deutsche Edition, kann aber auch auf englischen Betriebssystemen verwendet werden.
.NOTES
  Version:        1.0
  Autor:          Jörn Walter
  Erstellungsdatum:  2025-02-11

  Copyright (c) Jörn Walter. Alle Rechte vorbehalten.
#>

# Laden von Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Erstellen des Hauptformulars
$form = New-Object System.Windows.Forms.Form
$form.Text = "Set RDP-Window-Size"
$form.Size = New-Object System.Drawing.Size(600, 470)
$form.StartPosition = "CenterScreen"

# Erstellen der Steuerelemente
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "RDP-Datei(en) auswählen:"
$label1.Location = New-Object System.Drawing.Point(10, 20)
$label1.AutoSize = $true
$form.Controls.Add($label1)

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Location = New-Object System.Drawing.Point(10, 50)
$textBoxFile.Size = New-Object System.Drawing.Size(450, 20)
$form.Controls.Add($textBoxFile)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Durchsuchen"
$buttonBrowse.Location = New-Object System.Drawing.Point(470, 48)
$form.Controls.Add($buttonBrowse)

$label2 = New-Object System.Windows.Forms.Label
$label2.Text = "Ersetzungen (eine pro Zeile, Format: AlterText=NeuerText, Wildcards mit *):"
$label2.Location = New-Object System.Drawing.Point(10, 80)
$label2.AutoSize = $true
$form.Controls.Add($label2)

$textBoxReplacements = New-Object System.Windows.Forms.TextBox
$textBoxReplacements.Location = New-Object System.Drawing.Point(10, 110)
$textBoxReplacements.Size = New-Object System.Drawing.Size(560, 100)
$textBoxReplacements.Multiline = $true
$textBoxReplacements.Text = "desktopw*=desktopwidth:i:1920`r`ndesktoph*=desktopheight:i:1080"
$form.Controls.Add($textBoxReplacements)

$buttonReplace = New-Object System.Windows.Forms.Button
$buttonReplace.Text = "Ersetzen"
$buttonReplace.Location = New-Object System.Drawing.Point(250, 220)
$form.Controls.Add($buttonReplace)

$label3 = New-Object System.Windows.Forms.Label
$label3.Text = "Verarbeitungsausgabe:"
$label3.Location = New-Object System.Drawing.Point(10, 250)
$label3.AutoSize = $true
$form.Controls.Add($label3)

$textBoxOutput = New-Object System.Windows.Forms.TextBox
$textBoxOutput.Location = New-Object System.Drawing.Point(10, 280)
$textBoxOutput.Size = New-Object System.Drawing.Size(560, 100)
$textBoxOutput.Multiline = $true
$textBoxOutput.ReadOnly = $true
$textBoxOutput.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$textBoxOutput.WordWrap = $false
$form.Controls.Add($textBoxOutput)

# Copyright-Hinweis
$labelCopyright = New-Object System.Windows.Forms.Label
$labelCopyright.Text = "© 2025 Jörn Walter - https://www.der-windows-papst.de"
$labelCopyright.Location = New-Object System.Drawing.Point(10, 400)
$labelCopyright.AutoSize = $true
$form.Controls.Add($labelCopyright)

# Ereignishandler fÃ¼r die SchaltflÃ¤chen
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "RDP-Dateien (*.rdp)|*.rdp"
    $openFileDialog.Multiselect = $true
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxFile.Text = $openFileDialog.FileNames -join ";"
    }
})

$buttonReplace.Add_Click({
    $textBoxOutput.Text = ""  # Verarbeitungsausgabe leeren
    $files = $textBoxFile.Text -split ";"
    $replacements = $textBoxReplacements.Lines
    $output = ""

    foreach ($file in $files) {
        if (Test-Path $file) {
            $content = Get-Content -Path $file -Raw
            foreach ($replacement in $replacements) {
                if ($replacement -match "^(.+)=(.+)$") {
                    $findText = $matches[1]
                    $replaceText = $matches[2]
                    $regexPattern = [regex]::Escape($findText).Replace("\*", ".*")
                    $content = [regex]::Replace($content, $regexPattern, $replaceText)
                    $output += "Ersetzt '$findText' mit '$replaceText' in Datei '$file'`r`n"
                }
            }
            Set-Content -Path $file -Value $content
        }
    }

    $textBoxOutput.Text = $output
    [System.Windows.Forms.MessageBox]::Show("Ersetzung abgeschlossen!"+ [Environment]::NewLine)
})

# Formular anzeigen
$form.Add_Shown({$form.Activate()})
[void] $form.ShowDialog()
