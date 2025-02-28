<#
.SYNOPSIS
  JW COPY Master Tool
.DESCRIPTION
  Das Tool exportiert AD-Benutzerinformationen mit einer benutzerfreundlichen Oberfläche.
.PARAMETER language
    Das Tool hat eine deutsche Edition, kann aber auch auf englischen BS-Systemen verwendet werden.
.NOTES
  Version:        1.1
  Author:         Jörn Walter (GUI-Erweiterung)
  Creation Date:  2025-02-24

  Copyright (c) Jörn Walter. All rights reserved.
  Web: https://www.der-windows-papst.de
#>

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Setze Codepage auf 1252
chcp 1252

# Stellt sicher, dass die PowerShell-Konsole UTF-8 unterstützt
$OutputEncoding = [System.Text.Encoding]::UTF8

# Lade Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Log Funktion
function Get-LogFileName {
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $logDir = Join-Path $desktopPath "RCOPYLOGS"
        
        # Erstellt den RCOPYLOGS-Ordner, falls er nicht existiert
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            Write-Log "Log-Verzeichnis erstellt: $logDir"
        }
        
        return Join-Path $logDir "robocopy_$timestamp.log"
    }
    catch {
        throw
    }
}

# Pfade aktualisieren
function Update-PathComboBoxes {
    $paths = Load-Paths

    $cmbSource.Items.Clear()
    $cmbTarget.Items.Clear()

    foreach ($path in $paths) {
        $cmbSource.Items.Add($path.SourcePath)
        $cmbTarget.Items.Add($path.TargetPath)
    }
}

function Get-ConfigPath {
    try {
        $configDir = Join-Path $env:APPDATA "RobocopyGUI"
        if (-not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }
        return Join-Path $configDir "paths.json"
    }
    catch {
        throw
    }
}

function Save-Paths {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )

    try {
        $configPath = Get-ConfigPath
        $existingPaths = @()
        if (Test-Path $configPath) {
            $existingPaths = Get-Content -Path $configPath -Encoding UTF8 | ConvertFrom-Json
        }

        $newPath = @{
            "SourcePath" = $SourcePath
            "TargetPath" = $TargetPath
            "LastSaved" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

        $existingPaths = @($existingPaths | Where-Object {
            $_.SourcePath -ne $SourcePath -or $_.TargetPath -ne $TargetPath
        })

        $existingPaths = @($newPath) + $existingPaths

        if ($existingPaths.Count -gt 10) {
            $existingPaths = $existingPaths[0..9]
        }

        $json = $existingPaths | ConvertTo-Json
        Set-Content -Path $configPath -Value $json -Encoding UTF8

        Update-PathComboBoxes
    }
    catch {
        throw
    }
}

function Load-Paths {
    try {
        $configPath = Get-ConfigPath
        if (Test-Path $configPath) {
            $paths = Get-Content -Path $configPath -Encoding UTF8 | ConvertFrom-Json
            return $paths
        }
        return @()
    }
    catch {
        return @()
    }
}

function Write-Log {
    param($Message)

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] $Message"

        if ($txtOutput -ne $null) {
            $txtOutput.Invoke([Action]{
                $txtOutput.AppendText("$logMessage`r`n")
                $txtOutput.ScrollToCaret()
            })
        }

        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName

            # Setzt die Codepage auf 1252
            chcp 1252

            Add-Content -Path $logFile -Value $logMessage -Encoding Default
        }
    }
    catch {
        # Fallback für Encoding-Probleme
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $fallbackMessage = "[$timestamp] $($Message -replace '[^\x20-\x7E]', '?')"

        if ($txtOutput -ne $null) {
            $txtOutput.Invoke([Action]{
                $txtOutput.AppendText("$fallbackMessage`r`n")
                $txtOutput.ScrollToCaret()
            })
        }

        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName

            # Setze die Codepage auf 1252
            chcp 1252

            Add-Content -Path $logFile -Value $fallbackMessage -Encoding Default
        }
    }
}

function Parse-RobocopyOutput {
    param (
        [string[]]$Output
    )

    $stats = @{
        Files = 0
        Directories = 0
    }

    Write-Log "Parsing Robocopy Output..."

    foreach ($line in $Output) {
        Write-Log "Analyzing line: $line"

        if ($line -match '^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*$') {
        chcp 1252
            Write-Log "Found statistics line: $line"
            $stats.Files = [int]$Matches[1]
            Write-Log "Extracted total files: $($stats.Files)"
            break
        }

        # Alternative Erkennung für andere Formate
        if ($line -match 'Dateien\s*:\s*(\d+)' -or
            $line -match 'Files\s*:\s*(\d+)' -or
            $line -match 'Kopiert\s*:\s*(\d+)' -or
            $line -match 'Copied\s*:\s*(\d+)') {
            $stats.Files = [int]$Matches[1]
            chcp 1252
            Write-Log "Found files count: $($stats.Files)"
        }
    }

    Write-Log "Parsing complete. Total files: $($stats.Files)"
    return $stats
}

# Hauptformular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Robocopy GUI Tool'
$form.Size = New-Object System.Drawing.Size(1000,730)
$form.StartPosition = 'CenterScreen'

# Ausgabefenster für Logging
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20,480)
$txtOutput.Size = New-Object System.Drawing.Size(940,150)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.ReadOnly = $true
$txtOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOutput)

function Write-Log {
    param($Message)
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] $Message"
        
        if ($txtOutput -ne $null) {
            $txtOutput.AppendText("$logMessage`r`n")
            $txtOutput.Select($txtOutput.Text.Length, 0)
            $txtOutput.ScrollToCaret()
        }
        
        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName
            Add-Content -Path $logFile -Value $logMessage -Encoding Default
        }
    }
    catch {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $fallbackMessage = "[$timestamp] $($Message -replace '[^\x20-\x7E]', '?')"
        
        if ($txtOutput -ne $null) {
            $txtOutput.AppendText("$fallbackMessage`r`n")
            $txtOutput.Select($txtOutput.Text.Length, 0)
            $txtOutput.ScrollToCaret()
        }
    }
}

[System.Windows.Forms.Control]::DefaultMarginChanged
$txtOutput.Refresh()

# GUI-Basiselemente
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Location = New-Object System.Drawing.Point(20,20)
$lblSource.Size = New-Object System.Drawing.Size(100,20)
$lblSource.Text = 'Quellordner:'
$form.Controls.Add($lblSource)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(20,80)
$lblTarget.Size = New-Object System.Drawing.Size(100,20)
$lblTarget.Text = 'Zielordner:'
$form.Controls.Add($lblTarget)

# ComboBoxen für Pfade
$cmbSource = New-Object System.Windows.Forms.ComboBox
$cmbSource.Location = New-Object System.Drawing.Point(20,40)
$cmbSource.Size = New-Object System.Drawing.Size(800,20)
$cmbSource.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$form.Controls.Add($cmbSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(20,60)
$txtSource.Size = New-Object System.Drawing.Size(800,20)
$form.Controls.Add($txtSource)

$cmbTarget = New-Object System.Windows.Forms.ComboBox
$cmbTarget.Location = New-Object System.Drawing.Point(20,100)
$cmbTarget.Size = New-Object System.Drawing.Size(800,20)
$cmbTarget.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$form.Controls.Add($cmbTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(20,120)
$txtTarget.Size = New-Object System.Drawing.Size(800,20)
$form.Controls.Add($txtTarget)

# Browse Buttons
$btnSourceBrowse = New-Object System.Windows.Forms.Button
$btnSourceBrowse.Location = New-Object System.Drawing.Point(830,40)
$btnSourceBrowse.Size = New-Object System.Drawing.Size(100,20)
$btnSourceBrowse.Text = 'Durchsuchen'
$form.Controls.Add($btnSourceBrowse)

$btnTargetBrowse = New-Object System.Windows.Forms.Button
$btnTargetBrowse.Location = New-Object System.Drawing.Point(830,100)
$btnTargetBrowse.Size = New-Object System.Drawing.Size(100,20)
$btnTargetBrowse.Text = 'Durchsuchen'
$form.Controls.Add($btnTargetBrowse)

# Synchronisierungs-Bereich
$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Location = New-Object System.Drawing.Point(20,200)
$lblDays.Size = New-Object System.Drawing.Size(130,20)
$lblDays.Text = 'Die letzten x Tage:'
$form.Controls.Add($lblDays)

# Numerische Eingabe für Tage
$numDays = New-Object System.Windows.Forms.NumericUpDown
$numDays.Location = New-Object System.Drawing.Point(150,200)
$numDays.Size = New-Object System.Drawing.Size(60,20)
$numDays.Minimum = 1
$numDays.Maximum = 365
$numDays.Value = 7
$form.Controls.Add($numDays)

# Hauptfunktions-Buttons
$btnCreateStructure = New-Object System.Windows.Forms.Button
$btnCreateStructure.Location = New-Object System.Drawing.Point(20,160)
$btnCreateStructure.Size = New-Object System.Drawing.Size(190,30)
$btnCreateStructure.Text = 'Ordnerstruktur erstellen'
$form.Controls.Add($btnCreateStructure)

$btnCopyData = New-Object System.Windows.Forms.Button
$btnCopyData.Location = New-Object System.Drawing.Point(230,160)
$btnCopyData.Size = New-Object System.Drawing.Size(190,30)
$btnCopyData.Text = 'Daten kopieren'
$form.Controls.Add($btnCopyData)

$btnSync = New-Object System.Windows.Forms.Button
$btnSync.Location = New-Object System.Drawing.Point(230,200)
$btnSync.Size = New-Object System.Drawing.Size(190,30)
$btnSync.Text = 'Synchronisieren'
$form.Controls.Add($btnSync)

# Lösch-Button
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(20,240)
$btnDelete.Size = New-Object System.Drawing.Size(190,30)
$btnDelete.Text = 'Ziel bereinigen'
$form.Controls.Add($btnDelete)

# Vergleichs-Button
$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.Location = New-Object System.Drawing.Point(230,240)
$btnCompare.Size = New-Object System.Drawing.Size(190,30)
$btnCompare.Text = 'Ordner vergleichen'
$form.Controls.Add($btnCompare)

# Datm Lösch Controls
$lblDeleteDate = New-Object System.Windows.Forms.Label
$lblDeleteDate.Location = New-Object System.Drawing.Point(440,160)
$lblDeleteDate.Size = New-Object System.Drawing.Size(150,20)
$lblDeleteDate.Text = 'Datum für Löschung im Ziel:'
$form.Controls.Add($lblDeleteDate)

$dtpDeleteDate = New-Object System.Windows.Forms.DateTimePicker
$dtpDeleteDate.Location = New-Object System.Drawing.Point(440,180)
$dtpDeleteDate.Size = New-Object System.Drawing.Size(150,20)
$dtpDeleteDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$form.Controls.Add($dtpDeleteDate)

# Radio Buttons Auwahl
$gbDateType = New-Object System.Windows.Forms.GroupBox
$gbDateType.Location = New-Object System.Drawing.Point(600,160)
$gbDateType.Size = New-Object System.Drawing.Size(200,50)
$gbDateType.Text = "Datum-Typ"

$rbCreationDate = New-Object System.Windows.Forms.RadioButton
$rbCreationDate.Location = New-Object System.Drawing.Point(10,20)
$rbCreationDate.Size = New-Object System.Drawing.Size(90,20)
$rbCreationDate.Text = "Erstellung"
$rbCreationDate.Checked = $true
$gbDateType.Controls.Add($rbCreationDate)

$rbModifiedDate = New-Object System.Windows.Forms.RadioButton
$rbModifiedDate.Location = New-Object System.Drawing.Point(100,20)
$rbModifiedDate.Size = New-Object System.Drawing.Size(90,20)
$rbModifiedDate.Text = "Änderung"
$gbDateType.Controls.Add($rbModifiedDate)

$form.Controls.Add($gbDateType)

# Lösch Button nach Datum
$btnDeleteByDate = New-Object System.Windows.Forms.Button
$btnDeleteByDate.Location = New-Object System.Drawing.Point(440,215)
$btnDeleteByDate.Size = New-Object System.Drawing.Size(190,30)
$btnDeleteByDate.Text = 'Nach Datum löschen'
$form.Controls.Add($btnDeleteByDate)

# Löschen nach Pattern Controls
$lblNamePattern = New-Object System.Windows.Forms.Label
$lblNamePattern.Location = New-Object System.Drawing.Point(440,250)
$lblNamePattern.Size = New-Object System.Drawing.Size(150,20)
$lblNamePattern.Text = 'Pattern:'
$form.Controls.Add($lblNamePattern)

$txtNamePattern = New-Object System.Windows.Forms.TextBox
$txtNamePattern.Location = New-Object System.Drawing.Point(440,270)
$txtNamePattern.Size = New-Object System.Drawing.Size(180,20)
$form.Controls.Add($txtNamePattern)

# Löschen nach Pattern Button
$btnDeleteByPattern = New-Object System.Windows.Forms.Button
$btnDeleteByPattern.Location = New-Object System.Drawing.Point(440,300)
$btnDeleteByPattern.Size = New-Object System.Drawing.Size(190,30)
$btnDeleteByPattern.Text = 'Nach Pattern im Ziel löschen'
$form.Controls.Add($btnDeleteByPattern)

# Log-Optionen
$chkCreateLog = New-Object System.Windows.Forms.CheckBox
$chkCreateLog.Location = New-Object System.Drawing.Point(770,410)
$chkCreateLog.Size = New-Object System.Drawing.Size(300,20)
$chkCreateLog.Text = "Robocopy-Logdatei erstellen"
$chkCreateLog.Checked = $true
$form.Controls.Add($chkCreateLog)

# Copyright Label
$lblCopyright = New-Object System.Windows.Forms.Label
$lblCopyright.Location = New-Object System.Drawing.Point(690,640)
$lblCopyright.Size = New-Object System.Drawing.Size(400,20)
$lblCopyright.Text = '© 2025 Jörn Walter https://www.der-windows-papst.de'
$lblCopyright.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblCopyright)

# Statusleiste
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Event Handler für ComboBoxen
$cmbSource.Add_SelectedIndexChanged({
    if ($cmbSource.SelectedIndex -ge 0) {
        $paths = Load-Paths
        $selectedPath = $paths[$cmbSource.SelectedIndex]
        $txtSource.Text = $selectedPath.SourcePath
        $txtTarget.Text = $selectedPath.TargetPath
        $cmbTarget.SelectedIndex = $cmbSource.SelectedIndex
    }
})

$cmbTarget.Add_SelectedIndexChanged({
    if ($cmbTarget.SelectedIndex -ge 0) {
        $paths = Load-Paths
        $selectedPath = $paths[$cmbTarget.SelectedIndex]
        $txtTarget.Text = $selectedPath.TargetPath
    }
})

# Funktion zum Validieren von Pfaden (lokal und UNC)
function Test-ValidPath {
    param (
        [string]$Path
    )
    
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-Log "Leerer Pfad wurde übergeben"
            return $false
        }

        # Prüfe ob es sich um einen UNC-Pfad handelt
        if ($Path.StartsWith("\\")) {
            if ($Path -match '^\\\\[^\\]+\\[^\\]+(?:\\.*)?$') {
                if (Test-Path -Path $Path -ErrorAction SilentlyContinue) {
                    return $true
                }
                Write-Log "UNC-Pfad nicht erreichbar: $Path"
                return $false
            }
            Write-Log "Ungültiges UNC-Pfad Format: $Path"
            return $false
        }
        
        if (Test-Path -Path $Path -ErrorAction SilentlyContinue) {
            return $true
        }
        
        Write-Log "Pfad nicht gefunden: $Path"
        return $false
    }
    catch {
        Write-Log "Fehler bei der Pfadvalidierung: $_"
        return $false
    }
}

# Funktion zum Normalisieren von Pfaden
function Get-NormalizedPath {
    param (
        [string]$Path
    )
    
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    if ($Path.Length -gt 3 -and $Path.EndsWith('\')) {
        $Path = $Path.TrimEnd('\')
    }
    
    if ($Path.StartsWith('\\')) {
        $Path = '\\' + ($Path.Substring(2) -replace '\\+', '\')
    } else {
        $Path = $Path -replace '\\+', '\'
    }
    
    return $Path
}

# Aktualisierte Browse-Button Event Handler
$btnSourceBrowse.Add_Click({
    $result = Show-FolderDialog -Title "Quellordner auswählen" -InitialDirectory $txtSource.Text
    if ($result.Success) {
        $txtSource.Text = $result.SelectedPath
        $cmbSource.Text = $result.SelectedPath
    }
})

$btnTargetBrowse.Add_Click({
    $result = Show-FolderDialog -Title "Zielordner auswählen" -InitialDirectory $txtTarget.Text
    if ($result.Success) {
        $txtTarget.Text = $result.SelectedPath
        $cmbTarget.Text = $result.SelectedPath
    }
})

# Ordnerauswahl-Funktion
function Show-FolderDialog {
    param (
        [string]$Title,
        [string]$InitialDirectory
    )

    $customPathForm = New-Object System.Windows.Forms.Form
    $customPathForm.Text = $Title
    $customPathForm.Size = New-Object System.Drawing.Size(600,200)
    $customPathForm.StartPosition = 'CenterScreen'
    $customPathForm.FormBorderStyle = 'FixedDialog'
    $customPathForm.MaximizeBox = $false
    $customPathForm.MinimizeBox = $false

    # Beschreibung für UNC-Pfade
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(10,10)
    $lblInfo.Size = New-Object System.Drawing.Size(560,40)
    $lblInfo.Text = "Bitte gib einen Pfad ein oder wählen einen lokalen Ordner aus.`nUNC-Pfade sind erlaubt (z.B. \\PO\DFS\Ordner)"
    $customPathForm.Controls.Add($lblInfo)

    # Textfeld für Pfad
    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(10,60)
    $txtPath.Size = New-Object System.Drawing.Size(560,20)
    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
        $txtPath.Text = $InitialDirectory
    }
    $customPathForm.Controls.Add($txtPath)

    # Button für lokales Durchsuchen
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Location = New-Object System.Drawing.Point(10,90)
    $btnBrowse.Size = New-Object System.Drawing.Size(180,23)
    $btnBrowse.Text = "Lokalen Ordner durchsuchen..."
    $btnBrowse.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = $Title
        
        # Prüfe ob der aktuelle Pfad gültig ist
        if (-not [string]::IsNullOrWhiteSpace($txtPath.Text) -and (Test-Path -Path $txtPath.Text -ErrorAction SilentlyContinue)) {
            $folderBrowser.SelectedPath = $txtPath.Text
        }
        
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $txtPath.Text = $folderBrowser.SelectedPath
        }
    })
    $customPathForm.Controls.Add($btnBrowse)

    # OK Button
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(400,120)
    $btnOK.Size = New-Object System.Drawing.Size(75,23)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $customPathForm.Controls.Add($btnOK)

    # Abbrechen Button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(490,120)
    $btnCancel.Size = New-Object System.Drawing.Size(75,23)
    $btnCancel.Text = "Abbrechen"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $customPathForm.Controls.Add($btnCancel)

    $customPathForm.AcceptButton = $btnOK
    $customPathForm.CancelButton = $btnCancel

    $result = @{
        Success = $false
        SelectedPath = ""
    }

    $dialogResult = $customPathForm.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not [string]::IsNullOrWhiteSpace($txtPath.Text)) {
            $selectedPath = Get-NormalizedPath -Path $txtPath.Text
            if (Test-ValidPath -Path $selectedPath) {
                $result.Success = $true
                $result.SelectedPath = $selectedPath
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Der eingegebene Pfad ist ungültig oder nicht erreichbar.`nBitte überprüfe den Pfad und versuche es erneut.",
                    "Ungültiger Pfad",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Bitte gib einen Pfad ein oder wählen einen lokalen Ordner aus.",
                "Fehlender Pfad",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    }

    $customPathForm.Dispose()
    return $result
}

# Ordnerstruktur erstellen Button
$btnCreateStructure.Add_Click({
    if ([string]::IsNullOrEmpty($txtSource.Text) -or [string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle den Quell- und Zielordner aus.", "Fehler")
        return
    }

    try {
        chcp 1252
        Write-Log "Erstelle Ordnerstruktur..."
        Write-Log "Quelle: $($txtSource.Text)"
        Write-Log "Ziel: $($txtTarget.Text)"

        $roboParams = @($txtSource.Text, $txtTarget.Text, "/E", "/XF", "*", "/COPYALL")

        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName
            $roboParams += "/LOG:$logFile"
            Write-Log "Logdatei wird erstellt unter: $logFile"
        }

        $output = robocopy @roboParams

        foreach ($line in $output) {
            Write-Log $line
        }

        # Speichere die Pfade nur nach erfolgreicher Strukturerstellung
        Save-Paths -SourcePath $txtSource.Text -TargetPath $txtTarget.Text

        Write-Log "Ordnerstruktur erstellt"
        [System.Windows.Forms.MessageBox]::Show("Ordnerstruktur wurde erfolgreich erstellt!", "Erfolg")
    }
    catch {
        Write-Log "FEHLER: $_"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen der Ordnerstruktur: $_", "Fehler")
    }
})

# Progress Bar und Label erstellen
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 440)
$progressBar.Size = New-Object System.Drawing.Size(940, 20)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(20, 420)
$lblProgress.Size = New-Object System.Drawing.Size(940, 20)
$lblProgress.Text = "Bereit"
$form.Controls.Add($lblProgress)

# Hilfsfunktionen für den Fortschritt
function Update-Progress {
    param (
        [string]$Message,
        [int]$PercentComplete
    )
    
    $progressBar.Value = $PercentComplete
    $lblProgress.Text = "$Message - $PercentComplete%"
    [System.Windows.Forms.Application]::DoEvents()
}

function Reset-Progress {
    $progressBar.Value = 0
    $lblProgress.Text = "Bereit"
    [System.Windows.Forms.Application]::DoEvents()
}

# Kopieren Button
$btnCopyData.Add_Click({
    if ([string]::IsNullOrEmpty($txtSource.Text) -or [string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle den Quell- und Zielordner aus.", "Fehler")
        return
    }

    try {
        Reset-Progress
        Write-Log "Starte Kopiervorgang..."
        Write-Log "Quelle: $($txtSource.Text)"
        Write-Log "Ziel: $($txtTarget.Text)"

        $sourcePath = $txtSource.Text.TrimEnd('\')
        $targetPath = $txtTarget.Text.TrimEnd('\')
        
        # Erstelle kompletten Robocopy-Befehl als String
        $robocopyCommand = "robocopy `"$sourcePath`" `"$targetPath`" /E /COPY:DAT /R:1 /W:1 /DCOPY:DAT /MT:8"
        
        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName
            $robocopyCommand += " /LOG+:`"$logFile`""
            Write-Log "Logdatei wird erstellt unter: $logFile"
        }

        Write-Log "Ausführung: $robocopyCommand"

        # Job mit dem kompletten Befehl starten
        $job = Start-Job -ScriptBlock {
            param($cmd)
            Invoke-Expression $cmd
        } -ArgumentList $robocopyCommand

        while ($job.State -eq 'Running') {
            $progress = Get-Random -Minimum 10 -Maximum 90
            Update-Progress -Message "Kopiere Dateien" -PercentComplete $progress
            Start-Sleep -Milliseconds 500
        }

        $output = Receive-Job -Job $job -Wait
        Remove-Job -Job $job -Force

        foreach ($line in $output) {
            Write-Log $line
        }

        if ($LASTEXITCODE -lt 8) {
            Update-Progress -Message "Kopiervorgang abgeschlossen" -PercentComplete 100
            [System.Windows.Forms.MessageBox]::Show("Dateien wurden erfolgreich kopiert!", "Erfolg")
            Save-Paths -SourcePath $txtSource.Text -TargetPath $txtTarget.Text
        } else {
            throw "Robocopy beendet mit Fehlercode $LASTEXITCODE"
        }
    }
    catch {
        Write-Log "FEHLER: $_"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Kopieren der Daten: $_", "Fehler")
    }
    finally {
        Reset-Progress
    }
})

# Synchronisieren Button
$btnSync.Add_Click({
    if ([string]::IsNullOrEmpty($txtSource.Text) -or [string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle den Quell- und Zielordner aus.", "Fehler")
        return
    }

    try {
        Reset-Progress
        $daysBack = $numDays.Value
        $maxAge = (Get-Date).AddDays(-$daysBack)
        $maxAgeParam = "/MAXAGE:" + $maxAge.ToString("yyyyMMdd")

        Write-Log "Starte Synchronisierung..."
        Write-Log "Quelle: $($txtSource.Text)"
        Write-Log "Ziel: $($txtTarget.Text)"
        Write-Log "Zeitraum: Letzte $daysBack Tage"

        $sourcePath = $txtSource.Text.TrimEnd('\')
        $targetPath = $txtTarget.Text.TrimEnd('\')

        # Erstelle kompletten Robocopy-Befehl als String
        $robocopyCommand = "robocopy `"$sourcePath`" `"$targetPath`" /MIR /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /MT:8 $maxAgeParam"
        
        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName
            $robocopyCommand += " /LOG+:`"$logFile`""
            Write-Log "Logdatei wird erstellt unter: $logFile"
        }

        Write-Log "Ausführung: $robocopyCommand"

        # Job mit dem kompletten Befehl starten
        $job = Start-Job -ScriptBlock {
            param($cmd)
            Invoke-Expression $cmd
        } -ArgumentList $robocopyCommand

        while ($job.State -eq 'Running') {
            $progress = Get-Random -Minimum 10 -Maximum 90
            Update-Progress -Message "Synchronisiere Dateien" -PercentComplete $progress
            Start-Sleep -Milliseconds 500
        }

        $output = Receive-Job -Job $job -Wait
        Remove-Job -Job $job -Force

        foreach ($line in $output) {
            Write-Log $line
        }

        if ($LASTEXITCODE -lt 8) {
            Update-Progress -Message "Synchronisierung abgeschlossen" -PercentComplete 100
            [System.Windows.Forms.MessageBox]::Show("Synchronisierung erfolgreich abgeschlossen!", "Erfolg")
            Save-Paths -SourcePath $txtSource.Text -TargetPath $txtTarget.Text
        } else {
            throw "Robocopy beendet mit Fehlercode $LASTEXITCODE"
        }
    }
    catch {
        Write-Log "FEHLER: $_"
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Synchronisierung: $_", "Fehler")
    }
    finally {
        Reset-Progress
    }
})

# Vergleichs-Button
$btnCompare.Add_Click({
    if ([string]::IsNullOrEmpty($txtSource.Text) -or [string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle den Quell- und Zielordner aus.", "Fehler")
        return
    }

    try {
        Reset-Progress
        Write-Log "Starte Vergleich..."
        Write-Log "Quelle: $($txtSource.Text)"
        Write-Log "Ziel: $($txtTarget.Text)"

        $sourcePath = $txtSource.Text.TrimEnd('\')
        $targetPath = $txtTarget.Text.TrimEnd('\')

        # Erstelle kompletten Robocopy-Befehl als String
        $robocopyCommand = "robocopy `"$sourcePath`" `"$targetPath`" /E /L"
        
        if ($chkCreateLog.Checked) {
            $logFile = Get-LogFileName
            $robocopyCommand += " /LOG+:`"$logFile`""
            Write-Log "Logdatei wird erstellt unter: $logFile"
        }

        Write-Log "Ausführung: $robocopyCommand"

        # Job mit dem kompletten Befehl starten
        $job = Start-Job -ScriptBlock {
            param($cmd)
            Invoke-Expression $cmd
        } -ArgumentList $robocopyCommand

        while ($job.State -eq 'Running') {
            $progress = Get-Random -Minimum 10 -Maximum 90
            Update-Progress -Message "Vergleiche Dateien" -PercentComplete $progress
            Start-Sleep -Milliseconds 500
        }

        $output = Receive-Job -Job $job -Wait
        Remove-Job -Job $job -Force

        foreach ($line in $output) {
            Write-Log $line
        }

        Update-Progress -Message "Vergleich abgeschlossen" -PercentComplete 100
        Write-Log "Vergleich abgeschlossen"
    }
    catch {
        Write-Log "FEHLER: $_"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Vergleichen: $_", "Fehler")
    }
    finally {
        Reset-Progress
    }
})

# Event Handler für "Nach Datum löschen"
$btnDeleteByDate.Add_Click({
    if ([string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle einen Zielordner aus.", "Fehler")
        return
    }

    $selectedDate = $dtpDeleteDate.Value.Date  # Nur das Datum ohne Uhrzeit
    $dateType = if ($rbCreationDate.Checked) { "Erstellungsdatum" } else { "Änderungsdatum" }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Es werden alle Dateien im Zielverzeichnis gelöscht, die das Datum $($selectedDate.ToString('dd.MM.yyyy')) ($dateType) haben. Fortfahren?",
        "Warnung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Log "Starte Löschvorgang für Dateien mit Datum $($selectedDate.ToString('dd.MM.yyyy'))..."
            Write-Log "Zielordner: $($txtTarget.Text)"
            Write-Log "Datum: $($selectedDate.ToString('dd.MM.yyyy'))"
            Write-Log "Typ: $dateType"

            $files = Get-ChildItem -Path $txtTarget.Text -Recurse -File
            $count = 0

            foreach ($file in $files) {
                $dateToCheck = if ($rbCreationDate.Checked) {
                    $file.CreationTime.Date  # Nur das Datum ohne Uhrzeit
                } else {
                    $file.LastWriteTime.Date  # Nur das Datum ohne Uhrzeit
                }

                if ($dateToCheck -eq $selectedDate) {  # Exakte Übereinstimmung
                    Remove-Item $file.FullName -Force
                    $count++
                    Write-Log "Gelöscht: $($file.FullName) ($($dateToCheck.ToString('dd.MM.yyyy')))"
                }
            }

            Write-Log "Löschvorgang abgeschlossen. $count Dateien wurden gelöscht."
            [System.Windows.Forms.MessageBox]::Show("$count Dateien wurden erfolgreich gelöscht!", "Erfolg")
        }
        catch {
            Write-Log "FEHLER: $_"
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen der Dateien: $_", "Fehler")
        }
    }
})

# Lösch-Button Event Handler
$btnDelete.Add_Click({
    if ([string]::IsNullOrEmpty($txtSource.Text) -or [string]::IsNullOrEmpty($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle den Quell- und Zielordner aus.", "Fehler")
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Dies wird Dateien und Ordner im Zielverzeichnis löschen, die nicht in der Quelle existieren. Fortfahren?",
        "Warnung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Log "Starte Bereinigung..."
            Write-Log "Quelle: $($txtSource.Text)"
            Write-Log "Ziel: $($txtTarget.Text)"

            # Setze die Codepage auf 1252
            chcp 1252

            $roboParams = @($txtSource.Text, $txtTarget.Text, "/E", "/PURGE", "/XF", "*")

            if ($chkCreateLog.Checked) {
                $logFile = Get-LogFileName
                $roboParams += "/LOG:$logFile"
                Write-Log "Logdatei wird erstellt unter: $logFile"
            }

            $output = robocopy @roboParams

            foreach ($line in $output) {
                Write-Log $line
            }

            Write-Log "Bereinigung abgeschlossen"
            [System.Windows.Forms.MessageBox]::Show("Bereinigung erfolgreich abgeschlossen!", "Erfolg")
        }
        catch {
            Write-Log "FEHLER: $_"
            [System.Windows.Forms.MessageBox]::Show("Fehler bei der Bereinigung: $_", "Fehler")
        }
    }
})

# Löschen nach Pattern Event Handler
$btnDeleteByPattern.Add_Click({
    if ([string]::IsNullOrEmpty($txtTarget.Text) -or [string]::IsNullOrEmpty($txtNamePattern.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie einen Zielordner und geben Sie ein Pattern ein.", "Fehler")
        return
    }

    $pattern = $txtNamePattern.Text

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Dies wird alle Dateien im Zielverzeichnis löschen, die '$pattern' im Namen enthalten. Fortfahren?",
        "Warnung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Log "Starte Löschvorgang nach Namens-Pattern..."
            Write-Log "Zielordner: $($txtTarget.Text)"
            Write-Log "Pattern: $pattern"

            $files = Get-ChildItem -Path $txtTarget.Text -Recurse -File | Where-Object { $_.Name -like "*$pattern*" }

            $count = 0
            foreach ($file in $files) {
                Remove-Item $file.FullName -Force
                $count++
                Write-Log "Gelöscht: $($file.FullName)"
            }

            Write-Log "Löschvorgang abgeschlossen. $count Dateien wurden gelöscht."
            [System.Windows.Forms.MessageBox]::Show("$count Dateien wurden erfolgreich gelöscht!", "Erfolg")
        }
        catch {
            Write-Log "FEHLER: $_"
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen der Dateien: $_", "Fehler")
        }
    }
})

# Initialisiere die ComboBoxen beim Start
$form.Add_Shown({
    Write-Log "Anwendung gestartet"
    Update-PathComboBoxes
    $paths = Load-Paths
    if ($paths.Count -gt 0) {
        $txtSource.Text = $paths[0].SourcePath
        $txtTarget.Text = $paths[0].TargetPath
    }
})

# Kontext-Menü für das Ausgabefenster
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$copyMenuItem = $contextMenu.Items.Add("In Zwischenablage kopieren")
$copyMenuItem.Add_Click({
    if ($txtOutput.Text.Length -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($txtOutput.Text)
        $statusLabel.Text = "Ausgabe wurde in die Zwischenablage kopiert"
    }
})
$txtOutput.ContextMenuStrip = $contextMenu

# Starte die Windows Forms Anwendung
[System.Windows.Forms.Application]::Run($form)
