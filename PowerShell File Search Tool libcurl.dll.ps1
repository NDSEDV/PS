<#
.SYNOPSIS
  File Search Tool 
.DESCRIPTION
  Das Tool hilft bei der täglichen Arbeit
.NOTES
  Version:        1.2
  Author:         Jörn Walter
  Creation Date:  2025-07-29

  Copyright (c) Jörn Walter. All rights reserved.
  Web: https://www.der-windows-papst.de
#>

# Admin-Check
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -Window hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hauptfenster erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Search Tool - libcurl.dll"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)

# Titel Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "File Search Tool - libcurl.dll"
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(300, 30)
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($titleLabel)

# Suchbereich GroupBox
$searchGroup = New-Object System.Windows.Forms.GroupBox
$searchGroup.Text = "Suche"
$searchGroup.Location = New-Object System.Drawing.Point(20, 60)
$searchGroup.Size = New-Object System.Drawing.Size(840, 115)
$searchGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($searchGroup)

# Dateiname Label und TextBox
$filenameLabel = New-Object System.Windows.Forms.Label
$filenameLabel.Text = "Dateiname:"
$filenameLabel.Location = New-Object System.Drawing.Point(20, 30)
$filenameLabel.Size = New-Object System.Drawing.Size(80, 20)
$searchGroup.Controls.Add($filenameLabel)

$filenameTextBox = New-Object System.Windows.Forms.TextBox
$filenameTextBox.Location = New-Object System.Drawing.Point(110, 28)
$filenameTextBox.Size = New-Object System.Drawing.Size(200, 25)
$filenameTextBox.Text = "libcurl.dll"
$searchGroup.Controls.Add($filenameTextBox)

# Suchpfad Label und TextBox
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Suchpfad:"
$pathLabel.Location = New-Object System.Drawing.Point(330, 30)
$pathLabel.Size = New-Object System.Drawing.Size(70, 20)
$searchGroup.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(410, 28)
$pathTextBox.Size = New-Object System.Drawing.Size(200, 25)
$pathTextBox.Text = "C:\"
$searchGroup.Controls.Add($pathTextBox)

# Informations-Button für libcurl.dll
$infoButton = New-Object System.Windows.Forms.Button
$infoButton.Text = "Info libcurl.dll"
$infoButton.Location = New-Object System.Drawing.Point(720, 26)
$infoButton.Size = New-Object System.Drawing.Size(100, 30)
$infoButton.BackColor = [System.Drawing.Color]::FromArgb(255, 165, 0)
$infoButton.ForeColor = [System.Drawing.Color]::White
$infoButton.FlatStyle = "Flat"
$infoButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchGroup.Controls.Add($infoButton)

# Suchen Button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Suchen"
$searchButton.Location = New-Object System.Drawing.Point(630, 26)
$searchButton.Size = New-Object System.Drawing.Size(80, 30)
$searchButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$searchButton.ForeColor = [System.Drawing.Color]::White
$searchButton.FlatStyle = "Flat"
$searchButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchGroup.Controls.Add($searchButton)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 65)
$progressBar.Size = New-Object System.Drawing.Size(800, 20)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$searchGroup.Controls.Add($progressBar)

# Status Label für Live-Anzeige
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Bereit für Suche..."
$statusLabel.Location = New-Object System.Drawing.Point(20, 90)
$statusLabel.Size = New-Object System.Drawing.Size(800, 15)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$statusLabel.Visible = $false
$searchGroup.Controls.Add($statusLabel)

# Ergebnisse DataGridView
$resultsGroup = New-Object System.Windows.Forms.GroupBox
$resultsGroup.Text = "Suchergebnisse"
$resultsGroup.Location = New-Object System.Drawing.Point(20, 195)
$resultsGroup.Size = New-Object System.Drawing.Size(840, 385)
$resultsGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($resultsGroup)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(15, 25)
$dataGridView.Size = New-Object System.Drawing.Size(810, 345)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = "FullRowSelect"
$dataGridView.MultiSelect = $false
$dataGridView.AutoSizeColumnsMode = "Fill"
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.BorderStyle = "Fixed3D"
$resultsGroup.Controls.Add($dataGridView)

# Kontextmenü für DataGridView erstellen
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Menüeintrag: Ordner öffnen
$openFolderItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openFolderItem.Text = "📁 Ordner im Explorer öffnen"
$openFolderItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0]
        $filePath = $selectedRow.Cells[1].Value
        if ($filePath -and (Test-Path $filePath)) {
            # Ordner öffnen und Datei markieren
            Start-Process "explorer.exe" "/select,`"$filePath`""
        }
    }
})
$contextMenu.Items.Add($openFolderItem)

# Menüeintrag: Datei-Eigenschaften
$propertiesItem = New-Object System.Windows.Forms.ToolStripMenuItem
$propertiesItem.Text = "🔍 Datei-Eigenschaften anzeigen"
$propertiesItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0]
        $filePath = $selectedRow.Cells[1].Value
        if ($filePath -and (Test-Path $filePath)) {
            try {
                # Eigenschaften-Dialog über Shell COM-Objekt
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace((Get-Item $filePath).DirectoryName)
                $file = $folder.ParseName((Get-Item $filePath).Name)
                $file.InvokeVerb("properties")
            } catch {
                # Fallback: PowerShell Get-ItemProperty
                try {
                    $fileInfo = Get-ItemProperty $filePath
                    $details = "Datei-Eigenschaften:`n`nName: $($fileInfo.Name)`nPfad: $($fileInfo.FullName)`nGröße: $($fileInfo.Length) Bytes`nErstellt: $($fileInfo.CreationTime)"
                    [System.Windows.Forms.MessageBox]::Show($details, "Datei-Eigenschaften", "OK", "Information")
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Eigenschaften können nicht angezeigt werden.", "Fehler", "OK", "Warning")
                }
            }
        }
    }
})
$contextMenu.Items.Add($propertiesItem)

# Trennlinie
$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# Menüeintrag: Pfad kopieren
$copyPathItem = New-Object System.Windows.Forms.ToolStripMenuItem
$copyPathItem.Text = "📋 Vollständigen Pfad kopieren"
$copyPathItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0]
        $filePath = $selectedRow.Cells[1].Value
        if ($filePath) {
            [System.Windows.Forms.Clipboard]::SetText($filePath)
            [System.Windows.Forms.MessageBox]::Show("Pfad in Zwischenablage kopiert:`n$filePath", "Kopiert", "OK", "Information")
        }
    }
})
$contextMenu.Items.Add($copyPathItem)

# Menüeintrag: Dateiname kopieren
$copyNameItem = New-Object System.Windows.Forms.ToolStripMenuItem
$copyNameItem.Text = "📄 Dateiname kopieren"
$copyNameItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0]
        $fileName = $selectedRow.Cells[0].Value
        if ($fileName) {
            [System.Windows.Forms.Clipboard]::SetText($fileName)
            [System.Windows.Forms.MessageBox]::Show("Dateiname in Zwischenablage kopiert:`n$fileName", "Kopiert", "OK", "Information")
        }
    }
})
$contextMenu.Items.Add($copyNameItem)

# Trennlinie
$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

# Menüeintrag: Datei-Details anzeigen
$showDetailsItem = New-Object System.Windows.Forms.ToolStripMenuItem
$showDetailsItem.Text = "ℹ️ Detaillierte Datei-Informationen"
$showDetailsItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedRow = $dataGridView.SelectedRows[0]
        $filePath = $selectedRow.Cells[1].Value
        if ($filePath -and (Test-Path $filePath)) {
            try {
                $fileInfo = Get-Item $filePath
                $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($filePath)
                
                $details = @"
📁 DATEI-DETAILS

📄 Name: $($fileInfo.Name)
📂 Pfad: $($fileInfo.FullName)
📊 Größe: $([math]::Round($fileInfo.Length / 1024, 2)) KB ($($fileInfo.Length) Bytes)
📅 Erstellt: $($fileInfo.CreationTime)
📝 Geändert: $($fileInfo.LastWriteTime)
🔓 Zugriff: $($fileInfo.LastAccessTime)
🏷️ Attribute: $($fileInfo.Attributes)

📋 VERSIONS-INFORMATIONEN:
🔢 Dateiversion: $($versionInfo.FileVersion)
🔢 Produktversion: $($versionInfo.ProductVersion)
🏢 Firma: $($versionInfo.CompanyName)
📖 Beschreibung: $($versionInfo.FileDescription)
🔒 Copyright: $($versionInfo.LegalCopyright)
"@
                
                # Sicherheitswarnung für libcurl.dll hinzufügen
                if ($fileInfo.Name -like "*libcurl*") {
                    $riskLevel = "✅ OK"
                    if ($versionInfo.FileVersion -eq "0.0.0.0" -or [string]::IsNullOrEmpty($versionInfo.FileVersion)) {
                        $riskLevel = "🚨 KRITISCH - Version unbekannt oder 0.0.0.0"
                    } elseif ($fileInfo.CreationTime -lt [DateTime]"01.01.2018") {
                        $riskLevel = "⚠️ HOCH - Datei von vor 2018"
                    }
                    
                    $details += @"

🔒 SICHERHEITSBEWERTUNG (libcurl.dll):
$riskLevel

📥 AKTUELLE VERSION HERUNTERLADEN:
https://curl.se/download.html
"@
                }
                
                [System.Windows.Forms.MessageBox]::Show($details, "Datei-Details: $($fileInfo.Name)", "OK", "Information")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Abrufen der Datei-Details: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    }
})
$contextMenu.Items.Add($showDetailsItem)

# Kontextmenü der DataGridView zuweisen
$dataGridView.ContextMenuStrip = $contextMenu

# DataGridView Spalten erstellen
$dataGridView.Columns.Add("Dateiname", "Dateiname")
$dataGridView.Columns.Add("Pfad", "Vollständiger Pfad")
$dataGridView.Columns.Add("Version", "Version")
$dataGridView.Columns.Add("Größe", "Größe (KB)")
$dataGridView.Columns.Add("Erstellt", "Erstellt")

# Spaltenbreiten anpassen
$dataGridView.Columns[0].Width = 150
$dataGridView.Columns[1].Width = 400
$dataGridView.Columns[2].Width = 100
$dataGridView.Columns[3].Width = 80
$dataGridView.Columns[4].Width = 120

# HTML-Bericht Button
$reportButton = New-Object System.Windows.Forms.Button
$reportButton.Text = "HTML-Bericht erstellen"
$reportButton.Location = New-Object System.Drawing.Point(20, 600)
$reportButton.Size = New-Object System.Drawing.Size(150, 30)
$reportButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$reportButton.ForeColor = [System.Drawing.Color]::White
$reportButton.FlatStyle = "Flat"
$reportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$reportButton.Enabled = $false
$form.Controls.Add($reportButton)

# Copyright Label
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Text = "© 2025 Jörn Walter - https://www.der-windows-papst.de"
$copyrightLabel.Location = New-Object System.Drawing.Point(520, 605)
$copyrightLabel.Size = New-Object System.Drawing.Size(350, 20)
$copyrightLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$copyrightLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$copyrightLabel.TextAlign = "MiddleRight"
$form.Controls.Add($copyrightLabel)

# Globale Variable für Suchergebnisse
$global:searchResults = @()

# Funktion: libcurl.dll Information anzeigen
function Show-LibcurlInfo {
    $infoText = @"
🔒 LIBCURL.DLL - SICHERHEITSINFORMATIONEN

WAS IST LIBCURL.DLL?
libcurl.dll ist eine weit verbreitete C-Bibliothek für:
• HTTP/HTTPS-Datenübertragung
• FTP, SFTP, SMTP, POP3 Protokolle  
• SSL/TLS-Verschlüsselung
• Verwendet von vielen Anwendungen und Browsern

⚠️ SICHERHEITSRISIKEN ALTER VERSIONEN:

KRITISCHE PROBLEME bei Versionen vor 2018:
• CVE-2018-16839: Heap-basierte Pufferüberläufe
• CVE-2017-8816: NTLM-Authentifizierung Schwachstellen
• CVE-2016-8617: Out-of-bounds Zugriffe
• CVE-2014-3707: Duplikate Cookie-Header Angriffe

🚨 BESONDERS GEFÄHRLICH: Version 0.0.0.0 (2013)
Diese Versionen enthalten:
• Unverschlüsselte Datenübertragung
• Keine Zertifikatsprüfung
• Buffer-Overflow Schwachstellen
• Anfällig für Man-in-the-Middle Angriffe

EMPFOHLENE MASSNAHMEN:
✅ Aktualisierung auf libcurl 8.15.0 (Juli 2025) 
✅ Entfernung veralteter Versionen (besonders < 8.0)
✅ Überprüfung der Anwendungen die libcurl verwenden
✅ Regelmäßige Sicherheitsupdates

📥 AKTUELLE VERSION HERUNTERLADEN:
https://curl.se/download.html

🔗 Weitere Informationen:
https://curl.se/libcurl/security.html
"@

    [System.Windows.Forms.MessageBox]::Show($infoText, "libcurl.dll - Sicherheitsinformationen", "OK", "Information")
}

# Funktion: Sicherheitsanalyse der gefundenen libcurl.dll Dateien
function Analyze-LibcurlSecurity {
    $riskyFiles = @()
    $safeFiles = @()
    
    foreach ($file in $global:searchResults) {
        if ($file.Name -like "*libcurl*") {
            # Analyse der Versionsnummer und des Datums
            $isRisky = $false
            $riskLevel = "Niedrig"
            $riskReason = ""
            
            # Version 0.0.0.0 oder N/A sind hochriskant
            if ($file.Version -eq "0.0.0.0" -or $file.Version -eq "N/A") {
                $isRisky = $true
                $riskLevel = "KRITISCH"
                $riskReason = "Version unbekannt oder 0.0.0.0 - wahrscheinlich aus 2013 oder früher"
            }
            # Dateien von vor 2018 sind riskant
            elseif ([DateTime]::ParseExact($file.Created.Split(' ')[0], "dd.MM.yyyy", $null) -lt [DateTime]"01.01.2018") {
                $isRisky = $true
                $riskLevel = "HOCH"
                $riskReason = "Datei von vor 2018 - enthält bekannte Sicherheitslücken"
            }
            # Alte Versionspattern erkennen
            elseif ($file.Version -match "^[0-6]\." -or $file.Version -match "^7\.") {
                $isRisky = $true
                $riskLevel = "HOCH"
                $riskReason = "Stark veraltete libcurl Version (< 8.0) - dringend Update erforderlich"
            }
            # Version 8.0-8.14 als mittel riskant einstufen
            elseif ($file.Version -match "^8\.[0-9]\." -or $file.Version -match "^8\.1[0-4]\.") {
                $isRisky = $true
                $riskLevel = "MITTEL"
                $riskReason = "Veraltete libcurl 8.x Version - Update auf 8.15.0 empfohlen"
            }
            
            if ($isRisky) {
                $riskyFiles += [PSCustomObject]@{
                    Path = $file.Path
                    Version = $file.Version
                    Created = $file.Created
                    RiskLevel = $riskLevel
                    Reason = $riskReason
                }
            } else {
                $safeFiles += $file
            }
        }
    }
    
    return @{
        RiskyFiles = $riskyFiles
        SafeFiles = $safeFiles
    }
}

# Funktion: Dateien suchen
function Search-Files {
    param(
        [string]$FileName,
        [string]$SearchPath
    )
    
    $progressBar.Visible = $true
    $statusLabel.Visible = $true
    $statusLabel.Text = "Starte Suche..."
    $searchButton.Enabled = $false
    $dataGridView.Rows.Clear()
    $global:searchResults = @()
    
    # Aktualisierung der GUI
    $form.Refresh()
    
    try {
        Write-Host "Suche nach '$FileName' in '$SearchPath'..."
        $statusLabel.Text = "Starte Suche nach '$FileName' in '$SearchPath'..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Erst alle Verzeichnisse sammeln für bessere Performance
        $statusLabel.Text = "Sammle Verzeichnisse..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $directories = @()
        try {
            $directories = Get-ChildItem -Path $SearchPath -Directory -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            $directories = @($SearchPath) + $directories  # Startverzeichnis hinzufügen
        }
        catch {
            $directories = @($SearchPath)
        }
        
        $totalDirs = $directories.Count
        $currentDir = 0
        
        Write-Host "Durchsuche $totalDirs Verzeichnisse..."
        
        # Durch jedes Verzeichnis gehen
        foreach ($directory in $directories) {
            $currentDir++
            
            # Live-Status aktualisieren
            $statusLabel.Text = "[$currentDir/$totalDirs] Durchsuche: $directory"
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # In diesem Verzeichnis nach der Datei suchen (nicht rekursiv, da wir selbst rekursiv durchgehen)
                $files = Get-ChildItem -Path $directory -Name $FileName -File -ErrorAction SilentlyContinue
                
                foreach ($file in $files) {
                    try {
                        $fullPath = Join-Path $directory $file
                        $fileInfo = Get-Item $fullPath -ErrorAction SilentlyContinue
                        
                        if ($fileInfo) {
                            # Versionsinformationen abrufen
                            $version = "N/A"
                            try {
                                $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullPath)
                                if ($versionInfo.FileVersion) {
                                    $version = $versionInfo.FileVersion
                                }
                            }
                            catch {
                                $version = "N/A"
                            }
                            
                            # Dateigröße in KB
                            $sizeKB = [math]::Round($fileInfo.Length / 1024, 2)
                            
                            # Erstellungsdatum formatieren
                            $created = $fileInfo.CreationTime.ToString("dd.MM.yyyy HH:mm")
                            
                            # Ergebnis zur DataGridView hinzufügen
                            $row = @($fileInfo.Name, $fullPath, $version, $sizeKB, $created)
                            $dataGridView.Rows.Add($row)
                            
                            # Für HTML-Bericht speichern
                            $global:searchResults += [PSCustomObject]@{
                                Name = $fileInfo.Name
                                Path = $fullPath
                                Version = $version
                                Size = $sizeKB
                                Created = $created
                            }
                            
                            # Live-Update: Gefundene Dateien anzeigen
                            $statusLabel.Text = "[$currentDir/$totalDirs] Gefunden: $($global:searchResults.Count) Dateien | Aktuell: $directory"
                            
                            # GUI aktualisieren
                            [System.Windows.Forms.Application]::DoEvents()
                            
                            Write-Host "Gefunden: $fullPath (Version: $version)"
                        }
                    }
                    catch {
                        Write-Host "Fehler beim Verarbeiten von $file`: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                # Verzeichnis nicht zugänglich - ignorieren
            }
            
            # Alle 50 Verzeichnisse GUI aktualisieren (Performance)
            if ($currentDir % 50 -eq 0) {
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        
        # Finale Status-Nachricht
        $statusLabel.Text = "Suche abgeschlossen! $($global:searchResults.Count) Dateien gefunden."
        
        # Sicherheitsanalyse für libcurl.dll durchführen
        if ($FileName -like "*libcurl*") {
            $analysis = Analyze-LibcurlSecurity
            if ($analysis.RiskyFiles.Count -gt 0) {
                $warningText = "⚠️ SICHERHEITSWARNUNG ⚠️`n`n"
                $warningText += "$($analysis.RiskyFiles.Count) potentiell gefährliche libcurl.dll Dateien gefunden!`n`n"
                
                foreach ($risky in $analysis.RiskyFiles) {
                    $warningText += "• $($risky.RiskLevel): $($risky.Path)`n"
                    $warningText += "  Version: $($risky.Version) | Erstellt: $($risky.Created)`n"
                    $warningText += "  Grund: $($risky.Reason)`n`n"
                }
                
                $warningText += "EMPFEHLUNG: Aktualisiere die betroffenen Anwendungen!`n"
                $warningText += "📥 AKTUELLE VERSION HERUNTERLADEN:`n"
                $warningText += "https://curl.se/download.html"
                
                [System.Windows.Forms.MessageBox]::Show($warningText, "SICHERHEITSWARNUNG - libcurl.dll", "OK", "Warning")
            }
        }
        
        # Status anzeigen
        $statusText = "Suche abgeschlossen. $($global:searchResults.Count) Dateien gefunden."
        [System.Windows.Forms.MessageBox]::Show($statusText, "Suche beendet", "OK", "Information")
        
        # HTML-Bericht Button aktivieren wenn Ergebnisse vorhanden
        $reportButton.Enabled = ($global:searchResults.Count -gt 0)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Suche: $($_.Exception.Message)", "Fehler", "OK", "Error")
        $statusLabel.Text = "Fehler bei der Suche!"
    }
    finally {
        $progressBar.Visible = $false
        $searchButton.Enabled = $true
        # Status Label bleibt sichtbar - KEIN TIMER!
    }
}

# Funktion: HTML-Bericht erstellen
function Create-HTMLReport {
    if ($global:searchResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Suchergebnisse für Bericht verfügbar.", "Information", "OK", "Information")
        return
    }
    
    # Speicherdialog
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "HTML-Dateien (*.html)|*.html"
    $saveDialog.Title = "HTML-Bericht speichern"
    $saveDialog.FileName = "FileSearchReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Search Report - libcurl.dll</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 15px 35px rgba(13, 71, 161, 0.2);
            overflow: hidden;
            border: 1px solid #90caf9;
        }
        .header {
            background: linear-gradient(135deg, #0d47a1 0%, #1565c0 100%);
            color: white;
            padding: 40px;
            text-align: center;
            position: relative;
        }
        .header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
        }
        .header h1 {
            margin: 0;
            font-size: 2.8em;
            font-weight: 300;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        .header p {
            margin: 15px 0 0 0;
            opacity: 0.95;
            font-size: 1.2em;
        }
        .content {
            padding: 40px;
            background: linear-gradient(to bottom, #f8fbff 0%, #ffffff 100%);
        }
        .info-section {
            background: linear-gradient(135deg, #e3f2fd 0%, #f8fbff 100%);
            border: 2px solid #90caf9;
            border-left: 6px solid #1976d2;
            padding: 25px;
            margin-bottom: 35px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(25, 118, 210, 0.1);
        }
        .info-section h3 {
            margin: 0 0 15px 0;
            color: #0d47a1;
            font-size: 1.4em;
            font-weight: 600;
        }
        .info-section p {
            margin: 8px 0;
            color: #1565c0;
            font-size: 1.1em;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 25px;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 8px 25px rgba(13, 71, 161, 0.15);
            border: 2px solid #e3f2fd;
        }
        th {
            background: linear-gradient(135deg, #0d47a1 0%, #1976d2 100%);
            color: white;
            padding: 18px 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.95em;
            letter-spacing: 1.2px;
            text-shadow: 0 1px 2px rgba(0,0,0,0.2);
        }
        td {
            padding: 15px;
            border-bottom: 1px solid #e3f2fd;
            vertical-align: top;
            transition: background 0.3s ease;
        }
        tr:nth-child(even) {
            background: linear-gradient(to right, #f8fbff 0%, #ffffff 100%);
        }
        tr:hover {
            background: linear-gradient(to right, #e3f2fd 0%, #f0f8ff 100%);
            transform: scale(1.001);
            box-shadow: 0 2px 8px rgba(25, 118, 210, 0.1);
        }
        .path-cell {
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            word-break: break-all;
            color: #1565c0;
            background: rgba(227, 242, 253, 0.3);
            padding: 8px;
            border-radius: 4px;
        }
        .version-cell {
            font-weight: bold;
            color: #0d47a1;
            background: rgba(187, 222, 251, 0.2);
            padding: 6px 10px;
            border-radius: 15px;
            text-align: center;
        }
        .size-cell {
            text-align: right;
            font-weight: 600;
            color: #1976d2;
        }
        .footer {
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            padding: 25px 40px;
            text-align: center;
            border-top: 3px solid #1976d2;
            color: #0d47a1;
        }
        .footer a {
            color: #1565c0;
            text-decoration: none;
            font-weight: 600;
        }
        .footer a:hover {
            color: #0d47a1;
            text-decoration: underline;
        }
        .summary {
            display: flex;
            justify-content: space-around;
            margin-bottom: 30px;
            gap: 20px;
        }
        .summary-item {
            text-align: center;
            background: linear-gradient(135deg, #ffffff 0%, #f8fbff 100%);
            padding: 25px 20px;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(13, 71, 161, 0.12);
            border: 2px solid #e3f2fd;
            border-top: 4px solid #1976d2;
            flex: 1;
            transition: transform 0.3s ease;
        }
        .summary-item:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(13, 71, 161, 0.18);
        }
        .summary-item h4 {
            margin: 0;
            color: #0d47a1;
            font-size: 2.2em;
            font-weight: 700;
            text-shadow: 0 1px 3px rgba(13, 71, 161, 0.2);
        }
        .summary-item p {
            margin: 10px 0 0 0;
            color: #1565c0;
            font-weight: 500;
            font-size: 1em;
        }
        .security-warning {
            background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
            border: 2px solid #ff9800;
            border-left: 6px solid #f57c00;
            padding: 25px;
            margin-bottom: 35px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(245, 124, 0, 0.2);
        }
        .security-warning h3 {
            margin: 0 0 15px 0;
            color: #e65100;
            font-size: 1.4em;
            font-weight: 600;
        }
        .security-warning h4 {
            color: #f57c00;
            margin: 15px 0 10px 0;
        }
        .warning-details {
            margin: 15px 0;
        }
        .risk-item {
            padding: 12px;
            margin: 8px 0;
            border-radius: 6px;
            border-left: 4px solid;
        }
        .risk-kritisch {
            background: #ffebee;
            border-left-color: #f44336;
            color: #c62828;
        }
        .risk-hoch {
            background: #fff3e0;
            border-left-color: #ff9800;
            color: #e65100;
        }
        .risk-mittel {
            background: #fff8e1;
            border-left-color: #ffc107;
            color: #f57c00;
        }
        .security-advice {
            background: rgba(76, 175, 80, 0.1);
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #4caf50;
            margin-top: 15px;
        }
        .security-advice ul {
            margin: 10px 0;
            padding-left: 20px;
        }
        .security-advice li {
            margin: 5px 0;
            color: #2e7d32;
        }
        .download-link {
            background: linear-gradient(135deg, #e8f5e8 0%, #c8e6c9 100%);
            border: 2px solid #4caf50;
            border-left: 6px solid #2e7d32;
            padding: 20px;
            margin: 15px 0;
            border-radius: 8px;
            text-align: center;
        }
        .download-link a {
            color: #2e7d32;
            text-decoration: none;
            font-weight: bold;
            font-size: 1.1em;
        }
        .download-link a:hover {
            color: #1b5e20;
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>File Search Report</h1>
            <p>Generiert am $(Get-Date -Format 'dd.MM.yyyy um HH:mm:ss')</p>
        </div>
        
        <div class="content">
            <div class="info-section">
                <h3>Suchparameter</h3>
                <p><strong>Dateiname:</strong> $($filenameTextBox.Text)</p>
                <p><strong>Suchpfad:</strong> $($pathTextBox.Text)</p>
                <p><strong>Gefundene Dateien:</strong> $($global:searchResults.Count)</p>
            </div>
"@

            # Sicherheitsanalyse für libcurl.dll hinzufügen
            if ($filenameTextBox.Text -like "*libcurl*") {
                $analysis = Analyze-LibcurlSecurity
                if ($analysis.RiskyFiles.Count -gt 0) {
                    $htmlContent += @"
            <div class="security-warning">
                <h3>⚠️ SICHERHEITSWARNUNG - libcurl.dll</h3>
                <p><strong>$($analysis.RiskyFiles.Count) potentiell gefährliche Dateien gefunden!</strong></p>
                <div class="warning-details">
                    <h4>Riskante Dateien:</h4>
"@
                    foreach ($risky in $analysis.RiskyFiles) {
                        $htmlContent += @"
                    <div class="risk-item risk-$($risky.RiskLevel.ToLower())">
                        <strong>$($risky.RiskLevel):</strong> $($risky.Path)<br>
                        <small>Version: $($risky.Version) | Erstellt: $($risky.Created)<br>
                        Grund: $($risky.Reason)</small>
                    </div>
"@
                    }
                    $htmlContent += @"
                </div>
                <div class="security-advice">
                    <h4>🔒 Empfohlene Maßnahmen:</h4>
                    <ul>
                        <li>Aktualisierung auf libcurl 7.68.0+ (2020 oder neuer)</li>
                        <li>Entfernung veralteter Versionen (besonders 0.0.0.0 aus 2013)</li>
                        <li>Überprüfung der Anwendungen die libcurl verwenden</li>
                        <li>Regelmäßige Sicherheitsupdates implementieren</li>
                    </ul>
                    <div class="download-link">
                        <p><strong>📥 AKTUELLE VERSION HERUNTERLADEN:</strong><br>
                        <a href="https://curl.se/download.html" target="_blank">https://curl.se/download.html</a></p>
                    </div>
                    <p><strong>Weitere Informationen:</strong> <a href="https://curl.se/libcurl/security.html" target="_blank">curl.se/libcurl/security.html</a></p>
                </div>
            </div>
"@
                }
            }
            
            $htmlContent += @"
            <div class="summary">
                <div class="summary-item">
                    <h4>$($global:searchResults.Count)</h4>
                    <p>Gefundene Dateien</p>
                </div>
                <div class="summary-item">
                    <h4>$([math]::Round(($global:searchResults | Measure-Object -Property Size -Sum).Sum, 2))</h4>
                    <p>Gesamt KB</p>
                </div>
                <div class="summary-item">
                    <h4>$(($global:searchResults | Where-Object {$_.Version -ne "N/A"}).Count)</h4>
                    <p>Mit Versionsinformation</p>
                </div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>Dateiname</th>
                        <th>Pfad</th>
                        <th>Version</th>
                        <th>Größe (KB)</th>
                        <th>Erstellt</th>
                    </tr>
                </thead>
                <tbody>
"@
            
            foreach ($result in $global:searchResults) {
                $htmlContent += @"
                    <tr>
                        <td>$($result.Name)</td>
                        <td class="path-cell">$($result.Path)</td>
                        <td class="version-cell">$($result.Version)</td>
                        <td class="size-cell">$($result.Size)</td>
                        <td>$($result.Created)</td>
                    </tr>
"@
            }
            
            $htmlContent += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>© 2025 Jörn Walter - <a href="https://www.der-windows-papst.de" style="color: #1565c0;">https://www.der-windows-papst.de</a></p>
            <p>Erstellt mit File Search Tool</p>
        </div>
    </div>
</body>
</html>
"@
            
            $htmlContent | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("HTML-Bericht erfolgreich erstellt:`n$($saveDialog.FileName)", "Erfolg", "OK", "Information")
            
            # Fragen ob Bericht geöffnet werden soll
            $result = [System.Windows.Forms.MessageBox]::Show("Möchtest du den Bericht jetzt öffnen?", "Bericht öffnen", "YesNo", "Question")
            if ($result -eq "Yes") {
                Start-Process $saveDialog.FileName
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen des Berichts: $($_.Exception.Message)", "Fehler", "OK", "Error")
        }
    }
}

# Event Handlers
$searchButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($filenameTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gib einen Dateinamen ein.", "Eingabe erforderlich", "OK", "Warning")
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($pathTextBox.Text) -or -not (Test-Path $pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gib einen gültigen Suchpfad ein.", "Ungültiger Pfad", "OK", "Warning")
        return
    }
    
    Search-Files -FileName $filenameTextBox.Text -SearchPath $pathTextBox.Text
})

$reportButton.Add_Click({
    Create-HTMLReport
})

$infoButton.Add_Click({
    Show-LibcurlInfo
})

# Enter-Taste für Suche aktivieren
$filenameTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq "Enter") {
        $searchButton.PerformClick()
    }
})

$pathTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq "Enter") {
        $searchButton.PerformClick()
    }
})

# Copyright Label klickbar machen
$copyrightLabel.Add_Click({
    Start-Process "https://www.der-windows-papst.de"
})
$copyrightLabel.Cursor = "Hand"

# Formular anzeigen
Write-Host "File Search Tool wird gestartet..."
[System.Windows.Forms.Application]::Run($form)
