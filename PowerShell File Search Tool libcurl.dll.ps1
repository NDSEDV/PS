<#
.SYNOPSIS
  File Search Tool 
.DESCRIPTION
  Das Tool hilft bei der täglichen Arbeit
.NOTES
  Version:        1.3
  Author:         Jörn Walter
  Creation Date:  2025-07-31

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

# Funktion für den HTML Bericht
function Create-HTMLReport {
    if ($global:searchResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Suchergebnisse für Bericht verfügbar.", "Information", "OK", "Information")
        return
    }
    
    # Splitting-Parameter
    $maxRowsPerFile = 500  # Anzahl Zeilen pro HTML-Datei
    $totalFiles = $global:searchResults.Count
    $needsSplitting = $totalFiles -gt $maxRowsPerFile
    
    if ($needsSplitting) {
        $numberOfParts = [Math]::Ceiling($totalFiles / $maxRowsPerFile)
        
        $dialogResult = [System.Windows.Forms.MessageBox]::Show(
            "📊 BERICHT-OPTIONEN`n`nGefundene Dateien: $totalFiles`n`n🔄 SPLITTING EMPFOHLEN:`n• $numberOfParts Teildateien mit je max. $maxRowsPerFile Zeilen`n• Index-Datei mit Navigation zwischen Teilen`n• Bessere Performance im Browser`n`n📄 EINZELER BERICHT:`n• Eine große HTML-Datei mit $totalFiles Zeilen`n• Kann bei vielen Zeilen langsam laden`n`nMöchtest du den Bericht in $numberOfParts Teile aufteilen?", 
            "Bericht aufteilen? ($totalFiles Dateien)", 
            "YesNoCancel", 
            "Question"
        )
        
        if ($dialogResult -eq "Cancel") {
            return
        }
        
        $shouldSplit = ($dialogResult -eq "Yes")
    } else {
        $shouldSplit = $false
    }
    
    # Speicherdialog
    if ($shouldSplit) {
        # Ordner für geteilten Bericht
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Wähle einen Ordner für die geteilten Berichte"
        $folderDialog.ShowNewFolderButton = $true
        
        if ($folderDialog.ShowDialog() -eq "OK") {
            try {
                $reportFolder = $folderDialog.SelectedPath
                $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
                $reportName = "FileSearchReport_$timestamp"
                
                Create-SplitReportByRows -ReportFolder $reportFolder -ReportName $reportName -MaxRowsPerFile $maxRowsPerFile
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen des geteilten Berichts: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    } else {
        # Einzelne Datei
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "HTML-Dateien (*.html)|*.html"
        $saveDialog.Title = "HTML-Bericht speichern"
        $saveDialog.FileName = "FileSearchReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            try {
                Create-SingleReport -FilePath $saveDialog.FileName
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen des Berichts: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    }
}

# Funktion: Geteilter Bericht nach Zeilen
function Create-SplitReportByRows {
    param(
        [string]$ReportFolder,
        [string]$ReportName,
        [int]$MaxRowsPerFile
    )
    
    Write-Host "Erstelle geteilten Bericht mit max. $MaxRowsPerFile Zeilen pro Datei..."
    
    # Sortiere Ergebnisse alphabetisch
    $sortedResults = $global:searchResults | Sort-Object Name
    $totalFiles = $sortedResults.Count
    $numberOfParts = [Math]::Ceiling($totalFiles / $MaxRowsPerFile)
    
    # Erstelle Hauptverzeichnis
    $mainReportPath = Join-Path $ReportFolder $ReportName
    if (-not (Test-Path $mainReportPath)) {
        New-Item -ItemType Directory -Path $mainReportPath -Force | Out-Null
    }
    
    # Sicherheitsanalyse
    $analysis = $null
    if ($filenameTextBox.Text -like "*libcurl*") {
        $analysis = Analyze-LibcurlSecurity
    }
    
    # 1. ERSTELLE INDEX-DATEI
    $indexPath = Join-Path $mainReportPath "index.html"
    Create-IndexFileByRows -FilePath $indexPath -TotalFiles $totalFiles -NumberOfParts $numberOfParts -MaxRowsPerFile $MaxRowsPerFile -Analysis $analysis -ReportName $ReportName
    
    # 2. ERSTELLE TEIL-DATEIEN mit aussagekräftigen Namen
    $createdFiles = @()
    for ($i = 0; $i -lt $numberOfParts; $i++) {
        $startIndex = $i * $MaxRowsPerFile
        $endIndex = [Math]::Min(($startIndex + $MaxRowsPerFile - 1), ($totalFiles - 1))
        $partNumber = $i + 1
        
        $partFiles = $sortedResults[$startIndex..$endIndex]
        
        # Erstelle aussagekräftigen Dateinamen basierend auf ersten und letzten Dateinamen
        $firstFile = $partFiles[0].Name
        $lastFile = $partFiles[-1].Name
        
        # Kürze Namen falls zu lang
        $firstShort = if ($firstFile.Length -gt 20) { $firstFile.Substring(0, 17) + "..." } else { $firstFile }
        $lastShort = if ($lastFile.Length -gt 20) { $lastFile.Substring(0, 17) + "..." } else { $lastFile }
        
        # Erstelle Dateinamen: "Teil1_explorer-to-notepad.html"
        $partFileName = "Teil$partNumber" + "_" + ($firstShort -replace '[<>:"/\\|?*]', '_') + "-bis-" + ($lastShort -replace '[<>:"/\\|?*]', '_') + ".html"
        
        # Falls der Name zu lang wird, verwende einfache Nummerierung
        if ($partFileName.Length -gt 80) {
            $partFileName = "Teil$partNumber" + "_Zeilen$($startIndex + 1)-$($endIndex + 1).html"
        }
        
        $partFilePath = Join-Path $mainReportPath $partFileName
        
        Create-PartFile -FilePath $partFilePath -PartNumber $partNumber -Files $partFiles -StartIndex ($startIndex + 1) -EndIndex ($endIndex + 1) -TotalFiles $totalFiles -ReportName $ReportName -FirstFile $firstShort -LastFile $lastShort -PartFileName $partFileName
        $createdFiles += $partFileName
        
        Write-Host "Teil $partNumber erstellt: $($partFiles.Count) Dateien ($firstShort bis $lastShort)"
    }
    
    # 3. ERSTELLE SICHERHEITS-DATEI (falls notwendig)
    if ($analysis -and $analysis.RiskyFiles.Count -gt 0) {
        $securityPath = Join-Path $mainReportPath "security.html"
        Create-SecurityFile -FilePath $securityPath -Analysis $analysis -ReportName $ReportName
        $createdFiles += "security.html"
    }
    
    $successMessage = @"
✅ GETEILTER BERICHT ERFOLGREICH ERSTELLT!

📁 Hauptordner: $mainReportPath

📊 Aufteilung:
• $numberOfParts Teildateien mit je max. $MaxRowsPerFile Zeilen
• Gesamt: $totalFiles Dateien

📄 Erstellte Dateien:
• index.html (Hauptnavigation)
$(foreach ($file in $createdFiles) { "• $file`n" })

🌐 Öffne die 'index.html' um zu beginnen.
"@
    
    [System.Windows.Forms.MessageBox]::Show($successMessage, "Geteilter Bericht erstellt", "OK", "Information")
    
    # Fragen ob Index geöffnet werden soll
    $result = [System.Windows.Forms.MessageBox]::Show("Möchtest du den Index-Bericht jetzt öffnen?", "Bericht öffnen", "YesNo", "Question")
    if ($result -eq "Yes") {
        Start-Process $indexPath
    }
}

# Funktion: Index-Datei für zeilenbasiertes Splitting
function Create-IndexFileByRows {
    param(
        [string]$FilePath,
        [int]$TotalFiles,
        [int]$NumberOfParts,
        [int]$MaxRowsPerFile,
        [object]$Analysis,
        [string]$ReportName
    )
    
    $totalSize = [math]::Round(($global:searchResults | Measure-Object -Property Size -Sum).Sum / 1024, 2)
    $withVersion = ($global:searchResults | Where-Object {$_.Version -ne "N/A"}).Count
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportName - Index</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1200px; margin: 0 auto; background: white;
            border-radius: 15px; box-shadow: 0 15px 35px rgba(13, 71, 161, 0.2);
            overflow: hidden; border: 1px solid #90caf9;
        }
        .header {
            background: linear-gradient(135deg, #0d47a1 0%, #1565c0 100%);
            color: white; padding: 40px; text-align: center;
        }
        .header h1 { margin: 0; font-size: 2.8em; font-weight: 300; }
        .header p { margin: 15px 0 0 0; opacity: 0.95; font-size: 1.2em; }
        .content { padding: 40px; }
        
        .info-section {
            background: linear-gradient(135deg, #e3f2fd 0%, #f8fbff 100%);
            border: 2px solid #90caf9; border-left: 6px solid #1976d2;
            padding: 25px; margin-bottom: 35px; border-radius: 10px;
        }
        .info-section h3 { margin: 0 0 15px 0; color: #0d47a1; font-size: 1.4em; }
        .info-section p { margin: 8px 0; color: #1565c0; font-size: 1.1em; }
        
        .summary {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px; margin-bottom: 30px;
        }
        .summary-item {
            text-align: center; background: white; padding: 25px;
            border-radius: 12px; box-shadow: 0 6px 20px rgba(13, 71, 161, 0.12);
            border: 2px solid #e3f2fd; border-top: 4px solid #1976d2;
        }
        .summary-item h4 { margin: 0; color: #0d47a1; font-size: 2.2em; font-weight: 700; }
        .summary-item p { margin: 10px 0 0 0; color: #1565c0; font-weight: 500; }
        
        .parts-grid {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px; margin-top: 30px;
        }
        .part-card {
            background: white; border: 2px solid #e3f2fd; border-radius: 12px;
            padding: 20px; transition: all 0.3s ease; box-shadow: 0 4px 12px rgba(13, 71, 161, 0.1);
        }
        .part-card:hover { transform: translateY(-3px); box-shadow: 0 8px 20px rgba(13, 71, 161, 0.2); }
        .part-card h4 {
            margin: 0 0 15px 0; color: #0d47a1; font-size: 1.5em;
            padding: 10px 15px; background: #e3f2fd; border-radius: 8px; text-align: center;
        }
        .part-card p { margin: 8px 0; color: #666; }
        .part-card a {
            display: inline-block; margin-top: 15px; padding: 10px 20px;
            background: #1976d2; color: white; text-decoration: none;
            border-radius: 25px; font-weight: 600; transition: all 0.3s ease;
        }
        .part-card a:hover { background: #0d47a1; transform: scale(1.05); }
        
        .security-warning {
            background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%);
            border: 2px solid #ff9800; border-left: 6px solid #f57c00;
            padding: 25px; margin-bottom: 35px; border-radius: 10px;
        }
        .security-warning h3 { margin: 0 0 15px 0; color: #e65100; font-size: 1.4em; }
        .security-warning a {
            display: inline-block; margin-top: 15px; padding: 12px 25px;
            background: #f57c00; color: white; text-decoration: none;
            border-radius: 25px; font-weight: 600;
        }
        .security-warning a:hover { background: #e65100; }
        
        .footer {
            background: #e3f2fd; padding: 25px; text-align: center;
            border-top: 3px solid #1976d2; color: #0d47a1;
        }
        .footer a { color: #1565c0; text-decoration: none; font-weight: 600; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 File Search Report - Index</h1>
            <p>Geteilter Bericht | Erstellt am $(Get-Date -Format 'dd.MM.yyyy um HH:mm:ss')</p>
        </div>
        
        <div class="content">
            <div class="info-section">
                <h3>📋 Suchparameter</h3>
                <p><strong>Dateiname:</strong> $($filenameTextBox.Text)</p>
                <p><strong>Suchpfad:</strong> $($pathTextBox.Text)</p>
                <p><strong>Berichtstyp:</strong> Geteilter Bericht (max. $MaxRowsPerFile Zeilen pro Teil)</p>
                <p><strong>Aufteilung:</strong> $NumberOfParts Teildateien</p>
            </div>
            
            <div class="summary">
                <div class="summary-item">
                    <h4>$TotalFiles</h4>
                    <p>Gefundene Dateien</p>
                </div>
                <div class="summary-item">
                    <h4>$totalSize</h4>
                    <p>Gesamt MB</p>
                </div>
                <div class="summary-item">
                    <h4>$withVersion</h4>
                    <p>Mit Versionsinformation</p>
                </div>
                <div class="summary-item">
                    <h4>$NumberOfParts</h4>
                    <p>Teildateien</p>
                </div>
            </div>
"@

    # Sicherheitswarnung falls notwendig
    if ($Analysis -and $Analysis.RiskyFiles.Count -gt 0) {
        $htmlContent += @"
            <div class="security-warning">
                <h3>⚠️ SICHERHEITSWARNUNG</h3>
                <p><strong>$($Analysis.RiskyFiles.Count) potentiell gefährliche Dateien gefunden!</strong></p>
                <p>Überprüfe dringend die Sicherheitsbewertung für Details und Empfehlungen.</p>
                <a href="security.html">🔒 Sicherheitsbericht anzeigen</a>
            </div>
"@
    }

    $htmlContent += @"
            <h3>📄 Berichtsteile</h3>
            <div class="parts-grid">
"@

    for ($i = 1; $i -le $NumberOfParts; $i++) {
        $startRow = (($i - 1) * $MaxRowsPerFile) + 1
        $endRow = [Math]::Min(($i * $MaxRowsPerFile), $TotalFiles)
        $rowsInPart = $endRow - $startRow + 1
        
        # Bestimme ersten und letzten Dateinamen für diesen Teil
        $partStartIndex = ($i - 1) * $MaxRowsPerFile
        $partEndIndex = [Math]::Min(($partStartIndex + $MaxRowsPerFile - 1), ($TotalFiles - 1))
        $sortedResults = $global:searchResults | Sort-Object Name
        $firstFileName = $sortedResults[$partStartIndex].Name
        $lastFileName = $sortedResults[$partEndIndex].Name
        
        # Kürze Namen für Anzeige
        $firstShort = if ($firstFileName.Length -gt 25) { $firstFileName.Substring(0, 22) + "..." } else { $firstFileName }
        $lastShort = if ($lastFileName.Length -gt 25) { $lastFileName.Substring(0, 22) + "..." } else { $lastFileName }
        
        # Erstelle den tatsächlichen Dateinamen (wie in der Create-SplitReportByRows Funktion)
        $firstFileShort = if ($firstFileName.Length -gt 20) { $firstFileName.Substring(0, 17) + "..." } else { $firstFileName }
        $lastFileShort = if ($lastFileName.Length -gt 20) { $lastFileName.Substring(0, 17) + "..." } else { $lastFileName }
        $partFileName = "Teil$i" + "_" + ($firstFileShort -replace '[<>:"/\\|?*]', '_') + "-bis-" + ($lastFileShort -replace '[<>:"/\\|?*]', '_') + ".html"
        
        if ($partFileName.Length -gt 80) {
            $partFileName = "Teil$i" + "_Zeilen$startRow-$endRow.html"
        }
        
        $htmlContent += @"
                <div class="part-card">
                    <h4>📄 Teil $i</h4>
                    <p><strong>$firstShort</strong> bis <strong>$lastShort</strong></p>
                    <p><strong>Zeilen $startRow - $endRow</strong> ($rowsInPart Dateien)</p>
                    <p>Alle Dateien von "$firstShort" bis "$lastShort" alphabetisch sortiert</p>
                    <a href="$partFileName">Teil $i öffnen</a>
                </div>
"@
    }

    $htmlContent += @"
            </div>
        </div>
        
        <div class="footer">
            <p>© 2025 Jörn Walter - <a href="https://www.der-windows-papst.de">https://www.der-windows-papst.de</a></p>
            <p>File Search Tool v1.2 - Split Report (Zeilen-basiert)</p>
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
}

# Funktion: Teil-Datei erstellen
function Create-PartFile {
    param(
        [string]$FilePath,
        [int]$PartNumber,
        [array]$Files,
        [int]$StartIndex,
        [int]$EndIndex,
        [int]$TotalFiles,
        [string]$ReportName,
        [string]$FirstFile,
        [string]$LastFile,
        [string]$PartFileName
    )
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$ReportName - Teil $PartNumber</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1400px; margin: 0 auto; background: white;
            border-radius: 15px; box-shadow: 0 15px 35px rgba(13, 71, 161, 0.2);
            overflow: hidden; border: 1px solid #90caf9;
        }
        .header {
            background: linear-gradient(135deg, #0d47a1 0%, #1565c0 100%);
            color: white; padding: 30px; text-align: center;
        }
        .header h1 { margin: 0; font-size: 2.5em; font-weight: 300; }
        .header p { margin: 15px 0 0 0; opacity: 0.95; font-size: 1.1em; }
        .header .range-info { 
            background: rgba(255,255,255,0.1); padding: 10px 20px; 
            border-radius: 20px; margin-top: 15px; display: inline-block;
            font-family: 'Courier New', monospace; font-size: 0.9em;
        }
        
        .nav-bar {
            background: #f8fbff; padding: 15px 30px; border-bottom: 2px solid #e3f2fd;
            display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;
        }
        .nav-bar a {
            color: #1976d2; text-decoration: none; font-weight: 600;
            padding: 8px 16px; border-radius: 20px; transition: all 0.3s ease; margin: 2px;
        }
        .nav-bar a:hover { background: #1976d2; color: white; }
        .nav-info { color: #666; font-size: 0.9em; }
        
        .content { padding: 30px; }
        
        .table-container {
            width: 100%; overflow-x: auto; border-radius: 12px;
            box-shadow: 0 8px 25px rgba(13, 71, 161, 0.15); border: 2px solid #e3f2fd;
        }
        table {
            width: 100%; min-width: 1000px; border-collapse: collapse;
            background: white; table-layout: fixed;
        }
        /* OPTIMIERTE SPALTENBREITEN FÜR LANGE DATEINAMEN */
        table colgroup col:nth-child(1) { width: 20%; } /* Dateiname breiter */
        table colgroup col:nth-child(2) { width: 40%; } /* Pfad schmaler */
        table colgroup col:nth-child(3) { width: 15%; } /* Version */
        table colgroup col:nth-child(4) { width: 10%; } /* Größe */
        table colgroup col:nth-child(5) { width: 15%; } /* Erstellt */
        
        th {
            background: linear-gradient(135deg, #0d47a1 0%, #1976d2 100%);
            color: white; padding: 18px 15px; text-align: left;
            font-weight: 600; text-transform: uppercase; font-size: 0.95em;
            position: sticky; top: 0; z-index: 10;
        }
        td { 
            padding: 12px 8px; border-bottom: 1px solid #e3f2fd; 
            vertical-align: top; word-wrap: break-word; overflow-wrap: break-word;
        }
        tr:nth-child(even) { background: #f8fbff; }
        tr:hover { background: #e3f2fd; transform: scale(1.001); }
        
        /* VERBESSERTE ZELLEN-FORMATIERUNG */
        .filename-cell { 
            font-weight: 600; color: #0d47a1; font-size: 0.9em;
            max-width: 0; /* Ermöglicht Textumbruch */
            word-break: break-word;
            hyphens: auto;
            line-height: 1.3;
        }
        
        .filename-cell:hover {
            background: rgba(13, 71, 161, 0.1);
            cursor: help;
        }
        
        /* Tooltip für vollständigen Dateinamen */
        .filename-cell[title] {
            position: relative;
        }
        
        .path-cell {
            font-family: 'Courier New', monospace; font-size: 0.8em;
            color: #1565c0; word-break: break-all; max-width: 0;
            line-height: 1.2;
        }
        
        .version-cell {
            font-weight: bold; color: #0d47a1; text-align: center;
            background: rgba(187, 222, 251, 0.2); padding: 4px 8px; 
            border-radius: 12px; font-size: 0.85em;
        }
        
        .size-cell { 
            text-align: right; font-weight: 600; color: #1976d2; 
            font-family: 'Courier New', monospace; font-size: 0.9em;
        }
        
        .date-cell { 
            font-size: 0.85em; color: #666; white-space: nowrap;
        }
        
        /* RESPONSIVE VERBESSERUNGEN */
        @media screen and (max-width: 1200px) {
            table { min-width: 900px; }
            .filename-cell, .path-cell { font-size: 0.8em; }
        }
        
        @media screen and (max-width: 768px) {
            .nav-bar { flex-direction: column; text-align: center; }
            .nav-bar > div { margin: 5px 0; }
            table { min-width: 700px; }
            th, td { padding: 10px 6px; }
            .header h1 { font-size: 2em; }
            .header .range-info { font-size: 0.8em; padding: 8px 15px; }
        }
        
        .footer {
            background: #e3f2fd; padding: 20px; text-align: center;
            border-top: 3px solid #1976d2; color: #0d47a1;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 Teil $PartNumber</h1>
            <p>Zeilen $StartIndex - $EndIndex von $TotalFiles</p>
            <div class="range-info">
                Von: <strong>$FirstFile</strong><br>
                Bis: <strong>$LastFile</strong>
            </div>
        </div>
        
        <div class="nav-bar">
            <div>
                <a href="index.html">🏠 Zur Übersicht</a>
"@

    # Navigations-Links zu anderen Teilen (mit korrekten Dateinamen)
    $totalParts = [Math]::Ceiling($TotalFiles / 500)
    
    if ($PartNumber -gt 1) {
        # Bestimme Dateiname des vorherigen Teils
        $prevStartIndex = ($PartNumber - 2) * 500
        $prevEndIndex = [Math]::Min(($prevStartIndex + 499), ($TotalFiles - 1))
        $sortedResults = $global:searchResults | Sort-Object Name
        $prevFirstFile = $sortedResults[$prevStartIndex].Name
        $prevLastFile = $sortedResults[$prevEndIndex].Name
        
        $prevFirstShort = if ($prevFirstFile.Length -gt 20) { $prevFirstFile.Substring(0, 17) + "..." } else { $prevFirstFile }
        $prevLastShort = if ($prevLastFile.Length -gt 20) { $prevLastFile.Substring(0, 17) + "..." } else { $prevLastFile }
        $prevPartFileName = "Teil$($PartNumber - 1)" + "_" + ($prevFirstShort -replace '[<>:"/\\|?*]', '_') + "-bis-" + ($prevLastShort -replace '[<>:"/\\|?*]', '_') + ".html"
        
        if ($prevPartFileName.Length -gt 80) {
            $prevPartFileName = "Teil$($PartNumber - 1)" + "_Zeilen$(($prevStartIndex + 1))-$(($prevEndIndex + 1)).html"
        }
        
        $htmlContent += "<a href='$prevPartFileName'>⬅️ Teil $($PartNumber - 1)</a>"
    }
    
    if ($PartNumber -lt $totalParts) {
        # Bestimme Dateiname des nächsten Teils
        $nextStartIndex = $PartNumber * 500
        $nextEndIndex = [Math]::Min(($nextStartIndex + 499), ($TotalFiles - 1))
        $sortedResults = $global:searchResults | Sort-Object Name
        $nextFirstFile = $sortedResults[$nextStartIndex].Name
        $nextLastFile = $sortedResults[$nextEndIndex].Name
        
        $nextFirstShort = if ($nextFirstFile.Length -gt 20) { $nextFirstFile.Substring(0, 17) + "..." } else { $nextFirstFile }
        $nextLastShort = if ($nextLastFile.Length -gt 20) { $nextLastFile.Substring(0, 17) + "..." } else { $nextLastFile }
        $nextPartFileName = "Teil$($PartNumber + 1)" + "_" + ($nextFirstShort -replace '[<>:"/\\|?*]', '_') + "-bis-" + ($nextLastShort -replace '[<>:"/\\|?*]', '_') + ".html"
        
        if ($nextPartFileName.Length -gt 80) {
            $nextPartFileName = "Teil$($PartNumber + 1)" + "_Zeilen$(($nextStartIndex + 1))-$(($nextEndIndex + 1)).html"
        }
        
        $htmlContent += "<a href='$nextPartFileName'>Teil $($PartNumber + 1) ➡️</a>"
    }

    $htmlContent += @"
            </div>
            <div class="nav-info">
                $($Files.Count) Dateien | Teil $PartNumber von $totalParts
            </div>
        </div>
        
        <div class="content">
            <div class="table-container">
                <table>
                    <colgroup>
                        <col style="width: 20%;"><col style="width: 40%;"><col style="width: 15%;">
                        <col style="width: 10%;"><col style="width: 15%;">
                    </colgroup>
                    <thead>
                        <tr>
                            <th>Dateiname</th><th>Pfad</th><th>Version</th>
                            <th>Größe (KB)</th><th>Erstellt</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($file in $Files) {
        # Tooltip mit vollständigem Dateinamen falls gekürzt
        $tooltipTitle = if ($file.Name.Length -gt 50) { "title=`"$($file.Name)`"" } else { "" }
        
        $htmlContent += @"
                        <tr>
                            <td class="filename-cell" $tooltipTitle>$($file.Name)</td>
                            <td class="path-cell">$($file.Path)</td>
                            <td class="version-cell">$($file.Version)</td>
                            <td class="size-cell">$($file.Size)</td>
                            <td class="date-cell">$($file.Created)</td>
                        </tr>
"@
    }

    $htmlContent += @"
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>© 2025 Jörn Walter - File Search Tool v1.2</p>
            <p>Teil $PartNumber von $totalParts | $FirstFile bis $LastFile</p>
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
}

# Funktion: Einzelner Bericht (eine HTML-Datei)
function Create-SingleReport {
    param(
        [string]$FilePath
    )
    
    Write-Host "Erstelle einzelnen Bericht..."
    
    $sortedResults = $global:searchResults | Sort-Object Name
    $totalSize = [math]::Round(($global:searchResults | Measure-Object -Property Size -Sum).Sum / 1024, 2)
    $withVersion = ($global:searchResults | Where-Object {$_.Version -ne "N/A"}).Count
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Search Report</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px; background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
        }
        .container {
            max-width: 1400px; margin: 0 auto; background: white;
            border-radius: 15px; box-shadow: 0 15px 35px rgba(13, 71, 161, 0.2); overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #0d47a1 0%, #1565c0 100%);
            color: white; padding: 40px; text-align: center;
        }
        .header h1 { margin: 0; font-size: 2.8em; font-weight: 300; }
        .header p { margin: 15px 0 0 0; opacity: 0.95; font-size: 1.2em; }
        .content { padding: 40px; }
        
        .summary {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px; margin-bottom: 30px;
        }
        .summary-item {
            text-align: center; background: white; padding: 25px; border-radius: 12px;
            box-shadow: 0 6px 20px rgba(13, 71, 161, 0.12); border: 2px solid #e3f2fd;
            border-top: 4px solid #1976d2;
        }
        .summary-item h4 { margin: 0; color: #0d47a1; font-size: 2.2em; font-weight: 700; }
        .summary-item p { margin: 10px 0 0 0; color: #1565c0; font-weight: 500; }
        
        .table-container {
            width: 100%; overflow-x: auto; border-radius: 12px;
            box-shadow: 0 8px 25px rgba(13, 71, 161, 0.15); border: 2px solid #e3f2fd;
        }
        table {
            width: 100%; min-width: 1000px; border-collapse: collapse;
            background: white; table-layout: fixed;
        }
        table colgroup col:nth-child(1) { width: 15%; }
        table colgroup col:nth-child(2) { width: 45%; }
        table colgroup col:nth-child(3) { width: 15%; }
        table colgroup col:nth-child(4) { width: 10%; }
        table colgroup col:nth-child(5) { width: 15%; }
        
        th {
            background: linear-gradient(135deg, #0d47a1 0%, #1976d2 100%);
            color: white; padding: 18px 15px; text-align: left; font-weight: 600;
            text-transform: uppercase; font-size: 0.95em; position: sticky; top: 0; z-index: 10;
        }
        td { padding: 15px; border-bottom: 1px solid #e3f2fd; vertical-align: top; }
        tr:nth-child(even) { background: #f8fbff; }
        tr:hover { background: #e3f2fd; }
        
        .filename-cell { font-weight: 600; color: #0d47a1; }
        .path-cell { font-family: 'Courier New', monospace; font-size: 0.85em; color: #1565c0; word-break: break-all; }
        .version-cell { font-weight: bold; color: #0d47a1; text-align: center; background: rgba(187, 222, 251, 0.2); padding: 6px 10px; border-radius: 15px; }
        .size-cell { text-align: right; font-weight: 600; color: #1976d2; font-family: monospace; }
        .date-cell { font-size: 0.9em; color: #666; }
        
        .footer { background: #e3f2fd; padding: 25px; text-align: center; border-top: 3px solid #1976d2; color: #0d47a1; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 File Search Report</h1>
            <p>Einzelner Bericht | Erstellt am $(Get-Date -Format 'dd.MM.yyyy um HH:mm:ss')</p>
        </div>
        
        <div class="content">
            <div class="summary">
                <div class="summary-item">
                    <h4>$($sortedResults.Count)</h4>
                    <p>Gefundene Dateien</p>
                </div>
                <div class="summary-item">
                    <h4>$totalSize</h4>
                    <p>Gesamt MB</p>
                </div>
                <div class="summary-item">
                    <h4>$withVersion</h4>
                    <p>Mit Versionsinformation</p>
                </div>
            </div>
            
            <div class="table-container">
                <table>
                    <colgroup>
                        <col style="width: 15%;"><col style="width: 45%;"><col style="width: 15%;">
                        <col style="width: 10%;"><col style="width: 15%;">
                    </colgroup>
                    <thead>
                        <tr>
                            <th>Dateiname</th><th>Pfad</th><th>Version</th>
                            <th>Größe (KB)</th><th>Erstellt</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($result in $sortedResults) {
        $htmlContent += @"
                        <tr>
                            <td class="filename-cell">$($result.Name)</td>
                            <td class="path-cell">$($result.Path)</td>
                            <td class="version-cell">$($result.Version)</td>
                            <td class="size-cell">$($result.Size)</td>
                            <td class="date-cell">$($result.Created)</td>
                        </tr>
"@
    }

    $htmlContent += @"
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>© 2025 Jörn Walter - <a href="https://www.der-windows-papst.de" style="color: #1565c0; text-decoration: none;">https://www.der-windows-papst.de</a></p>
            <p>File Search Tool v1.2 - Einzelner Bericht</p>
        </div>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
    
    [System.Windows.Forms.MessageBox]::Show("Einzelner HTML-Bericht erstellt:`n$FilePath", "Erfolg", "OK", "Information")
    
    $result = [System.Windows.Forms.MessageBox]::Show("Möchtest du den Bericht jetzt öffnen?", "Bericht öffnen", "YesNo", "Question")
    if ($result -eq "Yes") {
        Start-Process $FilePath
    }
}

# Funktion: Sicherheitsbericht erstellen (wiederverwendet)
function Create-SecurityFile {
    param(
        [string]$FilePath,
        [object]$Analysis,
        [string]$ReportName
    )
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>$ReportName - Sicherheitsbericht</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; }
        .warning { background: #fff3e0; border: 2px solid #ff9800; padding: 20px; border-radius: 8px; margin: 20px 0; }
        .warning h2 { color: #e65100; margin: 0 0 15px 0; }
        .risk-high { background: #ffebee; border-left: 4px solid #f44336; padding: 15px; margin: 10px 0; }
        .risk-medium { background: #fff8e1; border-left: 4px solid #ffc107; padding: 15px; margin: 10px 0; }
        .nav-bar { padding: 15px; background: #e3f2fd; margin-bottom: 20px; border-radius: 8px; }
        .nav-bar a { color: #1976d2; text-decoration: none; font-weight: 600; }
    </style>
</head>
<body>
    <div class="container">
        <div class="nav-bar">
            <a href="index.html">🏠 Zurück zur Übersicht</a>
        </div>
        
        <h1>🔒 Sicherheitsbericht</h1>
        
        <div class="warning">
            <h2>⚠️ SICHERHEITSWARNUNG</h2>
            <p><strong>$($Analysis.RiskyFiles.Count) potentiell gefährliche Dateien gefunden!</strong></p>
        </div>
        
        <h3>🚨 Riskante Dateien:</h3>
"@

    foreach ($risky in $Analysis.RiskyFiles) {
        $cssClass = switch ($risky.RiskLevel) {
            "KRITISCH" { "risk-high" }
            "HOCH" { "risk-high" }
            default { "risk-medium" }
        }
        
        $htmlContent += @"
        <div class="$cssClass">
            <strong>$($risky.RiskLevel):</strong> $($risky.Path)<br>
            <small>Version: $($risky.Version) | Erstellt: $($risky.Created)<br>
            Grund: $($risky.Reason)</small>
        </div>
"@
    }

    $htmlContent += @"
        <h3>📋 Empfohlene Maßnahmen:</h3>
        <ul>
            <li>Aktualisierung auf libcurl 8.15.0+ (neueste Version)</li>
            <li>Entfernung veralteter Versionen (besonders 0.0.0.0 aus 2013)</li>
            <li>Überprüfung der Anwendungen die libcurl verwenden</li>
            <li>Regelmäßige Sicherheitsupdates implementieren</li>
        </ul>
        
        <p><strong>📥 AKTUELLE VERSION HERUNTERLADEN:</strong><br>
        <a href="https://curl.se/download.html" target="_blank">https://curl.se/download.html</a></p>
    </div>
</body>
</html>
"@

    $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
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
