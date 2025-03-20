<#
.SYNOPSIS
  Sort your files
.DESCRIPTION
  The tool has a German edition but can also be used on English OS systems.The tool is intended to help you with your daily business.
  This script allows you to copy files from a source folder to a destination folder with the option to sort the files by file extension or by year/month
.PARAMETER language
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-03-20
  Purpose/Change: Added recursive folder option
  
  Copyright (c) Jörn Walter. All rights reserved.
#>

# Funktion zum Überprüfen, ob das Skript mit Administratorrechten ausgeführt wird
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Alias für Test-Admin zur Konsistenz im Code
function Test-IsAdmin {
    return Test-Admin
}

# Überprüft, ob das Skript mit Administratorrechten ausgeführt wird
if (-not (Test-Admin)) {
    # If not, restart the script with administrative privileges
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

# GUI-Bibliothek laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Variable für Abbruch-Status
$global:abbrechenAngefordert = $false

# Hauptfenster erstellen
$formMain = New-Object System.Windows.Forms.Form
$formMain.Text = "Sort your files"
$formMain.Size = New-Object System.Drawing.Size(600, 630)
$formMain.StartPosition = "CenterScreen"
$formMain.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Quellordner-Auswahl
$labelQuelle = New-Object System.Windows.Forms.Label
$labelQuelle.Location = New-Object System.Drawing.Point(20, 20)
$labelQuelle.Size = New-Object System.Drawing.Size(150, 23)
$labelQuelle.Text = "Quellordner:"
$formMain.Controls.Add($labelQuelle)

$textBoxQuelle = New-Object System.Windows.Forms.TextBox
$textBoxQuelle.Location = New-Object System.Drawing.Point(20, 45)
$textBoxQuelle.Size = New-Object System.Drawing.Size(450, 23)
$textBoxQuelle.ReadOnly = $true
$formMain.Controls.Add($textBoxQuelle)

$buttonQuelle = New-Object System.Windows.Forms.Button
$buttonQuelle.Location = New-Object System.Drawing.Point(480, 45)
$buttonQuelle.Size = New-Object System.Drawing.Size(80, 23)
$buttonQuelle.Text = "Durchsuchen"
$buttonQuelle.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Quellordner auswählen"
    $dialogResult = $folderBrowser.ShowDialog()
    
    if ($dialogResult -eq "OK") {
        $textBoxQuelle.Text = $folderBrowser.SelectedPath
        # Wenn ein Quellordner ausgewählt wurde, aktivieren wir die Dateiliste
        AktualisiereDateiliste
    }
})
$formMain.Controls.Add($buttonQuelle)

# Zielordner-Auswahl
$labelZiel = New-Object System.Windows.Forms.Label
$labelZiel.Location = New-Object System.Drawing.Point(20, 80)
$labelZiel.Size = New-Object System.Drawing.Size(150, 23)
$labelZiel.Text = "Zielordner:"
$formMain.Controls.Add($labelZiel)

$textBoxZiel = New-Object System.Windows.Forms.TextBox
$textBoxZiel.Location = New-Object System.Drawing.Point(20, 105)
$textBoxZiel.Size = New-Object System.Drawing.Size(450, 23)
$textBoxZiel.ReadOnly = $true
$formMain.Controls.Add($textBoxZiel)

$buttonZiel = New-Object System.Windows.Forms.Button
$buttonZiel.Location = New-Object System.Drawing.Point(480, 105)
$buttonZiel.Size = New-Object System.Drawing.Size(80, 23)
$buttonZiel.Text = "Durchsuchen"
$buttonZiel.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Zielordner auswählen"
    $dialogResult = $folderBrowser.ShowDialog()
    
    if ($dialogResult -eq "OK") {
        $textBoxZiel.Text = $folderBrowser.SelectedPath
    }
})
$formMain.Controls.Add($buttonZiel)

# Rekursiv-Option
$checkBoxRekursiv = New-Object System.Windows.Forms.CheckBox
$checkBoxRekursiv.Location = New-Object System.Drawing.Point(20, 140)
$checkBoxRekursiv.Size = New-Object System.Drawing.Size(540, 30)
$checkBoxRekursiv.Text = "Unterordner einbeziehen (rekursiv)"
$checkBoxRekursiv.Checked = $true
$checkBoxRekursiv.Add_CheckedChanged({
    # Wenn sich der Status ändert und ein Quellordner ausgewählt ist, aktualisieren wir die Dateiliste
    if ($textBoxQuelle.Text -ne "") {
        $statusLabel.Text = "Lese Dateien neu ein..."
        $formMain.Refresh()
        AktualisiereDateiliste
    }
})
$formMain.Controls.Add($checkBoxRekursiv)

# Sortiermethode-Auswahl
$groupBoxSortierung = New-Object System.Windows.Forms.GroupBox
$groupBoxSortierung.Location = New-Object System.Drawing.Point(20, 180)
$groupBoxSortierung.Size = New-Object System.Drawing.Size(540, 70)
$groupBoxSortierung.Text = "Sortiermethode"
$formMain.Controls.Add($groupBoxSortierung)

$radioButtonDateiendung = New-Object System.Windows.Forms.RadioButton
$radioButtonDateiendung.Location = New-Object System.Drawing.Point(20, 25)
$radioButtonDateiendung.Size = New-Object System.Drawing.Size(250, 30)
$radioButtonDateiendung.Text = "Nach Dateiendung sortieren"
$radioButtonDateiendung.Checked = $true
$groupBoxSortierung.Controls.Add($radioButtonDateiendung)

$radioButtonDatum = New-Object System.Windows.Forms.RadioButton
$radioButtonDatum.Location = New-Object System.Drawing.Point(280, 25)
$radioButtonDatum.Size = New-Object System.Drawing.Size(250, 30)
$radioButtonDatum.Text = "Nach Jahr und Monat sortieren"
$groupBoxSortierung.Controls.Add($radioButtonDatum)

# Dateien überspringen Option
$checkBoxUeberspringen = New-Object System.Windows.Forms.CheckBox
$checkBoxUeberspringen.Location = New-Object System.Drawing.Point(30, 250)
$checkBoxUeberspringen.Size = New-Object System.Drawing.Size(540, 30)
$checkBoxUeberspringen.Text = "Bereits vorhandene Dateien überspringen (beschleunigt den Kopiervorgang)"
$checkBoxUeberspringen.Checked = $true
$formMain.Controls.Add($checkBoxUeberspringen)

# Dateiliste
$labelDateien = New-Object System.Windows.Forms.Label
$labelDateien.Location = New-Object System.Drawing.Point(20, 285)
$labelDateien.Size = New-Object System.Drawing.Size(200, 23)
$labelDateien.Text = "Gefundene Dateien:"
$formMain.Controls.Add($labelDateien)

$listBoxDateien = New-Object System.Windows.Forms.ListBox
$listBoxDateien.Location = New-Object System.Drawing.Point(20, 310)
$listBoxDateien.Size = New-Object System.Drawing.Size(540, 130)
$formMain.Controls.Add($listBoxDateien)

# Statusleiste
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit - Bitte Quellordner auswählen"
$statusStrip.Items.Add($statusLabel)
$formMain.Controls.Add($statusStrip)

# Fortschrittsbalken
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 455)
$progressBar.Size = New-Object System.Drawing.Size(200, 30)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$formMain.Controls.Add($progressBar)

# Fortschrittsanzeige-Label
$labelFortschritt = New-Object System.Windows.Forms.Label
$labelFortschritt.Location = New-Object System.Drawing.Point(230, 450)
$labelFortschritt.Size = New-Object System.Drawing.Size(330, 35)
$labelFortschritt.Text = "0 von 0 Dateien kopiert"
$formMain.Controls.Add($labelFortschritt)

# Copyright-Label hinzufügen
$labelCopyright = New-Object System.Windows.Forms.Label
$labelCopyright.Location = New-Object System.Drawing.Point(20, 540)
$labelCopyright.Size = New-Object System.Drawing.Size(540, 23)
$labelCopyright.Text = "© 2025 Jörn Walter - https://www-der-windows-papst.de"
$labelCopyright.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$formMain.Controls.Add($labelCopyright)

# Start-Button
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Location = New-Object System.Drawing.Point(180, 500)
$buttonStart.Size = New-Object System.Drawing.Size(120, 30)
$buttonStart.Text = "Starten"
$buttonStart.Add_Click({
    if ($textBoxQuelle.Text -eq "" -or $textBoxZiel.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle sowohl einen Quell- als auch einen Zielordner aus.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Abbruch-Status zurücksetzen
    $global:abbrechenAngefordert = $false
    
    # Abbrechen-Button aktivieren und Start-Button deaktivieren
    $buttonAbbrechen.Enabled = $true
    $buttonStart.Enabled = $false
    
    $statusLabel.Text = "Kopiere und sortiere Dateien..."
    [System.Windows.Forms.Application]::DoEvents()
    
    
    try {
        # Sortierung nach ausgewählter Methode
        $ergebnis = $null
        # Wenn keine Dateien gefunden wurden, direkt Meldung ausgeben
        if (($listBoxDateien.Items.Count -eq 0) -or 
            ($listBoxDateien.Items.Count -eq 1 -and $listBoxDateien.Items[0] -match "Keine Dateien gefunden")) {
            $ergebnis = @{
                Gesamt = 0
                Kopiert = 0
                Uebersprungen = 0
            }
        } else {
            if ($radioButtonDateiendung.Checked) {
                $ergebnis = SortiereNachDateiendung
            } else {
                $ergebnis = SortiereNachDatum
            }
        }
        
        if ($global:abbrechenAngefordert) {
            $statusLabel.Text = "Vorgang abgebrochen!"
            [System.Windows.Forms.MessageBox]::Show(
                "Vorgang wurde abgebrochen!`n`n" +
                "Gefundene Dateien: $($ergebnis.Gesamt)`n" +
                "Kopierte Dateien: $($ergebnis.Kopiert)`n" +
                "Übersprungene Dateien: $($ergebnis.Uebersprungen)", 
                "Abgebrochen", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "Fertig!"
            [System.Windows.Forms.MessageBox]::Show(
                "Verarbeitung abgeschlossen!`n`n" +
                "Gefundene Dateien: $($ergebnis.Gesamt)`n" +
                "Kopierte Dateien: $($ergebnis.Kopiert)`n" +
                "Übersprungene Dateien: $($ergebnis.Uebersprungen)", 
                "Erfolg", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
    catch {
        $statusLabel.Text = "Fehler aufgetreten"
        [System.Windows.Forms.MessageBox]::Show("Ein Fehler ist aufgetreten: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        # Abbrechen-Button deaktivieren und Start-Button aktivieren
        $buttonAbbrechen.Enabled = $false
        $buttonAbbrechen.Text = "Abbrechen"
        $buttonStart.Enabled = $true
        [System.Windows.Forms.Application]::DoEvents()
    }
})
$formMain.Controls.Add($buttonStart)

# Abbrechen-Button hinzufügen
$buttonAbbrechen = New-Object System.Windows.Forms.Button
$buttonAbbrechen.Location = New-Object System.Drawing.Point(310, 500)
$buttonAbbrechen.Size = New-Object System.Drawing.Size(120, 30)
$buttonAbbrechen.Text = "Abbrechen"
$buttonAbbrechen.Enabled = $false
$buttonAbbrechen.Add_Click({
    if (-not $global:abbrechenAngefordert) {
        $global:abbrechenAngefordert = $true
        $statusLabel.Text = "Abbruch angefordert... Bitte warten..."
        $buttonAbbrechen.Text = "Abbruch angefordert..."
        $buttonAbbrechen.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
    }
})
$formMain.Controls.Add($buttonAbbrechen)

# Funktionen

# Funktion zum Aktualisieren der Dateiliste
function AktualisiereDateiliste {
    $listBoxDateien.Items.Clear()
    $labelFortschritt.Text = "Lese Dateien ein..."
    $statusLabel.Text = "Suche Dateien..."
    $progressBar.Value = 0
    $formMain.Refresh()
    
    if ($textBoxQuelle.Text -ne "") {
        # Aktualisierung mit Live-Zählung
        $gefundeneDateien = 0
        $listBoxDateien.Items.Add("Dateien werden eingelesen...")
        $formMain.Refresh()
        
        # Zunächst die Dateien im Hauptverzeichnis einlesen
        $direkteDateien = Get-ChildItem -Path $textBoxQuelle.Text -File
        $gefundeneDateien += $direkteDateien.Count
        
        foreach ($datei in $direkteDateien) {
            $listBoxDateien.Items.Add($datei.FullName)
            $statusLabel.Text = "$gefundeneDateien Dateien gefunden"
            $labelFortschritt.Text = "Gefundene Dateien: $gefundeneDateien"
            $formMain.Refresh()
        }
        
        # Unterverzeichnisse durchsuchen, wenn rekursiv aktiviert ist
        if ($checkBoxRekursiv.Checked) {
            $unterordner = Get-ChildItem -Path $textBoxQuelle.Text -Directory -Recurse
            foreach ($ordner in $unterordner) {
                $dateienInOrdner = Get-ChildItem -Path $ordner.FullName -File
                $gefundeneDateien += $dateienInOrdner.Count
                
                foreach ($datei in $dateienInOrdner) {
                    $listBoxDateien.Items.Add($datei.FullName)
                    $statusLabel.Text = "$gefundeneDateien Dateien gefunden"
                    $labelFortschritt.Text = "Gefundene Dateien: $gefundeneDateien"
                    $formMain.Refresh()
                }
            }
        }
        
        if ($gefundeneDateien -eq 0) {
            $listBoxDateien.Items.Clear()
            $listBoxDateien.Items.Add("Keine Dateien gefunden.")
            $labelFortschritt.Text = "Keine Dateien gefunden"
        } else {
            $statusLabel.Text = "$gefundeneDateien Dateien gefunden"
            $labelFortschritt.Text = "Bereit: $gefundeneDateien Dateien gefunden"
        }
    }
}

# Funktion zum Sortieren nach Dateiendung
function SortiereNachDateiendung {
    $quellOrdner = $textBoxQuelle.Text
    $zielOrdner = $textBoxZiel.Text
    $ueberspringen = $checkBoxUeberspringen.Checked
    
    # Dateien wurden bereits in der ListBox erfasst - wir verwenden diese statt erneut zu scannen
    $gesamtDateien = $listBoxDateien.Items.Count
    # Sondermeldung "Keine Dateien gefunden" berücksichtigen
    if ($gesamtDateien -eq 1 -and $listBoxDateien.Items[0] -match "Keine Dateien gefunden") {
        $gesamtDateien = 0
    }
    $kopierteAnzahl = 0
    $uebersprungeneAnzahl = 0
    
    # Progressbar initialisieren
    $progressBar.Maximum = $gesamtDateien
    $progressBar.Value = 0
    $labelFortschritt.Text = "0 von $gesamtDateien Dateien verarbeitet"
    
    # Liste der zu verarbeitenden Dateien zusammenstellen
    $zuBearbeitendeDateien = @()
    
    # Alle validen Dateipfade aus der ListBox extrahieren
    for ($i = 0; $i -lt $listBoxDateien.Items.Count; $i++) {
        $dateiPfad = $listBoxDateien.Items[$i]
        # Prüfen auf Informationsmeldungen, die keine Dateien sind
        if ($dateiPfad -match "Dateien werden eingelesen" -or 
            $dateiPfad -match "Keine Dateien gefunden") {
            continue
        }
        
        if (Test-Path -Path $dateiPfad -PathType Leaf) {
            $zuBearbeitendeDateien += Get-Item -Path $dateiPfad
        }
    }
    
    # Sicherstellen, dass wir eine korrekte Anzahl haben
    $gesamtDateien = $zuBearbeitendeDateien.Count
    $progressBar.Maximum = $gesamtDateien
    
    foreach ($datei in $zuBearbeitendeDateien) {
        # Prüfen, ob Abbruch angefordert wurde
        if ($global:abbrechenAngefordert) {
            break
        }
        
        $dateiendung = $datei.Extension.TrimStart(".")
        
        if ([string]::IsNullOrEmpty($dateiendung)) {
            $dateiendung = "OhneEndung"
        }
        
        $neuerOrdner = Join-Path -Path $zielOrdner -ChildPath $dateiendung
        
        if (-not (Test-Path -Path $neuerOrdner)) {
            New-Item -Path $neuerOrdner -ItemType Directory | Out-Null
        }
        
        $zielDateiPfad = Join-Path -Path $neuerOrdner -ChildPath $datei.Name
        
        # Prüfen ob die Datei bereits existiert und gleich ist
        $dateiExistiert = Test-Path -Path $zielDateiPfad
        $kopiereNotwendig = $true
        
        if ($dateiExistiert -and $ueberspringen) {
            $zielDatei = Get-Item -Path $zielDateiPfad
            
            # Vergleich von Größe und Änderungsdatum 
            if (($zielDatei.Length -eq $datei.Length) -and 
                ($zielDatei.LastWriteTime -eq $datei.LastWriteTime)) {
                $kopiereNotwendig = $false
                $uebersprungeneAnzahl++
            }
        }
        
        if ($kopiereNotwendig) {
            Copy-Item -Path $datei.FullName -Destination $zielDateiPfad -Force
            $kopierteAnzahl++
        }
        
        # Fortschritt aktualisieren
        $bearbeiteteAnzahl = $kopierteAnzahl + $uebersprungeneAnzahl
        $progressBar.Value = $bearbeiteteAnzahl
        $labelFortschritt.Text = "$bearbeiteteAnzahl von $gesamtDateien Dateien verarbeitet ($kopierteAnzahl kopiert, $uebersprungeneAnzahl übersprungen)"
        
        # Windows Forms Message-Loop verarbeiten lassen, damit UI reagieren kann
        [System.Windows.Forms.Application]::DoEvents()
        
        # Prüfen, ob Abbruch angefordert wurde (nochmals prüfen nach DoEvents)
        if ($global:abbrechenAngefordert) {
            break
        }
    }
    
    return @{
        Gesamt = $gesamtDateien
        Kopiert = $kopierteAnzahl
        Uebersprungen = $uebersprungeneAnzahl
    }
}

# Funktion zum Sortieren nach Datum
function SortiereNachDatum {
    $quellOrdner = $textBoxQuelle.Text
    $zielOrdner = $textBoxZiel.Text
    $ueberspringen = $checkBoxUeberspringen.Checked
    
    # Dateien wurden bereits in der ListBox erfasst - wir verwenden diese statt erneut zu scannen
    $gesamtDateien = $listBoxDateien.Items.Count
    # Sondermeldung "Keine Dateien gefunden" berücksichtigen
    if ($gesamtDateien -eq 1 -and $listBoxDateien.Items[0] -match "Keine Dateien gefunden") {
        $gesamtDateien = 0
    }
    $kopierteAnzahl = 0
    $uebersprungeneAnzahl = 0
    
    # Progressbar initialisieren
    $progressBar.Maximum = $gesamtDateien
    $progressBar.Value = 0
    $labelFortschritt.Text = "0 von $gesamtDateien Dateien verarbeitet"
    
    # Liste der zu verarbeitenden Dateien zusammenstellen
    $zuBearbeitendeDateien = @()
    
    # Alle validen Dateipfade aus der ListBox extrahieren
    for ($i = 0; $i -lt $listBoxDateien.Items.Count; $i++) {
        $dateiPfad = $listBoxDateien.Items[$i]
        # Prüfen auf Informationsmeldungen, die keine Dateien sind
        if ($dateiPfad -match "Dateien werden eingelesen" -or 
            $dateiPfad -match "Keine Dateien gefunden") {
            continue
        }
        
        if (Test-Path -Path $dateiPfad -PathType Leaf) {
            $zuBearbeitendeDateien += Get-Item -Path $dateiPfad
        }
    }
    
    # Sicherstellen, dass wir eine korrekte Anzahl haben
    $gesamtDateien = $zuBearbeitendeDateien.Count
    $progressBar.Maximum = $gesamtDateien
    
    foreach ($datei in $zuBearbeitendeDateien) {
        # Prüfen, ob Abbruch angefordert wurde
        if ($global:abbrechenAngefordert) {
            break
        }
        
        $jahr = $datei.LastWriteTime.Year.ToString()
        $monat = $datei.LastWriteTime.Month.ToString("00")
        
        $jahrOrdner = Join-Path -Path $zielOrdner -ChildPath $jahr
        $monatOrdner = Join-Path -Path $jahrOrdner -ChildPath $monat
        
        if (-not (Test-Path -Path $jahrOrdner)) {
            New-Item -Path $jahrOrdner -ItemType Directory | Out-Null
        }
        
        if (-not (Test-Path -Path $monatOrdner)) {
            New-Item -Path $monatOrdner -ItemType Directory | Out-Null
        }
        
        $zielDateiPfad = Join-Path -Path $monatOrdner -ChildPath $datei.Name
        
        # Prüfen ob die Datei bereits existiert und gleich ist
        $dateiExistiert = Test-Path -Path $zielDateiPfad
        $kopiereNotwendig = $true
        
        if ($dateiExistiert -and $ueberspringen) {
            $zielDatei = Get-Item -Path $zielDateiPfad
            
            # Vergleich von Größe und Änderungsdatum 
            if (($zielDatei.Length -eq $datei.Length) -and 
                ($zielDatei.LastWriteTime -eq $datei.LastWriteTime)) {
                $kopiereNotwendig = $false
                $uebersprungeneAnzahl++
            }
        }
        
        if ($kopiereNotwendig) {
            Copy-Item -Path $datei.FullName -Destination $zielDateiPfad -Force
            $kopierteAnzahl++
        }
        
        # Fortschritt aktualisieren
        $bearbeiteteAnzahl = $kopierteAnzahl + $uebersprungeneAnzahl
        $progressBar.Value = $bearbeiteteAnzahl
        $labelFortschritt.Text = "$bearbeiteteAnzahl von $gesamtDateien Dateien verarbeitet ($kopierteAnzahl kopiert, $uebersprungeneAnzahl übersprungen)"
        
        # Windows Forms Message-Loop verarbeiten lassen, damit UI reagieren kann
        [System.Windows.Forms.Application]::DoEvents()
        
        # Prüfen, ob Abbruch angefordert wurde (nochmals prüfen nach DoEvents)
        if ($global:abbrechenAngefordert) {
            break
        }
    }
    
    return @{
        Gesamt = $gesamtDateien
        Kopiert = $kopierteAnzahl
        Uebersprungen = $uebersprungeneAnzahl
    }
}

# Hauptfenster anzeigen
[void]$formMain.ShowDialog()