<#
.SYNOPSIS
  Sign your Script
.DESCRIPTION
  The tool has a German edition but can also be used on English OS systems.The tool is intended to help you with your daily business.
.PARAMETER language
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-03-19
  Purpose/Change: Initial script development

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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Variablen
$global:isCreatingCertificate = $false
$global:certificateCreationCompleted = $false
$global:lastCertificateOperation = $null
$global:isDeleteInProgress = $false
$global:noInternetWarningShown = $false

# Funktion zum Ausführen von Befehlen mit Administratorrechten
function Invoke-AdminCommand {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Command,
        
        [Parameter(Mandatory=$false)]
        [ScriptBlock]$OnComplete = $null
    )
    
    # Status-Nachricht im Log
    $logTextBox.AppendText("Starte Befehl als separaten Prozess: $Command`r`n")
    
    # Führen wir den Befehl direkt im UI-Thread als Admin aus
    if (Test-IsAdmin) {
        # Skript läuft bereits als Admin, führe direkt aus
        try {
            $logTextBox.AppendText("Führe Befehl als Admin aus...`r`n")
            $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", "$Command" -Wait -PassThru -NoNewWindow
            
            if ($process.ExitCode -eq 0) {
                $resultObj = @{
                    Success = $true
                    Output = "Befehl erfolgreich ausgeführt."
                }
            }
            else {
                $resultObj = @{
                    Success = $false
                    Output = "Der Befehl wurde mit Fehlercode $($process.ExitCode) beendet."
                }
            }
        }
        catch {
            $resultObj = @{
                Success = $false
                Output = "Fehler beim Ausführen des Befehls: $($_.Exception.Message)"
            }
        }
    }
    else {
        # Skript braucht Admin-Rechte, starte einen neuen PowerShell-Prozess
        try {
            $logTextBox.AppendText("Starte neuen PowerShell-Prozess mit Admin-Rechten...`r`n")
            $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", "$Command" -Verb RunAs -PassThru -Wait
            
            if ($process.ExitCode -eq 0) {
                $resultObj = @{
                    Success = $true
                    Output = "Befehl erfolgreich als Admin ausgeführt."
                }
            }
            else {
                $resultObj = @{
                    Success = $false
                    Output = "Der Befehl wurde mit Fehlercode $($process.ExitCode) beendet."
                }
            }
        }
        catch {
            $resultObj = @{
                Success = $false
                Output = "Fehler beim Ausführen als Admin: $($_.Exception.Message)"
            }
        }
    }
    
    # Führe den Callback aus
    if ($OnComplete -ne $null) {
        $OnComplete.Invoke($resultObj)
    }
}

# Funktion zum Prüfen der Internetverbindung
function Test-InternetConnection {
    try {
        $response = Invoke-WebRequest -Uri "http://www.microsoft.com" -UseBasicParsing -TimeoutSec 5
        return $true
    }
    catch {
        return $false
    }
}

function Test-InternetConnection {
    try {
        $response = Invoke-WebRequest -Uri "http://www.microsoft.com" -UseBasicParsing -TimeoutSec 5
        return $true
    }
    catch {
        # Wenn keine Internetverbindung besteht und Warnung noch nicht angezeigt wurde
        if (-not $global:noInternetWarningShown) {
            [System.Windows.Forms.MessageBox]::Show(
                "Ohne Zeitstempel werden signierte Skripte nach Ablauf des Zertifikats unbrauchbar. Ein gesetzter Zeitstempel sorgt für eine permanente Nutzung.",
                "Warnung: Keine Internetverbindung",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $global:noInternetWarningShown = $true
        }
        return $false
    }
}

# Funktion zum Prüfen, ob ein Signierzertifikat existiert
function Get-SigningCertificate {
    # Use New-Object instead of static constructor
    try {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
            [System.Security.Cryptography.X509Certificates.StoreName]::My,
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
        )
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $store.Close()
    } catch {
        Write-Warning "Couldn't refresh certificate store: $($_.Exception.Message)"
    }
    
    # Get all code signing certificates
    $certs = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
    if ($certs.Count -eq 0) {
        return $null
    }
    else {
        return $certs
    }
}

# Funktion zum Signieren eines PowerShell-Skripts
function Sign-PowerShellScript {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,
        
        [Parameter(Mandatory=$true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        
        [Parameter(Mandatory=$false)]
        [bool]$UseTimestamp = $false,
        
        [Parameter(Mandatory=$false)]
        [string]$TimestampServer = "http://timestamp.digicert.com"
    )
    
    try {
        # Verify the file exists and is accessible
        if (-not (Test-Path -Path $ScriptPath -PathType Leaf)) {
            Write-Error "File not found or inaccessible: $ScriptPath"
            return $null
        }
        
        # Check if file is writable
        try {
            $fileInfo = [System.IO.FileInfo]::new($ScriptPath)
            $stream = $fileInfo.OpenWrite()
            $stream.Close()
        } catch {
            Write-Error "File is not writable: $ScriptPath. Error: $($_.Exception.Message)"
            return $null
        }
        
        # Attempt signing with detailed error reporting
        if ($UseTimestamp) {
            Write-Verbose "Signing with timestamp server: $TimestampServer"
            $signature = Set-AuthenticodeSignature -FilePath $ScriptPath -Certificate $Certificate -TimestampServer $TimestampServer -ErrorAction Stop
        } else {
            Write-Verbose "Signing without timestamp"
            $signature = Set-AuthenticodeSignature -FilePath $ScriptPath -Certificate $Certificate -ErrorAction Stop
        }
        
        # Verify the signature was created successfully
        if ($signature -eq $null) {
            Write-Error "Signature object is null"
            return $null
        }
        
        Write-Verbose "Signature status: $($signature.Status)"
        return $signature
    }
    catch {
        Write-Error "Exception during signing: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        return $null
    }
}

# PowerShell benötigt Runspaces für die asynchrone Ausführung
Add-Type -AssemblyName System.Management.Automation

# Erstellt das Hauptformulars
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell-Skript-Signierer"
$form.Size = New-Object System.Drawing.Size(800, 640)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Status-Panel erstellen (oben)
$statusPanel = New-Object System.Windows.Forms.GroupBox
$statusPanel.Text = "Status"
$statusPanel.Location = New-Object System.Drawing.Point(10, 10)
$statusPanel.Size = New-Object System.Drawing.Size(765, 120)
$form.Controls.Add($statusPanel)

# Internet-Status erstellen
$internetStatusLabel = New-Object System.Windows.Forms.Label
$internetStatusLabel.Location = New-Object System.Drawing.Point(10, 20)
$internetStatusLabel.Size = New-Object System.Drawing.Size(300, 20)
$internetStatusLabel.Text = "Prüfe Internetverbindung..."
$statusPanel.Controls.Add($internetStatusLabel)

# Aktualisieren-Button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(660, 80)
$refreshButton.Size = New-Object System.Drawing.Size(90, 25)
$refreshButton.Text = "Aktualisieren"
$refreshButton.BackColor = [System.Drawing.Color]::LightBlue
$statusPanel.Controls.Add($refreshButton)

# Zertifikat-Status erstellen
$certStatusLabel = New-Object System.Windows.Forms.Label
$certStatusLabel.Location = New-Object System.Drawing.Point(10, 50)
$certStatusLabel.Size = New-Object System.Drawing.Size(300, 40)
$certStatusLabel.Text = "Prüfe Signaturzertifikate..."
$certStatusLabel.AutoSize = $true
$statusPanel.Controls.Add($certStatusLabel)

# Zeitstempel-Checkbox erstellen
$timestampCheckBox = New-Object System.Windows.Forms.CheckBox
$timestampCheckBox.Location = New-Object System.Drawing.Point(430, 50)
$timestampCheckBox.Size = New-Object System.Drawing.Size(400, 20)
$timestampCheckBox.Text = "Mit Zeitstempel signieren (digicert.com)"
$timestampCheckBox.Enabled = $false  # Standardmäßig deaktiviert, wird später basierend auf Internetverbindung aktiviert
$statusPanel.Controls.Add($timestampCheckBox)

# Zertifikat-Auswahl-Label
$certSelectLabel = New-Object System.Windows.Forms.Label
$certSelectLabel.Location = New-Object System.Drawing.Point(310, 20)
$certSelectLabel.Size = New-Object System.Drawing.Size(120, 20)
$certSelectLabel.Text = "Zertifikat auswählen:"
$statusPanel.Controls.Add($certSelectLabel)

# Zertifikat-Auswahlfeld
$certComboBox = New-Object System.Windows.Forms.ComboBox
$certComboBox.Location = New-Object System.Drawing.Point(430, 18)
$certComboBox.Size = New-Object System.Drawing.Size(320, 20)
$certComboBox.DropDownStyle = "DropDownList"
$statusPanel.Controls.Add($certComboBox)

# Admin-Befehle Panel (neue Komponenten)
$adminCommandsLabel = New-Object System.Windows.Forms.Label
$adminCommandsLabel.Location = New-Object System.Drawing.Point(10, 80)
$adminCommandsLabel.Size = New-Object System.Drawing.Size(110, 20)
$adminCommandsLabel.Text = "Admin-Befehle:"
$statusPanel.Controls.Add($adminCommandsLabel)

# GPUpdate-Button
$gpUpdateButton = New-Object System.Windows.Forms.Button
$gpUpdateButton.Location = New-Object System.Drawing.Point(120, 80)
$gpUpdateButton.Size = New-Object System.Drawing.Size(130, 25)
$gpUpdateButton.Text = "GPUpdate"
$gpUpdateButton.BackColor = [System.Drawing.Color]::LightGreen
$statusPanel.Controls.Add($gpUpdateButton)

# CertUtil-Button
$certUtilButton = New-Object System.Windows.Forms.Button
$certUtilButton.Location = New-Object System.Drawing.Point(260, 80)
$certUtilButton.Size = New-Object System.Drawing.Size(140, 25)
$certUtilButton.Text = "CertUtil -pulse"
$certUtilButton.BackColor = [System.Drawing.Color]::LightGreen
$statusPanel.Controls.Add($certUtilButton)

# Admin-Status Label
$adminStatusLabel = New-Object System.Windows.Forms.Label
$adminStatusLabel.Location = New-Object System.Drawing.Point(430, 85)
$adminStatusLabel.Size = New-Object System.Drawing.Size(240, 25)
$adminStatusLabel.Text = "Admin-Status: Wird geprüft..."
$statusPanel.Controls.Add($adminStatusLabel)

# Datei-Auswahl-Panel (mitte)
$fileSelectionPanel = New-Object System.Windows.Forms.GroupBox
$fileSelectionPanel.Text = "Dateiauswahl"
$fileSelectionPanel.Location = New-Object System.Drawing.Point(10, 140)
$fileSelectionPanel.Size = New-Object System.Drawing.Size(765, 80)
$form.Controls.Add($fileSelectionPanel)

# Skript-Auswahl-Button
$selectButton = New-Object System.Windows.Forms.Button
$selectButton.Location = New-Object System.Drawing.Point(10, 20)
$selectButton.Size = New-Object System.Drawing.Size(200, 25)
$selectButton.Text = "Skript(e) auswählen"
$fileSelectionPanel.Controls.Add($selectButton)

# Skripte-Suchen-Button
$searchFolderButton = New-Object System.Windows.Forms.Button
$searchFolderButton.Location = New-Object System.Drawing.Point(10, 45)
$searchFolderButton.Size = New-Object System.Drawing.Size(200, 25)
$searchFolderButton.Text = "Ordner durchsuchen"
#$searchFolderButton.BackColor = [System.Drawing.Color]::LightBlue
$fileSelectionPanel.Controls.Add($searchFolderButton)

# Button zum Leeren der Skripte-Liste
$clearScriptsButton = New-Object System.Windows.Forms.Button
$clearScriptsButton.Location = New-Object System.Drawing.Point(220, 45)
$clearScriptsButton.Size = New-Object System.Drawing.Size(200, 25)
$clearScriptsButton.Text = "Ausgewählte Skripte leeren"
$fileSelectionPanel.Controls.Add($clearScriptsButton)

# Signieren-Button
$signButton = New-Object System.Windows.Forms.Button
$signButton.Location = New-Object System.Drawing.Point(220, 20)
$signButton.Size = New-Object System.Drawing.Size(200, 25)
$signButton.Text = "Ausgewählte Skripte signieren"
$signButton.Enabled = $false
$fileSelectionPanel.Controls.Add($signButton)

# Mit Notepad öffnen Button
$openWithNotepadButton = New-Object System.Windows.Forms.Button
$openWithNotepadButton.Location = New-Object System.Drawing.Point(430, 20)
$openWithNotepadButton.Size = New-Object System.Drawing.Size(150, 25)
$openWithNotepadButton.Text = "In Notepad öffnen"
$openWithNotepadButton.Enabled = $false
$fileSelectionPanel.Controls.Add($openWithNotepadButton)

# Dateilisten-Panel (unten)
$fileListPanel = New-Object System.Windows.Forms.GroupBox
$fileListPanel.Text = "Ausgewählte Skripte"
$fileListPanel.Location = New-Object System.Drawing.Point(10, 230)
$fileListPanel.Size = New-Object System.Drawing.Size(765, 150)
$form.Controls.Add($fileListPanel)

# ListBox für ausgewählte Dateien
$fileListBox = New-Object System.Windows.Forms.ListBox
$fileListBox.Location = New-Object System.Drawing.Point(10, 20)
$fileListBox.Size = New-Object System.Drawing.Size(745, 120)
$fileListBox.SelectionMode = "MultiExtended"
$fileListPanel.Controls.Add($fileListBox)

# Log-Panel (ganz unten)
$logPanel = New-Object System.Windows.Forms.GroupBox
$logPanel.Text = "Protokoll"
$logPanel.Location = New-Object System.Drawing.Point(10, 390)
$logPanel.Size = New-Object System.Drawing.Size(765, 180)
$form.Controls.Add($logPanel)

# Protokoll-TextBox
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 20)
$logTextBox.Size = New-Object System.Drawing.Size(745, 150)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = "Vertical"
$logTextBox.ReadOnly = $true
$logPanel.Controls.Add($logTextBox)

# Selbstsigniertes Zertifikat erstellen Button
$createCertButton = New-Object System.Windows.Forms.Button
$createCertButton.Location = New-Object System.Drawing.Point(590, 20)
$createCertButton.Size = New-Object System.Drawing.Size(160, 25)
$createCertButton.Text = "Zertifikat selbst erstellen"
$createCertButton.BackColor = [System.Drawing.Color]::LightYellow
$fileSelectionPanel.Controls.Add($createCertButton)

# Zertifikat löschen Buttons
$deleteCertButton = New-Object System.Windows.Forms.Button
$deleteCertButton.Location = New-Object System.Drawing.Point(590, 45)
$deleteCertButton.Size = New-Object System.Drawing.Size(160, 25)
$deleteCertButton.Text = "Zertifikat löschen"
$deleteCertButton.BackColor = [System.Drawing.Color]::LightCoral
$fileSelectionPanel.Controls.Add($deleteCertButton)

# Event-Handler für den "Zertifikat erstellen" Button
$createCertButton.Add_Click({
    # Dialog anzeigen
    $certParams = Show-CreateCertificateDialog
    
    if ($certParams -ne $null) {
        # Selbstsigniertes Zertifikat erstellen
        Create-SelfSignedCodeSigningCertificate -CertSubject $certParams.Subject -ValidityDays $certParams.ValidityDays
    }
})

# Event-Handler für den "Zertifikat löschen" Button
$deleteCertButton.Add_Click({
    # Prüfe, ob der Button bereits deaktiviert ist
    if (-not $deleteCertButton.Enabled) {
        $logTextBox.AppendText("Löschvorgang läuft bereits. Bitte warten...`r`n")
        return
    }
    
    $selectedCert = $certComboBox.SelectedItem
    
    if ($selectedCert -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle ein Zertifikat zum Löschen aus.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Speichere das Certificate-Objekt
    $certToDelete = $selectedCert.Certificate
    
    # Sicherheitsabfrage
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Möchtest du wirklich das folgende Zertifikat löschen?`n`n$($selectedCert.DisplayName)`n`nDiese Aktion kann nicht rückgängig gemacht werden!",
        "Zertifikat löschen bestätigen",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($confirmation -eq "Yes") {
        # Deaktiviere den Button während des Löschvorgangs
        $deleteCertButton.Enabled = $false
        $deleteCertButton.Text = "Wird gelöscht..."
        
        # Führe die Löschfunktion direkt aus (synchron, keine Callbacks)
        Remove-SelectedCertificate -Certificate $certToDelete
        
        # Hinweis: Die Button-Aktivierung erfolgt bereits in der Remove-SelectedCertificate Funktion
    }
})

# Funktion zum Aktualisieren des Admin-Status
function Update-AdminStatus {
    $isAdmin = Test-IsAdmin
    if ($isAdmin) {
        $adminStatusLabel.Text = "Admin-Status: ✓ Mit Admin-Rechten"
        $adminStatusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $adminStatusLabel.Text = "Admin-Status: ℹ️ Ohne Admin-Rechten"
        $adminStatusLabel.ForeColor = [System.Drawing.Color]::Blue
    }
}

# Event-Handler für GPUpdate-Button
$gpUpdateButton.Add_Click({
    # Deaktiviere den Button während der Ausführung
    $gpUpdateButton.Enabled = $false
    $gpUpdateButton.Text = "Wird ausgeführt..."
    
    $logTextBox.AppendText("--- GPUpdate gestartet $(Get-Date) ---`r`n")
    
    # Callback, der nach Abschluss des Befehls ausgeführt wird
    $onComplete = {
        param($result)
        
        # UI-Updates direkt ausführen, da wir nicht mehr in einem separaten Thread sind
        if ($result.Success) {
            $logTextBox.AppendText("Befehl: gpupdate /force`r`n")
            $logTextBox.AppendText("Ergebnis: Erfolgreich`r`n")
            if ($result.Output) {
                $logTextBox.AppendText("Ausgabe: $($result.Output)`r`n")
            }
        } else {
            $logTextBox.AppendText("Befehl: gpupdate /force`r`n")
            $logTextBox.AppendText("Ergebnis: Fehler`r`n")
            $logTextBox.AppendText("Fehlermeldung: $($result.Output)`r`n")
        }
        $logTextBox.AppendText("--- GPUpdate beendet $(Get-Date) ---`r`n`r`n")
        
        # Button wieder aktivieren
        $gpUpdateButton.Enabled = $true
        $gpUpdateButton.Text = "GPUpdate"
    }
    
    # Führe den Befehl aus
    Invoke-AdminCommand -Command "gpupdate /force" -OnComplete $onComplete
})

# Event-Handler für CertUtil-Button
$certUtilButton.Add_Click({
    # Deaktiviere den Button während der Ausführung
    $certUtilButton.Enabled = $false
    $certUtilButton.Text = "Wird ausgeführt..."
    
    $logTextBox.AppendText("--- CertUtil -pulse gestartet $(Get-Date) ---`r`n")
    
    # Callback, der nach Abschluss des Befehls ausgeführt wird
    $onComplete = {
        param($result)
        
        # UI-Updates direkt ausführen
        if ($result.Success) {
            $logTextBox.AppendText("Befehl: certutil -pulse`r`n")
            $logTextBox.AppendText("Ergebnis: Erfolgreich`r`n")
            if ($result.Output) {
                $logTextBox.AppendText("Ausgabe: $($result.Output)`r`n")
            }
        } else {
            $logTextBox.AppendText("Befehl: certutil -pulse`r`n")
            $logTextBox.AppendText("Ergebnis: Fehler`r`n")
            $logTextBox.AppendText("Fehlermeldung: $($result.Output)`r`n")
        }
        $logTextBox.AppendText("--- CertUtil -pulse beendet $(Get-Date) ---`r`n`r`n")
        
        # Button wieder aktivieren
        $certUtilButton.Enabled = $true
        $certUtilButton.Text = "CertUtil -pulse"
    }
    
    # Führe den Befehl aus
    Invoke-AdminCommand -Command "certutil -pulse" -OnComplete $onComplete
})

# Event-Handler für den "In Notepad öffnen" Button
$openWithNotepadButton.Add_Click({
    $selectedFiles = @()
    foreach ($index in $fileListBox.SelectedIndices) {
        $selectedFiles += $fileListBox.Items[$index]
    }
    
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle mindestens ein Skript zum Öffnen aus.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    foreach ($file in $selectedFiles) {
        $logTextBox.AppendText("Öffne $file in Notepad...`r`n")
        try {
            Start-Process -FilePath "notepad.exe" -ArgumentList $file
            $logTextBox.AppendText("Datei erfolgreich in Notepad geöffnet: $file`r`n")
        } catch {
            $logTextBox.AppendText("Fehler beim Öffnen der Datei in Notepad: $($_.Exception.Message)`r`n")
        }
    }
    
    $logTextBox.AppendText("`r`n")
})

# Funktion zum Aktualisieren der Status-Prüfungen
function Update-Status {

        if ($global:isCreatingCertificate) {
        $logTextBox.AppendText("Status-Update übersprungen - Zertifikatserstellung läuft...`r`n")
        return
    }
    # Internetverbindung prüfen
    $internetStatusLabel.Text = "Prüfe Internetverbindung..."
    $internetStatusLabel.ForeColor = [System.Drawing.Color]::Black
    $form.Refresh()
    
    $hasInternet = Test-InternetConnection
    if ($hasInternet) {
        $internetStatusLabel.Text = "✓ Internetverbindung verfügbar"
        $internetStatusLabel.ForeColor = [System.Drawing.Color]::Green
        # Aktiviere Zeitstempel-Option, da Internet verfügbar ist
        $timestampCheckBox.Enabled = $true
        $timestampCheckBox.Checked = $true  # Standardmäßig aktiviert
    }
    else {
        $internetStatusLabel.Text = "✗ Keine Internetverbindung"
        $internetStatusLabel.ForeColor = [System.Drawing.Color]::Red
        # Deaktiviere Zeitstempel-Option, da keine Internetverbindung besteht
        $timestampCheckBox.Enabled = $false
        $timestampCheckBox.Checked = $false
    }
    
    # Zertifikate prüfen
    $certStatusLabel.Text = "Prüfe Signaturzertifikate..."
    $certStatusLabel.ForeColor = [System.Drawing.Color]::Black
    $form.Refresh()
    
    # Zertifikate-Combobox leeren
    $certComboBox.Items.Clear()
    
    # Zertifikate erneut prüfen
    try {
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
            [System.Security.Cryptography.X509Certificates.StoreName]::My, 
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
        )
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $store.Close()
    } catch {
        $logTextBox.AppendText("Warnung: Konnte den Zertifikatspeicher nicht aktualisieren: $($_.Exception.Message)`r`n")
    }

    # Zertifikate erneut prüfen mit kurzer Verzögerung
    Start-Sleep -Milliseconds 500
    $certificates = Get-SigningCertificate
    
    if ($certificates -eq $null -or $certificates.Count -eq 0) {
        $certStatusLabel.Text = "✗ Keine Code-Signing-Zertifikate im CurrentUser-Speicher gefunden"
        $certStatusLabel.ForeColor = [System.Drawing.Color]::Red
    }
    else {
        $certStatusLabel.Text = "✓ $($certificates.Count) Code-Signing-Zertifikat(e) gefunden"
        $certStatusLabel.ForeColor = [System.Drawing.Color]::Green
        
        foreach ($cert in $certificates) {
            $item = New-Object PSObject -Property @{
                "Certificate" = $cert
                "DisplayName" = "$($cert.Subject) (gültig bis $($cert.NotAfter.ToString('dd.MM.yyyy')))"
            }
            $certComboBox.Items.Add($item)
        }
        
        if ($certComboBox.Items.Count -gt 0) {
            $certComboBox.DisplayMember = "DisplayName"
            $certComboBox.SelectedIndex = 0
        }
    }

    # Admin-Status aktualisieren
    Update-AdminStatus
    
    # Log-Eintrag zur Aktualisierung
    $logTextBox.AppendText("Status aktualisiert: $(Get-Date)`r`n")
    if ($hasInternet) {
        $logTextBox.AppendText("- Internetverbindung: Verfügbar`r`n")
    } else {
        $logTextBox.AppendText("- Internetverbindung: Nicht verfügbar`r`n")
    }
    
    if ($certificates -ne $null -and $certificates.Count -gt 0) {
        $logTextBox.AppendText("- $($certificates.Count) Code-Signing-Zertifikat(e) gefunden`r`n")
    } else {
        $logTextBox.AppendText("- Keine Code-Signing-Zertifikate gefunden`r`n")
    }
    
    $isAdmin = Test-IsAdmin
    if ($isAdmin) {
        $logTextBox.AppendText("- Admin-Status: Mit Admin-Rechten`r`n")
    } else {
        $logTextBox.AppendText("- Admin-Status: Ohne Admin-Rechten`r`n")
    }
    
    $logTextBox.AppendText("`r`n")
}

# Event-Handler für Dateiauswahl
$selectButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PowerShell-Skripte (*.ps1)|*.ps1|Alle Dateien (*.*)|*.*"
    $openFileDialog.Multiselect = $true
    $openFileDialog.Title = "PowerShell-Skript(e) zum Signieren auswählen"
    
    if ($openFileDialog.ShowDialog() -eq "OK") {
        foreach ($file in $openFileDialog.FileNames) {
            if (-not $fileListBox.Items.Contains($file)) {
                $fileListBox.Items.Add($file)
            }
        }
        
        if ($fileListBox.Items.Count -gt 0 -and $certComboBox.SelectedItem -ne $null) {
            $signButton.Enabled = $true
        }
        
        # Aktiviere den Notepad-Button, wenn Dateien ausgewählt sind
        if ($fileListBox.Items.Count -gt 0) {
            $openWithNotepadButton.Enabled = $true
        }
    }
})

# Event-Handler für den Ordner-Durchsuchen-Button
$searchFolderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Ordner mit PowerShell-Skripten auswählen"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $selectedFolder = $folderBrowser.SelectedPath
        $logTextBox.AppendText("Suche nach PowerShell-Skripten in: $selectedFolder`r`n")
        
        try {
            # Suche nach allen PS1-Dateien im ausgewählten Ordner
            $psFiles = Get-ChildItem -Path $selectedFolder -Filter "*.ps1" -File -Recurse -ErrorAction Stop
            
            if ($psFiles.Count -eq 0) {
                $logTextBox.AppendText("Keine PowerShell-Skripte im ausgewählten Ordner gefunden.`r`n")
                [System.Windows.Forms.MessageBox]::Show("Keine PowerShell-Skripte (*.ps1) im ausgewählten Ordner gefunden.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                $logTextBox.AppendText("$($psFiles.Count) PowerShell-Skript(e) gefunden.`r`n")
                
                # Frage den Benutzer, ob alle Skripte hinzugefügt werden sollen
                $addAllFiles = [System.Windows.Forms.MessageBox]::Show("$($psFiles.Count) PowerShell-Skript(e) gefunden. Alle zur Liste hinzufügen?", "Skripte gefunden", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($addAllFiles -eq "Yes") {
                    $addedCount = 0
                    
                    foreach ($file in $psFiles) {
                        $fullPath = $file.FullName
                        # Prüfe, ob die Datei bereits in der Liste ist
                        if (-not $fileListBox.Items.Contains($fullPath)) {
                            $fileListBox.Items.Add($fullPath)
                            $addedCount++
                        }
                    }
                    
                    $logTextBox.AppendText("$addedCount neue Skripte zur Liste hinzugefügt.`r`n")
                    
                    # Aktiviere den Signieren-Button, wenn Dateien hinzugefügt wurden und ein Zertifikat ausgewählt ist
                    if ($fileListBox.Items.Count -gt 0 -and $certComboBox.SelectedItem -ne $null) {
                        $signButton.Enabled = $true
                    }
                    
                    # Aktiviere den Notepad-Button, wenn Dateien in der Liste sind
                    if ($fileListBox.Items.Count -gt 0) {
                        $openWithNotepadButton.Enabled = $true
                    }
                } else {
                    $logTextBox.AppendText("Keine Skripte hinzugefügt.`r`n")
                }
            }
        } catch {
            $logTextBox.AppendText("Fehler beim Durchsuchen des Ordners: $($_.Exception.Message)`r`n")
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Durchsuchen des Ordners: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        
        $logTextBox.AppendText("`r`n")
    }
})

# Event-Handler für den "Ausgewählte Skripte leeren" Button
$clearScriptsButton.Add_Click({
    # Bestätigungsdialog anzeigen
    if ($fileListBox.Items.Count -gt 0) {
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie wirklich alle ausgewählten Skripte aus der Liste entfernen?",
            "Liste leeren bestätigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirmation -eq "Yes") {
            # Liste leeren
            $fileListBox.Items.Clear()
            
            # Log-Eintrag schreiben
            $logTextBox.AppendText("Liste der ausgewählten Skripte wurde geleert.`r`n`r`n")
            
            # Buttons deaktivieren, da keine Dateien mehr ausgewählt sind
            $signButton.Enabled = $false
            $openWithNotepadButton.Enabled = $false
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Die Liste ist bereits leer.",
            "Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# Event-Handler für Signatur-Button
$signButton.Add_Click({
    $selectedCert = $certComboBox.SelectedItem
    
    if ($selectedCert -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle ein Zertifikat aus.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $selectedFiles = @()
    foreach ($index in $fileListBox.SelectedIndices) {
        $selectedFiles += $fileListBox.Items[$index]
    }
    
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte wähle mindestens ein Skript zum Signieren aus.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Prüfen, ob mit Zeitstempel signiert werden soll
    $useTimestamp = $timestampCheckBox.Checked
    
    $logTextBox.AppendText("--- Signiervorgang gestartet $(Get-Date) ---`r`n")
    
    if ($useTimestamp) {
        $logTextBox.AppendText("Signierungsmodus: Mit Zeitstempel (timestamp.digicert.com)`r`n")
    } else {
        $logTextBox.AppendText("Signierungsmodus: Ohne Zeitstempel`r`n")
    }
    
    $logTextBox.AppendText("Verwende Zertifikat: $($selectedCert.DisplayName)`r`n")
    
    $signedFiles = @()
    foreach ($file in $selectedFiles) {
        $logTextBox.AppendText("Signiere $file...`r`n")
        
        # Check file access before attempting to sign
        try {
            if (-not (Test-Path -Path $file -PathType Leaf)) {
                $logTextBox.AppendText("Fehler: Datei nicht gefunden: $file`r`n")
                continue
            }
            
            # Test write permissions
            $fileInfo = [System.IO.FileInfo]::new($file)
            try {
                $stream = $fileInfo.OpenWrite()
                $stream.Close()
            } catch {
                $logTextBox.AppendText("Fehler: Keine Schreibberechtigung für die Datei: $file`r`n")
                $logTextBox.AppendText("Fehlermeldung: $($_.Exception.Message)`r`n")
                continue
            }
        } catch {
            $logTextBox.AppendText("Fehler beim Zugriff auf die Datei: $file`r`n")
            $logTextBox.AppendText("Fehlermeldung: $($_.Exception.Message)`r`n")
            continue
        }
        
        # Attempt to sign file with verbose error logging
        $signature = Sign-PowerShellScript -ScriptPath $file -Certificate $selectedCert.Certificate -UseTimestamp $useTimestamp
        
        if ($signature -ne $null) {
            $logTextBox.AppendText("Signaturstatus: $($signature.Status)`r`n")
            
            if ($signature.Status -eq "Valid") {
                $logTextBox.AppendText("Signatur erfolgreich: $file`r`n")
                $signedFiles += $file
                
                # Zusätzliche Informationen zum Zeitstempel, falls verwendet
                if ($useTimestamp -and $signature.TimeStamperCertificate -ne $null) {
                    $logTextBox.AppendText("  Zeitstempel: $($signature.TimeStamperCertificate.GetName())`r`n")
                    $logTextBox.AppendText("  Zeitstempel-Datum: $($signature.TimeStampInfo.Timestamp)`r`n")
                }
            }
            else {
                $logTextBox.AppendText("Signatur fehlgeschlagen: $file (Status: $($signature.Status))`r`n")
                $logTextBox.AppendText("Signaturdetails: $($signature | ConvertTo-Json -Depth 1)`r`n")
            }
        }
        else {
            $logTextBox.AppendText("Fehler beim Signieren: $file (Keine Signatur zurückgegeben)`r`n")
        }
    }
    
    # Aktiviere den Notepad-Button, wenn Dateien erfolgreich signiert wurden
    if ($signedFiles.Count -gt 0) {
        $openWithNotepadButton.Enabled = $true
    }
    
    $logTextBox.AppendText("--- Signiervorgang beendet $(Get-Date) ---`r`n`r`n")
})

# Event-Handler für Zertifikatsauswahl
$certComboBox.Add_SelectedIndexChanged({
    if ($certComboBox.SelectedItem -ne $null -and $fileListBox.Items.Count -gt 0) {
        $signButton.Enabled = $true
    }
    else {
        $signButton.Enabled = $false
    }
})

# Event-Handler für FileListBox SelectedIndexChanged
$fileListBox.Add_SelectedIndexChanged({
    # Aktiviere den Notepad-Button, wenn mindestens eine Datei ausgewählt ist
    if ($fileListBox.SelectedIndices.Count -gt 0) {
        $openWithNotepadButton.Enabled = $true
    } else {
        $openWithNotepadButton.Enabled = $false
    }
})

# Event-Handler für Aktualisieren-Button
$refreshButton.Add_Click({
    Update-Status
})

# Status-Prüfungen durchführen beim Laden des Formulars
$form.Add_Shown({
    Update-Status
})

# Funktion zum Erstellen eines selbstsignierten Code-Signing-Zertifikats
function Create-SelfSignedCodeSigningCertificate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CertSubject,

        [Parameter(Mandatory=$false)]
        [int]$ValidityDays = 365
    )

    # Prevent multiple instances
    if ($global:isCreatingCertificate) {
        $logTextBox.AppendText("Eine Zertifikatserstellung läuft bereits. Bitte warten...`r`n")
        return
    }

    $global:isCreatingCertificate = $true
    $global:certificateCreationCompleted = $false
    $global:lastCertificateOperation = Get-Date

    $logTextBox.AppendText("--- Selbstsigniertes Zertifikat wird erstellt $(Get-Date) ---`r`n")
    $logTextBox.AppendText("Betreff: $CertSubject`r`n")
    $logTextBox.AppendText("Gültigkeitsdauer: $ValidityDays Tage`r`n")

    # Define the script block for certificate creation
    $scriptBlock = {
        param ($CertSubject, $ValidityDays)

        try {
            $cert = New-SelfSignedCertificate -Subject $CertSubject -CertStoreLocation Cert:\CurrentUser\My -Type CodeSigningCert -KeyUsage DigitalSignature -KeyAlgorithm RSA -KeyLength 2048 -NotAfter (Get-Date).AddDays($ValidityDays)

            if ($cert -eq $null) {
                throw "Zertifikat konnte nicht erstellt werden."
            }

            Write-Output "INFO: Thumbprint: $($cert.Thumbprint)"

            $certPath = Join-Path -Path $env:TEMP -ChildPath "TempCert_$($cert.Thumbprint).cer"

            $null = Export-Certificate -Cert $cert -FilePath $certPath -Force -ErrorAction Stop
            Write-Output "INFO: Exportiert: $certPath"

            $null = Import-Certificate -FilePath $certPath -CertStoreLocation Cert:\LocalMachine\Root -ErrorAction Stop
            Write-Output "INFO: ImportRootStore: OK"

            Remove-Item -Path $certPath -Force -ErrorAction Stop
            Write-Output "INFO: TempDeleted: OK"
        }
        catch {
            Write-Output "ERROR: $($_.Exception.Message)"
        }
    }

    # Start the job to create the certificate
    $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $CertSubject, $ValidityDays

    # Wait for the job to complete
    Wait-Job -Job $job

    # Receive the job output
    $jobOutput = Receive-Job -Job $job -ErrorAction Stop

    # Remove the job
    Remove-Job -Job $job

    # Process the job output
    foreach ($line in $jobOutput) {
        if ($line -match "^INFO:(.+)$") {
            $logTextBox.AppendText("INFO: $($matches[1])`r`n")
        }
        elseif ($line -match "^ERROR:(.+)$") {
            $logTextBox.AppendText("ERROR: $($matches[1])`r`n")
        }
    }

    if ($job.State -eq "Failed") {
        $logTextBox.AppendText("❌ Fehler bei der Zertifikatserstellung.`r`n")
        $logTextBox.AppendText("--- Zertifikatserstellung fehlgeschlagen $(Get-Date) ---`r`n`r`n")

        # Setze Flags, um anzuzeigen, dass die Erstellung fehlgeschlagen ist
        $global:certificateCreationCompleted = $true
        $global:lastCertificateOperation = Get-Date
        $global:isCreatingCertificate = $false
    } else {
        $logTextBox.AppendText("--- Zertifikatserstellung beendet $(Get-Date) ---`r`n`r`n")

        # Setze Flags, um anzuzeigen, dass die Erstellung abgeschlossen ist
        $global:certificateCreationCompleted = $true
        $global:lastCertificateOperation = Get-Date
        $global:isCreatingCertificate = $false

        # Aktualisiere die Zertifikatsliste nach kurzer Verzögerung
        Start-Sleep -Milliseconds 1000
        Update-Status
    }
}

# Funktion zum Löschen eines Zertifikats
function Remove-SelectedCertificate {
    param (
        [Parameter(Mandatory=$true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    
    # Hole Informationen über das Zertifikat
    $thumbprint = $Certificate.Thumbprint
    $subject = $Certificate.Subject
    
    $logTextBox.AppendText("--- Zertifikat-Löschvorgang gestartet $(Get-Date) ---`r`n")
    $logTextBox.AppendText("Lösche Zertifikat: $subject (Thumbprint: $thumbprint)`r`n")
    
    # Direkter Ansatz zum Löschen des Zertifikats
    try {
        # Zuerst aus Current User\My löschen
        $logTextBox.AppendText("Lösche aus CurrentUser\My...`r`n")
        $store = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store -ArgumentList 'My', 'CurrentUser'
        $store.Open('ReadWrite')
        
        $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint } | Select-Object -First 1
        
        if ($certToRemove) {
            $store.Remove($certToRemove)
            $logTextBox.AppendText("✓ Zertifikat erfolgreich aus dem Benutzerspeicher (CurrentUser\My) gelöscht.`r`n")
        } else {
            $logTextBox.AppendText("ℹ️ Zertifikat nicht im Benutzerspeicher (CurrentUser\My) gefunden.`r`n")
        }
        
        $store.Close()
        
        # Dann aus Trusted Root löschen (erfordert Admin-Rechte)
        $logTextBox.AppendText("Lösche aus LocalMachine\Root...`r`n")
        
        # PowerShell-Befehl für den Admin-Teil (nur für Trusted Root) - in temporäre Datei schreiben
        $tempScriptFile = [System.IO.Path]::GetTempFileName() + ".ps1"
        
        $scriptContent = @"
try {
    `$storeRoot = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store -ArgumentList 'Root', 'LocalMachine'
    `$storeRoot.Open('ReadWrite')
    
    `$rootCertToRemove = `$storeRoot.Certificates | Where-Object { `$_.Thumbprint -eq '$thumbprint' } | Select-Object -First 1
    
    if (`$rootCertToRemove) {
        `$storeRoot.Remove(`$rootCertToRemove)
        exit 0  # Erfolg
    } else {
        exit 2  # Zertifikat nicht gefunden
    }
    
    `$storeRoot.Close()
} catch {
    exit 1  # Fehler
}
"@
        
        # Schreibe den Code in die temporäre Datei
        [System.IO.File]::WriteAllText($tempScriptFile, $scriptContent)
        
        # Führe das Skript mit Admin-Rechten aus
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$tempScriptFile`"" -Verb RunAs -PassThru -Wait -WindowStyle Hidden
        
        # Lösche temporäre Skriptdatei
        if (Test-Path $tempScriptFile) {
            Remove-Item -Path $tempScriptFile -Force
        }
        
        # Überprüfe Exit-Code
        switch ($process.ExitCode) {
            0 { $logTextBox.AppendText("✓ Zertifikat erfolgreich aus Vertrauenswürdige Stammzertifizierungsstellen gelöscht.`r`n") }
            1 { $logTextBox.AppendText("⚠️ Fehler beim Löschen aus Vertrauenswürdige Stammzertifizierungsstellen.`r`n") }
            2 { $logTextBox.AppendText("ℹ️ Zertifikat nicht in Vertrauenswürdige Stammzertifizierungsstellen gefunden.`r`n") }
            default { $logTextBox.AppendText("⚠️ Unbekannter Fehler beim Löschen aus Vertrauenswürdige Stammzertifizierungsstellen (Code: $($process.ExitCode)).`r`n") }
        }
        
        # Manuelle Aktualisierung der Zertifikatsliste
        $logTextBox.AppendText("Aktualisiere Zertifikatsliste...`r`n")
        
        # Kurze Verzögerung für OS-Aktualisierung
        Start-Sleep -Milliseconds 1000
        
        # Zertifikatsliste aktualisieren
        $certComboBox.Items.Clear()
        
        # Zertifikatspeicher neu laden
        $certificates = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue
        
        if ($certificates -eq $null -or $certificates.Count -eq 0) {
            $certStatusLabel.Text = "✗ Keine Code-Signing-Zertifikate im CurrentUser-Speicher gefunden"
            $certStatusLabel.ForeColor = [System.Drawing.Color]::Red
        }
        else {
            $certStatusLabel.Text = "✓ $($certificates.Count) Code-Signing-Zertifikat(e) gefunden"
            $certStatusLabel.ForeColor = [System.Drawing.Color]::Green
            
            foreach ($cert in $certificates) {
                $item = New-Object PSObject -Property @{
                    "Certificate" = $cert
                    "DisplayName" = "$($cert.Subject) (gültig bis $($cert.NotAfter.ToString('dd.MM.yyyy')))"
                }
                $certComboBox.Items.Add($item)
            }
            
            if ($certComboBox.Items.Count -gt 0) {
                $certComboBox.DisplayMember = "DisplayName"
                $certComboBox.SelectedIndex = 0
            }
        }
        
        $logTextBox.AppendText("✓ Zertifikatsliste wurde aktualisiert.`r`n")
        $logTextBox.AppendText("--- Zertifikat-Löschvorgang erfolgreich beendet $(Get-Date) ---`r`n`r`n")
        
        # Aktiviere den Button wieder
        $deleteCertButton.Enabled = $true
        $deleteCertButton.Text = "Zertifikat löschen"
        
        return $true
    }
    catch {
        $logTextBox.AppendText("❌ Fehler beim Löschen des Zertifikats: $($_.Exception.Message)`r`n")
        $logTextBox.AppendText("--- Zertifikat-Löschvorgang fehlgeschlagen $(Get-Date) ---`r`n`r`n")
        
        # Aktiviere den Button wieder
        $deleteCertButton.Enabled = $true
        $deleteCertButton.Text = "Zertifikat löschen"
        
        return $false
    }
}

# Timer for certificate refresh check
$certRefreshTimer = New-Object System.Windows.Forms.Timer
$certRefreshTimer.Interval = 3000  # 3 seconds
$certRefreshTimer.Add_Tick({
    if ($global:certificateCreationCompleted) {
        # Only refresh once, then reset the flag
        $global:certificateCreationCompleted = $false
        
        # Ensure enough time has passed since the last operation
        $timeSinceLastOp = (Get-Date) - $global:lastCertificateOperation
        if ($timeSinceLastOp.TotalSeconds -ge 2) {
            $logTextBox.AppendText("Aktualisiere Zertifikatsliste nach Erstellung...`r`n")
            
            # Force certificate store refresh
            try {
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
                    [System.Security.Cryptography.X509Certificates.StoreName]::My, 
                    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
                )
                $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
                $store.Close()
                
                # Clear and reload certificates
                $certificates = Get-SigningCertificate
                if ($certificates -ne $null -and $certificates.Count -gt 0) {
                    $certComboBox.Items.Clear()
                    foreach ($cert in $certificates) {
                        $item = New-Object PSObject -Property @{
                            "Certificate" = $cert
                            "DisplayName" = "$($cert.Subject) (gültig bis $($cert.NotAfter.ToString('dd.MM.yyyy')))"
                        }
                        $certComboBox.Items.Add($item)
                    }
                    if ($certComboBox.Items.Count -gt 0) {
                        $certComboBox.DisplayMember = "DisplayName"
                        $certComboBox.SelectedIndex = 0
                    }
                    $certStatusLabel.Text = "✓ $($certificates.Count) Code-Signing-Zertifikat(e) gefunden"
                    $certStatusLabel.ForeColor = [System.Drawing.Color]::Green
                }
                $logTextBox.AppendText("Zertifikatsliste wurde aktualisiert.`r`n`r`n")
            } catch {
                $logTextBox.AppendText("Fehler bei der Aktualisierung der Zertifikatsliste: $($_.Exception.Message)`r`n`r`n")
            }
        }
    }
})

# Funktion zum Anzeigen des Zertifikat-Dialog-Fensters
function Show-CreateCertificateDialog {
    $certDialog = New-Object System.Windows.Forms.Form
    $certDialog.Text = "Zertifikat selbst erstellen"
    $certDialog.Size = New-Object System.Drawing.Size(450, 250)
    $certDialog.StartPosition = "CenterParent"
    $certDialog.FormBorderStyle = "FixedDialog"
    $certDialog.MaximizeBox = $false
    $certDialog.MinimizeBox = $false
    
    # Beschreibung
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Location = New-Object System.Drawing.Point(10, 10)
    $descLabel.Size = New-Object System.Drawing.Size(420, 40)
    $descLabel.Text = "Dieses Tool erstellt ein selbstsigniertes Code-Signing-Zertifikat für Testzwecke. Im Produktiveinsatz sollte ein offizielles oder eins über die interne CA signiertes Zertifikat verwendet werden."
    $certDialog.Controls.Add($descLabel)
    
    # Zertifikatsname-Label
    $subjectLabel = New-Object System.Windows.Forms.Label
    $subjectLabel.Location = New-Object System.Drawing.Point(10, 60)
    $subjectLabel.Size = New-Object System.Drawing.Size(100, 20)
    $subjectLabel.Text = "Zertifikatsname:"
    $certDialog.Controls.Add($subjectLabel)
    
    # Zertifikatsname-Textfeld
    $subjectTextBox = New-Object System.Windows.Forms.TextBox
    $subjectTextBox.Location = New-Object System.Drawing.Point(120, 57)
    $subjectTextBox.Size = New-Object System.Drawing.Size(310, 20)
    $subjectTextBox.Text = "CN=PowerShell Code Signing"
    $certDialog.Controls.Add($subjectTextBox)
    
    # Gültigkeitsdauer-Label
    $validityLabel = New-Object System.Windows.Forms.Label
    $validityLabel.Location = New-Object System.Drawing.Point(10, 90)
    $validityLabel.Size = New-Object System.Drawing.Size(100, 20)
    $validityLabel.Text = "Gültig (Tage):"
    $certDialog.Controls.Add($validityLabel)
    
    # Gültigkeitsdauer-Textfeld
    $validityTextBox = New-Object System.Windows.Forms.TextBox
    $validityTextBox.Location = New-Object System.Drawing.Point(120, 87)
    $validityTextBox.Size = New-Object System.Drawing.Size(100, 20)
    $validityTextBox.Text = "365"
    $certDialog.Controls.Add($validityTextBox)
    
    # Warnhinweis
    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Location = New-Object System.Drawing.Point(10, 120)
    $warningLabel.Size = New-Object System.Drawing.Size(420, 40)
    $warningLabel.Text = "Hinweis: Selbstsignierte Zertifikate werden nicht von Windows vertraut. Sie können nur zum Testen verwendet werden."
    $warningLabel.ForeColor = [System.Drawing.Color]::Red
    $certDialog.Controls.Add($warningLabel)
    
    # Erstellen-Button
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Location = New-Object System.Drawing.Point(230, 170)
    $createButton.Size = New-Object System.Drawing.Size(100, 30)
    $createButton.Text = "Erstellen"
    $createButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $certDialog.Controls.Add($createButton)
    $certDialog.AcceptButton = $createButton
    
    # Abbrechen-Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(340, 170)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Abbrechen"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $certDialog.Controls.Add($cancelButton)
    $certDialog.CancelButton = $cancelButton
    
    # Event-Handler für Erstellen-Button
$createButton.Add_Click({
    # Prüfen, ob ein gültiger Zertifikatsname angegeben wurde
    if ([string]::IsNullOrWhiteSpace($subjectTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gib einen gültigen Zertifikatsnamen ein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $certDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
        return
    }
    
    # Prüfen, ob eine gültige Zahl für die Gültigkeitsdauer angegeben wurde
    $validityDays = 0
    if (-not [int]::TryParse($validityTextBox.Text, [ref]$validityDays)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte gib eine gültige Zahl für die Gültigkeitsdauer ein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $certDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
        return
    }
    
    if ($validityDays -le 0) {
        [System.Windows.Forms.MessageBox]::Show("Die Gültigkeitsdauer muss größer als 0 sein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $certDialog.DialogResult = [System.Windows.Forms.DialogResult]::None
        return
    }
})
    
    # Dialog anzeigen
    $result = $certDialog.ShowDialog()
    
    # Rückgabewerte, wenn OK gedrückt wurde
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return @{
            Subject = $subjectTextBox.Text
            ValidityDays = [int]::Parse($validityTextBox.Text)
        }
    }
    
    return $null
}

# Add a timer to check for certificate refresh
$refreshTimer = New-Object System.Windows.Forms.Timer
$refreshTimer.Interval = 2000  # Check every 2 seconds
$refreshTimer.Add_Tick({
    if ($script:needCertRefresh) {
        $script:needCertRefresh = $false
        # Force certificate store refresh
        try {
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
                [System.Security.Cryptography.X509Certificates.StoreName]::My, 
                [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
            )
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            $store.Close()
            
            # Now do a full UI refresh
            Update-Status
        } catch {
            $logTextBox.AppendText("Fehler beim Aktualisieren des Zertifikatspeichers: $($_.Exception.Message)`r`n")
        }
    }
})

# Initialize the flags
$script:isCreatingCert = $false
$script:needCertRefresh = $false

# Start the timer
$certRefreshTimer.Start()

# Copyright-Label hinzufügen
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Location = New-Object System.Drawing.Point(10, 580)
$copyrightLabel.Size = New-Object System.Drawing.Size(765, 20)
$copyrightLabel.Text = "© 2025 Jörn Walter - https://www.der-windows-papst.de"
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$copyrightLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($copyrightLabel)

# Form anzeigen
[void]$form.ShowDialog()