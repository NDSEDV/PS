<#
.SYNOPSIS
  AD Dashboard
.DESCRIPTION
  SMIME - PEMPEM to PFX Converter
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.0
  Author:         Jörn Walter
  Creation Date:  2025-03-13
  Purpose/Change: Initial script development

  Jörn Walter. All rights reserved.
#>

# Admin Funktion
function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    return (New-Object Security.Principal.WindowsPrincipal($currentUser)).IsInRole($adminRole)
}

if (-not (Test-Admin)) {
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Registry-Schlüssel für die Anwendung
$regPath = "HKCU:\Software\PEMtoPFXConverter"
$regKeyOpenSSL = "OpenSSLPath"

# Hauptformular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "PEM zu PFX Konverter / CSR Generator"
$form.Size = New-Object System.Drawing.Size(680, 530)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# TabControl erstellen
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(650, 450)

# Tab 1: PEM zu PFX
$tabPemToPfx = New-Object System.Windows.Forms.TabPage
$tabPemToPfx.Text = "PEM zu PFX"

# Tab 2: CSR-Generator für SMIME
$tabCsr = New-Object System.Windows.Forms.TabPage
$tabCsr.Text = "CSR Generator (SMIME)"

# OpenSSL-Pfad versuchen aus Registry zu laden
$openSslPath = ""
try {
    if (Test-Path $regPath) {
        $savedOpenSSLPath = Get-ItemProperty -Path $regPath -Name $regKeyOpenSSL -ErrorAction SilentlyContinue
        if ($savedOpenSSLPath -and $savedOpenSSLPath.$regKeyOpenSSL) {
            $openSslPath = $savedOpenSSLPath.$regKeyOpenSSL
        }
    }
} catch {
    # Ignorieren, wenn Registry-Eintrag nicht existiert
}

#------------------------------------
# Tab 1: PEM zu PFX Konverter
#------------------------------------

# OpenSSL-Pfad
$lblOpenSSL = New-Object System.Windows.Forms.Label
$lblOpenSSL.Location = New-Object System.Drawing.Point(20, 20)
$lblOpenSSL.Size = New-Object System.Drawing.Size(150, 20)
$lblOpenSSL.Text = "OpenSSL-Pfad:"

$txtOpenSSL = New-Object System.Windows.Forms.TextBox
$txtOpenSSL.Location = New-Object System.Drawing.Point(170, 20)
$txtOpenSSL.Size = New-Object System.Drawing.Size(350, 20)
$txtOpenSSL.Text = $openSslPath

$btnBrowseOpenSSL = New-Object System.Windows.Forms.Button
$btnBrowseOpenSSL.Location = New-Object System.Drawing.Point(530, 20)
$btnBrowseOpenSSL.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowseOpenSSL.Text = "Durchsuchen"
$btnBrowseOpenSSL.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "OpenSSL Executable (openssl.exe)|openssl.exe"
    $openFileDialog.Title = "OpenSSL-Pfad auswählen"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtOpenSSL.Text = $openFileDialog.FileName
        $txtOpenSSLCsr.Text = $openFileDialog.FileName
        
        # OpenSSL-Pfad in Registry speichern
        try {
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            Set-ItemProperty -Path $regPath -Name $regKeyOpenSSL -Value $txtOpenSSL.Text
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Registry-Pfad konnte nicht gespeichert werden: $_", "Warnung")
        }
    }
})

# Public-Key Datei
$lblPublicKey = New-Object System.Windows.Forms.Label
$lblPublicKey.Location = New-Object System.Drawing.Point(20, 60)
$lblPublicKey.Size = New-Object System.Drawing.Size(150, 20)
$lblPublicKey.Text = "Public-Key (PEM):"

$txtPublicKey = New-Object System.Windows.Forms.TextBox
$txtPublicKey.Location = New-Object System.Drawing.Point(170, 60)
$txtPublicKey.Size = New-Object System.Drawing.Size(350, 20)

$btnBrowsePublicKey = New-Object System.Windows.Forms.Button
$btnBrowsePublicKey.Location = New-Object System.Drawing.Point(530, 60)
$btnBrowsePublicKey.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowsePublicKey.Text = "Durchsuchen"
$btnBrowsePublicKey.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PEM-Dateien (*.pem;*.crt;*.cer)|*.pem;*.crt;*.cer|Alle Dateien (*.*)|*.*"
    $openFileDialog.Title = "Public-Key PEM-Datei auswählen"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtPublicKey.Text = $openFileDialog.FileName
    }
})

# Private-Key Datei
$lblPrivateKey = New-Object System.Windows.Forms.Label
$lblPrivateKey.Location = New-Object System.Drawing.Point(20, 100)
$lblPrivateKey.Size = New-Object System.Drawing.Size(150, 20)
$lblPrivateKey.Text = "Private-Key (PEM):"

$txtPrivateKey = New-Object System.Windows.Forms.TextBox
$txtPrivateKey.Location = New-Object System.Drawing.Point(170, 100)
$txtPrivateKey.Size = New-Object System.Drawing.Size(350, 20)

$btnBrowsePrivateKey = New-Object System.Windows.Forms.Button
$btnBrowsePrivateKey.Location = New-Object System.Drawing.Point(530, 100)
$btnBrowsePrivateKey.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowsePrivateKey.Text = "Durchsuchen"
$btnBrowsePrivateKey.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "PEM-Dateien (*.pem;*.key)|*.pem;*.key|Alle Dateien (*.*)|*.*"
    $openFileDialog.Title = "Private-Key PEM-Datei auswählen"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtPrivateKey.Text = $openFileDialog.FileName
    }
})

# Ausgabe-PFX-Datei
$lblOutputFile = New-Object System.Windows.Forms.Label
$lblOutputFile.Location = New-Object System.Drawing.Point(20, 140)
$lblOutputFile.Size = New-Object System.Drawing.Size(150, 20)
$lblOutputFile.Text = "Ausgabe-PFX-Datei:"

$txtOutputFile = New-Object System.Windows.Forms.TextBox
$txtOutputFile.Location = New-Object System.Drawing.Point(170, 140)
$txtOutputFile.Size = New-Object System.Drawing.Size(350, 20)

$btnBrowseOutputFile = New-Object System.Windows.Forms.Button
$btnBrowseOutputFile.Location = New-Object System.Drawing.Point(530, 140)
$btnBrowseOutputFile.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowseOutputFile.Text = "Durchsuchen"
$btnBrowseOutputFile.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "PFX-Dateien (*.pfx)|*.pfx|Alle Dateien (*.*)|*.*"
    $saveFileDialog.Title = "Speicherort für PFX-Datei auswählen"
    $saveFileDialog.DefaultExt = "pfx"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $txtOutputFile.Text = $saveFileDialog.FileName
    }
})

# Passwort für PFX
$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Location = New-Object System.Drawing.Point(20, 180)
$lblPassword.Size = New-Object System.Drawing.Size(150, 20)
$lblPassword.Text = "PFX-Passwort:"

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(170, 180)
$txtPassword.Size = New-Object System.Drawing.Size(350, 20)
$txtPassword.PasswordChar = '*'

# Status-Label für PFX
$lblStatusPfx = New-Object System.Windows.Forms.Label
$lblStatusPfx.Location = New-Object System.Drawing.Point(20, 260)
$lblStatusPfx.Size = New-Object System.Drawing.Size(610, 80)
$lblStatusPfx.Text = "Bereit."

# Konvertieren-Button
$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Location = New-Object System.Drawing.Point(170, 220)
$btnConvert.Size = New-Object System.Drawing.Size(120, 30)
$btnConvert.Text = "Konvertieren"
$btnConvert.Add_Click({
    # Eingaben überprüfen
    if (-not (Test-Path $txtOpenSSL.Text)) {
        $lblStatusPfx.Text = "Fehler: OpenSSL-Pfad nicht gefunden."
        return
    }
    if (-not (Test-Path $txtPublicKey.Text)) {
        $lblStatusPfx.Text = "Fehler: Public-Key-Datei nicht gefunden."
        return
    }
    if (-not (Test-Path $txtPrivateKey.Text)) {
        $lblStatusPfx.Text = "Fehler: Private-Key-Datei nicht gefunden."
        return
    }
    if ([string]::IsNullOrEmpty($txtOutputFile.Text)) {
        $lblStatusPfx.Text = "Fehler: Bitte Ausgabedatei angeben."
        return
    }
    if ([string]::IsNullOrEmpty($txtPassword.Text)) {
        $lblStatusPfx.Text = "Fehler: Bitte ein Passwort für die PFX-Datei angeben."
        return
    }

    # OpenSSL-Pfad in Registry speichern
    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name $regKeyOpenSSL -Value $txtOpenSSL.Text
    } catch {
        $lblStatusPfx.Text = "Warnung: Registry-Pfad konnte nicht gespeichert werden: $_"
    }

    # Temporäre Datei für Passwort erstellen
    $tempPassFile = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($tempPassFile, $txtPassword.Text)

    try {
        # OpenSSL-Befehl zusammenstellen
        $opensslPath = $txtOpenSSL.Text
        $publicKeyPath = $txtPublicKey.Text
        $privateKeyPath = $txtPrivateKey.Text
        $outputPath = $txtOutputFile.Text

        # PKCS12-Konvertierung mit OpenSSL durchführen - einfachere Methode
        $cmd = "`"$opensslPath`" pkcs12 -export -out `"$outputPath`" -inkey `"$privateKeyPath`" -in `"$publicKeyPath`" -passout file:`"$tempPassFile`""
        
        # Prozess starten und Ausgabe erfassen
        $result = cmd /c "$cmd 2>&1"
        
        # Ergebnis überprüfen (kein ExitCode, stattdessen Ausgabe prüfen)
        if (Test-Path $outputPath) {
            $lblStatusPfx.Text = "PFX-Konvertierung erfolgreich abgeschlossen.`nAusgabedatei: $outputPath"
        } else {
            $errorOutput = $result -join "`n"
            $lblStatusPfx.Text = "Fehler bei der Konvertierung:`n$errorOutput"
        }
    } catch {
        $lblStatusPfx.Text = "Fehler bei der Ausführung: $_"
    } finally {
        # Temporäre Datei aufräumen
        if (Test-Path $tempPassFile) {
            Remove-Item $tempPassFile -Force
        }
    }
})

#------------------------------------
# Tab 2: CSR Generator für SMIME
#------------------------------------

# OpenSSL-Pfad für CSR Tab
$lblOpenSSLCsr = New-Object System.Windows.Forms.Label
$lblOpenSSLCsr.Location = New-Object System.Drawing.Point(20, 20)
$lblOpenSSLCsr.Size = New-Object System.Drawing.Size(150, 20)
$lblOpenSSLCsr.Text = "OpenSSL-Pfad:"

$txtOpenSSLCsr = New-Object System.Windows.Forms.TextBox
$txtOpenSSLCsr.Location = New-Object System.Drawing.Point(170, 20)
$txtOpenSSLCsr.Size = New-Object System.Drawing.Size(350, 20)
$txtOpenSSLCsr.Text = $openSslPath

$btnBrowseOpenSSLCsr = New-Object System.Windows.Forms.Button
$btnBrowseOpenSSLCsr.Location = New-Object System.Drawing.Point(530, 20)
$btnBrowseOpenSSLCsr.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowseOpenSSLCsr.Text = "Durchsuchen"
$btnBrowseOpenSSLCsr.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "OpenSSL Executable (openssl.exe)|openssl.exe"
    $openFileDialog.Title = "OpenSSL-Pfad auswählen"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $txtOpenSSLCsr.Text = $openFileDialog.FileName
        $txtOpenSSL.Text = $openFileDialog.FileName

        # OpenSSL-Pfad in Registry speichern
        try {
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
            }
            Set-ItemProperty -Path $regPath -Name $regKeyOpenSSL -Value $txtOpenSSLCsr.Text
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Registry-Pfad konnte nicht gespeichert werden: $_", "Warnung")
        }
    }
})

# SMIME Felder
$lblEmailAddress = New-Object System.Windows.Forms.Label
$lblEmailAddress.Location = New-Object System.Drawing.Point(20, 60)
$lblEmailAddress.Size = New-Object System.Drawing.Size(150, 20)
$lblEmailAddress.Text = "E-Mail-Adresse:"

$txtEmailAddress = New-Object System.Windows.Forms.TextBox
$txtEmailAddress.Location = New-Object System.Drawing.Point(170, 60)
$txtEmailAddress.Size = New-Object System.Drawing.Size(350, 20)

$lblCommonName = New-Object System.Windows.Forms.Label
$lblCommonName.Location = New-Object System.Drawing.Point(20, 90)
$lblCommonName.Size = New-Object System.Drawing.Size(150, 20)
$lblCommonName.Text = "Common Name (Name):"

$txtCommonName = New-Object System.Windows.Forms.TextBox
$txtCommonName.Location = New-Object System.Drawing.Point(170, 90)
$txtCommonName.Size = New-Object System.Drawing.Size(350, 20)

$lblOrganization = New-Object System.Windows.Forms.Label
$lblOrganization.Location = New-Object System.Drawing.Point(20, 120)
$lblOrganization.Size = New-Object System.Drawing.Size(150, 20)
$lblOrganization.Text = "Organisation:"

$txtOrganization = New-Object System.Windows.Forms.TextBox
$txtOrganization.Location = New-Object System.Drawing.Point(170, 120)
$txtOrganization.Size = New-Object System.Drawing.Size(350, 20)

$lblOrganizationalUnit = New-Object System.Windows.Forms.Label
$lblOrganizationalUnit.Location = New-Object System.Drawing.Point(20, 150)
$lblOrganizationalUnit.Size = New-Object System.Drawing.Size(150, 20)
$lblOrganizationalUnit.Text = "Abteilung:"

$txtOrganizationalUnit = New-Object System.Windows.Forms.TextBox
$txtOrganizationalUnit.Location = New-Object System.Drawing.Point(170, 150)
$txtOrganizationalUnit.Size = New-Object System.Drawing.Size(350, 20)

$lblLocality = New-Object System.Windows.Forms.Label
$lblLocality.Location = New-Object System.Drawing.Point(20, 180)
$lblLocality.Size = New-Object System.Drawing.Size(150, 20)
$lblLocality.Text = "Stadt:"

$txtLocality = New-Object System.Windows.Forms.TextBox
$txtLocality.Location = New-Object System.Drawing.Point(170, 180)
$txtLocality.Size = New-Object System.Drawing.Size(350, 20)

$lblState = New-Object System.Windows.Forms.Label
$lblState.Location = New-Object System.Drawing.Point(20, 210)
$lblState.Size = New-Object System.Drawing.Size(150, 20)
$lblState.Text = "Bundesland:"

$txtState = New-Object System.Windows.Forms.TextBox
$txtState.Location = New-Object System.Drawing.Point(170, 210)
$txtState.Size = New-Object System.Drawing.Size(350, 20)

$lblCountry = New-Object System.Windows.Forms.Label
$lblCountry.Location = New-Object System.Drawing.Point(20, 240)
$lblCountry.Size = New-Object System.Drawing.Size(150, 20)
$lblCountry.Text = "Land (2-stellig):"

$txtCountry = New-Object System.Windows.Forms.TextBox
$txtCountry.Location = New-Object System.Drawing.Point(170, 240)
$txtCountry.Size = New-Object System.Drawing.Size(350, 20)
$txtCountry.MaxLength = 2
$txtCountry.Text = "DE"

$lblKeySize = New-Object System.Windows.Forms.Label
$lblKeySize.Location = New-Object System.Drawing.Point(20, 270)
$lblKeySize.Size = New-Object System.Drawing.Size(150, 20)
$lblKeySize.Text = "Schlüsselgröße:"

$cboKeySize = New-Object System.Windows.Forms.ComboBox
$cboKeySize.Location = New-Object System.Drawing.Point(170, 270)
$cboKeySize.Size = New-Object System.Drawing.Size(350, 20)
$cboKeySize.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboKeySize.Items.AddRange(@("2048", "3072", "4096"))
$cboKeySize.SelectedIndex = 1  # 3072 als Standard

$lblCsrOutput = New-Object System.Windows.Forms.Label
$lblCsrOutput.Location = New-Object System.Drawing.Point(20, 300)
$lblCsrOutput.Size = New-Object System.Drawing.Size(150, 20)
$lblCsrOutput.Text = "Ausgabe-Verzeichnis:"

$txtCsrOutput = New-Object System.Windows.Forms.TextBox
$txtCsrOutput.Location = New-Object System.Drawing.Point(170, 300)
$txtCsrOutput.Size = New-Object System.Drawing.Size(350, 20)

$btnBrowseCsrOutput = New-Object System.Windows.Forms.Button
$btnBrowseCsrOutput.Location = New-Object System.Drawing.Point(530, 300)
$btnBrowseCsrOutput.Size = New-Object System.Drawing.Size(80, 23)
$btnBrowseCsrOutput.Text = "Durchsuchen"
$btnBrowseCsrOutput.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Ausgabeverzeichnis für CSR und Private Key auswählen"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $txtCsrOutput.Text = $folderBrowser.SelectedPath
    }
})

# Status-Label für CSR
$lblStatusCsr = New-Object System.Windows.Forms.Label
$lblStatusCsr.Location = New-Object System.Drawing.Point(20, 370)
$lblStatusCsr.Size = New-Object System.Drawing.Size(610, 50)
$lblStatusCsr.Text = "Bereit."

# Generieren-Button
$btnGenerateCsr = New-Object System.Windows.Forms.Button
$btnGenerateCsr.Location = New-Object System.Drawing.Point(170, 330)
$btnGenerateCsr.Size = New-Object System.Drawing.Size(120, 30)
$btnGenerateCsr.Text = "CSR generieren"
$btnGenerateCsr.Add_Click({
    # Eingaben überprüfen
    if (-not (Test-Path $txtOpenSSLCsr.Text)) {
        $lblStatusCsr.Text = "Fehler: OpenSSL-Pfad nicht gefunden."
        return
    }
    if ([string]::IsNullOrEmpty($txtEmailAddress.Text)) {
        $lblStatusCsr.Text = "Fehler: E-Mail-Adresse ist erforderlich."
        return
    }
    if ([string]::IsNullOrEmpty($txtCommonName.Text)) {
        $lblStatusCsr.Text = "Fehler: Common Name ist erforderlich."
        return
    }
    if ([string]::IsNullOrEmpty($txtCsrOutput.Text) -or (-not (Test-Path $txtCsrOutput.Text -PathType Container))) {
        $lblStatusCsr.Text = "Fehler: Gültiges Ausgabeverzeichnis ist erforderlich."
        return
    }
    if ([string]::IsNullOrEmpty($txtCountry.Text) -or ($txtCountry.Text.Length -ne 2)) {
        $lblStatusCsr.Text = "Fehler: Länderkürzel muss aus zwei Buchstaben bestehen."
        return
    }
    
    # Stellen Sie sicher, dass alle benötigten Felder ausgefüllt sind
    if ([string]::IsNullOrWhiteSpace($txtState.Text)) {
        $txtState.Text = "N/A"
    }
    if ([string]::IsNullOrWhiteSpace($txtLocality.Text)) {
        $txtLocality.Text = "N/A"
    }
    if ([string]::IsNullOrWhiteSpace($txtOrganization.Text)) {
        $txtOrganization.Text = "N/A"
    }
    if ([string]::IsNullOrWhiteSpace($txtOrganizationalUnit.Text)) {
        $txtOrganizationalUnit.Text = "N/A"
    }

    # OpenSSL-Pfad in Registry speichern
    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name $regKeyOpenSSL -Value $txtOpenSSLCsr.Text
    } catch {
        $lblStatusCsr.Text = "Warnung: Registry-Pfad konnte nicht gespeichert werden: $_"
    }

    try {
        # Dateipfade erstellen
        $baseName = $txtEmailAddress.Text -replace '@', '-at-'
        $baseName = $baseName -replace '\.', '-'
        $configFile = Join-Path $txtCsrOutput.Text "$baseName-openssl.cnf"
        $keyFile = Join-Path $txtCsrOutput.Text "$baseName-key.pem"
        $csrFile = Join-Path $txtCsrOutput.Text "$baseName-csr.pem"
        
        # OpenSSL Config generieren mit -addtrust=emailProtection und verbesserter v3_req Sektion
        $config = @"
[ req ]
default_bits = $($cboKeySize.SelectedItem)
distinguished_name = req_distinguished_name
req_extensions = v3_req
prompt = no
string_mask = utf8only

[ req_distinguished_name ]
countryName = $($txtCountry.Text)
stateOrProvinceName = $($txtState.Text)
localityName = $($txtLocality.Text)
organizationName = $($txtOrganization.Text)
organizationalUnitName = $($txtOrganizationalUnit.Text)
commonName = $($txtCommonName.Text)
emailAddress = $($txtEmailAddress.Text)

[ v3_req ]
basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature, keyEncipherment, nonRepudiation, dataEncipherment
extendedKeyUsage = clientAuth, emailProtection
subjectAltName = @alt_names

[ alt_names ]
email.1 = $($txtEmailAddress.Text)
"@
        
        # Config-Datei erstellen
        $config | Out-File -FilePath $configFile -Encoding ASCII
        
        # OpenSSL-Pfad
        $openSslPath = $txtOpenSSLCsr.Text
        $keySize = $cboKeySize.SelectedItem
        
        # Status aktualisieren
        $lblStatusCsr.Text = "Erstelle Private Key..."
        $form.Refresh()
        
        # 1. Schlüssel generieren mit Invoke-Expression
        $keyCmd = "& '$openSslPath' genrsa -out '$keyFile' $keySize"
        Invoke-Expression $keyCmd
        
        # Prüfen, ob der Key erstellt wurde
        if (Test-Path $keyFile) {
            # Status aktualisieren
            $lblStatusCsr.Text = "Private Key erstellt. Erstelle nun CSR..."
            $form.Refresh()
            
            # 2. CSR mit dem erstellten Key erzeugen mit Debug-Ausgabe
            $tempOutput = Join-Path $txtCsrOutput.Text "openssl_debug.log"
            $csrCmd = "& '$openSslPath' req -new -key '$keyFile' -config '$configFile' -out '$csrFile' *> '$tempOutput'"
            Invoke-Expression $csrCmd
            
            # Überprüfen, ob CSR erstellt wurde
            if (Test-Path $csrFile) {
                $lblStatusCsr.Text = "CSR und Private Key erfolgreich erstellt.`nCSR: $csrFile`nKey: $keyFile"
                
                # Temporäre Dateien löschen
                if (Test-Path $tempOutput) {
                    Remove-Item -Path $tempOutput -Force
                }
                if (Test-Path $configFile) {
                    Remove-Item -Path $configFile -Force
                }
            } else {
                # Lese Debug-Informationen falls verfügbar
                if (Test-Path $tempOutput) {
                    $errorLog = Get-Content -Path $tempOutput -Raw
                    $lblStatusCsr.Text = "Fehler: CSR konnte nicht erstellt werden.`n$errorLog"
                } else {
                    $lblStatusCsr.Text = "Fehler: CSR konnte nicht erstellt werden."
                }
            }
        } else {
            $lblStatusCsr.Text = "Fehler: Private Key konnte nicht erstellt werden."
        }
    } catch {
        $lblStatusCsr.Text = "Fehler bei der Ausführung: $_"
    }
})

# Beenden-Button für PFX Tab
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(310, 220)
$btnExit.Size = New-Object System.Drawing.Size(120, 30)
$btnExit.Text = "Beenden"
$btnExit.Add_Click({
    $form.Close()
})

# Copyright Label hinzufügen
$lblCopyright = New-Object System.Windows.Forms.Label
$lblCopyright.Location = New-Object System.Drawing.Point(20, 470)
$lblCopyright.Size = New-Object System.Drawing.Size(640, 20)
$lblCopyright.Text = "Copyright © 2025 Jörn Walter - https://www.der-windows-papst.de"
$lblCopyright.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$lblCopyright.ForeColor = [System.Drawing.Color]::DarkBlue

# Elemente zu den Tabs hinzufügen
# Tab 1: PEM zu PFX
$tabPemToPfx.Controls.Add($lblOpenSSL)
$tabPemToPfx.Controls.Add($txtOpenSSL)
$tabPemToPfx.Controls.Add($btnBrowseOpenSSL)
$tabPemToPfx.Controls.Add($lblPublicKey)
$tabPemToPfx.Controls.Add($txtPublicKey)
$tabPemToPfx.Controls.Add($btnBrowsePublicKey)
$tabPemToPfx.Controls.Add($lblPrivateKey)
$tabPemToPfx.Controls.Add($txtPrivateKey)
$tabPemToPfx.Controls.Add($btnBrowsePrivateKey)
$tabPemToPfx.Controls.Add($lblOutputFile)
$tabPemToPfx.Controls.Add($txtOutputFile)
$tabPemToPfx.Controls.Add($btnBrowseOutputFile)
$tabPemToPfx.Controls.Add($lblPassword)
$tabPemToPfx.Controls.Add($txtPassword)
$tabPemToPfx.Controls.Add($btnConvert)
$tabPemToPfx.Controls.Add($btnExit)
$tabPemToPfx.Controls.Add($lblStatusPfx)

# Tab 2: CSR Generator
$tabCsr.Controls.Add($lblOpenSSLCsr)
$tabCsr.Controls.Add($txtOpenSSLCsr)
$tabCsr.Controls.Add($btnBrowseOpenSSLCsr)
$tabCsr.Controls.Add($lblEmailAddress)
$tabCsr.Controls.Add($txtEmailAddress)
$tabCsr.Controls.Add($lblCommonName)
$tabCsr.Controls.Add($txtCommonName)
$tabCsr.Controls.Add($lblOrganization)
$tabCsr.Controls.Add($txtOrganization)
$tabCsr.Controls.Add($lblOrganizationalUnit)
$tabCsr.Controls.Add($txtOrganizationalUnit)
$tabCsr.Controls.Add($lblLocality)
$tabCsr.Controls.Add($txtLocality)
$tabCsr.Controls.Add($lblState)
$tabCsr.Controls.Add($txtState)
$tabCsr.Controls.Add($lblCountry)
$tabCsr.Controls.Add($txtCountry)
$tabCsr.Controls.Add($lblKeySize)
$tabCsr.Controls.Add($cboKeySize)
$tabCsr.Controls.Add($lblCsrOutput)
$tabCsr.Controls.Add($txtCsrOutput)
$tabCsr.Controls.Add($btnBrowseCsrOutput)
$tabCsr.Controls.Add($btnGenerateCsr)
$tabCsr.Controls.Add($lblStatusCsr)

# Tabs zum TabControl hinzufügen
$tabControl.Controls.Add($tabPemToPfx)
$tabControl.Controls.Add($tabCsr)

# TabControl und Copyright-Label zum Formular hinzufügen
$form.Controls.Add($tabControl)
$form.Controls.Add($lblCopyright)

# Formular anzeigen
[void]$form.ShowDialog()