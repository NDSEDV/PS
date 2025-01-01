<#
.SYNOPSIS
  Create your own Security/Certificates
.DESCRIPTION
  The tool is intended to help you with your dailiy business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.3
  Author:         Jörn Walter
  Creation Date:  2024-12-23
  Purpose/Change: Initial script development
  Update Date:    2024-12-24
  Purpose/Change: Add new SMB Tab

  Copyright (c) Jörn Walter. All rights reserved.
#>
# Function to check if the script is running with administrative privileges
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check if the script is running with administrative privileges
if (-not (Test-Admin)) {
    # If not, restart the script with administrative privileges
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

# Load Assemblies
function Load-RequiredAssemblies {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Security
        Write-Output "Assemblies erfolgreich geladen."
    } catch {
        Write-Error "Fehler beim Laden der Assemblies: $_"
        exit 1
    }
}
Load-RequiredAssemblies

# Function to generate a random password
function Generate-RandomPassword {
    $length = 12
    $chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()"
    $password = ""
    for ($i = 0; $i -lt $length; $i++) {
        $password += $chars[(Get-Random -Maximum $chars.Length)]
    }
    return $password
}

# Create GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create your own Security"
$form.Size = New-Object System.Drawing.Size(470, 890)
$form.StartPosition = "CenterScreen"

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(430, 830)

# Create Tabs
$tabPage1 = New-Object System.Windows.Forms.TabPage
$tabPage1.Text = "Zertifikat erstellen"

# Add Tabs to TabControl
$tabControl.TabPages.Add($tabPage1)

# Add TabControl to Form
$form.Controls.Add($tabControl)

# Copyright Label
    $lblCopyright = New-Object System.Windows.Forms.Label
    $lblCopyright.Text = "© 2025 Jörn Walter. Alle Rechte vorbehalten."
    $lblCopyright.AutoSize = $true
    $lblCopyright.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
    $lblCopyright.Location = New-Object System.Drawing.Point(10, 740)
    $tabPage1.Controls.Add($lblCopyright)

# Additional Label for URL
    $lblURL = New-Object System.Windows.Forms.Label
    $lblURL.Text = "https://www.der-windows-papst.de"
    $lblURL.AutoSize = $true
    $lblURL.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $lblURL.ForeColor = 'Blue'
    $lblURL.Location = New-Object System.Drawing.Point(10, 760)
    $tabPage1.Controls.Add($lblURL)

# Tab 1: Certificate Creation

    # Label and ComboBox for Algorithmus
    $lblAlgorithm = New-Object System.Windows.Forms.Label
    $lblAlgorithm.Text = "Algorithmus:"
    $lblAlgorithm.AutoSize = $true
    $lblAlgorithm.Location = New-Object System.Drawing.Point(10, 20)
    $tabPage1.Controls.Add($lblAlgorithm)

    $comboAlgorithm = New-Object System.Windows.Forms.ComboBox
    $comboAlgorithm.Items.AddRange(@("RSA", "ECDSA"))
    $comboAlgorithm.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboAlgorithm.Location = New-Object System.Drawing.Point(150, 20)
    $tabPage1.Controls.Add($comboAlgorithm)

    # Label and ComboBox for Optionen (RSA Key Size oder ECDSA Kurven)
    $lblOption = New-Object System.Windows.Forms.Label
    $lblOption.Text = "Option:"
    $lblOption.AutoSize = $true
    $lblOption.Location = New-Object System.Drawing.Point(10, 60)
    $tabPage1.Controls.Add($lblOption)

    $comboOption = New-Object System.Windows.Forms.ComboBox
    $comboOption.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboOption.Location = New-Object System.Drawing.Point(150, 60)
    $tabPage1.Controls.Add($comboOption)

    # Update options based on Algorithmus
    $comboAlgorithm.add_SelectedIndexChanged({
        $comboOption.Items.Clear()
        if ($comboAlgorithm.SelectedItem -eq "RSA") {
            $comboOption.Items.AddRange(@("2048", "3072", "4096"))
        } elseif ($comboAlgorithm.SelectedItem -eq "ECDSA") {
            $comboOption.Items.AddRange(@("NISTP256", "NISTP384", "brainpoolP256r1", "brainpoolP384r1", "brainpoolP512r1", "secP256r1", "secP384r1", "secP521r1"))
        }
    })

    # Label and TextBox for Laufzeit der Zertifikate
    $lblLaufzeit = New-Object System.Windows.Forms.Label
    $lblLaufzeit.Text = "Laufzeit (in Tagen):"
    $lblLaufzeit.AutoSize = $true
    $lblLaufzeit.Location = New-Object System.Drawing.Point(10, 100)
    $tabPage1.Controls.Add($lblLaufzeit)

    $txtLaufzeit = New-Object System.Windows.Forms.TextBox
    $txtLaufzeit.Location = New-Object System.Drawing.Point(150, 100)
    $txtLaufzeit.Size = New-Object System.Drawing.Size(100, 20)
    $tabPage1.Controls.Add($txtLaufzeit)

    # Label and TextBox for Subject Name
    $lblSubjectName = New-Object System.Windows.Forms.Label
    $lblSubjectName.Text = "Subject Names (kommagetrennt): Erster Name ist = CN "
    $lblSubjectName.AutoSize = $true
    $lblSubjectName.Location = New-Object System.Drawing.Point(10, 140)
    $tabPage1.Controls.Add($lblSubjectName)

    $txtSubjectName = New-Object System.Windows.Forms.TextBox
    $txtSubjectName.Location = New-Object System.Drawing.Point(10, 160)
    $txtSubjectName.Size = New-Object System.Drawing.Size(360, 20)
    $tabPage1.Controls.Add($txtSubjectName)

    # Label and ComboBox for Certificate Usage
    $lblUsage = New-Object System.Windows.Forms.Label
    $lblUsage.Text = "Zweck:"
    $lblUsage.AutoSize = $true
    $lblUsage.Location = New-Object System.Drawing.Point(10, 200)
    $tabPage1.Controls.Add($lblUsage)

    $comboUsage = New-Object System.Windows.Forms.ComboBox
    $comboUsage.Items.AddRange(@("Clientauthentifizierung", "Serverauthentifizierung", "Signing"))
    $comboUsage.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboUsage.Location = New-Object System.Drawing.Point(150, 200)
    $tabPage1.Controls.Add($comboUsage)

    # Label and ComboBox for Signature Hash Algorithm
    $lblHashAlgorithm = New-Object System.Windows.Forms.Label
    $lblHashAlgorithm.Text = "Hash-Algorithmus:"
    $lblHashAlgorithm.AutoSize = $true
    $lblHashAlgorithm.Location = New-Object System.Drawing.Point(10, 240)
    $tabPage1.Controls.Add($lblHashAlgorithm)

    $comboHashAlgorithm = New-Object System.Windows.Forms.ComboBox
    $comboHashAlgorithm.Items.AddRange(@("SHA256", "SHA384", "SHA512"))
    $comboHashAlgorithm.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboHashAlgorithm.Location = New-Object System.Drawing.Point(150, 240)
    $tabPage1.Controls.Add($comboHashAlgorithm)

    # Label and TextBox for PFX Password
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Text = "PFX Passwort:"
    $lblPassword.AutoSize = $true
    $lblPassword.Location = New-Object System.Drawing.Point(10, 280)
    $tabPage1.Controls.Add($lblPassword)

    $txtPassword = New-Object System.Windows.Forms.TextBox
    $txtPassword.Location = New-Object System.Drawing.Point(150, 280)
    $txtPassword.Size = New-Object System.Drawing.Size(220, 20)
    $txtPassword.UseSystemPasswordChar = $true
    $tabPage1.Controls.Add($txtPassword)

    # Radio buttons for export options
    $lblExportOptions = New-Object System.Windows.Forms.Label
    $lblExportOptions.Text = "Exportoptionen:"
    $lblExportOptions.AutoSize = $true
    $lblExportOptions.Location = New-Object System.Drawing.Point(10, 320)
    $tabPage1.Controls.Add($lblExportOptions)

    $radioPfx = New-Object System.Windows.Forms.RadioButton
    $radioPfx.Text = "PFX exportieren"
    $radioPfx.AutoSize = $true
    $radioPfx.Location = New-Object System.Drawing.Point(150, 320)
    $tabPage1.Controls.Add($radioPfx)

    $radioPublicKey = New-Object System.Windows.Forms.RadioButton
    $radioPublicKey.Text = "Öffentlichen Schlüssel exportieren"
    $radioPublicKey.AutoSize = $true
    $radioPublicKey.Location = New-Object System.Drawing.Point(150, 350)
    $tabPage1.Controls.Add($radioPublicKey)

    $radioBoth = New-Object System.Windows.Forms.RadioButton
    $radioBoth.Text = "Beides exportieren"
    $radioBoth.AutoSize = $true
    $radioBoth.Location = New-Object System.Drawing.Point(150, 380)
    $tabPage1.Controls.Add($radioBoth)

    # Buttons for Create and Reset
    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Zertifikat erstellen"
    $btnCreate.Location = New-Object System.Drawing.Point(120, 420)
    $btnCreate.Size = New-Object System.Drawing.Size(130, 30)
    $tabPage1.Controls.Add($btnCreate)

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "Zurücksetzen"
    $btnReset.Location = New-Object System.Drawing.Point(260, 420)
    $btnReset.Size = New-Object System.Drawing.Size(110, 30)
    $tabPage1.Controls.Add($btnReset)

    # Output TextBox
    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Location = New-Object System.Drawing.Point(10, 460)
    $outputBox.Size = New-Object System.Drawing.Size(360, 150)
    $outputBox.Multiline = $true
    $outputBox.ScrollBars = "Vertical"
    $tabPage1.Controls.Add($outputBox)

    # Button to show supported ECC curves
    $btnShowCurves = New-Object System.Windows.Forms.Button
    $btnShowCurves.Text = "Die vom System unterstützten ECC-Kurven anzeigen"
    $btnShowCurves.Location = New-Object System.Drawing.Point(90, 620)
    $btnShowCurves.AutoSize = $true
    $tabPage1.Controls.Add($btnShowCurves)

    # Button to show supported Cipher Suites
    $btnShowCipherSuites = New-Object System.Windows.Forms.Button
    $btnShowCipherSuites.Text = "Die vom System unterstützten Cipher Suiten anzeigen"
    $btnShowCipherSuites.Location = New-Object System.Drawing.Point(90, 650)
    $btnShowCipherSuites.AutoSize = $true
    $tabPage1.Controls.Add($btnShowCipherSuites)

    # Button to copy output to clipboard
    $btnCopyOutput = New-Object System.Windows.Forms.Button
    $btnCopyOutput.Text = "Ausgabe in Zwischenablage kopieren"
    $btnCopyOutput.Location = New-Object System.Drawing.Point(90, 680)
    $btnCopyOutput.AutoSize = $true
    $tabPage1.Controls.Add($btnCopyOutput)

    # Reset form logic
$btnReset.Add_Click({
    $comboAlgorithm.SelectedIndex = -1
    $comboOption.Items.Clear()
    $txtLaufzeit.Text = ""
    $comboHashAlgorithm.SelectedIndex = -1
    $txtSubjectName.Text = ""
    $comboUsage.SelectedIndex = -1
    $txtPassword.Text = ""
    $outputBox.Text = ""
    $radioPfx.Checked = $false
    $radioPublicKey.Checked = $false
    $radioBoth.Checked = $false
})

# Show supported ECC curves logic
$btnShowCurves.Add_Click({
    $outputBox.Clear()
    $outputBox.AppendText("Unterstützte ECC Kurven in Reihenfolge:`n" + [Environment]::NewLine)
    try {
        $curves = Get-TlsEccCurve
        if ($curves.Count -eq 0) {
            $outputBox.AppendText("Es wurden keine unterstützten ECC Kurven gefunden. Dies könnte durch eine Gruppenrichtlinie eingeschränkt sein.`n")
        } else {
           foreach ($EccCurve in $Curves) {
                $outputBox.AppendText($EccCurve  + [Environment]::NewLine)
            }
        }
    } catch {
        $outputBox.AppendText("Fehler beim Abrufen der ECC Kurven: " + $_.Exception.Message)
    }
})

# Show supported Cipher Suites logic
$btnShowCipherSuites.Add_Click({
    $outputBox.Clear()
    $outputBox.AppendText("Unterstützte Cipher Suiten:`n" + [Environment]::NewLine)
    try {
        $cipherSuites = Get-TlsCipherSuite
        if ($cipherSuites.Count -eq 0) {
            $outputBox.AppendText("Es wurden keine unterstützten Cipher Suiten gefunden. Dies könnte durch eine Gruppenrichtlinie eingeschränkt sein.`n")

            # Read the "Functions" value from the registry
            $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002"
            $functionsValue = Get-ItemProperty -Path $regPath -Name "Functions" -ErrorAction Stop

            $outputBox.AppendText("Registry-Pfad: $regPath`n" + [Environment]::NewLine)
            $outputBox.AppendText("Functions-Wert:`n" + $functionsValue.Functions -replace ',', [Environment]::NewLine)
        } else {
            foreach ($cipherSuite in $cipherSuites) {
                $outputBox.AppendText($cipherSuite.Name + [Environment]::NewLine)
            }
        }
    } catch {
        $outputBox.AppendText("Fehler beim Abrufen der Cipher Suiten: " + $_.Exception.Message)
    }
})

# Copy output to clipboard logic
$btnCopyOutput.Add_Click({
    try {
        $outputText = $outputBox.Text
        if ([string]::IsNullOrWhiteSpace($outputText)) {
            [System.Windows.Forms.MessageBox]::Show("Die Ausgabe ist leer. Es gibt nichts zu kopieren.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.Clipboard]::SetText($outputText)
            [System.Windows.Forms.MessageBox]::Show("Ausgabe erfolgreich in die Zwischenablage kopiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Kopieren der Ausgabe in die Zwischenablage: " + $_.Exception.Message, "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Create Certificate logic
$btnCreate.Add_Click({
    $algorithm = $comboAlgorithm.SelectedItem
    $option = $comboOption.SelectedItem
    $hashAlgorithm = $comboHashAlgorithm.SelectedItem
    $subjectNames = $txtSubjectName.Text -split ","
    $usage = $comboUsage.SelectedItem
    $password = $txtPassword.Text
    $laufzeit = [int]$txtLaufzeit.Text

    if (-not $algorithm -or -not $option -or -not $hashAlgorithm -or -not $subjectNames -or -not $usage -or $subjectNames.Count -eq 0 -or -not $laufzeit) {
        [System.Windows.Forms.MessageBox]::Show("Bitte alle Felder ausfüllen und mindestens einen DNS-Namen sowie ein Passwort angeben.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if ([string]::IsNullOrEmpty($password)) {
        $password = Generate-RandomPassword
    } elseif ($password.Length -lt 8) {
        [System.Windows.Forms.MessageBox]::Show("Das Passwort muss mindestens 8 Zeichen lang sein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        $enhancedKeyUsage = switch ($usage) {
            "Clientauthentifizierung" { "1.3.6.1.5.5.7.3.2" }
            "Serverauthentifizierung" { "1.3.6.1.5.5.7.3.1" }
            "Signing" { "1.3.6.1.5.5.7.3.3" }
        }

        $storeLocation = if ($usage -eq "Signing") { "Cert:\CurrentUser\My" } else { "Cert:\LocalMachine\My" }

        if ($algorithm -eq "RSA") {
            $keySize = [int]$option
            $cert = New-SelfSignedCertificate -DnsName $subjectNames -KeyAlgorithm RSA -HashAlgorithm $hashAlgorithm -KeyLength $keySize -CertStoreLocation $storeLocation -TextExtension @("2.5.29.37={text}$enhancedKeyUsage") -NotAfter (Get-Date).AddDays($laufzeit)
        } elseif ($algorithm -eq "ECDSA") {
            $curve = switch ($option) {
                "NISTP256" { "ECDSA_NISTP256" }
                "NISTP384" { "ECDSA_NISTP384" }
                "brainpoolP256r1" { "ECDSA_brainpoolP256r1" }
                "brainpoolP384r1" { "ECDSA_brainpoolP384r1" }
                "brainpoolP512r1" { "ECDSA_brainpoolP512r1" }
                "secP256r1" { "secP256r1" }
                "secP384r1" { "secP384r1" }
                "secP521r1" { "secP521r1" }
            }
            $cert = New-SelfSignedCertificate -DnsName $subjectNames -KeyAlgorithm $curve -HashAlgorithm $hashAlgorithm -CertStoreLocation $storeLocation -TextExtension @("2.5.29.37={text}$enhancedKeyUsage") -NotAfter (Get-Date).AddDays($laufzeit)
        }

        $outputBox.Text = ("Zertifikat erfolgreich erstellt:`n" + $cert.Thumbprint  + [Environment]::NewLine)
        $outputBox.AppendText("`nVerwendete Optionen:`n")
        $outputBox.AppendText("Algorithmus: $algorithm`n" + [Environment]::NewLine)
        $outputBox.AppendText("Option: $option`n" + [Environment]::NewLine)
        $outputBox.AppendText("Hash-Algorithmus: $hashAlgorithm`n" + [Environment]::NewLine)
        foreach ($subjectName in $subjectNames) {
            $outputBox.AppendText("Subject Name: $subjectName`n" + [Environment]::NewLine)
        }
        $outputBox.AppendText("Zweck: $usage`n" + [Environment]::NewLine)
        $outputBox.AppendText("Laufzeit (in Tagen): $laufzeit`n" + [Environment]::NewLine)
        $outputBox.AppendText("PFX Passwort: $password`n" + [Environment]::NewLine)

        $desktopPath = [System.Environment]::GetFolderPath("Desktop")
        if ($radioPfx.Checked -or $radioBoth.Checked) {
            $pfxFilePath = [System.IO.Path]::Combine($desktopPath, "ExportedCert_$($subjectNames[0]).pfx")
            Export-PfxCertificate -Cert $cert -FilePath $pfxFilePath -Password (ConvertTo-SecureString -String $password -Force -AsPlainText)
            $outputBox.Text += ("`nPFX exportiert nach: $pfxFilePath" + [Environment]::NewLine)
        }
        if ($radioPublicKey.Checked -or $radioBoth.Checked) {
            $publicKeyFilePath = [System.IO.Path]::Combine($desktopPath, "PublicKey_$($subjectNames[0]).cer")
            Export-Certificate -Cert $cert -FilePath $publicKeyFilePath
            $outputBox.Text += "`nÖffentlicher Schlüssel exportiert nach: $publicKeyFilePath"

            # Prompt to import the public key into the trusted root certification authorities container
            $result = [System.Windows.Forms.MessageBox]::Show("Möchten Sie den öffentlichen Schlüssel in den Container für vertrauenswürdige Stammzertifizierungsstellen importieren?", "Bestätigung", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                try {
                    Import-Certificate -FilePath $publicKeyFilePath -CertStoreLocation "Cert:\LocalMachine\Root"
                    $outputBox.Text += "`nÖffentlicher Schlüssel erfolgreich in den Container für vertrauenswürdige Stammzertifizierungsstellen importiert."
                } catch {
                    $outputBox.Text += "`nFehler beim Importieren des öffentlichen Schlüssels in den Container für vertrauenswürdige Stammzertifizierungsstellen: " + $_.Exception.Message
                }
            }
        }
    } catch {
        $outputBox.Text = "Fehler beim Erstellen des Zertifikats:`n" + $_.Exception.Message
    }
})

    # Ensure the form is shown
[void]$form.ShowDialog()
