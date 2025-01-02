<#
.SYNOPSIS
  Manage your Windows Firewall-Profiles
.DESCRIPTION
  The tool is intended to help you with your dailiy business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.0
  Author:         Jörn Walter
  Creation Date:  2025-01-02
  Purpose/Change: Initial script development

  Jörn Walter. All rights reserved.
#>

# Erforderliche Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Ermitteln der Netzwerkverbindungen und ihrer Firewall-Profile
function Get-NetworkConnectionProfiles {
    Get-NetConnectionProfile
}

# Funktion zum Setzen des Firewall-Profils für eine bestimmte Verbindung
function Set-NetworkConnectionProfile($InterfaceAlias, $NewCategory) {
    Set-NetConnectionProfile -InterfaceAlias $InterfaceAlias -NetworkCategory $NewCategory
}

# Hashtables zum Speichern der Labels und Radio Buttons
$profileLabels = @{}
$radioButtons = @{}

# Hauptformular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "Firewall-Profile verwalten"
$form.Size = New-Object System.Drawing.Size(600,470)
$form.StartPosition = "CenterScreen"

# Panel für die dynamische Anzeige der Verbindungen
$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(10,10)
$panel.Size = New-Object System.Drawing.Size(560,300)
$panel.AutoScroll = $true

# Button zum Anwenden der Änderungen
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Location = New-Object System.Drawing.Point(10,320)
$applyButton.Size = New-Object System.Drawing.Size(100,30)
$applyButton.Text = "Anwenden"

# Button zum Aktualisieren des Status
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(120,320)
$refreshButton.Size = New-Object System.Drawing.Size(100,30)
$refreshButton.Text = "Aktualisieren"

# Status-Label erstellen
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10,360)
$statusLabel.Size = New-Object System.Drawing.Size(560,20)
$statusLabel.Text = ""
$statusLabel.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)

# Label für 'Jörn Walter'
$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Location = New-Object System.Drawing.Point(10, 380)
$authorLabel.AutoSize = $true
$authorLabel.TextAlign = 'BottomLeft'
$authorLabel.Text = "Jörn Walter - Alle Rechte vorbehalten."
$authorLabel.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

# Label für 'https://www.der-windows-papst.de'
$websiteLabel = New-Object System.Windows.Forms.LinkLabel
$websiteLabel.Location = New-Object System.Drawing.Point(10, 390)
$websiteLabel.Size = New-Object System.Drawing.Size(180, 20)
$websiteLabel.TextAlign = 'BottomLeft'
$websiteLabel.Text = "https://www.der-windows-papst.de"
$websiteLabel.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Regular)

# Funktion zum Erstellen der GUI-Komponenten für die Netzwerkadapter
function Build-AdapterGUI {
    # Hashtables leeren
    $profileLabels.Clear()
    $radioButtons.Clear()

    # Controls aus dem Panel entfernen
    $panel.Controls.Clear()

    # y-Position zurücksetzen
    $yPos = 10

    foreach ($conn in $connections) {
        $interfaceAlias = $conn.InterfaceAlias

        # Label für den Namen der Verbindung
        $connLabel = New-Object System.Windows.Forms.Label
        $connLabel.Location = New-Object System.Drawing.Point(10,$yPos)
        $connLabel.Size = New-Object System.Drawing.Size(400,20)
        $connLabel.Text = "Adapter: $($conn.InterfaceAlias) (Netzwerk: $($conn.Name))"
        $panel.Controls.Add($connLabel)
        $yPos += 25

        # Label für das aktuelle Firewall-Profil
        $profileLabel = New-Object System.Windows.Forms.Label
        $profileLabel.Location = New-Object System.Drawing.Point(30,$yPos)
        $profileLabel.Size = New-Object System.Drawing.Size(200,20)
        $profileLabel.Text = "Aktuelles Profil: $($conn.NetworkCategory)"
        $panel.Controls.Add($profileLabel)

        # Labels in der Hashtable speichern
        $profileLabels[$interfaceAlias] = $profileLabel

        # Radio Buttons erstellen und in der Hashtable speichern
        $rbPrivate = New-Object System.Windows.Forms.RadioButton
        $rbPrivate.Location = New-Object System.Drawing.Point(250,$yPos)
        $rbPrivate.Size = New-Object System.Drawing.Size(70,20)
        $rbPrivate.Text = "Privat"

        $rbPublic = New-Object System.Windows.Forms.RadioButton
        $rbPublic.Location = New-Object System.Drawing.Point(330,$yPos)
        $rbPublic.Size = New-Object System.Drawing.Size(80,20)
        $rbPublic.Text = "Öffentlich"

        $rbDomain = New-Object System.Windows.Forms.RadioButton
        $rbDomain.Location = New-Object System.Drawing.Point(420,$yPos)
        $rbDomain.Size = New-Object System.Drawing.Size(80,20)
        $rbDomain.Text = "Domäne"

        # Aktuelles Profil vorab auswählen
        $rbPrivate.Checked = $conn.NetworkCategory -eq "Private"
        $rbPublic.Checked = $conn.NetworkCategory -eq "Public"
        $rbDomain.Checked = $conn.NetworkCategory -eq "DomainAuthenticated"

        # Radio Buttons zum Panel hinzufügen
        $panel.Controls.Add($rbPrivate)
        $panel.Controls.Add($rbPublic)
        $panel.Controls.Add($rbDomain)

        # Radio Buttons in der Hashtable speichern
        $radioButtons[$interfaceAlias] = @{
            "Private" = $rbPrivate
            "Public" = $rbPublic
            "DomainAuthenticated" = $rbDomain
        }

        $yPos += 30
    }
}

# Netzwerkverbindungen abrufen
$connections = Get-NetworkConnectionProfiles

# GUI-Komponenten erstellen
Build-AdapterGUI

# Ereignis beim Klicken auf den "Anwenden"-Button
$applyButton.Add_Click({
    # Netzwerkverbindungen aktualisiert abrufen
    $connections = Get-NetworkConnectionProfiles

    foreach ($conn in $connections) {
        $interfaceAlias = $conn.InterfaceAlias
        $selectedProfile = ""

        # Überprüfen, ob die Verbindung in den Hashtables vorhanden ist
        if ($radioButtons.ContainsKey($interfaceAlias)) {
            # Überprüfen, welcher Radio Button ausgewählt ist
            if ($radioButtons[$interfaceAlias]["Private"].Checked) {
                $selectedProfile = "Private"
            } elseif ($radioButtons[$interfaceAlias]["Public"].Checked) {
                $selectedProfile = "Public"
            } elseif ($radioButtons[$interfaceAlias]["DomainAuthenticated"].Checked) {
                $selectedProfile = "DomainAuthenticated"
            }

            # Profil nur ändern, wenn es verschieden ist
            if ($selectedProfile -and ($selectedProfile -ne $conn.NetworkCategory)) {
                Set-NetworkConnectionProfile -InterfaceAlias $interfaceAlias -NewCategory $selectedProfile
            }
        }
    }

    # Statusmeldung aktualisieren
    $statusLabel.Text = "Die Firewall-Profile wurden erfolgreich aktualisiert."
})

# Ereignis beim Klicken auf den "Aktualisieren"-Button
$refreshButton.Add_Click({
    # Netzwerkverbindungen aktualisiert abrufen
    $connections = Get-NetworkConnectionProfiles

    # GUI-Komponenten neu erstellen
    Build-AdapterGUI

    # Statusmeldung aktualisieren
    $statusLabel.Text = "Der Status wurde erfolgreich aktualisiert."
})

# Steuerelemente zum Formular hinzufügen
$form.Controls.Add($panel)
$form.Controls.Add($applyButton)
$form.Controls.Add($refreshButton)
$form.Controls.Add($statusLabel)
$form.Controls.Add($authorLabel)
$form.Controls.Add($websiteLabel)

# Formular anzeigen
$form.Topmost = $true
[void]$form.ShowDialog()
