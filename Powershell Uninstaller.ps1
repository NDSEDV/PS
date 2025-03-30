<#
.SYNOPSIS
  Powershell Uninstaller
.DESCRIPTION
  Das Tool ermöglicht die Deinstallation von Software auf Basis der Uninstall-Strings in der Registry
.PARAMETER language
.NOTES
  Version:        1.0
  Author:         Jörn Walter
  Creation Date:  2025-03-24
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
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Abrufen aller installierten Programme mit Uninstall-Strings
function Get-InstalledSoftware {
    $software = @()
    
    # 32-bit Software in 64-bit System
    $uninstallKeys = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $software += Get-ItemProperty $uninstallKeys | 
                Where-Object { $_.DisplayName -and $_.UninstallString } | 
                Select-Object DisplayName, Publisher, DisplayVersion, InstallDate, UninstallString
    
    # 64-bit Software
    $uninstallKeys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $software += Get-ItemProperty $uninstallKeys | 
                Where-Object { $_.DisplayName -and $_.UninstallString } | 
                Select-Object DisplayName, Publisher, DisplayVersion, InstallDate, UninstallString
    
    # Benutzerspezifische Software (HKCU)
    $uninstallKeys = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $software += Get-ItemProperty $uninstallKeys | 
                Where-Object { $_.DisplayName -and $_.UninstallString } | 
                Select-Object DisplayName, Publisher, DisplayVersion, InstallDate, UninstallString
    
    # Sortieren nach Displayname
    return $software | Sort-Object DisplayName
}

# Funktion zum Deinstallieren der ausgewählten Software
function Uninstall-Software {
    param (
        [string]$uninstallString,
        [string]$displayName
    )
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Möchten du die Software '$displayName' wirklich deinstallieren?",
        "Bestätigung",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            # Überprüfen, ob der Uninstall-String ein MSI-Deinstallationscode ist
            if ($uninstallString -like "*MsiExec.exe*" -or $uninstallString -match "/I{[A-Z0-9\-]+}") {
                $msiCode = ""
                
                # Extrahieren des MSI-Produktcodes
                if ($uninstallString -match "{[A-Z0-9\-]+}") {
                    $msiCode = $matches[0]
                }
                
                if ($msiCode) {
                    # MSI-Deinstallation mit Quiet-Parameter
                    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $msiCode /qb" -Wait -PassThru
                } else {
                    # Wenn kein MSI-Code gefunden wurde, modifizieren wir den String
                    $modifiedString = $uninstallString -replace "/I", "/X"
                    $modifiedString += " /qb"
                    
                    # Starten des modifizierten Befehls
                    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $modifiedString" -Wait -PassThru
                }
            } else {
                # Für andere Deinstallationsmethoden
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString" -Wait -PassThru
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Deinstallation von '$displayName' abgeschlossen. Exit-Code: $($process.ExitCode)",
                "Information",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # GUI aktualisieren
            RefreshSoftwareList
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Deinstallation von '$displayName': $($_.Exception.Message)",
                "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

# Hauptformular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "Powershell Uninstaller"
$form.Size = New-Object System.Drawing.Size(1360, 610)
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Application

# ListView für die Software-Liste erstellen
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 50)
$listView.Size = New-Object System.Drawing.Size(1330, 450)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $false

# Spalten für die ListView hinzufügen
$listView.Columns.Add("Name", 300)
$listView.Columns.Add("Hersteller", 150)
$listView.Columns.Add("Version", 100)
$listView.Columns.Add("Installationsdatum", 110)
$listView.Columns.Add("Uninstall-String", 640)

# Suchfeld hinzufügen
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Suche:"
$searchLabel.Location = New-Object System.Drawing.Point(10, 15)
$searchLabel.Size = New-Object System.Drawing.Size(50, 20)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(70, 15)
$searchBox.Size = New-Object System.Drawing.Size(300, 20)

# Refresh-Button hinzufügen
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(380, 15)
$refreshButton.Size = New-Object System.Drawing.Size(100, 23)
$refreshButton.Text = "Aktualisieren"

# Export-Button hinzufügen
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(490, 15)
$exportButton.Size = New-Object System.Drawing.Size(100, 23)
$exportButton.Text = "Exportieren"

# Uninstall-Button hinzufügen
$uninstallButton = New-Object System.Windows.Forms.Button
$uninstallButton.Location = New-Object System.Drawing.Point(10, 510)
$uninstallButton.Size = New-Object System.Drawing.Size(200, 30)
$uninstallButton.Text = "Ausgewählte Software deinstallieren"
$uninstallButton.Enabled = $false

# Statusleiste hinzufügen
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Bereit"
$statusStrip.Items.Add($statusLabel)

# Copyright-Label hinzufügen
$copyrightLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$copyrightLabel.Text = "© 2025 Jörn Walter - https://www.der-windows-papst.de"
$copyrightLabel.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
$statusStrip.Items.Add($copyrightLabel)

# Funktion zum Aktualisieren der Software-Liste
function RefreshSoftwareList {
    $listView.Items.Clear()
    $statusLabel.Text = "Lade installierte Software..."
    $form.Update()
    
    $software = Get-InstalledSoftware
    
    $searchText = $searchBox.Text.ToLower()
    
    foreach ($app in $software) {
        # Nur Einträge anzeigen, die der Suche entsprechen
        if ([string]::IsNullOrEmpty($searchText) -or 
            $app.DisplayName -like "*$searchText*" -or 
            $app.Publisher -like "*$searchText*") {
            
            $item = New-Object System.Windows.Forms.ListViewItem($app.DisplayName)
            $item.SubItems.Add($(if ([string]::IsNullOrEmpty($app.Publisher)) { "-" } else { $app.Publisher }))
            $item.SubItems.Add($(if ([string]::IsNullOrEmpty($app.DisplayVersion)) { "-" } else { $app.DisplayVersion }))
            $item.SubItems.Add($(if ([string]::IsNullOrEmpty($app.InstallDate)) { "-" } else { $app.InstallDate }))
            $item.SubItems.Add($app.UninstallString)
            $item.Tag = $app
            
            $listView.Items.Add($item)
        }
    }
    
    $statusLabel.Text = "Fertig. $($listView.Items.Count) Programme gefunden."
}

# Event-Handler für Suchfeld
$searchBox.Add_TextChanged({
    RefreshSoftwareList
})

# Event-Handler für Refresh-Button
$refreshButton.Add_Click({
    RefreshSoftwareList
})

# Event-Handler für Export-Button
$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
    $saveFileDialog.Title = "Software-Liste exportieren"
    $saveFileDialog.FileName = "Installed_Software_$(Get-Date -Format 'yyyy-MM-dd').csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $software = Get-InstalledSoftware
            $software | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show(
                "Die Software-Liste wurde erfolgreich exportiert nach: $($saveFileDialog.FileName)",
                "Export erfolgreich",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Exportieren der Software-Liste: $($_.Exception.Message)",
                "Fehler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Event-Handler für Uninstall-Button
$uninstallButton.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        $app = $selectedItem.Tag
        
        Uninstall-Software -uninstallString $app.UninstallString -displayName $app.DisplayName
    }
})

# Event-Handler für ListView-Auswahl
$listView.Add_SelectedIndexChanged({
    $uninstallButton.Enabled = $listView.SelectedItems.Count -gt 0
})

# Doppelklick auf ein Element führt zur Deinstallation
$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        $app = $selectedItem.Tag
        
        Uninstall-Software -uninstallString $app.UninstallString -displayName $app.DisplayName
    }
})

# Controls zum Formular hinzufügen
$form.Controls.Add($searchLabel)
$form.Controls.Add($searchBox)
$form.Controls.Add($refreshButton)
$form.Controls.Add($exportButton)
$form.Controls.Add($listView)
$form.Controls.Add($uninstallButton)
$form.Controls.Add($statusStrip)

# Initial Software-Liste laden
RefreshSoftwareList

# Formular anzeigen
$form.ShowDialog()