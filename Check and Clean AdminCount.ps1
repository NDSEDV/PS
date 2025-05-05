<#
.SYNOPSIS
  Check and Clean AdminCount
.DESCRIPTION
  GUI für das Cleanup-AdminCount.ps1 Skript von Mark Heitbrink
.PARAMETER language
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-05-02
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

# Funktion zum Überprüfen, ob der Computer Teil einer Domäne ist
function Test-IsDomainJoined {
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        return ($computerSystem.PartOfDomain)
    }
    catch {
        return $false
    }
}

# Überprüft, ob der Computer Teil einer Domäne ist
if (-not (Test-IsDomainJoined)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Dieses Tool kann nur in einer Domänenumgebung verwendet werden.`nDer Computer ist nicht Teil einer Domäne.",
        "Fehler: Keine Domäne gefunden",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
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

# Erstellen eines Protokollverzeichnisses, falls nicht vorhanden
$LogDir = "$PSScriptRoot\Logs"
if (-not (Test-Path -Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# Logdatei mit Zeitstempel
$LogFile = "$LogDir\AdminCount-Cleanup_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

# Funktion zum Protokollieren
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    
    # In Datei schreiben
    Add-Content -Path $LogFile -Value $LogMessage
    
    # Zur Anzeige im GUI hinzufügen - mit Fehlerbehandlung
    try {
        if ($global:logTextBox -ne $null) {
            $global:logTextBox.AppendText("$LogMessage`r`n")
            $global:logTextBox.SelectionStart = $global:logTextBox.Text.Length
            $global:logTextBox.ScrollToCaret()
            
            # UI explizit aktualisieren
            $global:logTextBox.Refresh()
        }
    } catch {
        # Fehler beim Aktualisieren der Textbox abfangen
        Write-Host "Fehler bei der GUI-Aktualisierung: $($_.Exception.Message)"
    }
    
    # Färbung je nach Typ
    switch ($Level) {
        "INFO"    { $color = "Black" }
        "WARNING" { $color = "Orange" }
        "ERROR"   { $color = "Red" }
        "SUCCESS" { $color = "Green" }
    }
    
    try {
        if ($global:statusLabel -ne $null) {
            $global:statusLabel.Text = $Message
            $global:statusLabel.ForeColor = $color
            $global:statusLabel.Refresh()
        }
    } catch {
        Write-Host "Fehler beim Aktualisieren des Status-Labels: $($_.Exception.Message)"
    }
    
    # UI aktualisieren
    [System.Windows.Forms.Application]::DoEvents()
}

# Hauptfunktion, die das ursprüngliche Skript ausführt
function Start-AdminCountCleanup {
    param (
        [switch]$Cleanup
    )
    
    Write-Log "Script gestartet im $(if ($Cleanup) {'Cleanup'} else {'Report-Only'}) Modus" -Level "INFO"
    
    try {
        # 1. Determine AdminSDHolder protected objects
        Write-Log "Ermittle AdminSDHolder-geschützte Objekte..." -Level "INFO"
        
        # 1.1 Get Domain SID and set BuiltIn
        $RootSID = ((Get-ADForest).Domains | Get-ADDomain).DomainSID
        $DomSID = (Get-ADDomain).DomainSID
        $BuiltIn = "S-1-5-32"
        
        Write-Log "Domain SID: $DomSID, Root SID: $RootSID" -Level "INFO"
        
        # 1.2 Determine Protected Groups
        Write-Log "Identifiziere geschützte Gruppen..." -Level "INFO"
        
        # BUILTIN_ADMINISTRATORS, S-1-5-32-544
        $BAdmins = Get-ADGroup "$BuiltIn-544"
        
        # ACCOUNT_OPERATORS, S-1-5-32-548
        $BAccounts = Get-ADGroup "$BuiltIn-548"
        
        # SERVER_OPERATORS, S-1-5-32-549
        $BServerOps = Get-ADGroup "$BuiltIn-549"
        
        # PRINTER_OPERATORS, S-1-5-32-550
        $BPrinterOps = Get-ADGroup "$BuiltIn-550"
        
        # BACKUP_OPERATORS, S-1-5-32-551
        $BBackupOps = Get-ADGroup "$BuiltIn-551"
        
        # REPLICATOR, S-1-5-32-552
        $BReplicator = Get-ADGroup "$BuiltIn-552"
        
        # ADMINISTRATOR, S-1-5-21-<machine>-500
        $Administrator = Get-ADUser "$DomSID-500"
        
        # KRBTGT, S-1-5-21-<domain>-502
        $krbtgt = Get-ADUser "$DomSID-502"
        
        # DOMAIN_ADMINS, S-1-5-21-<domain>-512
        $DomAdmins = Get-ADGroup "$DomSID-512"
        
        # DOMAIN_CONTROLLERS, S-1-5-21-<domain>-516
        try {
            $DomCon = Get-ADGroup "$DomSID-516"
        } catch {
            Write-Log "Warnung: Domain Controllers Gruppe nicht gefunden (SID: $DomSID-516)" -Level "WARNING"
        }
        
        # SCHEMA_ADMINISTRATORS, S-1-5-21-<root-domain>-518
        try {
            $SchemaAdmins = Get-ADGroup "$RootSID-518"
        } catch {
            Write-Log "Warnung: Schema Admins Gruppe nicht gefunden (SID: $RootSID-518)" -Level "WARNING"
        }
        
        # ENTERPRISE_ADMINS, S-1-5-21-<root-domain>-519
        try {
            $EntAdmins = Get-ADGroup "$RootSID-519"
        } catch {
            Write-Log "Warnung: Enterprise Admins Gruppe nicht gefunden (SID: $RootSID-519)" -Level "WARNING"
        }
        
        # READONLY_DOMAIN_CONTROLLERS, S-1-5-21-<domain>-521
        try {
            $RODC = Get-ADGroup "$DomSID-521"
        } catch {
            Write-Log "Warnung: Read-Only Domain Controllers Gruppe nicht gefunden (SID: $DomSID-521)" -Level "WARNING"
        }
        
        # KEY_ADMINS, S-1-5-21-<domain>-526
        try {
            $KeyAdmins = Get-ADGroup "$DomSID-526"
        } catch {
            Write-Log "Warnung: Key Admins Gruppe nicht gefunden (SID: $DomSID-526)" -Level "WARNING"
        }
        
        # ENTERPRISE_KEY_ADMINS, S-1-5-21-<domain>-527
        try {
            $EntKeyAdmins = Get-ADGroup "$DomSID-527"
        } catch {
            Write-Log "Warnung: Enterprise Key Admins Gruppe nicht gefunden (SID: $DomSID-527)" -Level "WARNING"
        }
        
        # 1.3 All AdminSDHolder Objects
        $AllAdminSD = @(
            $BAdmins.Name,
            $BAccounts.Name,
            $BServerOps.Name,
            $BPrinterOps.Name,
            $BBackupOps.Name,
            $BReplicator.Name,
            $Administrator.Name,
            $krbtgt.Name,
            $DomAdmins.Name
        )
        
        # Füge optionale Gruppen hinzu, wenn vorhanden
        if ($DomCon) { $AllAdminSD += $DomCon.Name }
        if ($SchemaAdmins) { $AllAdminSD += $SchemaAdmins.Name }
        if ($EntAdmins) { $AllAdminSD += $EntAdmins.Name }
        if ($RODC) { $AllAdminSD += $RODC.Name }
        if ($KeyAdmins) { $AllAdminSD += $KeyAdmins.Name }
        if ($EntKeyAdmins) { $AllAdminSD += $EntKeyAdmins.Name }
        
        Write-Log "Geschützte Gruppen identifiziert: $($AllAdminSD.Count) Gruppen gefunden" -Level "INFO"
        
        # 2. Collect all Users, where AdminCount = 1
        Write-Log "Suche nach Benutzern mit AdminCount = 1..." -Level "INFO"
        $AllAdminCount = Get-ADUser -Filter {AdminCount -eq "1"} -Properties MemberOf, DistinguishedName
        
        Write-Log "Gefunden: $($AllAdminCount.Count) Benutzer mit AdminCount = 1" -Level "INFO"
        
        # Aktualisiere die ProgressBar
        $global:progressBar.Maximum = $AllAdminCount.Count
        $global:progressBar.Value = 0
        
        # Leere die Ergebnisse-ListView
        $global:resultListView.Items.Clear()
        
        $ProtectedCount = 0
        $OrphanedCount = 0
        $ProcessedCount = 0
        
        # 2.1 Report and Process all Users, where AdminCount = 1
        foreach ($User in $AllAdminCount) {
            $global:progressBar.Value++
            $ProcessedCount++
            
            Write-Log "Verarbeite Benutzer: $($User.Name) ($ProcessedCount von $($AllAdminCount.Count))" -Level "INFO"
            
            # Collect Group Memberships of the Users
            $AllGroups = (Get-ADPrincipalGroupMembership $User).Name
            
            # Combine User Groups and Protected Groups
            $AllTogether = @($AllGroups + $AllAdminSD)
            
            # Find Duplicates/Matches
            $Duplicates = $AllTogether | Group-Object | Where-Object { $_.Count -gt 1 } 
            
            $status = ""
            $action = ""
            
            # Exclude Administrator (RID-500) and krbtgt from Processing
            if (($User.Name -eq $Administrator.Name) -or ($User.Name -eq $krbtgt.Name)) {
                $status = "Geschützt (Standard)"
                $action = "Keine Aktion - bleibt geschützt"
                Write-Log "$($User.Name): Administrator oder krbtgt Konto - bleibt geschützt" -Level "INFO"
            } 
            else {
                if ($Duplicates) {
                    # User is in protected group(s)
                    $ProtectedCount++
                    $protectedGroups = ($Duplicates.Name -join ", ")
                    $status = "Geschützt (Gruppenmitglied)"
                    $action = "Keine Aktion - Mitglied in: $protectedGroups"
                    Write-Log "$($User.Name): Ist Mitglied in geschützten Gruppen: $protectedGroups" -Level "SUCCESS"
                } 
                else {
                    # User is not in any protected group
                    $OrphanedCount++
                    $status = "Verwaist"
                    
                    if ($Cleanup) {
                        # Reset admincount to "not set"
                        try {
                            Set-ADUser -Identity $User -Clear adminCount
                            
                            # Read ACL, set new ACL and write back ACL
                            $CN = $User.DistinguishedName
                            $GetAcl = Get-Acl -Path "AD:$CN"
                            $GetAcl.SetAccessRuleProtection($false, $true)
                            Set-Acl -Path "AD:$CN" -AclObject $GetAcl
                            
                            $action = "AdminCount zurückgesetzt und Vererbung aktiviert"
                            Write-Log "$($User.Name): AdminCount zurückgesetzt und Vererbung aktiviert" -Level "SUCCESS"
                        }
                        catch {
                            $action = "FEHLER: $($_.Exception.Message)"
                            Write-Log "$($User.Name): Fehler beim Zurücksetzen: $($_.Exception.Message)" -Level "ERROR"
                        }
                    }
                    else {
                        $action = "AdminCount sollte zurückgesetzt werden (Cleanup-Modus aktivieren)"
                        Write-Log "$($User.Name): Verwaister AdminCount = 1 sollte zurückgesetzt werden" -Level "WARNING"
                    }
                }
            }
            
            # Füge zu ListView hinzu
            $item = New-Object System.Windows.Forms.ListViewItem($User.Name)
            $item.SubItems.Add($status)
            $item.SubItems.Add($action)
            
            # Färbung je nach Status
            if ($status -eq "Geschützt (Standard)") {
                $item.BackColor = [System.Drawing.Color]::LightGray
            }
            elseif ($status -eq "Geschützt (Gruppenmitglied)") {
                $item.BackColor = [System.Drawing.Color]::LightGreen
            }
            elseif ($status -eq "Verwaist") {
                if ($Cleanup) {
                    $item.BackColor = [System.Drawing.Color]::LightBlue
                }
                else {
                    $item.BackColor = [System.Drawing.Color]::LightSalmon
                }
            }
            
            $global:resultListView.Items.Add($item)
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Abschlussbericht
        Write-Log "Abschlussbericht:" -Level "INFO"
        Write-Log "Gesamtzahl der verarbeiteten Benutzer: $($AllAdminCount.Count)" -Level "INFO"
        Write-Log "Geschützte Benutzer (in Gruppen): $ProtectedCount" -Level "INFO"
        Write-Log "Verwaiste AdminCount-Benutzer: $OrphanedCount" -Level "INFO"
        
        if ($Cleanup) {
            Write-Log "Cleanup abgeschlossen: $OrphanedCount Benutzer bereinigt" -Level "SUCCESS"
        }
        else {
            Write-Log "Report abgeschlossen: $OrphanedCount Benutzer könnten bereinigt werden" -Level "INFO"
        }
        
        # Aktiviere UI-Elemente wieder
        $global:reportButton.Enabled = $true
        $global:cleanupButton.Enabled = $true
        $global:exportButton.Enabled = $true
        
    }
    catch {
        Write-Log "Kritischer Fehler: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
        
        # Aktiviere UI-Elemente wieder
        $global:reportButton.Enabled = $true
        $global:cleanupButton.Enabled = $true
    }
}

# Funktion zum Exportieren der Ergebnisse
function Export-Results {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
    $SaveFileDialog.Title = "Speichern der Ergebnisse"
    $SaveFileDialog.FileName = "AdminCount-Cleanup_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    
    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $Results = @()
            
            foreach ($Item in $global:resultListView.Items) {
                $Results += [PSCustomObject]@{
                    Benutzername = $Item.Text
                    Status = $Item.SubItems[1].Text
                    Aktion = $Item.SubItems[2].Text
                }
            }
            
            $Results | Export-Csv -Path $SaveFileDialog.FileName -NoTypeInformation -Encoding UTF8
            
            Write-Log "Ergebnisse erfolgreich exportiert nach: $($SaveFileDialog.FileName)" -Level "SUCCESS"
        }
        catch {
            Write-Log "Fehler beim Exportieren: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

# GUI erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "AdminCount Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(900, 730)
$form.StartPosition = "CenterScreen"
$form.Icon = [System.Drawing.SystemIcons]::Shield

# Hauptpanel
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$form.Controls.Add($mainPanel)

# Header
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Active Directory AdminCount Cleanup Tool"
$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$headerLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerLabel.Height = 40
$mainPanel.Controls.Add($headerLabel)

# Beschreibung
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "Dieses Tool identifiziert und bereinigt Benutzer mit gesetztem AdminCount-Attribut, die nicht mehr in geschützten Gruppen sind."
$descriptionLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$descriptionLabel.Height = 30
$mainPanel.Controls.Add($descriptionLabel)

# Button-Panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$buttonPanel.Height = 50
$mainPanel.Controls.Add($buttonPanel)

# Report-Button
$global:reportButton = New-Object System.Windows.Forms.Button
$global:reportButton.Text = "Nur Bericht"
$global:reportButton.Location = New-Object System.Drawing.Point(0, 10)
$global:reportButton.Size = New-Object System.Drawing.Size(150, 30)
$global:reportButton.Add_Click({
    $global:reportButton.Enabled = $false
    $global:cleanupButton.Enabled = $false
    $global:exportButton.Enabled = $false
    Start-AdminCountCleanup
})
$buttonPanel.Controls.Add($global:reportButton)

# Cleanup-Button
$global:cleanupButton = New-Object System.Windows.Forms.Button
$global:cleanupButton.Text = "Bereinigen und Reparieren"
$global:cleanupButton.Location = New-Object System.Drawing.Point(160, 10)
$global:cleanupButton.Size = New-Object System.Drawing.Size(200, 30)
$global:cleanupButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Bist du sicher, dass du alle verwaisten AdminCount-Attribute bereinigen möchtest?", 
        "Warnung", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $global:reportButton.Enabled = $false
        $global:cleanupButton.Enabled = $false
        $global:exportButton.Enabled = $false
        Start-AdminCountCleanup -Cleanup
    }
})
$buttonPanel.Controls.Add($global:cleanupButton)

# Export-Button
$global:exportButton = New-Object System.Windows.Forms.Button
$global:exportButton.Text = "Ergebnisse exportieren"
$global:exportButton.Location = New-Object System.Drawing.Point(370, 10)
$global:exportButton.Size = New-Object System.Drawing.Size(170, 30)
$global:exportButton.Add_Click({
    Export-Results
})
$buttonPanel.Controls.Add($global:exportButton)

# Open Log-Button
$openLogButton = New-Object System.Windows.Forms.Button
$openLogButton.Text = "Log-Verzeichnis öffnen"
$openLogButton.Location = New-Object System.Drawing.Point(550, 10)
$openLogButton.Size = New-Object System.Drawing.Size(170, 30)
$openLogButton.Add_Click({
    Start-Process "explorer.exe" -ArgumentList $LogDir
})
$buttonPanel.Controls.Add($openLogButton)

# Status-Panel
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$statusPanel.Height = 40
$mainPanel.Controls.Add($statusPanel)

# Status-Label
$statusLabelDesc = New-Object System.Windows.Forms.Label
$statusLabelDesc.Text = "Status:"
$statusLabelDesc.Location = New-Object System.Drawing.Point(0, 10)
$statusLabelDesc.AutoSize = $true
$statusLabelDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$statusPanel.Controls.Add($statusLabelDesc)

$global:statusLabel = New-Object System.Windows.Forms.Label
$global:statusLabel.Text = "Bereit."
$global:statusLabel.Location = New-Object System.Drawing.Point(70, 10)
$global:statusLabel.AutoSize = $true
$statusPanel.Controls.Add($global:statusLabel)

# Fortschrittsanzeige
$global:progressBar = New-Object System.Windows.Forms.ProgressBar
$global:progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
$global:progressBar.Height = 20
$mainPanel.Controls.Add($global:progressBar)

# Ergebnisse-Label
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Text = "Ergebnisse:"
$resultsLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$resultsLabel.Height = 20
$resultsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$mainPanel.Controls.Add($resultsLabel)

# Ergebnisse-ListView
$global:resultListView = New-Object System.Windows.Forms.ListView
$global:resultListView.View = [System.Windows.Forms.View]::Details
$global:resultListView.FullRowSelect = $true
$global:resultListView.GridLines = $true
$global:resultListView.Height = 200
$global:resultListView.Dock = [System.Windows.Forms.DockStyle]::Top

# Spalten hinzufügen
$global:resultListView.Columns.Add("Benutzername", 200)
$global:resultListView.Columns.Add("Status", 200)
$global:resultListView.Columns.Add("Aktion", 450)

$mainPanel.Controls.Add($global:resultListView)

# Log-Label
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Text = "Protokoll:"
$logLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$logLabel.Height = 20
$logLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$mainPanel.Controls.Add($logLabel)

# Manuelle Positionierung des Log-Panels
$logPanel = New-Object System.Windows.Forms.Panel
$logPanel.Dock = [System.Windows.Forms.DockStyle]::None
# Feste Werte für Position und Größe
$panelTop = 450  # Position von oben
$logPanel.Location = New-Object System.Drawing.Point(10, $panelTop)
$logPanel.Size = New-Object System.Drawing.Size(860, 200)
$logPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                   [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                   [System.Windows.Forms.AnchorStyles]::Left -bor 
                   [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($logPanel)

# Log TextBox
$global:logTextBox = New-Object System.Windows.Forms.RichTextBox
$global:logTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$global:logTextBox.ReadOnly = $true
$global:logTextBox.BackColor = [System.Drawing.Color]::White
$global:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logPanel.Controls.Add($global:logTextBox)

# Copyright-Panel mit fester Position
$copyrightPanel = New-Object System.Windows.Forms.Panel
$copyrightPanel.Dock = [System.Windows.Forms.DockStyle]::None
$copyrightPanel.Location = New-Object System.Drawing.Point(10, 655)
$copyrightPanel.Size = New-Object System.Drawing.Size(860, 30)  # Breite 860, Höhe 30
$copyrightPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                         [System.Windows.Forms.AnchorStyles]::Left -bor 
                         [System.Windows.Forms.AnchorStyles]::Right
$mainPanel.Controls.Add($copyrightPanel)

# Copyright-Label
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Text = "$(Get-Date -Format 'yyyy') Jörn Walter - https://www.der-windows-papst.de."
$copyrightLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$copyrightLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
$copyrightPanel.Controls.Add($copyrightLabel)

# Domänenname anzeigen
$domainInfoLabel = New-Object System.Windows.Forms.Label
$domainInfoLabel.Text = "Aktuelle Domäne: $((Get-ADDomain).DNSRoot)"
$domainInfoLabel.Location = New-Object System.Drawing.Point(10, 625)
$domainInfoLabel.Size = New-Object System.Drawing.Size(500, 20)
$domainInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$domainInfoLabel.ForeColor = [System.Drawing.Color]::Navy
$domainInfoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$mainPanel.Controls.Add($domainInfoLabel)

# Initialisiere die Textbox mit Text, um zu prüfen, ob sie funktioniert
$global:logTextBox.AppendText("AdminCount Cleanup Tool bereit. Protokollierung aktiv.`r`n")

# Warten bis das Formular vollständig geladen ist, bevor die erste Meldung angezeigt wird
$form.Add_Shown({
    # Startmeldung schreiben - erst wenn Formular sichtbar ist
    Write-Log "AdminCount Cleanup Tool gestartet. Logdatei: $LogFile" -Level "INFO"
    Write-Log "Aktuelle Domäne: $((Get-ADDomain).DNSRoot)" -Level "INFO"
    Write-Log "Bereit. Wähle eine Aktion aus." -Level "INFO"
})

# Start des Formulars
$form.Add_Shown({
    $form.Activate()
    
    # Kurz Moment warten, damit alle Steuerelemente initialisiert sind
    Start-Sleep -Milliseconds 100
    
    # Test-Eintrag in das Log schreiben, um zu prüfen, ob es korrekt angezeigt wird
    $global:logTextBox.AppendText("GUI gestartet und bereit...`r`n")
    $global:logTextBox.SelectionStart = $global:logTextBox.Text.Length
    $global:logTextBox.ScrollToCaret()
    $global:logTextBox.Refresh()
})

# Hauptformular anzeigen
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
