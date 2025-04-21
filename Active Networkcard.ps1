<#
.SYNOPSIS
  Manage the Networkcards from DC
.DESCRIPTION
  Netzwerkkarten-Verwaltung mit besonderer Unterstützung für Domain Controller
.NOTES
  Version:        1.2
  Author:         Jörn Walter
  Creation Date:  2025-04-19
#>

# Funktion zum Überprüfen, ob das Skript mit Administratorrechten ausgeführt wird
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Überprüft, ob das Skript mit Administratorrechten ausgeführt wird
if (-not (Test-Admin)) {
    Write-Host "Starte mit Administratorrechten neu..." -ForegroundColor Yellow
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Variable für den aktuellen Anzeigemodus
$global:isGridView = $true
$global:adapters = @()
$global:selectedAdapter = $null
$global:isDomainController = $false
$global:debugMode = $true  # Debugging-Modus aktivieren

# Überprüfen, ob es sich um einen Domain Controller handelt
function Test-IsDomainController {
    try {
        $role = Get-WmiObject -Class Win32_ComputerSystem -Property DomainRole -ErrorAction Stop
        return ($role.DomainRole -eq 4 -or $role.DomainRole -eq 5)
    }
    catch {
        return $false
    }
}
# Überprüfen, ob das Skript auf einem DC läuft
$global:isDomainController = Test-IsDomainController
if ($global:isDomainController) {
    Write-Host "Läuft auf einem Domain Controller." -ForegroundColor Cyan
}

# Funktion für Debug-Logging
function Write-DebugLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Color = "Gray"
    )
    
    if ($global:debugMode) {
        Write-Host "DEBUG: $Message" -ForegroundColor $Color
    }
}

#region UI-Erstellung
# Erstellen des Hauptfensters
$form = New-Object System.Windows.Forms.Form
$form.Text = "Aktive Netzwerkkarten"
$form.Size = New-Object System.Drawing.Size(1100, 600)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Erstellen des Ausgabebereichs
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 40)
$outputBox.Size = New-Object System.Drawing.Size(1065, 420)
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($outputBox)

# Erstellen der Listbox für die Netzwerkadapter
$adapterListBox = New-Object System.Windows.Forms.ListBox
$adapterListBox.Location = New-Object System.Drawing.Point(10, 40)
$adapterListBox.Size = New-Object System.Drawing.Size(250, 420)
$adapterListBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($adapterListBox)

# Erstellen des Detail-Panels für die vertikale Anzeige
$detailPanel = New-Object System.Windows.Forms.Panel
$detailPanel.Location = New-Object System.Drawing.Point(270, 40)
$detailPanel.Size = New-Object System.Drawing.Size(805, 420)
$detailPanel.BackColor = [System.Drawing.Color]::White
$detailPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($detailPanel)

# Erstellen der Beschriftung
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(300, 20)
$label.Text = "Netzwerkkarten-Informationen:"
$form.Controls.Add($label)

# Erstellen des Aktualisierungsbuttons
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(10, 470)
$refreshButton.Size = New-Object System.Drawing.Size(150, 30)
$refreshButton.Text = "Aktualisieren"
$form.Controls.Add($refreshButton)

# Erstellen des Detailansicht-Buttons
$detailButton = New-Object System.Windows.Forms.Button
$detailButton.Location = New-Object System.Drawing.Point(170, 470)
$detailButton.Size = New-Object System.Drawing.Size(150, 30)
$detailButton.Text = "Text-Ansicht"
$form.Controls.Add($detailButton)

# Erstellen des Export-Buttons
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(330, 470)
$exportButton.Size = New-Object System.Drawing.Size(150, 30)
$exportButton.Text = "Exportieren"
$form.Controls.Add($exportButton)

# Erstellen des Neustart-Buttons
$restartButton = New-Object System.Windows.Forms.Button
$restartButton.Location = New-Object System.Drawing.Point(490, 470)
$restartButton.Size = New-Object System.Drawing.Size(150, 30)
$restartButton.Text = "Adapter neustarten"
$form.Controls.Add($restartButton)

# Erstellen des "Autostart-Task"-Buttons
$autostartButton = New-Object System.Windows.Forms.Button
$autostartButton.Location = New-Object System.Drawing.Point(650, 470)
$autostartButton.Size = New-Object System.Drawing.Size(200, 30)
$autostartButton.Text = "Autostart-Task erstellen"
$form.Controls.Add($autostartButton)

# Diagnose-Button hinzufügen
$diagnoseButton = New-Object System.Windows.Forms.Button
$diagnoseButton.Location = New-Object System.Drawing.Point(860, 470)
$diagnoseButton.Size = New-Object System.Drawing.Size(150, 30)
$diagnoseButton.Text = "Diagnose"
$form.Controls.Add($diagnoseButton)

# Erstellen des Status-Labels
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 520)
$statusLabel.Size = New-Object System.Drawing.Size(1065, 20)
$statusLabel.Text = "Bereit"
$form.Controls.Add($statusLabel)

# Erstellen des Copyright-Labels
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Location = New-Object System.Drawing.Point(10, 545)
$copyrightLabel.Size = New-Object System.Drawing.Size(1065, 20)
$copyrightLabel.Text = "© 2025 Jörn Walter - https://www.der-windows-papst.de"
$copyrightLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($copyrightLabel)
#endregion

# Funktion zum Aktualisieren des Detail-Panels links
function Update-DetailPanel {
    try {
        # Leeren des Detail-Panels
        $detailPanel.Controls.Clear()
        
        Write-DebugLog "Selected index: $($adapterListBox.SelectedIndex), Adapter count: $($global:adapters.Count)" -Color "Yellow"
        
        if ($adapterListBox.SelectedIndex -ge 0 -and $adapterListBox.SelectedIndex -lt $global:adapters.Count) {
            $adapter = $global:adapters[$adapterListBox.SelectedIndex]
            Write-DebugLog "Selected adapter: $($adapter.Name)" -Color "Green"
            
            # Debug-Info anzeigen
            $debugInfo = New-Object System.Windows.Forms.Label
            $debugInfo.Location = New-Object System.Drawing.Point(20, 380)
            $debugInfo.Size = New-Object System.Drawing.Size(760, 24)
            $propertyNames = $adapter | Get-Member -MemberType Property,NoteProperty | Select-Object -ExpandProperty Name
            if ($propertyNames) {
            $debugInfo.Text = "Available: $([string]::Join(", ", $propertyNames))"
            } else {
            $debugInfo.Text = "Debug: No adapter properties available"
            }
            $debugInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8)
            $debugInfo.ForeColor = [System.Drawing.Color]::Blue
            $detailPanel.Controls.Add($debugInfo)
            
            # Name - Label und Wert
            $nameLabelTitle = New-Object System.Windows.Forms.Label
            $nameLabelTitle.Location = New-Object System.Drawing.Point(20, 20)
            $nameLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $nameLabelTitle.Text = "Name:"
            $nameLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($nameLabelTitle)
            
            $nameValue = New-Object System.Windows.Forms.Label
            $nameValue.Location = New-Object System.Drawing.Point(180, 20)
            $nameValue.Size = New-Object System.Drawing.Size(600, 24)
            $nameValue.Text = if ($adapter.Name) { $adapter.Name } else { "Nicht verfügbar" }
            $nameValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $detailPanel.Controls.Add($nameValue)
            
            # Beschreibung - Label und Wert
            $descLabelTitle = New-Object System.Windows.Forms.Label
            $descLabelTitle.Location = New-Object System.Drawing.Point(20, 65)
            $descLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $descLabelTitle.Text = "Beschreibung:"
            $descLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($descLabelTitle)
            
            $descValue = New-Object System.Windows.Forms.Label
            $descValue.Location = New-Object System.Drawing.Point(180, 65)
            $descValue.Size = New-Object System.Drawing.Size(600, 24)
            $descValue.Text = if ($adapter.InterfaceDescription) { $adapter.InterfaceDescription } else { "Nicht verfügbar" }
            $descValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $detailPanel.Controls.Add($descValue)
            
            # Status - Label und Wert
            $statusLabelTitle = New-Object System.Windows.Forms.Label
            $statusLabelTitle.Location = New-Object System.Drawing.Point(20, 110)
            $statusLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $statusLabelTitle.Text = "Status:"
            $statusLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($statusLabelTitle)
            
            $statusValue = New-Object System.Windows.Forms.Label
            $statusValue.Location = New-Object System.Drawing.Point(180, 110)
            $statusValue.Size = New-Object System.Drawing.Size(600, 24)
            $statusValue.Text = if ($adapter.ConnectionState) { $adapter.ConnectionState } else { 
                if ($adapter.Status) { $adapter.Status } else { "Unbekannt" }
            }
            $statusValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            if ($statusValue.Text -eq "Verbunden" -or $statusValue.Text -eq "Up") {
                $statusValue.ForeColor = [System.Drawing.Color]::Green
                $statusValue.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            } else {
                $statusValue.ForeColor = [System.Drawing.Color]::Red
            }
            $detailPanel.Controls.Add($statusValue)
            
            # MAC-Adresse - Label und Wert
            $macLabelTitle = New-Object System.Windows.Forms.Label
            $macLabelTitle.Location = New-Object System.Drawing.Point(20, 155)
            $macLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $macLabelTitle.Text = "MAC-Adresse:"
            $macLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($macLabelTitle)
            
            $macValue = New-Object System.Windows.Forms.Label
            $macValue.Location = New-Object System.Drawing.Point(180, 155)
            $macValue.Size = New-Object System.Drawing.Size(600, 24)
            $macValue.Text = if ($adapter.MacAddress) { $adapter.MacAddress } else { "Nicht verfügbar" }
            $macValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $detailPanel.Controls.Add($macValue)
            
            # Link-Geschwindigkeit - Label und Wert
            $speedLabelTitle = New-Object System.Windows.Forms.Label
            $speedLabelTitle.Location = New-Object System.Drawing.Point(20, 200)
            $speedLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $speedLabelTitle.Text = "Link-Geschwindigkeit:"
            $speedLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($speedLabelTitle)
            
            $speedValue = New-Object System.Windows.Forms.Label
            $speedValue.Location = New-Object System.Drawing.Point(180, 200)
            $speedValue.Size = New-Object System.Drawing.Size(600, 24)
            $speedValue.Text = if ($adapter.LinkSpeed) { $adapter.LinkSpeed } else { "Nicht verfügbar" }
            $speedValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $detailPanel.Controls.Add($speedValue)
            
            # IPv4-Adresse - Label und Wert
            $ipv4LabelTitle = New-Object System.Windows.Forms.Label
            $ipv4LabelTitle.Location = New-Object System.Drawing.Point(20, 245)
            $ipv4LabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $ipv4LabelTitle.Text = "IPv4-Adresse:"
            $ipv4LabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($ipv4LabelTitle)
            
            $ipv4Value = New-Object System.Windows.Forms.Label
            $ipv4Value.Location = New-Object System.Drawing.Point(180, 245)
            $ipv4Value.Size = New-Object System.Drawing.Size(600, 24)
            $ipv4Value.Text = if ($adapter.IPv4Address) { $adapter.IPv4Address } else { "Nicht verfügbar" }
            $ipv4Value.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $detailPanel.Controls.Add($ipv4Value)
            
            # IPv6-Adresse - Label und Wert
            $ipv6LabelTitle = New-Object System.Windows.Forms.Label
            $ipv6LabelTitle.Location = New-Object System.Drawing.Point(20, 290)
            $ipv6LabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $ipv6LabelTitle.Text = "IPv6-Adresse:"
            $ipv6LabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($ipv6LabelTitle)
            
            $ipv6Value = New-Object System.Windows.Forms.Label
            $ipv6Value.Location = New-Object System.Drawing.Point(180, 290)
            $ipv6Value.Size = New-Object System.Drawing.Size(600, 40)
            $ipv6Value.Text = if ($adapter.IPv6Address) { $adapter.IPv6Address } else { "Nicht verfügbar" }
            $ipv6Value.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $ipv6Value.AutoSize = $true
            $ipv6Value.MaximumSize = New-Object System.Drawing.Size(600, 0)
            $detailPanel.Controls.Add($ipv6Value)
            
            # Gateway - Label und Wert
            $gatewayLabelTitle = New-Object System.Windows.Forms.Label
            $gatewayLabelTitle.Location = New-Object System.Drawing.Point(20, 340)
            $gatewayLabelTitle.Size = New-Object System.Drawing.Size(150, 24)
            $gatewayLabelTitle.Text = "Gateway:"
            $gatewayLabelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $detailPanel.Controls.Add($gatewayLabelTitle)
            
            $gatewayValue = New-Object System.Windows.Forms.Label
            $gatewayValue.Location = New-Object System.Drawing.Point(180, 340)
            $gatewayValue.Size = New-Object System.Drawing.Size(600, 30)
            $gatewayValue.Text = if ($adapter.Gateway) { $adapter.Gateway } else { "Nicht verfügbar" }
            $gatewayValue.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $gatewayValue.AutoSize = $true
            $gatewayValue.MaximumSize = New-Object System.Drawing.Size(600, 0)
            $detailPanel.Controls.Add($gatewayValue)
        }
        else {
            Write-DebugLog "Kein Adapter ausgewählt oder ungültiger Index" -Color "Red"
            $errorMsg = New-Object System.Windows.Forms.Label
            $errorMsg.Location = New-Object System.Drawing.Point(20, 20)
            $errorMsg.Size = New-Object System.Drawing.Size(760, 380)
            $errorMsg.Text = "Kein Adapter ausgewählt oder keine Adapter gefunden."
            $errorMsg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
            $errorMsg.ForeColor = [System.Drawing.Color]::Red
            $detailPanel.Controls.Add($errorMsg)
        }
    }
    catch {
        Write-DebugLog "Fehler in Update-DetailPanel: $_" -Color "Red"
        $errorMsg = New-Object System.Windows.Forms.Label
        $errorMsg.Location = New-Object System.Drawing.Point(20, 20)
        $errorMsg.Size = New-Object System.Drawing.Size(760, 380)
        $errorMsg.Text = "Fehler beim Anzeigen der Adapterdetails: $_"
        $errorMsg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $errorMsg.ForeColor = [System.Drawing.Color]::Red
        $detailPanel.Controls.Add($errorMsg)
        
        $statusLabel.Text = "Fehler beim Anzeigen der Adapterdetails: $_"
    }
}

# Funktion zum Abrufen der Netzwerkkarten-Informationen
function Get-NetworkAdapterInfo {
    try {
        $statusLabel.Text = "Suche nach Netzwerkadaptern..."
        Write-DebugLog "Starte Suche nach Netzwerkadaptern..." -Color "Cyan"
        
        # Verwenden einer robusteren Methode, die auch auf DCs funktioniert
        $adapterResults = @()
        
        # Versuchen, Get-NetAdapter zu verwenden (bevorzugt)
        try {
            Write-DebugLog "Versuche Get-NetAdapter zu verwenden..." -Color "Yellow"
            $netAdapters = Get-NetAdapter | Where-Object { $_.Status -ne "Disabled" } -ErrorAction Stop
            Write-DebugLog "Get-NetAdapter Erfolg: $($netAdapters.Count) Adapter gefunden" -Color "Green"
            
            foreach ($adapter in $netAdapters) {
                try {
                    Write-DebugLog "Verarbeite Adapter: $($adapter.Name)" -Color "White"
                    
                    # Alle IP-Adressen abrufen, nicht nur IPv4
                    $ipAddresses = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
                    
                    # IPv4-Adressen filtern
                    $ipv4 = $ipAddresses | Where-Object { $_.AddressFamily -eq "IPv4" } | Select-Object -ExpandProperty IPAddress
                    
                    # IPv6-Adressen filtern
                    $ipv6 = $ipAddresses | Where-Object { $_.AddressFamily -eq "IPv6" } | Select-Object -ExpandProperty IPAddress
                    
                    $gateway = Get-NetRoute -InterfaceIndex $adapter.ifIndex -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue | 
                               Select-Object -ExpandProperty NextHop
                    
                    $dns = Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue | 
                           Where-Object {$_.AddressFamily -eq 2} | 
                           Select-Object -ExpandProperty ServerAddresses
                    
                    # Statistik holen - mit Fehlerbehandlung
                    try {
                        $stats = Get-NetAdapterStatistics -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
                        $activity = if ($stats) {
                            "↑ $([math]::Round($stats.SentBytes/1MB, 2)) MB / ↓ $([math]::Round($stats.ReceivedBytes/1MB, 2)) MB"
                        } else {
                            "Keine Daten"
                        }
                    } catch {
                        Write-DebugLog "Fehler bei Statistik für $($adapter.Name): $_" -Color "Yellow"
                        $activity = "Nicht verfügbar"
                    }
                    
                    $connectionState = if ($adapter.Status -eq "Up") { "Verbunden" } 
                                    elseif ($adapter.Status -eq "Disconnected") { "Getrennt" }
                                    else { $adapter.Status }
                    
                    Write-DebugLog "Adapter $($adapter.Name) Status: $connectionState" -Color "White"
                    
                    $adapterResults += [PSCustomObject]@{
                        Name = $adapter.Name
                        InterfaceDescription = $adapter.InterfaceDescription
                        Status = $adapter.Status
                        ConnectionState = $connectionState
                        MacAddress = $adapter.MacAddress
                        LinkSpeed = $adapter.LinkSpeed
                        IPv4Address = ($ipv4 -join ", ")
                        IPv6Address = ($ipv6 -join ", ")
                        Gateway = ($gateway -join ", ")
                        DNS = ($dns -join ", ")
                        Aktivität = $activity
                    }
                } catch {
                    # Bei Fehler bei einzelnem Adapter fortfahren, aber Logging
                    Write-DebugLog "Fehler bei Adapter $($adapter.Name): $_" -Color "Yellow"
                }
            }
        } catch {
            Write-DebugLog "Get-NetAdapter nicht verfügbar oder Fehler: $_" -Color "Yellow"
            # Wir werden die alternativen Methoden unten verwenden
        }
        
        # Wenn Get-NetAdapter keine Ergebnisse liefert, alternativen Weg verwenden
        # Wenn immer noch keine Adapter gefunden wurden, letzte Alternative versuchen
if ($adapterResults.Count -eq 0) {
    Write-DebugLog "Keine Adapter gefunden, versuche letzte Alternative..." -Color "Yellow"
    
    # Eine einfache Methode, die fast immer funktioniert
    $simpleAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()
    Write-DebugLog "System.Net.NetworkInformation: $($simpleAdapters.Count) Adapter gefunden" -Color "Green"
    
    foreach ($adapter in $simpleAdapters) {
        try {
            Write-DebugLog "Verarbeite einfachen Adapter: $($adapter.Name)" -Color "White"
            
            $ipProps = $adapter.GetIPProperties()
            $ipv4Addresses = $ipProps.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.Address.ToString() }
            $ipv6Addresses = $ipProps.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetworkV6' } | ForEach-Object { $_.Address.ToString() }
            $gateways = $ipProps.GatewayAddresses | ForEach-Object { $_.Address.ToString() }
            $dnsServers = $ipProps.DnsAddresses | ForEach-Object { $_.ToString() }
            
            $connectionState = if ($adapter.OperationalStatus -eq 'Up') { "Verbunden" } else { $adapter.OperationalStatus.ToString() }
            
            Write-DebugLog "Einfacher Adapter $($adapter.Name) Status: $connectionState" -Color "White"
            
            $adapterResults += [PSCustomObject]@{
                Name = $adapter.Name
                InterfaceDescription = $adapter.Description
                Status = $adapter.OperationalStatus
                ConnectionState = $connectionState
                MacAddress = $adapter.GetPhysicalAddress().ToString() -replace '(..)','$1-' -replace '-$',''
                LinkSpeed = if ($adapter.Speed -gt 0) { "$([math]::Round($adapter.Speed / 1000000, 0)) Mbps" } else { "Unbekannt" }
                IPv4Address = ($ipv4Addresses -join ", ")
                IPv6Address = ($ipv6Addresses -join ", ")
                Gateway = ($gateways -join ", ")
                DNS = ($dnsServers -join ", ")
                Aktivität = "Nicht verfügbar"
            }
        } catch {
            Write-DebugLog "Fehler bei einfachem Adapter $($adapter.Name): $_" -Color "Yellow"
        }
    }
}
            
            # Eine einfache Methode, die fast immer funktioniert
            $simpleAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()
            Write-DebugLog "System.Net.NetworkInformation: $($simpleAdapters.Count) Adapter gefunden" -Color "Green"
            
            foreach ($adapter in $simpleAdapters) {
                try {
                    Write-DebugLog "Verarbeite einfachen Adapter: $($adapter.Name)" -Color "White"
                    
                    $ipProps = $adapter.GetIPProperties()
                    $ipv4Addresses = $ipProps.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.Address.ToString() }
                    $ipv6Addresses = $ipProps.UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetworkV6' } | ForEach-Object { $_.Address.ToString() }
                    $gateways = $ipProps.GatewayAddresses | ForEach-Object { $_.Address.ToString() }
                    $dnsServers = $ipProps.DnsAddresses | ForEach-Object { $_.ToString() }
                    
                    $connectionState = if ($adapter.OperationalStatus -eq 'Up') { "Verbunden" } else { $adapter.OperationalStatus.ToString() }
                    
                    Write-DebugLog "Einfacher Adapter $($adapter.Name) Status: $connectionState" -Color "White"
                    
                    $adapterResults += [PSCustomObject]@{
                        Name = $adapter.Name
                        InterfaceDescription = $adapter.Description
                        Status = $adapter.OperationalStatus
                        ConnectionState = $connectionState
                        MacAddress = $adapter.GetPhysicalAddress().ToString() -replace '(..)','$1-' -replace '-$',''
                        LinkSpeed = if ($adapter.Speed -gt 0) { "$([math]::Round($adapter.Speed / 1000000, 0)) Mbps" } else { "Unbekannt" }
                        IPv4Address = ($ipv4Addresses -join ", ")
                        IPv6Address = ($ipv6Addresses -join ", ")
                        Gateway = ($gateways -join ", ")
                        DNS = ($dnsServers -join ", ")
                        Aktivität = "Nicht verfügbar"
                    }
                } catch {
                    Write-DebugLog "Fehler bei einfachem Adapter $($adapter.Name): $_" -Color "Yellow"
                }
            }
        
        
        # Suche speziell nach "Ethernet0"
if ($adapterResults.Count -eq 0) {
    Write-DebugLog "Suche nach 'ethernet0'..." -Color "Magenta"
    
    # Versuche, einen minimalen Eintrag für ethernet0 zu erstellen
    $ethernet0 = [PSCustomObject]@{
        Name = "ethernet0"
        InterfaceDescription = "Netzwerkkarte"
        Status = "Unbekannt"
        ConnectionState = "Unbekannt"
        MacAddress = "Nicht verfügbar"
        LinkSpeed = "Nicht verfügbar"
        IPv4Address = "Nicht verfügbar"
        IPv6Address = "Nicht verfügbar"
        Gateway = "Nicht verfügbar"
        DNS = "Nicht verfügbar"
        Aktivität = "Nicht verfügbar"
    }
    
    # Versuche, weitere Informationen über WMI zu holen
    try {
        $wmiEthernet0 = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { 
            $_.Name -eq "ethernet0" -or $_.NetConnectionID -eq "ethernet0" 
        }
        
        if ($wmiEthernet0) {
            Write-DebugLog "Gefunden über WMI: $($wmiEthernet0.NetConnectionID)" -Color "Green"
            
            $config = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.Index -eq $wmiEthernet0.Index }
            
            if ($config) {
                $ethernet0.InterfaceDescription = $wmiEthernet0.Description
                $ethernet0.Status = if ($wmiEthernet0.NetEnabled) { "Up" } else { "Down" }
                $ethernet0.ConnectionState = if ($wmiEthernet0.NetEnabled) { "Verbunden" } else { "Getrennt" }
                $ethernet0.MacAddress = $config.MACAddress
                $ethernet0.IPv4Address = ($config.IPAddress | Where-Object { $_ -like "*.*" }) -join ", "
                $ethernet0.IPv6Address = ($config.IPAddress | Where-Object { $_ -like "*:*" }) -join ", "
                $ethernet0.Gateway = ($config.DefaultIPGateway) -join ", "
                $ethernet0.DNS = ($config.DNSServerSearchOrder) -join ", "
            }
        }
    } catch {
        Write-DebugLog "Fehler bei der Suche nach ethernet0: $_" -Color "Red"
    }
    
    $adapterResults += $ethernet0
}

# Filtere doppelte Adapter heraus
$uniqueAdapters = @{}
foreach ($adapter in $adapterResults) {
    if (-not $uniqueAdapters.ContainsKey($adapter.Name)) {
        $uniqueAdapters[$adapter.Name] = $adapter
    }
}

# Alphabetisch nach Namen sortieren
$adapters = $uniqueAdapters.Values | Sort-Object -Property Name

Write-DebugLog "Finale Adapter-Liste: $($adapters.Count) Adapter" -Color "Green"
foreach ($adapter in $adapters) {
    Write-DebugLog "  - $($adapter.Name) ($($adapter.ConnectionState))" -Color "White"
}

$statusLabel.Text = "Netzwerkadapter wurden gefunden: $($adapters.Count)"
return $adapters
}
catch {
    Write-DebugLog "Kritischer Fehler: $_" -Color "Red"
    Write-Host "Kritischer Fehler: $_" -ForegroundColor Red
    return "Fehler beim Abrufen der Netzwerkkarten-Informationen: $_"
}
}

# Funktion zum Neustarten eines Netzwerkadapters
function Restart-NetworkAdapter {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AdapterName
    )
    
    try {
        # Administratorrechte überprüfen
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            $statusLabel.Text = "Fehler: Administratorrechte erforderlich, um Netzwerkadapter neu zu starten!"
            return $false
        }
        
        $statusLabel.Text = "Deaktiviere Netzwerkadapter '$AdapterName'..."
        Write-DebugLog "Deaktiviere Adapter: $AdapterName" -Color "Yellow"
        Disable-NetAdapter -Name $AdapterName -Confirm:$false
        Start-Sleep -Seconds 2
        
        $statusLabel.Text = "Aktiviere Netzwerkadapter '$AdapterName'..."
        Write-DebugLog "Aktiviere Adapter: $AdapterName" -Color "Yellow"
        Enable-NetAdapter -Name $AdapterName -Confirm:$false
        Start-Sleep -Seconds 3
        
        # Automatische Aktualisierung nach dem Neustart
        Update-Display
        
        $statusLabel.Text = "Netzwerkadapter '$AdapterName' wurde erfolgreich neu gestartet."
        return $true
    }
    catch {
        Write-DebugLog "Fehler beim Neustarten des Adapters: $_" -Color "Red"
        $statusLabel.Text = "Fehler beim Neustarten des Netzwerkadapters: $_"
        return $false
    }
}

# Funktion zum Erstellen einer geplanten Aufgabe für den Neustart eines Netzwerkadapters beim Systemstart
function Create-AdapterRestartTask {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AdapterName
    )
    
    try {
        # Administratorrechte überprüfen
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin) {
            $statusLabel.Text = "Fehler: Administratorrechte erforderlich, um die Aufgabe zu erstellen!"
            return $false
        }
        
        # Skriptpfad für die temporäre Datei
        $scriptFolder = "$env:ProgramData\AdapterRestart"
        $scriptPath = "$scriptFolder\RestartAdapter_$($AdapterName.Replace(' ', '_')).ps1"
        
        # Ordner erstellen, falls er nicht existiert
        if (-not (Test-Path $scriptFolder)) {
            New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null
        }
        
        # Sicheren Adaptername für Dateinamen erstellen
        $safeAdapterName = $AdapterName.Replace(' ', '_')
        
        # PS1-Skript zum Neustarten des Adapters erstellen
        $scriptContent = @"
# Skript zum Neustart des Netzwerkadapters mit netsh
# Datei: C:\ProgramData\AdapterRestart\RestartAdapter_$safeAdapterName.ps1

# Transkript starten
Start-Transcript -Path "C:\ProgramData\AdapterRestart\RestartAdapter_Log_$safeAdapterName.txt" -Append

# Startmeldung mit aktuellem Datum
Write-Output "Neustart des Netzwerkadapters '$AdapterName' gestartet am `$(Get-Date -Format 'MM.dd.yyyy HH:mm:ss')"

try {
    # Netzwerkadapter mit netsh deaktivieren
    Write-Output "Deaktiviere Netzwerkadapter '$AdapterName'..."
    netsh interface set interface name="$AdapterName" admin=disabled
    
    # Netzwerkadapter mit netsh aktivieren
    Write-Output "Aktiviere Netzwerkadapter '$AdapterName'..."
    netsh interface set interface name="$AdapterName" admin=enabled
    
    # Erfolg melden
    Write-Output "Netzwerkadapter '$AdapterName' wurde erfolgreich neu gestartet."
}
catch {
    # Fehler ausgeben
    Write-Output "Fehler beim Neustarten des Netzwerkadapters:"
    Write-Output "Fehlermeldung: `$(`$_.Exception.Message)"
}
finally {
    # Transkript beenden
    Stop-Transcript
   
    # Beenden mit Exit-Code 0
    exit 0
}
"@
        
        # Skript in die Datei schreiben
        $scriptContent | Out-File -FilePath $scriptPath -Encoding utf8 -Force
        
        # Erstellen der geplanten Aufgabe
        $taskName = "RestartAdapter_$safeAdapterName"
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        
        # Vorhandene Aufgabe entfernen, falls sie existiert
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Neue Aufgabe registrieren
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal
        
        $statusLabel.Text = "Autostart-Task für Netzwerkadapter '$AdapterName' wurde erfolgreich erstellt."
        return $true
    }
    catch {
        Write-DebugLog "Fehler beim Erstellen der Autostart-Aufgabe: $_" -Color "Red"
        $statusLabel.Text = "Fehler beim Erstellen der Autostart-Aufgabe: $_"
        return $false
    }
}

# Diagnose-Funktion
function Show-DiagnosticInfo {
    $diagForm = New-Object System.Windows.Forms.Form
    $diagForm.Text = "Netzwerkadapter-Diagnose"
    $diagForm.Size = New-Object System.Drawing.Size(800, 600)
    $diagForm.StartPosition = "CenterScreen"
    
    $diagTextBox = New-Object System.Windows.Forms.RichTextBox
    $diagTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $diagTextBox.Size = New-Object System.Drawing.Size(765, 505)
    $diagTextBox.ReadOnly = $true
    $diagTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $diagForm.Controls.Add($diagTextBox)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(350, 525)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Schließen"
    $closeButton.Add_Click({ $diagForm.Close() })
    $diagForm.Controls.Add($closeButton)
    
    $diagTextBox.AppendText("--- SYSTEM-INFORMATIONEN ---`r`n")
    $diagTextBox.AppendText("Computername: $env:COMPUTERNAME`r`n")
    $diagTextBox.AppendText("Betriebssystem: $(Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption)`r`n")
    $diagTextBox.AppendText("Domain Controller: $($global:isDomainController)`r`n")
    $diagTextBox.AppendText("PowerShell Version: $($PSVersionTable.PSVersion)`r`n`r`n")
    
    $diagTextBox.AppendText("--- NETZWERKADAPTER-VERFÜGBARKEIT ---`r`n")
    try {
        $netAdapters = Get-NetAdapter -ErrorAction Stop
        $diagTextBox.AppendText("Get-NetAdapter verfügbar: JA`r`n")
        $diagTextBox.AppendText("Anzahl der Adapter: $($netAdapters.Count)`r`n")
        foreach ($adapter in $netAdapters) {
            $diagTextBox.AppendText("- $($adapter.Name) (Status: $($adapter.Status))`r`n")
        }
    }
    catch {
        $diagTextBox.AppendText("Get-NetAdapter verfügbar: NEIN`r`n")
        $diagTextBox.AppendText("Fehler: $_`r`n")
    }
    
    $diagTextBox.AppendText("`r`n--- SYSTEM.NET NETZWERKADAPTER ---`r`n")
    try {
        $sysNetAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()
        $diagTextBox.AppendText("Anzahl der System.Net Adapter: $($sysNetAdapters.Count)`r`n")
        foreach ($adapter in $sysNetAdapters) {
            $diagTextBox.AppendText("- $($adapter.Name) (Status: $($adapter.OperationalStatus))`r`n")
        }
    }
    catch {
        $diagTextBox.AppendText("Fehler bei System.Net-Abfrage: $_`r`n")
    }
    
    $diagTextBox.AppendText("`r`n--- WMI-NETZWERKADAPTER ---`r`n")
    try {
        $wmiAdapters = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.NetEnabled -eq $true -or $_.Name -like "*ethernet*" }
        $diagTextBox.AppendText("Anzahl der WMI-Adapter: $($wmiAdapters.Count)`r`n")
        foreach ($adapter in $wmiAdapters) {
            $diagTextBox.AppendText("- $($adapter.NetConnectionID) (Index: $($adapter.Index), Enabled: $($adapter.NetEnabled))`r`n")
        }
    }
    catch {
        $diagTextBox.AppendText("Fehler bei WMI-Abfrage: $_`r`n")
    }
    
    $diagTextBox.AppendText("`r`n--- ETHERNET0 SUCHE ---`r`n")
    # WMI-Methode
    try {
        $ethernet0WMI = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { 
            $_.Name -eq "ethernet0" -or $_.NetConnectionID -eq "ethernet0" 
        }
        
        if ($ethernet0WMI) {
            $diagTextBox.AppendText("ethernet0 über WMI gefunden: JA`r`n")
            $diagTextBox.AppendText("Description: $($ethernet0WMI.Description)`r`n")
            $diagTextBox.AppendText("DeviceID: $($ethernet0WMI.DeviceID)`r`n")
            $diagTextBox.AppendText("Index: $($ethernet0WMI.Index)`r`n")
            $diagTextBox.AppendText("NetEnabled: $($ethernet0WMI.NetEnabled)`r`n")
            
            $config = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.Index -eq $ethernet0WMI.Index }
            if ($config) {
                $diagTextBox.AppendText("MAC: $($config.MACAddress)`r`n")
                $diagTextBox.AppendText("IP: $($config.IPAddress -join ', ')`r`n")
            }
        } else {
            $diagTextBox.AppendText("ethernet0 über WMI gefunden: NEIN`r`n")
        }
    } catch {
        $diagTextBox.AppendText("Fehler bei ethernet0 WMI-Suche: $_`r`n")
    }
    
    # NetAdapter-Methode
    try {
        $ethernet0Net = Get-NetAdapter | Where-Object { $_.Name -eq "ethernet0" } -ErrorAction SilentlyContinue
        if ($ethernet0Net) {
            $diagTextBox.AppendText("ethernet0 über Get-NetAdapter gefunden: JA`r`n")
            $diagTextBox.AppendText("Status: $($ethernet0Net.Status)`r`n")
            $diagTextBox.AppendText("InterfaceDescription: $($ethernet0Net.InterfaceDescription)`r`n")
            $diagTextBox.AppendText("ifIndex: $($ethernet0Net.ifIndex)`r`n")
        } else {
            $diagTextBox.AppendText("ethernet0 über Get-NetAdapter gefunden: NEIN`r`n")
        }
    } catch {
        $diagTextBox.AppendText("Fehler bei ethernet0 NetAdapter-Suche: $_`r`n")
    }
    
    $diagTextBox.AppendText("`r`n--- UI-STATUS ---`r`n")
    $diagTextBox.AppendText("Listbox sichtbar: $($adapterListBox.Visible)`r`n")
    $diagTextBox.AppendText("DetailPanel sichtbar: $($detailPanel.Visible)`r`n")
    $diagTextBox.AppendText("ListBox Items: $($adapterListBox.Items.Count)`r`n")
    $diagTextBox.AppendText("Ausgewählter Index: $($adapterListBox.SelectedIndex)`r`n")
    $diagTextBox.AppendText("Global Adapters Count: $($global:adapters.Count)`r`n")
    
    $diagForm.ShowDialog()
}

# Funktion zum Aktualisieren der Anzeige
function Update-Display {
    $statusLabel.Text = "Lade Netzwerkkarten-Informationen..."
    
    try {
        $global:adapters = Get-NetworkAdapterInfo
        
        if ($global:isGridView) {
            # Listbox leeren und neu füllen
            $adapterListBox.Items.Clear()
            
            # Prüfen, ob die Adapterliste ein String ist (Fehlermeldung)
            if ($global:adapters -is [string]) {
                # Fehlermeldung anzeigen
                $errorMsg = New-Object System.Windows.Forms.Label
                $errorMsg.Location = New-Object System.Drawing.Point(20, 20)
                $errorMsg.Size = New-Object System.Drawing.Size(760, 380)
                $errorMsg.Text = $global:adapters
                $errorMsg.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $errorMsg.ForeColor = [System.Drawing.Color]::Red
                $detailPanel.Controls.Clear()
                $detailPanel.Controls.Add($errorMsg)
                
                $statusLabel.Text = "Fehler beim Abrufen der Netzwerkadapter"
                return
            }
            
            # Debugging-Information
            Write-DebugLog "Gefundene Adapter: $($global:adapters.Count)" -Color "Cyan"
            
            foreach ($adapter in $global:adapters) {
                $status = ""
                if ($adapter.ConnectionState -eq "Verbunden" -or $adapter.ConnectionState -eq "Up") {
                    $status = "✓ "
                } else {
                    $status = "✗ "
                }
                $adapterListBox.Items.Add($status + $adapter.Name)
            }
            
            # Das Detail-Panel, die Listbox und den Text-Button anzeigen, die RichTextBox ausblenden
            $detailPanel.Visible = $true
            $adapterListBox.Visible = $true
            $outputBox.Visible = $false
            $detailButton.Text = "Text-Ansicht"
            
            # Ersten Eintrag auswählen, falls vorhanden
            if ($adapterListBox.Items.Count -gt 0) {
                $adapterListBox.SelectedIndex = 0
                Update-DetailPanel
            }
        }
        else {
            # Text-Ansicht aktualisieren
            $outputBox.Clear()
            
            if ($global:adapters -is [string]) {
                $outputBox.AppendText($global:adapters)
            }
            else {
                $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
                $outputBox.AppendText("AKTIVE NETZWERKKARTEN`r`n")
                $outputBox.AppendText("=====================`r`n`r`n")
                
                if ($global:adapters.Count -eq 0) {
                    $outputBox.SelectionColor = [System.Drawing.Color]::Red
                    $outputBox.AppendText("Keine aktiven Netzwerkkarten gefunden!`r`n")
                    $outputBox.SelectionColor = [System.Drawing.Color]::Black
                }
                else {
                    foreach ($adapter in $global:adapters) {
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("Name: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.Name)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("Beschreibung: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.InterfaceDescription)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("Status: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        
                        if ($adapter.ConnectionState -eq "Verbunden" -or $adapter.ConnectionState -eq "Up") {
                            $outputBox.SelectionColor = [System.Drawing.Color]::Green
                        }
                        else {
                            $outputBox.SelectionColor = [System.Drawing.Color]::Red
                        }
                        
                        $outputBox.AppendText("$($adapter.ConnectionState)`r`n")
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("MAC-Adresse: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.MacAddress)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("IPv4-Adresse: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.IPv4Address)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("IPv6-Adresse: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.IPv6Address)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("Gateway: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.Gateway)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("DNS-Server: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.DNS)`r`n")
                        
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Blue
                        $outputBox.AppendText("Linkgeschwindigkeit: ")
                        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Regular)
                        $outputBox.SelectionColor = [System.Drawing.Color]::Black
                        $outputBox.AppendText("$($adapter.LinkSpeed)`r`n")
                        
                        $outputBox.AppendText("`r`n")
                    }
                }
            }
            
            # Die RichTextBox anzeigen, das Detail-Panel und die Listbox ausblenden
            $outputBox.Visible = $true
            $detailPanel.Visible = $false
            $adapterListBox.Visible = $false
            $detailButton.Text = "Detail-Ansicht"
        }
        
        $statusLabel.Text = "Bereit: Letzte Aktualisierung: $(Get-Date -Format 'HH:mm:ss dd.MM.yyyy')"
    }
    catch {
        Write-DebugLog "Fehler in Update-Display: $_" -Color "Red"
        $statusLabel.Text = "Fehler: $_"
    }
	}

# Event-Handler für den Aktualisierungsbutton
$refreshButton.Add_Click({
    Update-Display
})

# Event-Handler für den Detailansicht-Button
$detailButton.Add_Click({
    $global:isGridView = !$global:isGridView
    Update-Display
})

# Event-Handler für die Listbox
$adapterListBox.Add_SelectedIndexChanged({
    if ($adapterListBox.SelectedItem) {
        try {
            if ($global:isDomainController) {
                Write-Host "Adapter ausgewählt: $($adapterListBox.SelectedItem) (Index: $($adapterListBox.SelectedIndex))" -ForegroundColor Cyan
            }
            Write-DebugLog "SelectedIndexChanged: $($adapterListBox.SelectedIndex)" -Color "Cyan"
            Update-DetailPanel
        }
        catch {
            Write-Host "Fehler beim Adapter-Wechsel: $_" -ForegroundColor Red
            Write-DebugLog "Fehler beim Adapter-Wechsel: $_" -Color "Red"
            $statusLabel.Text = "Fehler beim Wechseln des Adapters: $_"
        }
    }
})

# Event-Handler für den Export-Button
$exportButton.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV-Datei (*.csv)|*.csv|Textdatei (*.txt)|*.txt"
    $saveDialog.Title = "Netzwerkkarten-Informationen exportieren"
    $saveDialog.FileName = "Netzwerkkarten_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $adapters = Get-NetworkAdapterInfo
            
            if ($saveDialog.FileName.EndsWith(".csv")) {
                $adapters | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            }
            else {
                $output = "AKTIVE NETZWERKKARTEN`r`n=====================`r`n`r`n"
                
                foreach ($adapter in $adapters) {
                    $output += "Name: $($adapter.Name)`r`n"
                    $output += "Beschreibung: $($adapter.InterfaceDescription)`r`n"
                    $output += "Status: $($adapter.ConnectionState)`r`n"
                    $output += "MAC-Adresse: $($adapter.MacAddress)`r`n"
                    $output += "IPv4-Adresse: $($adapter.IPv4Address)`r`n"
                    $output += "IPv6-Adresse: $($adapter.IPv6Address)`r`n"
                    $output += "Linkgeschwindigkeit: $($adapter.LinkSpeed)`r`n`r`n"
                }
                
                [System.IO.File]::WriteAllText($saveDialog.FileName, $output)
            }
            
            $statusLabel.Text = "Erfolgreich exportiert nach: $($saveDialog.FileName)"
        }
        catch {
            $statusLabel.Text = "Fehler beim Exportieren: $_"
        }
    }
})

# Event-Handler für den Neustart-Button
$restartButton.Add_Click({
    if ($adapterListBox.SelectedIndex -ge 0 -and $adapterListBox.SelectedIndex -lt $global:adapters.Count) {
        $adapter = $global:adapters[$adapterListBox.SelectedIndex]
        
        # Bestätigungsdialog anzeigen
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie den Netzwerkadapter '$($adapter.Name)' wirklich neu starten?",
            "Netzwerkadapter neu starten",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $success = Restart-NetworkAdapter -AdapterName $adapter.Name
            
            if ($success) {
                # Aktualisierung wird bereits in der Restart-Funktion durchgeführt
            }
            else {
                # Ein Fehler ist aufgetreten - eventuell Adminrechte erforderlich
                $elevateResult = [System.Windows.Forms.MessageBox]::Show(
                    "Der Neustart des Netzwerkadapters erfordert Administratorrechte. Möchten Sie das Skript mit erhöhten Rechten neu starten?",
                    "Administratorrechte erforderlich",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                    
                if ($elevateResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    # Skript mit Admin-Rechten neu starten
                    $scriptPath = $MyInvocation.MyCommand.Path
                    if ($scriptPath) {
                        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
                        $form.Close()
                    }
                    else {
                        $statusLabel.Text = "Fehler: Skriptpfad konnte nicht ermittelt werden."
                    }
                }
            }
        }
    }
    else {
        $statusLabel.Text = "Kein Netzwerkadapter ausgewählt."
    }
})

# Event-Handler für den Autostart-Task-Button
$autostartButton.Add_Click({
    if ($adapterListBox.SelectedIndex -ge 0 -and $adapterListBox.SelectedIndex -lt $global:adapters.Count) {
        $adapter = $global:adapters[$adapterListBox.SelectedIndex]
        
        # Bestätigungsdialog anzeigen
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie eine geplante Aufgabe erstellen, um den Netzwerkadapter '$($adapter.Name)' beim Systemstart automatisch neu zu starten?",
            "Autostart-Task erstellen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $success = Create-AdapterRestartTask -AdapterName $adapter.Name
            
            if (-not $success) {
                # Ein Fehler ist aufgetreten - eventuell Adminrechte erforderlich
                $elevateResult = [System.Windows.Forms.MessageBox]::Show(
                    "Das Erstellen der geplanten Aufgabe erfordert Administratorrechte. Möchten Sie das Skript mit erhöhten Rechten neu starten?",
                    "Administratorrechte erforderlich",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                    
                if ($elevateResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    # Skript mit Admin-Rechten neu starten
                    $scriptPath = $MyInvocation.MyCommand.Path
                    if ($scriptPath) {
                        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
                        $form.Close()
                    }
                    else {
                        $statusLabel.Text = "Fehler: Skriptpfad konnte nicht ermittelt werden."
                    }
                }
            }
            else {
                # Information über erfolgreiche Erstellung anzeigen
                [System.Windows.Forms.MessageBox]::Show(
                    "Die geplante Aufgabe zum automatischen Neustart des Netzwerkadapters '$($adapter.Name)' wurde erfolgreich erstellt.`n`nDie Aufgabe wird bei jedem Systemstart ausgeführt.",
                    "Autostart-Task erstellt",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    }
    else {
        $statusLabel.Text = "Kein Netzwerkadapter ausgewählt."
    }
})

# Event-Handler für den Diagnose-Button
$diagnoseButton.Add_Click({
    Show-DiagnosticInfo
})

# Initialisierung der Anzeige beim Start
try {
    Write-DebugLog "Initialisiere Anzeige" -Color "Cyan"
    Update-Display
    
    # Fehlermeldung auf dem Hauptfenster anzeigen, wenn auf einem DC und keine Adapter gefunden wurden
    if ($global:isDomainController -and ($global:adapters -isnot [array] -or $global:adapters.Count -eq 0)) {
        $outputBox.Clear()
        $outputBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
        $outputBox.SelectionColor = [System.Drawing.Color]::Red
        $outputBox.AppendText("Domain Controller Modus`r`n")
        $outputBox.AppendText("=====================`r`n`r`n")
        $outputBox.AppendText("Keine Netzwerkadapter gefunden. Bitte klicken Sie auf 'Diagnose' für mehr Informationen.`r`n")
        $outputBox.AppendText("Die Text-Ansicht sollte dennoch funktionieren.`r`n`r`n")
        $outputBox.AppendText("Hinweis: Versuchen Sie das Skript mit administrativen Rechten auszuführen.`r`n")
        
        # Textmodus als Standard aktivieren, wenn auf einem DC gestartet
        $global:isGridView = $false
        Update-Display
    }
}
catch {
    $errorMessage = "Fehler bei der Initialisierung: $_"
    Write-Host $errorMessage -ForegroundColor Red
    Write-DebugLog $errorMessage -Color "Red"
    $statusLabel.Text = $errorMessage
}

# Anzeigen des Formulars
$form.ShowDialog()