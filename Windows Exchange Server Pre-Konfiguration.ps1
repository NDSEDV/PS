<#
.SYNOPSIS
  Microsoft Exchange Server Installation Tool
.DESCRIPTION
  The tool is intended to help you with your daily business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-03-05
  Purpose/Change: Initial script development

  Copyright (c) Jörn Walter. All rights reserved.
#>

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Neustart als Administrator
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

# Initialisieren mit Platzhaltern, tatsächliche Daten später laden
$script:form = $null
$script:tabControl = $null
$script:tabMain = $null
$script:tabUCMA = $null
$script:tabExchange = $null
$script:systemResources = $null

# UI-Komponenten
$script:buttonInstallFeatures = $null
$script:buttonDisableRealtime = $null
$script:buttonInstallVC = $null
$script:buttonInstallVC2013 = $null
$script:buttonInstallURLRewrite = $null
$script:pathTextBox = $null
$script:statusText = $null
$script:resultLabel = $null
$script:installButton = $null
$script:exchangePathTextBox = $null
$script:exchangeStatusText = $null
$script:exchangeResultLabel = $null
$script:exchangeInstallButton = $null
$script:rebootStatusLabel = $null

# Statusverfolgung
$script:ucmaInitialized = $false
$script:exchangeInitialized = $false
$script:statusChecksLoaded = $false
$script:uninstallCache = $null

# Registry-Hilfsfunktionen
function Ensure-RegistryPath {
    if (-not (Test-Path "HKCU:\Software\ExchangeRequirements")) {
        New-Item -Path "HKCU:\Software" -Name "ExchangeRequirements" -Force | Out-Null
    }
}

# Autostart-Funktionen für Neustarts
function Set-AutoStartAfterReboot {
    try {
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { throw "Skriptpfad nicht gefunden." }
        
        Ensure-RegistryPath
        $batchContent = "@echo off`npowershell.exe -WindowStyle hidden -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        $startupFolder = [System.Environment]::GetFolderPath('Startup')
        $batchPath = Join-Path -Path $startupFolder -ChildPath "ExchangeRequirements.bat"
        Set-Content -Path $batchPath -Value $batchContent -Force
        Set-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -Value $true -Type DWORD -Force
        return $true
    } catch {
        Write-Error "Autostart-Setup-Fehler: $_"
        return $false
    }
}

function Remove-AutoStartAfterReboot {
    try {
        $startupFolder = [System.Environment]::GetFolderPath('Startup')
        $batchPath = Join-Path -Path $startupFolder -ChildPath "ExchangeRequirements.bat"
        if (Test-Path $batchPath) { Remove-Item -Path $batchPath -Force }
        
        if (Test-Path "HKCU:\Software\ExchangeRequirements") {
            Remove-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -ErrorAction SilentlyContinue
        }
        return $true
    } catch {
        return $false
    }
}

# SYSTEMINFORMATIONEN
function Get-SystemResources {
    # Informationen nur einmal abrufen
    if ($script:systemResources) { return $script:systemResources }
    
    try {
        # Alle Prozessoren abrufen und deren logische Prozessoren summieren
        $processors = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue
        $coreCount = ($processors | Measure-Object -Property NumberOfLogicalProcessors -Sum -ErrorAction SilentlyContinue).Sum
        
        if (-not $coreCount) { $coreCount = "N/A" }
        
        $totalRAM = "N/A"
        $memoryInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($memoryInfo -and $memoryInfo.TotalPhysicalMemory) {
            $totalRAM = [math]::Round($memoryInfo.TotalPhysicalMemory / 1GB, 2)
        }
        
        $script:systemResources = @{
            CoreCount = $coreCount
            TotalRAM = $totalRAM
        }
        
        return $script:systemResources
    } catch {
        $script:systemResources = @{ CoreCount = "N/A"; TotalRAM = "N/A" }
        return $script:systemResources
    }
}

# SUCHE - Nur eingehängte Laufwerke mit Volumes überprüfen
function Find-UCMASetupExe {
    try {
        # Alle verfügbaren Laufwerke ermitteln
        $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match "^\w:\\" }
        
        foreach ($drive in $drives) {
            $driveLetter = $drive.Root.Substring(0, 2)
            $setupPath = "$driveLetter\UCMARedist\setup.exe"
            
            # Prüfen, ob die Datei existiert
            if (Test-Path $setupPath) {
                return $setupPath
            }
        }
        
        # Nichts gefunden
        return $null
    }
    catch {
        Write-Error "Fehler bei der Suche nach UCMA Setup: $_"
        return $null
    }
}

# SUCHE - Exchange-Setup effizienter finden
function Find-ExchangeSetupExe {
    param (
        [string]$UcmaPath = $null
    )
    
    try {
        # Wenn UCMA-Pfad angegeben wurde, daraus den Exchange-Pfad ableiten
        if (-not [string]::IsNullOrEmpty($UcmaPath) -and $UcmaPath -match "^([A-Z]:)\\UCMARedist\\setup\.exe$") {
            $driveLetter = $matches[1]
            $setupPath = "$driveLetter\setup.exe"
            
            if (Test-Path $setupPath) {
                return $setupPath
            }
        }
        
        # Ansonsten alle verfügbaren Laufwerke durchsuchen
        $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match "^\w:\\" }
        
        foreach ($drive in $drives) {
            $driveLetter = $drive.Root.Substring(0, 2)
            $setupPath = "$driveLetter\setup.exe"
            
            # Prüfen, ob die Datei existiert und ob es sich um das Exchange-Setup handelt
            if (Test-Path $setupPath) {
                $fileInfo = Get-Item $setupPath
                
                # Wenn die Datei größer als 20 KB ist, handelt es sich wahrscheinlich um das Exchange-Setup
                if ($fileInfo.Length -gt 20KB) {
                    # Versuche zusätzlich zu prüfen, ob es wirklich das Exchange-Setup ist
                    $fileVersion = $fileInfo.VersionInfo
                    if ($fileVersion.ProductName -like "*Exchange*" -or $fileVersion.FileDescription -like "*Exchange*") {
                        return $setupPath
                    }
                    
                    # Alternativ: Prüfen, ob es die nötigen Dateien im selben Verzeichnis gibt
                    $exchangeFiles = @("setup.exe", "ExSetup.exe", "bin", "Client Access")
                    $foundFiles = 0
                    
                    foreach ($file in $exchangeFiles) {
                        if (Test-Path "$driveLetter\$file") {
                            $foundFiles++
                        }
                    }
                    
                    # Wenn mindestens 3 der typischen Exchange-Dateien/Ordner gefunden wurden
                    if ($foundFiles -ge 2) {
                        return $setupPath
                    }
                }
            }
        }
        
        # Nichts gefunden
        return $null
    }
    catch {
        Write-Error "Fehler bei der Suche nach Exchange Setup: $_"
        return $null
    }
}

# Funktion zum Aufräumen von Exchange-Installations-Ressourcen
function Cleanup-ExchangeInstallationResources {
    # Timer stoppen
    if ($script:exchangeInstallStatusTimer) {
        try {
            $script:exchangeInstallStatusTimer.Stop()
            $script:exchangeInstallStatusTimer.Dispose()
            $script:exchangeInstallStatusTimer = $null
        } catch {
            Write-Host "Fehler beim Stoppen des Exchange-Installations-Timers: $_"
        }
    }
    
    # Flag-Datei entfernen
    try {
        $installFlagPath = "$env:TEMP\ExchangeInstallationRunning.flag"
        if (Test-Path $installFlagPath) {
            Remove-Item -Path $installFlagPath -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "Fehler beim Entfernen der Exchange-Installations-Flag-Datei: $_"
    }
}

# Vereinfachte Windows-Features Überprüfung mit Flag-Datei
function Test-WindowsFeatures {
    param ([string[]]$FeatureNames)
    
    # Flag-Datei im Temp-Verzeichnis
    $flagFilePath = "$env:TEMP\ExchangeRequirements_Features_Installed.flag"
    
    # Wenn die Flag-Datei existiert, gelten die Features als installiert
    if (Test-Path $flagFilePath) {
        return $true
    }
    
    # Andernfalls prüfe kurz, ob die Features installiert sind
    try {
        # Fehler und Warnungen unterdrücken
        $ErrorActionPreference = 'SilentlyContinue'
        $WarningPreference = 'SilentlyContinue'
        
        $installedCount = 0
        $totalCount = $FeatureNames.Count
        
        # Prüfen, ob Server- oder Client-Betriebssystem (schneller als beide Befehle zu versuchen)
        $isServer = [bool](Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue)
        
        if ($isServer) {
            # Server-Betriebssystem - Features in einem Aufruf abrufen
            $features = Get-WindowsFeature -Name $FeatureNames -ErrorAction SilentlyContinue
            $installedCount = ($features | Where-Object { $_.Installed -eq $true }).Count
        } else {
            # Client-Betriebssystem - Get-WindowsOptionalFeature verwenden (vereinfacht)
            foreach ($feature in $FeatureNames) {
                $featureInfo = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
                if ($featureInfo -and $featureInfo.State -eq "Enabled") { $installedCount++ }
            }
        }
        
        $result = ($installedCount -eq $totalCount)
        
        # Bei erfolgreicher Installation die Flag-Datei erstellen
        if ($result) {
            Set-Content -Path $flagFilePath -Value "Windows Features installiert am $(Get-Date)" -Force
        }
        
        return $result
    } catch {
        # Im Fehlerfall Standardwert zurückgeben
        return $false
    } finally {
        # Präferenzen zurücksetzen
        $ErrorActionPreference = 'Continue'
        $WarningPreference = 'Continue'
    }
}

# VC++ PRÜFUNG - Cache-Registry-Abfragen
function Test-VCRedistInstalled {
    param ([string]$DisplayNamePattern)
    
    try {
        # Registry-Werte zwischenspeichern, um wiederholte Abfragen zu vermeiden
        if (-not $script:uninstallCache) {
            $script:uninstallCache = @()
            # 64-Bit- und 32-Bit-Registry in einem Vorgang kombinieren
            $script:uninstallCache += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
            $script:uninstallCache += Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
        }
        
        return ($script:uninstallCache | Where-Object { $_.DisplayName -like $DisplayNamePattern }) -ne $null
    } catch {
        return $false
    }
}

# URL REWRITE PRÜFUNG - Einfache Dateiexistenz
function Test-URLRewriteInstalled {
    try {
        # Nur auf die DLL prüfen - schnellste Methode
        return (Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll")
    } catch {
        return $false
    }
}

# ECHTZEITSCHUTZ-PRÜFUNG - Direkter API-Aufruf
function Test-RealtimeProtectionDisabled {
    try {
        # Status in einem Aufruf abrufen
        return (Get-MpPreference -ErrorAction SilentlyContinue).DisableRealtimeMonitoring -eq $true
    } catch {
        return $false
    }
}

# NEUSTART-PRÜFUNG - Nur kritische Indikatoren prüfen
function Test-PendingReboot {
    try {
        # Nur die häufigsten und zuverlässigsten Indikatoren prüfen
        return (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") -or
               (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -or
               ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -ErrorAction SilentlyContinue).PendingFileRenameOperations)
    } catch {
        return $false
    }
}

# UI basierend auf Neustart-Status aktualisieren
function Update-RebootStatus {
    if (-not $script:rebootStatusLabel) { return }
    
    if (Test-PendingReboot) {
        $script:rebootStatusLabel.Text = "HINWEIS: Es steht ein Systemneustart aus!"
        $script:rebootStatusLabel.ForeColor = [System.Drawing.Color]::Red
    } else {
        $script:rebootStatusLabel.Text = "Kein Systemneustart ausstehend."
        $script:rebootStatusLabel.ForeColor = [System.Drawing.Color]::Green
    }
}

# UI-Status-Prüfung verwalten - nur bei Bedarf prüfen
function Update-UIStatus {
    # Überspringen, wenn keine Schaltflächen zu aktualisieren sind
    if (-not $script:buttonInstallFeatures -and 
        -not $script:buttonDisableRealtime -and
        -not $script:buttonInstallVC -and
        -not $script:buttonInstallVC2013 -and
        -not $script:buttonInstallURLRewrite) {
        return
    }
    
    # Echtzeitschutz prüfen
    if ($script:buttonDisableRealtime) {
        if (Test-RealtimeProtectionDisabled) {
            $script:buttonDisableRealtime.Enabled = $false
            $script:buttonDisableRealtime.Text = "Echtzeitschutz bereits deaktiviert"
        }
    }
    
    # Windows-Features prüfen - vereinfachte Prüfung über Flag-Datei
    if ($script:buttonInstallFeatures -and $script:buttonInstallFeatures.Enabled) {
        $flagFilePath = "$env:TEMP\ExchangeRequirements_Features_Installed.flag"
        if (Test-Path $flagFilePath) {
            $script:buttonInstallFeatures.Enabled = $false
            $script:buttonInstallFeatures.Text = "Windows-Features bereits installiert"
        }
    }
    
    # VC++-Installationen prüfen
    if ($script:buttonInstallVC -and $script:buttonInstallVC.Enabled) {
        if (Test-VCRedistInstalled -DisplayNamePattern "*Visual C++ 2012*") {
            $script:buttonInstallVC.Enabled = $false
            $script:buttonInstallVC.Text = "Visual C++ 2012 Redistributable bereits installiert"
        }
    }
    
    if ($script:buttonInstallVC2013 -and $script:buttonInstallVC2013.Enabled) {
        if (Test-VCRedistInstalled -DisplayNamePattern "*Visual C++ 2013*") {
            $script:buttonInstallVC2013.Enabled = $false
            $script:buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable bereits installiert"
        }
    }
    
    # URL Rewrite prüfen
    if ($script:buttonInstallURLRewrite -and $script:buttonInstallURLRewrite.Enabled) {
        if (Test-URLRewriteInstalled) {
            $script:buttonInstallURLRewrite.Enabled = $false
            $script:buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul bereits installiert"
        }
    }
    
    # Neustart-Status-Label aktualisieren, falls vorhanden
    if ($script:rebootStatusLabel) {
        Update-RebootStatus
    }
}

# Initialisierung des Hauptformulars - das Herzstück der Anwendung
function Initialize-MainForm {
    # Formular erstellen
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Microsoft Exchange Server Installation Tool"
    $form.Size = New-Object System.Drawing.Size(520, 605)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

        # Formular-Schließen-Ereignishandler hinzufügen
    $form.Add_FormClosing({
        # Ressourcen aufräumen
        Cleanup-ExchangeInstallationResources
    })

    # Systeminfo laden, wenn der Header erstellt wird, nicht früher
    $sysInfo = Get-SystemResources
    
    # Header-Panel erstellen
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(500, 40)
    $headerPanel.BackColor = [System.Drawing.Color]::LightGray

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Location = New-Object System.Drawing.Point(10, 10)
    $headerLabel.Size = New-Object System.Drawing.Size(480, 20)
    $headerLabel.Text = "System: $($sysInfo.CoreCount) virtuelle Kerne | $($sysInfo.TotalRAM) GB RAM"
    $headerLabel.Font = New-Object System.Drawing.Font("Verdana", 9, [System.Drawing.FontStyle]::Bold)
    $headerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $headerPanel.Controls.Add($headerLabel)
    $form.Controls.Add($headerPanel)

    # Tab-Steuerung erstellen
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(0, 40)
    $tabControl.Size = New-Object System.Drawing.Size(500, 490)
    $tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                         [System.Windows.Forms.AnchorStyles]::Left -bor 
                         [System.Windows.Forms.AnchorStyles]::Right -bor 
                         [System.Windows.Forms.AnchorStyles]::Bottom
    $form.Controls.Add($tabControl)

    # Alle Tabs erstellen
    $tabMain = New-Object System.Windows.Forms.TabPage
    $tabMain.Text = "Hauptfunktionen"
    $tabControl.Controls.Add($tabMain)
    
    $tabUCMA = New-Object System.Windows.Forms.TabPage
    $tabUCMA.Text = "UCMA Installation"
    $tabControl.Controls.Add($tabUCMA)
    
    $tabExchange = New-Object System.Windows.Forms.TabPage
    $tabExchange.Text = "Exchange Installation"
    $tabControl.Controls.Add($tabExchange)

    # Copyright-Label erstellen
    $copyrightLabel = New-Object System.Windows.Forms.Label
    $copyrightLabel.Location = New-Object System.Drawing.Point(10, 535)
    $copyrightLabel.Size = New-Object System.Drawing.Size(480, 30)
    $copyrightLabel.Text = "© 2025 Jörn Walter https://www.der-windows-papst.de"
    $copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $copyrightLabel.Font = New-Object System.Drawing.Font("Verdana", 8)
    $copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($copyrightLabel)

    # Ereignis für Tab-Wechsel hinzufügen - Status nur prüfen, wenn sich Tab ändert
    $tabControl.Add_SelectedIndexChanged({
        Update-UIStatus
    })

    # Globale Referenzen speichern
    $script:form = $form
    $script:tabControl = $tabControl
    $script:tabMain = $tabMain
    $script:tabUCMA = $tabUCMA
    $script:tabExchange = $tabExchange
    
    # Zuerst den Haupt-Tab initialisieren - dieser wird sofort sichtbar sein
    Initialize-MainTab -TabPage $tabMain
    
    # Andere Tabs noch nicht initialisieren - erst bei Auswahl
    $tabUCMA.Add_Enter({
        if (-not $script:ucmaInitialized) {
            Initialize-UCMATab -TabPage $tabUCMA
            $script:ucmaInitialized = $true
        }
    })
    
    $tabExchange.Add_Enter({
        if (-not $script:exchangeInitialized) {
            Initialize-ExchangeTab -TabPage $tabExchange
            $script:exchangeInitialized = $true
        }
    })
    
    return $form
}

# Funktion zur Überprüfung, ob UCMA bereits installiert ist
function Test-UCMAInstalled {
    try {
        # UCMA in der Registry überprüfen
        $ucmaInstalled = $false
        
        # Sowohl in 64-Bit- als auch in 32-Bit-Deinstallationsorten prüfen
        $uninstallPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        
        foreach ($path in $uninstallPaths) {
            $installed = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -like "*Microsoft Unified Communications Managed API*" -or 
                              $_.DisplayName -like "*UCMA*" }
            
            if ($installed) {
                $ucmaInstalled = $true
                break
            }
        }
        
        # Alternative Prüfung: Nach spezifischen UCMA-Assemblies oder Dateien suchen
        if (-not $ucmaInstalled) {
            $ucmaAssemblyPaths = @(
                "$env:ProgramFiles\Microsoft UCMA 4.0\Bin\Microsoft.Rtc.Collaboration.dll",
                "${env:ProgramFiles(x86)}\Microsoft UCMA 4.0\Bin\Microsoft.Rtc.Collaboration.dll",
                "$env:windir\Microsoft.NET\assembly\GAC_MSIL\Microsoft.Rtc.Collaboration"
            )
            
            foreach ($path in $ucmaAssemblyPaths) {
                if (Test-Path $path) {
                    $ucmaInstalled = $true
                    break
                }
            }
        }
        
        return $ucmaInstalled
    }
    catch {
        # Bei einem Fehler davon ausgehen, dass UCMA nicht installiert ist
        return $false
    }
}

# Den Haupt-Tab initialisieren - Steuerelemente für den Hauptfunktionen-Tab erstellen
function Initialize-MainTab {
    param([System.Windows.Forms.TabPage]$TabPage)
    
    # Echtzeitschutz-Button erstellen
    $buttonDisableRealtime = New-Object System.Windows.Forms.Button
    $buttonDisableRealtime.Location = New-Object System.Drawing.Point(50, 50)
    $buttonDisableRealtime.Size = New-Object System.Drawing.Size(400, 40)
    $buttonDisableRealtime.Text = "Echtzeitschutz deaktivieren"
    $buttonDisableRealtime.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $buttonDisableRealtime.Add_Click({
        try {
            Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("Echtzeitschutz wurde erfolgreich deaktiviert.", "Erfolg", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Button-Status aktualisieren
            $buttonDisableRealtime.Enabled = $false
            $buttonDisableRealtime.Text = "Echtzeitschutz bereits deaktiviert"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Deaktivieren des Echtzeitschutzes: $_", "Fehler", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $TabPage.Controls.Add($buttonDisableRealtime)
    $script:buttonDisableRealtime = $buttonDisableRealtime

    # Windows-Features-Button erstellen
    $buttonInstallFeatures = New-Object System.Windows.Forms.Button
    $buttonInstallFeatures.Location = New-Object System.Drawing.Point(50, 120)
    $buttonInstallFeatures.Size = New-Object System.Drawing.Size(400, 40)
    $buttonInstallFeatures.Text = "Windows-Features installieren"
    $buttonInstallFeatures.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)

#  Windows Feature Installation
# Modifizierte Button-Handler für Windows-Feature Installation
$buttonInstallFeatures.Add_Click({
    try {
        # Definiere erforderliche Features
        $requiredFeatures = @(
            "NET-Framework-45-Features", 
            "Web-Server", 
            "Web-Asp-Net45",
            "Web-Basic-Auth", 
            "Web-Client-Auth", 
            "Web-Digest-Auth", 
            "Web-Dir-Browsing", 
            "Web-Dyn-Compression", 
            "Web-Http-Errors", 
            "Web-Http-Logging", 
            "Web-Http-Redirect", 
            "Web-Http-Tracing", 
            "Web-ISAPI-Ext", 
            "Web-ISAPI-Filter", 
            "Web-Mgmt-Console", 
            "Web-Mgmt-Service", 
            "Web-Net-Ext45", 
            "Web-Static-Content", 
            "Web-Windows-Auth", 
            "Web-WMI",
            "Server-Media-Foundation", 
            "RPC-over-HTTP-proxy", 
            "RSAT-Clustering", 
            "RSAT-Clustering-CmdInterface", 
            "RSAT-Clustering-Mgmt", 
            "RSAT-Clustering-PowerShell", 
            "WAS-Process-Model", 
            "Web-Metabase", 
            "Web-Request-Monitor", 
            "Web-Stat-Compression", 
            "Windows-Identity-Foundation", 
            "RSAT-ADDS"
        )
        
        # Erstelle PowerShell-Skript für die Feature-Installation
        $scriptPath = "$env:TEMP\InstallFeatures.ps1"
        $scriptContent = @"
# Aktiviere detaillierte Ausgabe
`$VerbosePreference = 'Continue'
`$ProgressPreference = 'Continue'

# Features installieren
Write-Host "Starte Installation der Windows-Features..." -ForegroundColor Green

# Features als Array definieren
`$features = @(
    "NET-Framework-45-Features", 
    "Web-Server", 
    "Web-Asp-Net45",
    "Web-Basic-Auth", 
    "Web-Client-Auth", 
    "Web-Digest-Auth", 
    "Web-Dir-Browsing", 
    "Web-Dyn-Compression", 
    "Web-Http-Errors", 
    "Web-Http-Logging", 
    "Web-Http-Redirect", 
    "Web-Http-Tracing", 
    "Web-ISAPI-Ext", 
    "Web-ISAPI-Filter", 
    "Web-Mgmt-Console", 
    "Web-Mgmt-Service", 
    "Web-Net-Ext45", 
    "Web-Static-Content", 
    "Web-Windows-Auth", 
    "Web-WMI",
    "Server-Media-Foundation", 
    "RPC-over-HTTP-proxy", 
    "RSAT-Clustering", 
    "RSAT-Clustering-CmdInterface", 
    "RSAT-Clustering-Mgmt", 
    "RSAT-Clustering-PowerShell", 
    "WAS-Process-Model", 
    "Web-Metabase", 
    "Web-Request-Monitor", 
    "Web-Stat-Compression", 
    "Windows-Identity-Foundation", 
    "RSAT-ADDS"
)

Write-Host "Installiere die folgenden Features:" -ForegroundColor Cyan
`$features | ForEach-Object { Write-Host "  - `$_" }

# Führe die Installation durch
`$result = Install-WindowsFeature -Name `$features -IncludeManagementTools -Verbose

# Analysiere das Ergebnis
Write-Host "`nInstallationsergebnis:" -ForegroundColor Green
Write-Host "Erfolg: `$(`$result.Success)" -ForegroundColor Yellow
Write-Host "Exitcode: `$(`$result.ExitCode)" -ForegroundColor Yellow
Write-Host "Neustart erforderlich: `$(`$result.RestartNeeded)" -ForegroundColor Yellow

# Liste installierte Features auf
Write-Host "`nInstallierte Features:" -ForegroundColor Green
Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true } | Format-Table Name,DisplayName -AutoSize

# Erstelle Log-Datei
Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true } | Out-File "$env:TEMP\installed_features.txt"

# Erstelle Flag-Datei für erfolgreiche Installation
if (`$result.Success) {
    Set-Content -Path "$env:TEMP\ExchangeRequirements_Features_Installed.flag" -Value "Windows Features installiert am $(Get-Date)" -Force
    Write-Host "`nFlag-Datei für erfolgreiche Installation erstellt." -ForegroundColor Green
}

Write-Host "`nInstallation abgeschlossen." -ForegroundColor Green
Start-Sleep 10
exit
"@
        
        # Schreibe das PowerShell-Skript
        Set-Content -Path $scriptPath -Value $scriptContent -Force
        
        # Bestätigungsdialog
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Die Windows-Features werden jetzt installiert.`n`nEs wird ein PowerShell-Fenster angezeigt, das den Fortschritt der Installation zeigt.`nBitte schließe dieses Fenster NICHT, bis die Installation abgeschlossen ist.`n`nMöchtest du fortfahren?", 
            "Windows-Features installieren", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
        
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) {
            return
        }
        
        # Deaktiviere den Button während der Installation
        $buttonInstallFeatures.Enabled = $false
        $buttonInstallFeatures.Text = "Installation läuft..."
        
        # Aktualisiere die UI
        [System.Windows.Forms.Application]::DoEvents()
        
        # Starte das PowerShell-Skript in einem neuen Fenster, damit der Benutzer den Fortschritt sehen kann
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$scriptPath`"" -Wait
        
        # Nach Abschluss: Lösche das Skript
        if (Test-Path $scriptPath) {
            Remove-Item $scriptPath -Force -ErrorAction SilentlyContinue
        }
        
        # Prüfe ob die Flag-Datei existiert (vereinfachte Überprüfung)
        $flagFilePath = "$env:TEMP\ExchangeRequirements_Features_Installed.flag"
        if (Test-Path $flagFilePath) {
            $buttonInstallFeatures.Enabled = $false
            $buttonInstallFeatures.Text = "Windows-Features bereits installiert"
        } else {
            $buttonInstallFeatures.Enabled = $true
            $buttonInstallFeatures.Text = "Windows-Features installieren"
        }
        
        # Frage nach Systemneustart
        $restartResult = [System.Windows.Forms.MessageBox]::Show(
            "Ein Neustart wird empfohlen, um die Installation der Windows-Features abzuschließen.`n`nMöchtest du den Computer jetzt neu starten?", 
            "Neustart erforderlich", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Richte Autostart vor dem Neustart ein
            Set-AutoStartAfterReboot
            
            # Gib dem Benutzer einen Moment, um die Bestätigung zu sehen
            Start-Sleep -Seconds 2
            
            # Starte Computer neu
            Restart-Computer -Force
        }
    }
    catch {
        # Fehlerbehandlung
        $buttonInstallFeatures.Enabled = $true
        $buttonInstallFeatures.Text = "Windows-Features installieren"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Fehler bei der Installation der Windows-Features:`n$_", 
            "Fehler", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

    $TabPage.Controls.Add($buttonInstallFeatures)
    $script:buttonInstallFeatures = $buttonInstallFeatures

    # VC++ 2012 Installations-Button
    $buttonInstallVC = New-Object System.Windows.Forms.Button
    $buttonInstallVC.Location = New-Object System.Drawing.Point(50, 190)
    $buttonInstallVC.Size = New-Object System.Drawing.Size(400, 40)
    $buttonInstallVC.Text = "Visual C++ 2012 Redistributable herunterladen und installieren"
    $buttonInstallVC.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $buttonInstallVC.Add_Click({
        try {
            $tempDir = [System.IO.Path]::GetTempPath()
            $vcRedistPath = Join-Path -Path $tempDir -ChildPath "vcredist_x64.exe"
            $url = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe"
            
            # Status-Fenster
            $statusForm = New-Object System.Windows.Forms.Form
            $statusForm.Text = "Download und Installation..."
            $statusForm.Size = New-Object System.Drawing.Size(240, 100)
            $statusForm.StartPosition = "CenterScreen"
            $statusForm.FormBorderStyle = "FixedDialog"
            $statusForm.ControlBox = $false
            
            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
            $statusLabel.Size = New-Object System.Drawing.Size(220, 60)
            $statusLabel.Text = "Visual C++ 2012 Redistributable wird heruntergeladen und installiert..."
            $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
            $statusForm.Controls.Add($statusLabel)
            
            $statusForm.Show()
            $script:form.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            
            # Datei mit besserer Fehlerbehandlung herunterladen
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($url, $vcRedistPath)
            } catch {
                throw "Download fehlgeschlagen: $_"
            }
            
            # Prüfen, ob die Datei heruntergeladen wurde
            if (Test-Path $vcRedistPath) {
                # Paket im Hintergrund installieren
                $process = Start-Process -FilePath $vcRedistPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
                
                # Status-Fenster schließen
                $statusForm.Close()
                $script:form.Enabled = $true
                
                # Installationsergebnis prüfen
                if ($process.ExitCode -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Visual C++ 2012 Redistributable wurde erfolgreich installiert.", 
                        "Erfolg", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                    
                    $buttonInstallVC.Enabled = $false
                    $buttonInstallVC.Text = "Visual C++ 2012 Redistributable bereits installiert"
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Die Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen.", 
                        "Information", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
                
                # Aufräumen
                if (Test-Path $vcRedistPath) {
                    Remove-Item -Path $vcRedistPath -Force -ErrorAction SilentlyContinue
                }
            } else {
                $statusForm.Close()
                $script:form.Enabled = $true
                throw "Heruntergeladene Datei nicht gefunden"
            }
        } catch {
            if ($statusForm) {
                $statusForm.Close()
                $script:form.Enabled = $true
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei Visual C++ 2012 Redistributable: $_", 
                "Fehler", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $TabPage.Controls.Add($buttonInstallVC)
    $script:buttonInstallVC = $buttonInstallVC

    # VC++ 2013 Installations-Button
    $buttonInstallVC2013 = New-Object System.Windows.Forms.Button
    $buttonInstallVC2013.Location = New-Object System.Drawing.Point(50, 250)
    $buttonInstallVC2013.Size = New-Object System.Drawing.Size(400, 40)
    $buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable herunterladen und installieren"
    $buttonInstallVC2013.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $buttonInstallVC2013.Add_Click({
        try {
            $tempDir = [System.IO.Path]::GetTempPath()
            $vcRedistPath = Join-Path -Path $tempDir -ChildPath "vcredist_x64_2013.exe"
            $url = "https://download.visualstudio.microsoft.com/download/pr/10912041/cee5d6bca2ddbcd039da727bf4acb48a/vcredist_x64.exe"
            
            # Status-Fenster
            $statusForm = New-Object System.Windows.Forms.Form
            $statusForm.Text = "Download und Installation..."
            $statusForm.Size = New-Object System.Drawing.Size(240, 100)
            $statusForm.StartPosition = "CenterScreen"
            $statusForm.FormBorderStyle = "FixedDialog"
            $statusForm.ControlBox = $false
            
            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
            $statusLabel.Size = New-Object System.Drawing.Size(220, 60)
            $statusLabel.Text = "Visual C++ 2013 Redistributable wird heruntergeladen und installiert..."
            $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
            $statusForm.Controls.Add($statusLabel)
            
            $statusForm.Show()
            $script:form.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            
            # Datei mit besserer Fehlerbehandlung herunterladen
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($url, $vcRedistPath)
            } catch {
                throw "Download fehlgeschlagen: $_"
            }
            
            # Prüfen, ob die Datei heruntergeladen wurde
            if (Test-Path $vcRedistPath) {
                # Paket im Hintergrund installieren
                $process = Start-Process -FilePath $vcRedistPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
                
                # Status-Fenster schließen
                $statusForm.Close()
                $script:form.Enabled = $true
                
                # Installationsergebnis prüfen
                if ($process.ExitCode -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Visual C++ 2013 Redistributable wurde erfolgreich installiert.", 
                        "Erfolg", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                    
                    $buttonInstallVC2013.Enabled = $false
                    $buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable bereits installiert"
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Die Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen.", 
                        "Information", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
                
                # Aufräumen
                if (Test-Path $vcRedistPath) {
                    Remove-Item -Path $vcRedistPath -Force -ErrorAction SilentlyContinue
                }
            } else {
                $statusForm.Close()
                $script:form.Enabled = $true
                throw "Heruntergeladene Datei nicht gefunden"
            }
        } catch {
            if ($statusForm) {
                $statusForm.Close()
                $script:form.Enabled = $true
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei Visual C++ 2013 Redistributable: $_", 
                "Fehler", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $TabPage.Controls.Add($buttonInstallVC2013)
    $script:buttonInstallVC2013 = $buttonInstallVC2013

    # URL Rewrite Installations-Button
    $buttonInstallURLRewrite = New-Object System.Windows.Forms.Button
    $buttonInstallURLRewrite.Location = New-Object System.Drawing.Point(50, 310)
    $buttonInstallURLRewrite.Size = New-Object System.Drawing.Size(400, 50)
    $buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul herunterladen und installieren"
    $buttonInstallURLRewrite.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $buttonInstallURLRewrite.Add_Click({
        try {
            $tempDir = [System.IO.Path]::GetTempPath()
            $urlRewritePath = Join-Path -Path $tempDir -ChildPath "rewrite_amd64_en-US.msi"
            $url = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
            
            # Status-Fenster
            $statusForm = New-Object System.Windows.Forms.Form
            $statusForm.Text = "Download und Installation..."
            $statusForm.Size = New-Object System.Drawing.Size(240, 90)
            $statusForm.StartPosition = "CenterScreen"
            $statusForm.FormBorderStyle = "FixedDialog"
            $statusForm.ControlBox = $false
            
            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
            $statusLabel.Size = New-Object System.Drawing.Size(220, 60)
            $statusLabel.Text = "IIS URL Rewrite Modul wird heruntergeladen und installiert..."
            $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
            $statusForm.Controls.Add($statusLabel)
            
            $statusForm.Show()
            $script:form.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            
            # Datei herunterladen
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($url, $urlRewritePath)
            
            # MSI-Paket im Hintergrund installieren
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$urlRewritePath`" /quiet /norestart" -PassThru -Wait
            
            $statusForm.Close()
            $script:form.Enabled = $true
            
            if ($process.ExitCode -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "IIS URL Rewrite Modul wurde erfolgreich installiert.", 
                    "Erfolg", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                
                $buttonInstallURLRewrite.Enabled = $false
                $buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul bereits installiert"
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Die Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen.", 
                    "Information", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            
            # Aufräumen
            if (Test-Path $urlRewritePath) {
                Remove-Item -Path $urlRewritePath -Force
            }
        } catch {
            if ($statusForm) {
                $statusForm.Close()
                $script:form.Enabled = $true
            }
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Herunterladen oder Installieren des IIS URL Rewrite Moduls: $_", 
                "Fehler", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $TabPage.Controls.Add($buttonInstallURLRewrite)
    $script:buttonInstallURLRewrite = $buttonInstallURLRewrite

    # Hilfe-Symbol mit Tooltip
    $helpIcon = New-Object System.Windows.Forms.PictureBox
    $helpIcon.Location = New-Object System.Drawing.Point(50, 375)
    $helpIcon.Size = New-Object System.Drawing.Size(16, 16)
    $helpIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage

    # Fragezeichen-Symbol effizienter erstellen
    $bitmap = New-Object System.Drawing.Bitmap 16, 16
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Blue)
    $graphics.DrawString("?", $font, $brush, 2, -2)
    $graphics.Dispose()
    $helpIcon.Image = $bitmap

    # Tooltip hinzufügen
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($helpIcon, "UCMARedist wird vom eingelegten Exchange-Datenträger installiert.")

    # Kontext-Label hinzufügen
    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Location = New-Object System.Drawing.Point(75, 375)
    $helpLabel.Size = New-Object System.Drawing.Size(375, 20)
    $helpLabel.Text = "UCMARedist-Installation"
    $helpLabel.Font = New-Object System.Drawing.Font("Verdana", 9)

    $TabPage.Controls.Add($helpIcon)
    $TabPage.Controls.Add($helpLabel)

    # Statusmeldung hinzufügen
    $statusInfoLabel = New-Object System.Windows.Forms.Label
    $statusInfoLabel.Location = New-Object System.Drawing.Point(10, 405)
    $statusInfoLabel.Size = New-Object System.Drawing.Size(480, 40)
    $statusInfoLabel.Text = "Info: Das Tool kann nicht zaubern, aber dich unterstützen."
    $statusInfoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $statusInfoLabel.Font = New-Object System.Drawing.Font("Verdana", 8)
    $TabPage.Controls.Add($statusInfoLabel)
}

# UCMA-Tab initialisieren - Wird nur aufgerufen, wenn der Tab ausgewählt wird
function Initialize-UCMATab {
    param([System.Windows.Forms.TabPage]$TabPage)
    
    # Info-Label
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 20)
    $infoLabel.Size = New-Object System.Drawing.Size(460, 40)
    $infoLabel.Text = "Dieser Tab ermöglicht die Installation des UCMA Redistributable Pakets."
    $infoLabel.Font = New-Object System.Drawing.Font("Verdana", 10)
    $TabPage.Controls.Add($infoLabel)

    # Anweisungs-Label
    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Location = New-Object System.Drawing.Point(20, 70)
    $instructionLabel.Size = New-Object System.Drawing.Size(460, 40)
    $instructionLabel.Text = "Lege die Exchange-DVD ein oder mounte das ISO."
    $instructionLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($instructionLabel)

    # Status-Label und Text
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 120)
    $statusLabel.Size = New-Object System.Drawing.Size(100, 25)
    $statusLabel.Text = "Status:"
    $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($statusLabel)

    $statusText = New-Object System.Windows.Forms.Label
    $statusText.Location = New-Object System.Drawing.Point(120, 120)
    $statusText.Size = New-Object System.Drawing.Size(360, 25)
    $statusText.Text = "Bereit"
    $statusText.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($statusText)
    $script:statusText = $statusText

    # Pfad-Label und Textbox
    $pathLabel = New-Object System.Windows.Forms.Label
    $pathLabel.Location = New-Object System.Drawing.Point(20, 160)
    $pathLabel.Size = New-Object System.Drawing.Size(100, 25)
    $pathLabel.Text = "Setup-Pfad:"
    $pathLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($pathLabel)

    $pathTextBox = New-Object System.Windows.Forms.TextBox
    $pathTextBox.Location = New-Object System.Drawing.Point(120, 160)
    $pathTextBox.Size = New-Object System.Drawing.Size(360, 25)
    $pathTextBox.ReadOnly = $true
    $pathTextBox.BackColor = [System.Drawing.SystemColors]::Window
    $TabPage.Controls.Add($pathTextBox)
    $script:pathTextBox = $pathTextBox

    # Such-Button
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(20, 200)
    $searchButton.Size = New-Object System.Drawing.Size(200, 40)
    $searchButton.Text = "Nach UCMA suchen"
    $searchButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($searchButton)

    # Installations-Button
    $installButton = New-Object System.Windows.Forms.Button
    $installButton.Location = New-Object System.Drawing.Point(280, 200)
    $installButton.Size = New-Object System.Drawing.Size(200, 40)
    $installButton.Text = "UCMA installieren"
    $installButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($installButton)
    $script:installButton = $installButton

    # Ergebnis-Label
    $resultLabel = New-Object System.Windows.Forms.Label
    $resultLabel.Location = New-Object System.Drawing.Point(20, 260)
    $resultLabel.Size = New-Object System.Drawing.Size(460, 25)
    $resultLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $resultLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($resultLabel)
    $script:resultLabel = $resultLabel

    # Prüfen, ob UCMA bereits installiert ist
    if (Test-UCMAInstalled) {
        $statusText.Text = "UCMA ist bereits installiert"
        $resultLabel.Text = "UCMA ist bereits auf diesem System installiert."
        $resultLabel.ForeColor = [System.Drawing.Color]::Green
        $installButton.Enabled = $false
        $installButton.Text = "UCMA bereits installiert"
    }

    # Such-Button-Klick-Handler
    $searchButton.Add_Click({
        # Status aktualisieren
        $statusText.Text = "Suche nach UCMA Setup..."
        $resultLabel.Text = ""
        $pathTextBox.Text = ""
        
        # UI aktualisieren
        [System.Windows.Forms.Application]::DoEvents()
        
        # Nach setup.exe suchen
        $setupPath = Find-UCMASetupExe
        
        if ($setupPath -ne $null) {
            $statusText.Text = "UCMA Setup gefunden"
            $pathTextBox.Text = $setupPath
            $resultLabel.Text = "Setup.exe gefunden. Klicke auf 'UCMA installieren'."
            $resultLabel.ForeColor = [System.Drawing.Color]::Green
            
            # Exchange-Pfad-Aktualisierung nur, wenn Exchange-Tab existiert
            if ($script:exchangePathTextBox -ne $null -and $script:exchangeInstallButton -ne $null) {
                $possibleExchangePath = $setupPath -replace "UCMARedist\\setup\.exe$", "setup.exe"
                if (Test-Path $possibleExchangePath) {
                    $script:exchangePathTextBox.Text = $possibleExchangePath
                    $script:exchangeStatusText.Text = "Exchange Setup gefunden"
                    $script:exchangeResultLabel.Text = "Exchange Setup.exe wurde automatisch gefunden. Wechsle zum Exchange-Tab, um die Installation zu starten."
                    $script:exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
                }
            }
        } else {
            $statusText.Text = "UCMA Setup nicht gefunden"
            $resultLabel.Text = "Setup.exe wurde nicht gefunden. Bitte lege die Exchange-DVD ein oder mounte das ISO."
            $resultLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # Installations-Button-Klick-Handler
    $installButton.Add_Click({
        $setupPath = $pathTextBox.Text
        
        if (-not [string]::IsNullOrEmpty($setupPath) -and (Test-Path $setupPath)) {
            # UI-Status aktualisieren
            $statusText.Text = "Installation läuft..."
            $resultLabel.Text = "UCMA wird installiert, bitte warten..."
            $resultLabel.ForeColor = [System.Drawing.Color]::Blue
            
            # UI aktualisieren
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # Bestätigungsdialog
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    "Die UCMA-Installation wird jetzt gestartet.`n`nMöchtest du fortfahren?", 
                    "UCMA installieren", 
                    [System.Windows.Forms.MessageBoxButtons]::YesNo, 
                    [System.Windows.Forms.MessageBoxIcon]::Question)
                
                if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) {
                    $statusText.Text = "Installation abgebrochen"
                    $resultLabel.Text = "Die Installation wurde abgebrochen."
                    $resultLabel.ForeColor = [System.Drawing.Color]::Red
                    return
                }
                
                # Installation starten
                $process = Start-Process -FilePath $setupPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
                
                # Ergebnis anzeigen
                if ($process.ExitCode -eq 0) {
                    $statusText.Text = "Installation erfolgreich"
                    $resultLabel.Text = "UCMA Redistributable wurde erfolgreich installiert."
                    $resultLabel.ForeColor = [System.Drawing.Color]::Green
                    
                    # Deaktiviere den Installations-Button nach erfolgreicher Installation
                    $installButton.Enabled = $false
                    $installButton.Text = "UCMA bereits installiert"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "UCMA Redistributable wurde erfolgreich installiert.", 
                        "Erfolg", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    $statusText.Text = "Installationsfehler"
                    $resultLabel.Text = "Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen."
                    $resultLabel.ForeColor = [System.Drawing.Color]::Red
                    $searchButton.Enabled = $true
                    [System.Windows.Forms.MessageBox]::Show(
                        "Die Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen. Möglicherweise sind zusätzliche Maßnahmen erforderlich.", 
                        "Warnung", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            }
            catch {
                # Fehlerbehandlung
                $statusText.Text = "Fehler"
                $resultLabel.Text = "Fehler bei der Installation: $_"
                $resultLabel.ForeColor = [System.Drawing.Color]::Red
                [System.Windows.Forms.MessageBox]::Show(
                    "Fehler bei der Installation: $_", 
                    "Fehler", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            $resultLabel.Text = "Setup-Pfad ist nicht mehr gültig. Bitte erneut suchen."
            $resultLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # Automatische Suche beim Betreten des Tabs (nur wenn UCMA noch nicht installiert ist)
    if (-not (Test-UCMAInstalled)) {
        $searchButton.PerformClick()
    }
}

# Exchange-Tab initialisieren - Wird nur aufgerufen, wenn der Tab ausgewählt wird
function Initialize-ExchangeTab {
    param([System.Windows.Forms.TabPage]$TabPage)
    
    # Exchange-Header-Label
    $exchangeHeaderLabel = New-Object System.Windows.Forms.Label
    $exchangeHeaderLabel.Location = New-Object System.Drawing.Point(20, 20)
    $exchangeHeaderLabel.Size = New-Object System.Drawing.Size(460, 40)
    $exchangeHeaderLabel.Text = "Exchange Server Installation"
    $exchangeHeaderLabel.Font = New-Object System.Drawing.Font("Verdana", 12, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($exchangeHeaderLabel)

    # Info-Label
    $exchangeInfoLabel = New-Object System.Windows.Forms.Label
    $exchangeInfoLabel.Location = New-Object System.Drawing.Point(20, 70)
    $exchangeInfoLabel.Size = New-Object System.Drawing.Size(460, 40)
    $exchangeInfoLabel.Text = "Bitte stelle sicher, dass alle Voraussetzungen erfüllt sind, bevor du mit der Exchange-Installation fortfährst."
    $exchangeInfoLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($exchangeInfoLabel)

    # Setup-Pfad-Label
    $exchangePathLabel = New-Object System.Windows.Forms.Label
    $exchangePathLabel.Location = New-Object System.Drawing.Point(20, 120)
    $exchangePathLabel.Size = New-Object System.Drawing.Size(120, 25)
    $exchangePathLabel.Text = "Exchange Setup:"
    $exchangePathLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($exchangePathLabel)

    # Setup-Pfad-Textbox
    $exchangePathTextBox = New-Object System.Windows.Forms.TextBox
    $exchangePathTextBox.Location = New-Object System.Drawing.Point(150, 120)
    $exchangePathTextBox.Size = New-Object System.Drawing.Size(330, 25)
    $exchangePathTextBox.ReadOnly = $true
    $exchangePathTextBox.BackColor = [System.Drawing.SystemColors]::Window
    $TabPage.Controls.Add($exchangePathTextBox)
    $script:exchangePathTextBox = $exchangePathTextBox

    # Status-Label
    $exchangeStatusLabel = New-Object System.Windows.Forms.Label
    $exchangeStatusLabel.Location = New-Object System.Drawing.Point(20, 160)
    $exchangeStatusLabel.Size = New-Object System.Drawing.Size(120, 25)
    $exchangeStatusLabel.Text = "Status:"
    $exchangeStatusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($exchangeStatusLabel)

    # Status-Text
    $exchangeStatusText = New-Object System.Windows.Forms.Label
    $exchangeStatusText.Location = New-Object System.Drawing.Point(150, 160)
    $exchangeStatusText.Size = New-Object System.Drawing.Size(330, 25)
    $exchangeStatusText.Text = "Bereit"
    $exchangeStatusText.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($exchangeStatusText)
    $script:exchangeStatusText = $exchangeStatusText

    # Such-Button
    $exchangeSearchButton = New-Object System.Windows.Forms.Button
    $exchangeSearchButton.Location = New-Object System.Drawing.Point(20, 200)
    $exchangeSearchButton.Size = New-Object System.Drawing.Size(200, 40)
    $exchangeSearchButton.Text = "Nach Exchange suchen"
    $exchangeSearchButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($exchangeSearchButton)

    # Installations-Button
    $exchangeInstallButton = New-Object System.Windows.Forms.Button
    $exchangeInstallButton.Location = New-Object System.Drawing.Point(280, 200)
    $exchangeInstallButton.Size = New-Object System.Drawing.Size(200, 40)
    $exchangeInstallButton.Text = "Exchange installieren"
    $exchangeInstallButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($exchangeInstallButton)
    $script:exchangeInstallButton = $exchangeInstallButton

    # Ergebnis-Label
    $exchangeResultLabel = New-Object System.Windows.Forms.Label
    $exchangeResultLabel.Location = New-Object System.Drawing.Point(20, 260)
    $exchangeResultLabel.Size = New-Object System.Drawing.Size(460, 50)
    $exchangeResultLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $exchangeResultLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
    $TabPage.Controls.Add($exchangeResultLabel)
    $script:exchangeResultLabel = $exchangeResultLabel

    # Neustart-Status-Label
    $rebootStatusLabel = New-Object System.Windows.Forms.Label
    $rebootStatusLabel.Location = New-Object System.Drawing.Point(20, 320)
    $rebootStatusLabel.Size = New-Object System.Drawing.Size(460, 30)
    $rebootStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $rebootStatusLabel.Font = New-Object System.Drawing.Font("Verdana", 9, [System.Drawing.FontStyle]::Bold)
    $TabPage.Controls.Add($rebootStatusLabel)
    $script:rebootStatusLabel = $rebootStatusLabel
    
    # Neustart-Status aktualisieren
    Update-RebootStatus

    # Such-Button-Klick-Handler
    $exchangeSearchButton.Add_Click({
        # Status aktualisieren
        $exchangeStatusText.Text = "Suche nach Exchange Setup..."
        $exchangeResultLabel.Text = ""
        $exchangePathTextBox.Text = ""
        
        # UI aktualisieren
        [System.Windows.Forms.Application]::DoEvents()
        
        # UCMA-Pfad verwenden, um Exchange zu finden
        $ucmaPath = $script:pathTextBox.Text
        
        # Nach Setup.exe suchen
        $setupPath = Find-ExchangeSetupExe -UcmaPath $ucmaPath
        
        if ($setupPath -ne $null) {
            $exchangeStatusText.Text = "Exchange Setup gefunden"
            $exchangePathTextBox.Text = $setupPath
            $exchangeResultLabel.Text = "Setup.exe gefunden. Klicke auf 'Exchange installieren', um die Installation zu starten."
            $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
        } else {
            $exchangeStatusText.Text = "Exchange Setup nicht gefunden"
            $exchangeResultLabel.Text = "Setup.exe wurde nicht gefunden. Bitte lege die Exchange-DVD ein oder mounte das ISO."
            $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # Installations-Button-Klick-Handler
# Exchange Installations-Button-Klick-Handler - deutlich verbessert
$exchangeInstallButton.Add_Click({
    $setupPath = $exchangePathTextBox.Text
    
    if (-not [string]::IsNullOrEmpty($setupPath) -and (Test-Path $setupPath)) {
        # Neustart-Status vor Beginn der Installation aktualisieren
        Update-RebootStatus
        
        # Auf ausstehenden Neustart prüfen
        if (Test-PendingReboot) {
            $rebootResult = [System.Windows.Forms.MessageBox]::Show(
                "Es steht ein Systemneustart aus. Für eine erfolgreiche Exchange-Installation wird dringend empfohlen, den Computer zuerst neu zu starten.`n`nMöchtest du den Computer jetzt neu starten?", 
                "Neustart erforderlich", 
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, 
                [System.Windows.Forms.MessageBoxIcon]::Warning)
                
            if ($rebootResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Autostart vor dem Neustart einrichten
                Set-AutoStartAfterReboot
                
                # Dem Benutzer einen Moment Zeit geben, die Bestätigung zu sehen
                Start-Sleep -Seconds 2
                        
                # Computer neu starten
                Restart-Computer -Force
                return
            }
            elseif ($rebootResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                # Installation abbrechen
                return
            }
            # Wenn "Nein", mit der Installation trotz ausstehendem Neustart fortfahren
        }
        
        # Nach Installationsmodus fragen
        $installMode = [System.Windows.Forms.MessageBox]::Show(
            "Möchtest du die Exchange-Installation im unbeaufsichtigten Modus ausführen?`n`nWähle 'Ja' für die unbeaufsichtigte Installation oder 'Nein' für den interaktiven Setup-Assistenten.", 
            "Installationsmodus", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        # Parameter für unbeaufsichtigte oder interaktive Installation
        $arguments = ""
        $installationType = ""
        
        if ($installMode -eq [System.Windows.Forms.DialogResult]::Yes) {
            $installationType = "unbeaufsichtigt"
            # Installations-Parameter-Formular anzeigen
            $paramForm = New-Object System.Windows.Forms.Form
            $paramForm.Text = "Exchange Installationsparameter"
            $paramForm.Size = New-Object System.Drawing.Size(500, 400)
            $paramForm.StartPosition = "CenterScreen"
            $paramForm.FormBorderStyle = "FixedDialog"
            $paramForm.MaximizeBox = $false
            $paramForm.MinimizeBox = $false
            
            # Server- und Organisationsname
            $serverNameLabel = New-Object System.Windows.Forms.Label
            $serverNameLabel.Location = New-Object System.Drawing.Point(20, 20)
            $serverNameLabel.Size = New-Object System.Drawing.Size(460, 20)
            $serverNameLabel.Text = "Bitte gib den Servernamen ein:"
            $paramForm.Controls.Add($serverNameLabel)
            
            $serverNameTextBox = New-Object System.Windows.Forms.TextBox
            $serverNameTextBox.Location = New-Object System.Drawing.Point(20, 45)
            $serverNameTextBox.Size = New-Object System.Drawing.Size(440, 25)
            $serverNameTextBox.Text = $env:COMPUTERNAME
            $paramForm.Controls.Add($serverNameTextBox)
            
            $organizationLabel = New-Object System.Windows.Forms.Label
            $organizationLabel.Location = New-Object System.Drawing.Point(20, 80)
            $organizationLabel.Size = New-Object System.Drawing.Size(460, 20)
            $organizationLabel.Text = "Bitte gib den Organisationsnamen ein:"
            $paramForm.Controls.Add($organizationLabel)
            
            $organizationTextBox = New-Object System.Windows.Forms.TextBox
            $organizationTextBox.Location = New-Object System.Drawing.Point(20, 105)
            $organizationTextBox.Size = New-Object System.Drawing.Size(440, 25)
            $organizationTextBox.Text = "Organisation"
            $paramForm.Controls.Add($organizationTextBox)
            
            # Installationspfad
            $installPathLabel = New-Object System.Windows.Forms.Label
            $installPathLabel.Location = New-Object System.Drawing.Point(20, 140)
            $installPathLabel.Size = New-Object System.Drawing.Size(460, 20)
            $installPathLabel.Text = "Installationspfad für Exchange:"
            $paramForm.Controls.Add($installPathLabel)
            
            $installPathTextBox = New-Object System.Windows.Forms.TextBox
            $installPathTextBox.Location = New-Object System.Drawing.Point(20, 165)
            $installPathTextBox.Size = New-Object System.Drawing.Size(350, 25)
            $installPathTextBox.Text = "C:\Program Files\Microsoft\Exchange Server\V15"
            $paramForm.Controls.Add($installPathTextBox)
            
            $installPathBrowseButton = New-Object System.Windows.Forms.Button
            $installPathBrowseButton.Location = New-Object System.Drawing.Point(380, 165)
            $installPathBrowseButton.Size = New-Object System.Drawing.Size(80, 25)
            $installPathBrowseButton.Text = "Durchsuchen"
            $installPathBrowseButton.Add_Click({
                $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderBrowser.Description = "Wähle den Installationspfad für Exchange aus"
                $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
                
                if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $installPathTextBox.Text = $folderBrowser.SelectedPath
                }
            })
            $paramForm.Controls.Add($installPathBrowseButton)
            
            # Datenbankpfad
            $dbPathLabel = New-Object System.Windows.Forms.Label
            $dbPathLabel.Location = New-Object System.Drawing.Point(20, 200)
            $dbPathLabel.Size = New-Object System.Drawing.Size(460, 20)
            $dbPathLabel.Text = "Pfad für die Exchange-Datenbank:"
            $paramForm.Controls.Add($dbPathLabel)
            
            $dbPathTextBox = New-Object System.Windows.Forms.TextBox
            $dbPathTextBox.Location = New-Object System.Drawing.Point(20, 225)
            $dbPathTextBox.Size = New-Object System.Drawing.Size(350, 25)
            $dbPathTextBox.Text = "C:\ExchangeDatabases\MDB"
            $paramForm.Controls.Add($dbPathTextBox)
            
            $dbPathBrowseButton = New-Object System.Windows.Forms.Button
            $dbPathBrowseButton.Location = New-Object System.Drawing.Point(380, 225)
            $dbPathBrowseButton.Size = New-Object System.Drawing.Size(80, 25)
            $dbPathBrowseButton.Text = "Durchsuchen"
            $dbPathBrowseButton.Add_Click({
                $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderBrowser.Description = "Wähle den Pfad für die Exchange-Datenbank aus"
                $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
                
                if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $dbPathTextBox.Text = $folderBrowser.SelectedPath
                }
            })
            $paramForm.Controls.Add($dbPathBrowseButton)
            
            # Log-Pfad
            $logPathLabel = New-Object System.Windows.Forms.Label
            $logPathLabel.Location = New-Object System.Drawing.Point(20, 260)
            $logPathLabel.Size = New-Object System.Drawing.Size(460, 20)
            $logPathLabel.Text = "Pfad für die Exchange-Logdateien:"
            $paramForm.Controls.Add($logPathLabel)
            
            $logPathTextBox = New-Object System.Windows.Forms.TextBox
            $logPathTextBox.Location = New-Object System.Drawing.Point(20, 285)
            $logPathTextBox.Size = New-Object System.Drawing.Size(350, 25)
            $logPathTextBox.Text = "C:\ExchangeDatabases\MDB\Logs"
            $paramForm.Controls.Add($logPathTextBox)
            
            $logPathBrowseButton = New-Object System.Windows.Forms.Button
            $logPathBrowseButton.Location = New-Object System.Drawing.Point(380, 285)
            $logPathBrowseButton.Size = New-Object System.Drawing.Size(80, 25)
            $logPathBrowseButton.Text = "Durchsuchen"
            $logPathBrowseButton.Add_Click({
                $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                $folderBrowser.Description = "Wähle den Pfad für die Exchange-Logdateien aus"
                $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
                
                if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $logPathTextBox.Text = $folderBrowser.SelectedPath
                }
            })
            $paramForm.Controls.Add($logPathBrowseButton)
            
            # OK/Abbrechen-Buttons
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Location = New-Object System.Drawing.Point(300, 330)
            $okButton.Size = New-Object System.Drawing.Size(75, 30)
            $okButton.Text = "OK"
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $paramForm.Controls.Add($okButton)
            $paramForm.AcceptButton = $okButton
            
            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Location = New-Object System.Drawing.Point(385, 330)
            $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
            $cancelButton.Text = "Abbrechen"
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $paramForm.Controls.Add($cancelButton)
            $paramForm.CancelButton = $cancelButton
            
            $result = $paramForm.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $serverName = $serverNameTextBox.Text.Trim()
                $organizationName = $organizationTextBox.Text.Trim()
                $installPath = $installPathTextBox.Text.Trim()
                $dbPath = $dbPathTextBox.Text.Trim()
                $logPath = $logPathTextBox.Text.Trim()
                
                # Verzeichnisse erstellen, falls sie nicht existieren
                try {
                    if (-not (Test-Path $dbPath)) {
                        New-Item -Path $dbPath -ItemType Directory -Force | Out-Null
                    }
                    
                    if (-not (Test-Path $logPath)) {
                        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
                    }
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Fehler beim Erstellen der Verzeichnisse: $_", 
                        "Fehler", 
                        [System.Windows.Forms.MessageBoxButtons]::OK, 
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
                
                # Datenbankdateiname
                $dbFile = Join-Path -Path $dbPath -ChildPath "MDB.edb"
                
                # Parameter für unbeaufsichtigte Installation
                $arguments = "/Mode:Install /Role:Mailbox /OrganizationName:`"$organizationName`" /IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF /InstallWindowsComponents /TargetDir:`"$installPath`" /MdbName:Exchange /DbFilePath:`"$dbFile`" /LogFolderPath:`"$logPath`""
            }
            else {
                # Installation abbrechen
                return
            }
        } else {
            $installationType = "interaktiv"
        }
        
        # Installation vorbereiten
        # Zuerst UI-Status aktualisieren
        $exchangeStatusText.Text = "Installation läuft..."
        $exchangeResultLabel.Text = "Exchange wird installiert, bitte warten... Dies kann länger dauern."
        $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Blue
        
        # UI aktualisieren
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Speichere den Installationstyp in einer globalen Variablen
            $script:exchangeInstallationType = $installationType
            
            # Beginne die Installation
            if ([string]::IsNullOrEmpty($arguments)) {
                # Interaktiver Modus - starte ohne zu warten
                $process = Start-Process -FilePath $setupPath -PassThru
                
                # Setze den Status auf "Läuft" und halte ihn dort
                $exchangeStatusText.Text = "Installation läuft"
                $exchangeResultLabel.Text = "Exchange-Setup läuft im interaktiven Modus. Bitte folge den Anweisungen im Setup-Assistenten."
                $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
                
                # Speichere die Prozess-ID zur späteren Überprüfung
                $script:exchangeInstallProcessId = $process.Id
            }
            else {
                # Unbeaufsichtigter Modus - starte ohne zu warten
                $process = Start-Process -FilePath $setupPath -ArgumentList $arguments -PassThru
                
                # Speichere die Prozess-ID zur späteren Überprüfung
                $script:exchangeInstallProcessId = $process.Id
                
                # Setze den Status auf "Läuft" und halte ihn dort
                $exchangeStatusText.Text = "Installation läuft"
                $exchangeResultLabel.Text = "Exchange wird unbeaufsichtigt installiert. Dies kann länger dauern (typischerweise 30-60 Minuten). Bitte gedulde dich."
                $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
            }
            
            # Erstelle eine Flag-Datei, um anzuzeigen, dass eine Installation läuft
            $installFlagPath = "$env:TEMP\ExchangeInstallationRunning.flag"
            Set-Content -Path $installFlagPath -Value "Exchange-Installation im $installationType Modus gestartet am $(Get-Date)" -Force
            
            # Erstelle einen Timer, um den Installationsstatus regelmäßig zu aktualisieren
            if (-not [System.Windows.Forms.Application]::OpenForms.Count -eq 0) {
                $statusTimer = New-Object System.Windows.Forms.Timer
                $statusTimer.Interval = 5000  # Prüfe alle 5 Sekunden
                $statusTimer.Add_Tick({
                    # Prüfe, ob die Flag-Datei existiert
                    if (Test-Path "$env:TEMP\ExchangeInstallationRunning.flag") {
                        # Stelle sicher, dass der Status auf "Läuft" bleibt
                        if ($exchangeStatusText.Text -ne "Installation läuft") {
                            $exchangeStatusText.Text = "Installation läuft"
                            $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
                            
                            if ($script:exchangeInstallationType -eq "interaktiv") {
                                $exchangeResultLabel.Text = "Exchange-Setup läuft im interaktiven Modus. Bitte folge den Anweisungen im Setup-Assistenten."
                            } else {
                                $exchangeResultLabel.Text = "Exchange wird unbeaufsichtigt installiert. Dies kann länger dauern. Bitte gedulde dich."
                            }
                        }
                    }
                })
                $statusTimer.Start()
                
                # Speichere den Timer in einer globalen Variable
                $script:exchangeInstallStatusTimer = $statusTimer
            }
            
        }
        catch {
            # Fehlerbehandlung
            $exchangeStatusText.Text = "Fehler"
            $exchangeResultLabel.Text = "Fehler beim Starten der Installation: $_"
            $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler beim Starten der Installation: $_", 
                "Fehler", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        $exchangeResultLabel.Text = "Setup-Pfad ist nicht mehr gültig. Bitte erneut suchen."
        $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Red
        $exchangeInstallButton.Enabled = $false
    }
})

    # Prüfen, ob wir bereits einen UCMA-Pfad haben
    $ucmaPath = $script:pathTextBox.Text
    
    if (-not [string]::IsNullOrEmpty($ucmaPath) -and $ucmaPath -match "^([A-Z]:)\\UCMARedist\\setup\.exe$") {
        # Prüfen, ob das Exchange-Setup am erwarteten Speicherort existiert
        $driveLetter = $matches[1]
        $possibleSetupPath = "$driveLetter\setup.exe"
        
        if (Test-Path $possibleSetupPath) {
            $exchangePathTextBox.Text = $possibleSetupPath
            $exchangeStatusText.Text = "Exchange Setup automatisch gefunden"
            $exchangeResultLabel.Text = "Exchange Setup.exe wurde automatisch gefunden. Klicke auf 'Exchange installieren', um die Installation zu starten."
            $exchangeResultLabel.ForeColor = [System.Drawing.Color]::Green
        }
        else {
            # Automatische Suche, wenn UCMA bekannt ist, aber Exchange nicht am erwarteten Speicherort gefunden wurde
            $exchangeSearchButton.PerformClick()
        }
    }
    else {
        # Automatische Suche, wenn noch nichts bekannt ist
        $exchangeSearchButton.PerformClick()
    }
}

# HAUPTSKRIPT - Ausführung beginnt hier

# Zuerst prüfen, ob wir nach einem Neustart laufen
$isPostReboot = $false
if (Test-Path "HKCU:\Software\ExchangeRequirements") {
    try {
        $isPostReboot = (Get-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -ErrorAction SilentlyContinue).PostRebootExecution -eq $true
    } catch {
        $isPostReboot = $false
    }
}

# Autostart-Eintrag entfernen, um zukünftige automatische Starts zu verhindern
Remove-AutoStartAfterReboot

# Hauptformular initialisieren
$form = Initialize-MainForm

# Formular-Anzeige-Ereignis
$form.Add_Shown({
    # Zuerst UI anzeigen, dann Status prüfen
    $form.Activate()
    
    # Neustart-Benachrichtigung nur anzeigen, wenn wir nach einem Neustart automatisch gestartet wurden
    if ($isPostReboot) {
        [System.Windows.Forms.MessageBox]::Show(
            "Das Skript wurde nach dem Neustart automatisch gestartet.", 
            "Nach Neustart", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    # Timer-Variable im richtigen Gültigkeitsbereich erstellen
    $script:statusCheckTimer = New-Object System.Windows.Forms.Timer
    $script:statusCheckTimer.Interval = 500  # Alle halbe Sekunde prüfen
    
    # Hintergrundjob starten, um den Status zu laden, während der Benutzer mit der UI interagiert
    Start-Job -ScriptBlock {
        param ($features, $patterns)
        
        # Windows-Features prüfen
        $featuresInstalled = $false
        try {
            $isServer = [bool](Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue)
            if ($isServer) {
                $installedCount = (Get-WindowsFeature -Name $features -ErrorAction SilentlyContinue | 
                                  Where-Object { $_.Installed -eq $true }).Count
                $featuresInstalled = ($installedCount -eq $features.Count)
            } else {
                $installedCount = 0
                foreach ($feature in $features) {
                    $featureInfo = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
                    if ($featureInfo -and $featureInfo.State -eq "Enabled") { $installedCount++ }
                }
                $featuresInstalled = ($installedCount -eq $features.Count)
            }
        } catch {}
        
        # VC++ Redistributables prüfen
        $uninstallInfo = @()
        try {
            $uninstallInfo += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
            $uninstallInfo += Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue
        } catch {}
        
        $vc2012Installed = ($uninstallInfo | Where-Object { $_.DisplayName -like $patterns[0] }) -ne $null
        $vc2013Installed = ($uninstallInfo | Where-Object { $_.DisplayName -like $patterns[1] }) -ne $null
        
        # URL Rewrite prüfen
        $urlRewriteInstalled = (Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll")
        
        # Echtzeitschutz prüfen
        $realtimeDisabled = $false
        try {
            $realtimeDisabled = (Get-MpPreference -ErrorAction SilentlyContinue).DisableRealtimeMonitoring -eq $true
        } catch {}
        
        # Ergebnisse zurückgeben
        @{
            FeaturesInstalled = $featuresInstalled
            VC2012Installed = $vc2012Installed
            VC2013Installed = $vc2013Installed
            URLRewriteInstalled = $urlRewriteInstalled
            RealtimeDisabled = $realtimeDisabled
        }
    } -ArgumentList @(
        @("NET-Framework-45-Features", "Web-Server", "Web-Asp-Net45"), 
        @("*Visual C++ 2012*", "*Visual C++ 2013*")
    ) -Name "StatusCheck" | Out-Null
    
    # Mit der UI-Interaktion fortfahren, während die Statusprüfung im Hintergrund läuft
    
    # Sichererer Event-Handler mit Timer-Variable im Skript-Gültigkeitsbereich
    $script:statusCheckTimer.Add_Tick({
        # Dies muss eine sichere Referenz sein
        if (-not $script:statusCheckTimer) { return }
        
        $job = Get-Job -Name "StatusCheck" -ErrorAction SilentlyContinue
        
        if ($job -and $job.State -eq "Completed") {
            # Ergebnisse abrufen
            $results = Receive-Job -Job $job
            Remove-Job -Job $job
            
            # Button-Zustände aktualisieren
            if ($script:buttonInstallFeatures -and $results.FeaturesInstalled) {
                $script:buttonInstallFeatures.Enabled = $false
                $script:buttonInstallFeatures.Text = "Windows-Features bereits installiert"
            }
            
            if ($script:buttonDisableRealtime -and $results.RealtimeDisabled) {
                $script:buttonDisableRealtime.Enabled = $false
                $script:buttonDisableRealtime.Text = "Echtzeitschutz bereits deaktiviert"
            }
            
            if ($script:buttonInstallVC -and $results.VC2012Installed) {
                $script:buttonInstallVC.Enabled = $false
                $script:buttonInstallVC.Text = "Visual C++ 2012 Redistributable bereits installiert"
            }
            
            if ($script:buttonInstallVC2013 -and $results.VC2013Installed) {
                $script:buttonInstallVC2013.Enabled = $false
                $script:buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable bereits installiert"
            }
            
            if ($script:buttonInstallURLRewrite -and $results.URLRewriteInstalled) {
                $script:buttonInstallURLRewrite.Enabled = $false
                $script:buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul bereits installiert"
            }
            
            # Timer stoppen und freigeben - sicher
            try {
                # Sicherstellen, dass der Timer noch existiert, bevor er gestoppt wird
                if ($script:statusCheckTimer) {
                    $script:statusCheckTimer.Stop()
                    $script:statusCheckTimer.Dispose()
                    $script:statusCheckTimer = $null
                }
            } catch {
                # Fehler beim Stoppen des Timers ignorieren
                Write-Host "Warnung: Fehler beim Aufräumen des Timers: $_"
            }
        }
    })
    $script:statusCheckTimer.Start()
})

# Formular anzeigen - dies ist ein blockierender Aufruf
[void]$form.ShowDialog()
