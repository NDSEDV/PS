<#
.SYNOPSIS
  Windows Exchange Server Pre-Konfiguration
.DESCRIPTION
  The tool is intended to help you with your daily business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-03-03

  Copyright (c) Jörn Walter. All rights reserved.
#>

# Funktion zum Einrichten des Autostarts nach dem Neustart
function Set-AutoStartAfterReboot {
    try {
        # Skriptpfad ermitteln
        $scriptPath = $PSCommandPath
        if (-not $scriptPath) {
            throw "Skriptpfad konnte nicht ermittelt werden."
        }

        Ensure-RegistryPath

        # Batch-Datei erstellen, die das PowerShell-Skript ausführt
        $batchContent = @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$scriptPath"
"@

        # Startordner-Pfad definieren
        $startupFolder = [System.Environment]::GetFolderPath('Startup')
        $batchPath = Join-Path -Path $startupFolder -ChildPath "ExchangeRequirements.bat"

        # Batch-Datei erstellen
        Set-Content -Path $batchPath -Value $batchContent -Force

        # Flag in der Registry setzen, um anzuzeigen, dass es sich um eine Ausführung nach dem Neustart handelt
        Set-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -Value $true -Type DWORD -Force

        return $true
    }
    catch {
        Write-Error "Fehler beim Einrichten des Autostarts: $_"
        return $false
    }
}

# Funktion zum Entfernen des Autostart-Eintrags
function Remove-AutoStartAfterReboot {
    try {
        # Batch-Datei aus dem Startordner entfernen
        $startupFolder = [System.Environment]::GetFolderPath('Startup')
        $batchPath = Join-Path -Path $startupFolder -ChildPath "ExchangeRequirements.bat"
        
        if (Test-Path $batchPath) {
            Remove-Item -Path $batchPath -Force
        }

        # Registry-Schlüssel bereinigen, falls vorhanden
        if (Test-Path "HKCU:\Software\ExchangeRequirements") {
            Remove-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -ErrorAction SilentlyContinue
        }

        return $true
    }
    catch {
        Write-Error "Fehler beim Entfernen des Autostarts: $_"
        return $false
    }
}

# Prüfen, ob Registry-Pfad existiert, ggf. erstellen
function Ensure-RegistryPath {
    if (-not (Test-Path "HKCU:\Software\ExchangeRequirements")) {
        New-Item -Path "HKCU:\Software" -Name "ExchangeRequirements" -Force | Out-Null
    }
}

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
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

# Funktion zum Suchen nach UCMARedist\setup.exe auf allen Laufwerken
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Abrufen der Systemressourcen
function Get-SystemResources {
    try {
        # Abrufen der CPU-Informationen
        $cpuInfo = Get-WmiObject -Class Win32_Processor
        $coreCount = ($cpuInfo | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
        
        # Abrufen der RAM-Informationen
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        $totalRAM = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
        
        return @{
            CoreCount = $coreCount
            TotalRAM = $totalRAM
        }
    }
    catch {
        Write-Error "Fehler beim Abrufen der Systemressourcen: $_"
        return @{
            CoreCount = "Unbekannt"
            TotalRAM = "Unbekannt"
        }
    }
}

# Prüfen, ob Windows-Features installiert sind
function Test-WindowsFeatures {
    param (
        [string[]]$FeatureNames
    )
    
    # Warnmeldungen unterdrücken
    $WarningPreference = 'SilentlyContinue'
    
    # Logdatei für Debugging
    $logFile = "$env:TEMP\feature_check_log.txt"
    "Feature-Überprüfung gestartet am $(Get-Date)" | Out-File -FilePath $logFile
    
    # Zähler für installierte Features
    $installedCount = 0
    $totalCount = $FeatureNames.Count
    
    "Zu prüfende Features: $totalCount" | Out-File -FilePath $logFile -Append
    
    foreach ($feature in $FeatureNames) {
        try {
            "Prüfe Feature: $feature" | Out-File -FilePath $logFile -Append
            
            # Verwende Get-WindowsFeature wenn verfügbar (Windows Server)
            if (Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue) {
                $featureStatus = Get-WindowsFeature -Name $feature -ErrorAction SilentlyContinue
                
                if ($featureStatus -and $featureStatus.Installed) {
                    $installedCount++
                    "Feature $feature ist installiert (Get-WindowsFeature)." | Out-File -FilePath $logFile -Append
                } else {
                    "Feature $feature ist NICHT installiert (Get-WindowsFeature)." | Out-File -FilePath $logFile -Append
                }
            }
            # Alternative: Verwende DISM
            else {
                $tempFile = "$env:TEMP\feature_check_$feature.txt"
                $process = Start-Process -FilePath "DISM.exe" -ArgumentList "/Online /Get-FeatureInfo /FeatureName:$feature" -PassThru -Wait -NoNewWindow -RedirectStandardOutput $tempFile
                
                if ($process.ExitCode -eq 0) {
                    $output = Get-Content $tempFile -ErrorAction SilentlyContinue
                    
                    if ($output -match "State : Enabled" -or $output -match "Status : Aktiviert") {
                        $installedCount++
                        "Feature $feature ist installiert (DISM)." | Out-File -FilePath $logFile -Append
                    } else {
                        "Feature $feature ist NICHT installiert (DISM)." | Out-File -FilePath $logFile -Append
                    }
                } else {
                    "DISM-Befehl für Feature $feature fehlgeschlagen mit Exit-Code $($process.ExitCode)." | Out-File -FilePath $logFile -Append
                }
                
                # Bereinigen
                if (Test-Path $tempFile) {
                    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            "Fehler bei der Prüfung von Feature $feature $_" | Out-File -FilePath $logFile -Append
        }
    }
    
    # Auswertung: Sind alle erforderlichen Features installiert?
    $allFeaturesInstalled = ($installedCount -eq $totalCount)
    
    # Zusammenfassung in Log schreiben
    "Installierte Features: $installedCount von $totalCount" | Out-File -FilePath $logFile -Append
    "Alle Features installiert: $allFeaturesInstalled" | Out-File -FilePath $logFile -Append
    
    # Warnmeldungen zurücksetzen
    $WarningPreference = 'Continue'
    
    # Rückgabe true wenn ALLE Features installiert sind, sonst false
    return $allFeaturesInstalled
}

# Funktion zum Prüfen, ob Visual C++ Redistributable installiert ist
function Test-VCRedistInstalled {
    param (
        [string]$DisplayNamePattern
    )
    
    $installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                    Where-Object { $_.DisplayName -like $DisplayNamePattern } | 
                    Select-Object DisplayName, DisplayVersion
    
    if (-not $installed) {
        # Prüfe auch 32-Bit-Registry auf 64-Bit-Systemen
        $installed = Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                        Where-Object { $_.DisplayName -like $DisplayNamePattern } | 
                        Select-Object DisplayName, DisplayVersion
    }
    
    return $installed -ne $null
}

# Funktion zum Prüfen, ob IIS URL Rewrite installiert ist
function Test-URLRewriteInstalled {
    try {
        # Prüfe über Registry mit Anführungszeichen für den Pfad
        $regPath = "HKLM:\SOFTWARE\Microsoft\IIS Extensions\URL Rewrite"
        $installed = Get-ItemProperty -Path $regPath -ErrorAction Stop
        return $true
    }
    catch {
        try {
            # Alternative Methode über IIS Module
            Import-Module WebAdministration -ErrorAction SilentlyContinue
            $installed = Get-WebGlobalModule -Name "RewriteModule" -ErrorAction Stop
            return ($installed -ne $null)
        }
        catch {
            try {
                # Dritte Methode: Prüfe, ob die DLL existiert
                $iisRewriteDll = "$env:SystemRoot\System32\inetsrv\rewrite.dll"
                return (Test-Path $iisRewriteDll)
            }
            catch {
                return $false
            }
        }
    }
}

# Funktion zum Prüfen, ob der Echtzeitschutz deaktiviert ist
function Test-RealtimeProtectionDisabled {
    try {
        $mpPreference = Get-MpPreference -ErrorAction SilentlyContinue
        return $mpPreference.DisableRealtimeMonitoring -eq $true
    }
    catch {
        # Windows Defender möglicherweise nicht installiert oder kein Zugriff
        return $false
    }
}

# Erstellen des Hauptfensters
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Exchange Server Pre-Konfiguration"
$form.Size = New-Object System.Drawing.Size(520, 605)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Systeminformationen abrufen
$systemResources = Get-SystemResources

# Header mit Systeminformationen erstellen
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(500, 40)
$headerPanel.BackColor = [System.Drawing.Color]::LightGray

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(10, 10)
$headerLabel.Size = New-Object System.Drawing.Size(480, 20)
$headerLabel.Text = "System: $($systemResources.CoreCount) virtuelle Kerne | $($systemResources.TotalRAM) GB RAM"
$headerLabel.Font = New-Object System.Drawing.Font("Verdana", 9, [System.Drawing.FontStyle]::Bold)
$headerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$headerPanel.Controls.Add($headerLabel)
$form.Controls.Add($headerPanel)

# TabControl erstellen
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(0, 40) # Unterhalb des Header-Panels
$tabControl.Size = New-Object System.Drawing.Size(500, 490)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($tabControl)

# Haupttab erstellen
$tabMain = New-Object System.Windows.Forms.TabPage
$tabMain.Text = "Hauptfunktionen"
$tabControl.Controls.Add($tabMain)

# UCMA-Tab erstellen
$tabUCMA = New-Object System.Windows.Forms.TabPage
$tabUCMA.Text = "UCMA Installation"
$tabControl.Controls.Add($tabUCMA)

# Erstellt den Buttons für die Deaktivierung des Echtzeitschutzes
$buttonDisableRealtime = New-Object System.Windows.Forms.Button
$buttonDisableRealtime.Location = New-Object System.Drawing.Point(50, 50)
$buttonDisableRealtime.Size = New-Object System.Drawing.Size(400, 40)
$buttonDisableRealtime.Text = "Echtzeitschutz deaktivieren"
$buttonDisableRealtime.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$buttonDisableRealtime.Add_Click({
    try {
        Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Echtzeitschutz wurde erfolgreich deaktiviert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Prüfe den Status nach der Aktion und aktualisiere den Button
        if (Test-RealtimeProtectionDisabled) {
            $buttonDisableRealtime.Enabled = $false
            $buttonDisableRealtime.Text = "Echtzeitschutz bereits deaktiviert"
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Deaktivieren des Echtzeitschutzes: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$tabMain.Controls.Add($buttonDisableRealtime)

# Erstellen des Buttons für die Installation der Windows-Features
$buttonInstallFeatures = New-Object System.Windows.Forms.Button
$buttonInstallFeatures.Location = New-Object System.Drawing.Point(50, 120)
$buttonInstallFeatures.Size = New-Object System.Drawing.Size(400, 40)
$buttonInstallFeatures.Text = "Windows-Features installieren"
$buttonInstallFeatures.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$buttonInstallFeatures.Add_Click({
    try {
        # Definiere erforderliche Features für die Installation
        $requiredFeatures = @(
            "NET-Framework-45-Features", 
            "Web-Server", 
            "Web-Asp-Net45"
        )
        
        # Statusfenster anzeigen (um 40% verkleinert)
        $statusForm = New-Object System.Windows.Forms.Form
        $statusForm.Text = "Installation läuft..."
        $statusForm.Size = New-Object System.Drawing.Size(240, 90)
        $statusForm.StartPosition = "CenterScreen"
        $statusForm.FormBorderStyle = "FixedDialog"
        $statusForm.ControlBox = $false
        
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
        $statusLabel.Size = New-Object System.Drawing.Size(220, 60)
        $statusLabel.Text = "Die Windows-Features werden installiert. Dies kann einige Zeit dauern..."
        $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
        $statusForm.Controls.Add($statusLabel)
        
        $statusForm.Show()
        $form.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        
        # Warnmeldungen unterdrücken
        $WarningPreference = 'SilentlyContinue'
        $ErrorActionPreference = 'SilentlyContinue'
        
        # Erstelle eine PowerShell-Datei für die Installation mit Administratorrechten
        $installScript = @"
# Führt die Installation aller benötigten Windows-Features aus
Import-Module ServerManager

# Server-Features Liste
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

# Direktes Installieren aller Features mit Add-WindowsFeature
Add-WindowsFeature -Name `$features -IncludeManagementTools -Verbose

# Schreibe Ergebnis in Datei
`$result = Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }
`$result | Out-File "$env:TEMP\installed_features.txt"
"@

        # Schreibe das Installationsskript in eine temporäre Datei
        $scriptPath = "$env:TEMP\Install-WindowsFeatures.ps1"
        Set-Content -Path $scriptPath -Value $installScript -Force
        
        # Installationslog-Datei definieren
        $logFilePath = "$env:TEMP\installed_features.txt"
        
        # Führe das Skript mit PowerShell als Administrator aus
        $statusLabel.Text = "Installiere Windows-Features mit PowerShell..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Direktes PowerShell-Kommando zur Installation mit PowerShell
        $powerShellPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        
        $process = Start-Process -FilePath $powerShellPath -ArgumentList $arguments -Wait -PassThru
        
        # Alternative Installationsmethode, falls die erste fehlgeschlagen ist
        if ($process.ExitCode -ne 0) {
            $statusLabel.Text = "Alternative Methode: Installiere mit ServerManagerCmd..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Versuche mit ServerManagerCmd.exe zu installieren
            foreach ($feature in $features) {
                Start-Process -FilePath "ServerManagerCmd.exe" -ArgumentList "-install $feature" -Wait -NoNewWindow
            }
        }
        
        # Bereinigen
        if (Test-Path $scriptPath) {
            Remove-Item -Path $scriptPath -Force
        }
        
        # Überprüfe, ob alle erforderlichen Features installiert wurden
        $featuresInstalled = Test-WindowsFeatures -FeatureNames $requiredFeatures
        
        # Warnmeldungen zurücksetzen
        $WarningPreference = 'Continue'
        $ErrorActionPreference = 'Continue'
        
        $statusForm.Close()
        $form.Enabled = $true
        
        # Button je nach Installationsstatus deaktivieren
        if ($featuresInstalled) {
            $buttonInstallFeatures.Enabled = $false
            $buttonInstallFeatures.Text = "Windows-Features bereits installiert"
            Write-Host "Features wurden installiert, Button deaktiviert."
        } else {
            Write-Host "Features möglicherweise nicht vollständig installiert."
        }
        
        # Frage, ob das Log geöffnet werden soll
        $openLogResult = [System.Windows.Forms.MessageBox]::Show(
            "Die Windows-Features wurden installiert.`n`nMöchtest du das Installationslog öffnen?", 
            "Installation abgeschlossen", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
            
        # Öffne Log-Datei, wenn gewünscht
        if ($openLogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            if (Test-Path $logFilePath) {
                Start-Process "notepad.exe" -ArgumentList $logFilePath
            } else {
                [System.Windows.Forms.MessageBox]::Show("Die Log-Datei konnte nicht gefunden werden.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
        
        # Frage nach Neustart des Systems
        $restartResult = [System.Windows.Forms.MessageBox]::Show(
            "Ein Neustart wird empfohlen, um die Installation der Windows-Features abzuschließen.`n`nMöchtest du den Computer jetzt neu starten?", 
            "Neustart erforderlich", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        # Führe Neustart durch, wenn gewünscht
        if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Autostart vor dem Neustart einrichten
            $autoStartSetup = Set-AutoStartAfterReboot
            
            if ($autoStartSetup) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Das Skript wird nach dem Neustart automatisch fortgesetzt.", 
                    "Autostart aktiviert", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            
            # Kurze Verzögerung, damit die Meldung angezeigt werden kann
            Start-Sleep -Seconds 2
                    
            # Starte Neustart-Prozess
            Restart-Computer -Force
        }
        else {
            # Wenn kein Neustart gewünscht, aktualisiere den UI-Status
            if (-not $featuresInstalled) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Bitte beachten Sie, dass einige Features möglicherweise erst nach einem Neustart vollständig verfügbar sind.", 
                    "Hinweis", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    }
    catch {
        # Warnmeldungen zurücksetzen
        $WarningPreference = 'Continue'
        $ErrorActionPreference = 'Continue'
        
        if ($statusForm -ne $null) {
            $statusForm.Close()
            $form.Enabled = $true
        }
        
        # Fehlermeldung mit Details anzeigen
        [System.Windows.Forms.MessageBox]::Show("Fehler bei der Installation der Windows-Features: $_ `n`nBitte versuche folgendes: `n1. Starte PowerShell als Administrator `n2. Führe 'Install-WindowsFeature Web-Server,NET-Framework-45-Features' aus", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$tabMain.Controls.Add($buttonInstallFeatures)

# Erstellen des Buttons für die Installation von Visual C++ 2012 Redistributable
$buttonInstallVC = New-Object System.Windows.Forms.Button
$buttonInstallVC.Location = New-Object System.Drawing.Point(50, 190)
$buttonInstallVC.Size = New-Object System.Drawing.Size(400, 40)
$buttonInstallVC.Text = "Visual C++ 2012 Redistributable herunterladen und installieren"
$buttonInstallVC.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$buttonInstallVC.Add_Click({
    try {
        # Temporäres Verzeichnis für den Download
        $tempDir = [System.IO.Path]::GetTempPath()
        $vcRedistPath = Join-Path -Path $tempDir -ChildPath "vcredist_x64.exe"
        $url = "https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe"
        
        # Statusfenster für VC++ 2012 Installation
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
        $form.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        
        # Download der Datei
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $vcRedistPath)
        
        # Installation starten
        $process = Start-Process -FilePath $vcRedistPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
        
        $statusForm.Close()
        $form.Enabled = $true
        
        if ($process.ExitCode -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Visual C++ 2012 Redistributable wurde erfolgreich installiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Button deaktivieren nach erfolgreicher Installation
            $buttonInstallVC.Enabled = $false
            $buttonInstallVC.Text = "Visual C++ 2012 Redistributable bereits installiert"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Die Installation wurde abgeschlossen, möglicherweise mit Warnungen (Exit Code: $($process.ExitCode)).", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        
        # Aufräumen
        if (Test-Path $vcRedistPath) {
            Remove-Item -Path $vcRedistPath -Force
        }
    }
    catch {
        if ($statusForm -ne $null) {
            $statusForm.Close()
            $form.Enabled = $true
        }
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Herunterladen oder Installieren von Visual C++ 2012 Redistributable: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$tabMain.Controls.Add($buttonInstallVC)

# Erstellen des Buttons für die Installation von Visual C++ 2013 Redistributable
$buttonInstallVC2013 = New-Object System.Windows.Forms.Button
$buttonInstallVC2013.Location = New-Object System.Drawing.Point(50, 260)
$buttonInstallVC2013.Size = New-Object System.Drawing.Size(400, 40)
$buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable herunterladen und installieren"
$buttonInstallVC2013.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$buttonInstallVC2013.Add_Click({
    try {
        # Temporäres Verzeichnis für den Download
        $tempDir = [System.IO.Path]::GetTempPath()
        $vcRedistPath = Join-Path -Path $tempDir -ChildPath "vcredist_x64_2013.exe"
        $url = "https://download.visualstudio.microsoft.com/download/pr/10912041/cee5d6bca2ddbcd039da727bf4acb48a/vcredist_x64.exe"
        
        # Statusfenster für VC++ 2013 Installation
        $statusForm = New-Object System.Windows.Forms.Form
        $statusForm.Text = "Download und Installation..."
        $statusForm.Size = New-Object System.Drawing.Size(240, 100)
        $statusForm.StartPosition = "CenterScreen"
        $statusForm.FormBorderStyle = "FixedDialog"
        $statusForm.ControlBox = $false
        
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
        $statusLabel.Size = New-Object System.Drawing.Size(220, 70)
        $statusLabel.Text = "Visual C++ 2013 Redistributable wird heruntergeladen und installiert..."
        $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
        $statusForm.Controls.Add($statusLabel)
        
        $statusForm.Show()
        $form.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        
        # Download der Datei
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $vcRedistPath)
        
        # Installation starten
        $process = Start-Process -FilePath $vcRedistPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
        
        $statusForm.Close()
        $form.Enabled = $true
        
        if ($process.ExitCode -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Visual C++ 2013 Redistributable wurde erfolgreich installiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Button deaktivieren nach erfolgreicher Installation
            $buttonInstallVC2013.Enabled = $false
            $buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable bereits installiert"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Die Installation wurde abgeschlossen, möglicherweise mit Warnungen (Exit Code: $($process.ExitCode)).", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        
        # Aufräumen
        if (Test-Path $vcRedistPath) {
            Remove-Item -Path $vcRedistPath -Force
        }
    }
    catch {
        if ($statusForm -ne $null) {
            $statusForm.Close()
            $form.Enabled = $true
        }
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Herunterladen oder Installieren von Visual C++ 2013 Redistributable: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$tabMain.Controls.Add($buttonInstallVC2013)

# Erstellen des Buttons für die Installation des IIS URL Rewrite Moduls
$buttonInstallURLRewrite = New-Object System.Windows.Forms.Button
$buttonInstallURLRewrite.Location = New-Object System.Drawing.Point(50, 330)
$buttonInstallURLRewrite.Size = New-Object System.Drawing.Size(400, 50)
$buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul herunterladen und installieren"
$buttonInstallURLRewrite.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$buttonInstallURLRewrite.Add_Click({
    try {
        # Temporäres Verzeichnis für den Download
        $tempDir = [System.IO.Path]::GetTempPath()
        $urlRewritePath = Join-Path -Path $tempDir -ChildPath "rewrite_amd64_en-US.msi"
        $url = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
        
        # Statusfenster für IIS URL Rewrite Installation
        $statusForm = New-Object System.Windows.Forms.Form
        $statusForm.Text = "Download und Installation..."
        $statusForm.Size = New-Object System.Drawing.Size(240, 90) # Verkleinert um 40%
        $statusForm.StartPosition = "CenterScreen"
        $statusForm.FormBorderStyle = "FixedDialog"
        $statusForm.ControlBox = $false
        
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Location = New-Object System.Drawing.Point(10, 15)
        $statusLabel.Size = New-Object System.Drawing.Size(220, 70)
        $statusLabel.Text = "IIS URL Rewrite Modul wird heruntergeladen und installiert..."
        $statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
        $statusForm.Controls.Add($statusLabel)
        
        $statusForm.Show()
        $form.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        
        # Download der Datei
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $urlRewritePath)
        
        # Installation starten
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$urlRewritePath`" /quiet /norestart" -PassThru -Wait
        
        $statusForm.Close()
        $form.Enabled = $true
        
        if ($process.ExitCode -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("IIS URL Rewrite Modul wurde erfolgreich installiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Button deaktivieren nach erfolgreicher Installation
            $buttonInstallURLRewrite.Enabled = $false
            $buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul bereits installiert"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Die Installation wurde abgeschlossen, möglicherweise mit Warnungen (Exit Code: $($process.ExitCode)).", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        
        # Aufräumen
        if (Test-Path $urlRewritePath) {
            Remove-Item -Path $urlRewritePath -Force
        }
    }
    catch {
        if ($statusForm -ne $null) {
            $statusForm.Close()
            $form.Enabled = $true
        }
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Herunterladen oder Installieren des IIS URL Rewrite Moduls: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$tabMain.Controls.Add($buttonInstallURLRewrite)

# Hilfe-Icons mit Tooltip
$helpIcon = New-Object System.Windows.Forms.PictureBox
$helpIcon.Location = New-Object System.Drawing.Point(50, 395)
$helpIcon.Size = New-Object System.Drawing.Size(16, 16)
$helpIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage

# Fragezeichen-Icons
$bitmap = New-Object System.Drawing.Bitmap 16, 16
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::Transparent)
$font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Blue)
$graphics.DrawString("?", $font, $brush, 2, -2)
$graphics.Dispose()
$helpIcon.Image = $bitmap

# Hinzufügen eines Tooltips
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($helpIcon, "UCMARedist wird vom eingelegten Exchange-Datenträger installiert.")

# Hinzufügen eines Labels für mehr Kontext
$helpLabel = New-Object System.Windows.Forms.Label
$helpLabel.Location = New-Object System.Drawing.Point(75, 395)
$helpLabel.Size = New-Object System.Drawing.Size(375, 20)
$helpLabel.Text = "UCMARedist-Installation"
$helpLabel.Font = New-Object System.Drawing.Font("Verdana", 9)

$tabMain.Controls.Add($helpIcon)
$tabMain.Controls.Add($helpLabel)

# Hinzufügen eines Hinweises über Statusmeldungen
$statusInfoLabel = New-Object System.Windows.Forms.Label
$statusInfoLabel.Location = New-Object System.Drawing.Point(10, 420)
$statusInfoLabel.Size = New-Object System.Drawing.Size(480, 40)
$statusInfoLabel.Text = "Info: Einige Warnmeldungen bezüglich IIS werden unterdrückt. Diese haben keinen Einfluss auf die Funktionalität."
$statusInfoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusInfoLabel.Font = New-Object System.Drawing.Font("Verdana", 8)
$tabMain.Controls.Add($statusInfoLabel)

# Copyright Label erstellen
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Location = New-Object System.Drawing.Point(10, 535)
$copyrightLabel.Size = New-Object System.Drawing.Size(480, 30)
$copyrightLabel.Text = "© 2025 Jörn Walter https://www.der-windows-papst.de"
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$copyrightLabel.Font = New-Object System.Drawing.Font("Verdana", 8)
$copyrightLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($copyrightLabel)

# UCMA-Tab Steuerelemente erstellen
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(20, 20)
$infoLabel.Size = New-Object System.Drawing.Size(460, 40)
$infoLabel.Text = "Dieser Tab ermöglicht die Installation des UCMA Redistributable Pakets."
$infoLabel.Font = New-Object System.Drawing.Font("Verdana", 10)
$tabUCMA.Controls.Add($infoLabel)

$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Location = New-Object System.Drawing.Point(20, 70)
$instructionLabel.Size = New-Object System.Drawing.Size(460, 40)
$instructionLabel.Text = "Lege die Exchange-DVD ein oder mounte die ISO-Datei."
$instructionLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
$tabUCMA.Controls.Add($instructionLabel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 120)
$statusLabel.Size = New-Object System.Drawing.Size(100, 25)
$statusLabel.Text = "Status:"
$statusLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
$tabUCMA.Controls.Add($statusLabel)

$statusText = New-Object System.Windows.Forms.Label
$statusText.Location = New-Object System.Drawing.Point(120, 120)
$statusText.Size = New-Object System.Drawing.Size(360, 25)
$statusText.Text = "Bereit"
$statusText.Font = New-Object System.Drawing.Font("Verdana", 9)
$tabUCMA.Controls.Add($statusText)

$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(20, 160)
$pathLabel.Size = New-Object System.Drawing.Size(100, 25)
$pathLabel.Text = "Setup-Pfad:"
$pathLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
$tabUCMA.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(120, 160)
$pathTextBox.Size = New-Object System.Drawing.Size(360, 25)
$pathTextBox.ReadOnly = $true
$pathTextBox.BackColor = [System.Drawing.SystemColors]::Window
$tabUCMA.Controls.Add($pathTextBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(20, 200)
$searchButton.Size = New-Object System.Drawing.Size(200, 40)
$searchButton.Text = "Nach UCMA suchen"
$searchButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$tabUCMA.Controls.Add($searchButton)

$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(280, 200)
$installButton.Size = New-Object System.Drawing.Size(200, 40)
$installButton.Text = "UCMA installieren"
$installButton.Font = New-Object System.Drawing.Font("Verdana", 10, [System.Drawing.FontStyle]::Bold)
$installButton.Enabled = $false  # Initial deaktiviert
$tabUCMA.Controls.Add($installButton)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 260)
$progressBar.Size = New-Object System.Drawing.Size(460, 25)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 0  # Anfangs ausgeschaltet
$tabUCMA.Controls.Add($progressBar)

$resultLabel = New-Object System.Windows.Forms.Label
$resultLabel.Location = New-Object System.Drawing.Point(20, 300)
$resultLabel.Size = New-Object System.Drawing.Size(460, 25)
$resultLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$resultLabel.Font = New-Object System.Drawing.Font("Verdana", 9)
$tabUCMA.Controls.Add($resultLabel)

# Suchbutton-Handler
$searchButton.Add_Click({
    # Status aktualisieren
    $statusText.Text = "Suche nach UCMA Setup..."
    $resultLabel.Text = ""
    $pathTextBox.Text = ""
    $installButton.Enabled = $false
    $progressBar.MarqueeAnimationSpeed = 30  # Fortschrittsanzeige starten
    
    # UI aktualisieren
    [System.Windows.Forms.Application]::DoEvents()
    
    # Suche nach Setup.exe
    $setupPath = Find-UCMASetupExe
    
    # Fortschrittsanzeige stoppen
    $progressBar.MarqueeAnimationSpeed = 0
    
    if ($setupPath -ne $null) {
        $statusText.Text = "UCMA Setup gefunden"
        $pathTextBox.Text = $setupPath
        $installButton.Enabled = $true
        $resultLabel.Text = "Setup.exe gefunden. Klicken Sie auf 'UCMA installieren'."
        $resultLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $statusText.Text = "UCMA Setup nicht gefunden"
        $resultLabel.Text = "Setup.exe wurde nicht gefunden. Bitte lege die Exchange-DVD ein oder mounte die ISO."
        $resultLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

# Installationsbutton-Handler
$installButton.Add_Click({
    $setupPath = $pathTextBox.Text
    
    if (-not [string]::IsNullOrEmpty($setupPath) -and (Test-Path $setupPath)) {
        # UI-Status aktualisieren
        $statusText.Text = "Installation läuft..."
        $resultLabel.Text = "UCMA wird installiert, bitte warten..."
        $resultLabel.ForeColor = [System.Drawing.Color]::Blue
        $progressBar.MarqueeAnimationSpeed = 30
        $searchButton.Enabled = $false
        $installButton.Enabled = $false
        
        # UI aktualisieren
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Installation starten
            $process = Start-Process -FilePath $setupPath -ArgumentList "/quiet", "/norestart" -PassThru -Wait
            
            # Fortschrittsanzeige stoppen
            $progressBar.MarqueeAnimationSpeed = 0
            
            # Ergebnis anzeigen
            if ($process.ExitCode -eq 0) {
                $statusText.Text = "Installation erfolgreich"
                $resultLabel.Text = "UCMA Redistributable wurde erfolgreich installiert."
                $resultLabel.ForeColor = [System.Drawing.Color]::Green
                $searchButton.Enabled = $true
                [System.Windows.Forms.MessageBox]::Show("UCMA Redistributable wurde erfolgreich installiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                $statusText.Text = "Installationsfehler"
                $resultLabel.Text = "Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen."
                $resultLabel.ForeColor = [System.Drawing.Color]::Red
                $searchButton.Enabled = $true
                $installButton.Enabled = $true
                [System.Windows.Forms.MessageBox]::Show("Die Installation wurde mit Exit-Code $($process.ExitCode) abgeschlossen. Möglicherweise sind zusätzliche Maßnahmen erforderlich.", "Warnung", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        catch {
            # Fehlerfall
            $progressBar.MarqueeAnimationSpeed = 0
            $statusText.Text = "Fehler"
            $resultLabel.Text = "Fehler bei der Installation: $_"
            $resultLabel.ForeColor = [System.Drawing.Color]::Red
            $searchButton.Enabled = $true
            $installButton.Enabled = $true
            [System.Windows.Forms.MessageBox]::Show("Fehler bei der Installation: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        $resultLabel.Text = "Setup-Pfad ist nicht mehr gültig. Bitte erneut suchen."
        $resultLabel.ForeColor = [System.Drawing.Color]::Red
        $installButton.Enabled = $false
    }
})

# Tab automatisch prüfen, wenn er ausgewählt wird
$tabUCMA.Add_Enter({
    # Auto-Suche beim Betreten des Tabs, falls noch nicht durchgeführt
    if ([string]::IsNullOrEmpty($pathTextBox.Text)) {
        $searchButton.PerformClick()
    }
})

# Anzeigen des Formulars
$form.Add_Shown({
    # Prüfen, ob es sich um eine Ausführung nach dem Neustart handelt
    $isPostReboot = $false
    if (Test-Path "HKCU:\Software\ExchangeRequirements") {
        $isPostReboot = (Get-ItemProperty -Path "HKCU:\Software\ExchangeRequirements" -Name "PostRebootExecution" -ErrorAction SilentlyContinue).PostRebootExecution -eq $true
    }

    # Autostart-Eintrag entfernen, um zukünftige automatische Starts zu verhindern
    Remove-AutoStartAfterReboot
        
    # Definiere erforderliche Features für die Prüfung
    $requiredFeatures = @(
        "NET-Framework-45-Features", 
        "Web-Server", 
        "Web-Asp-Net45"
    )
    
    # Prüfe explizit den Status der Windows-Features und aktualisiere den Button
    $featuresInstalled = Test-WindowsFeatures -FeatureNames $requiredFeatures
    if ($featuresInstalled) {
        $buttonInstallFeatures.Enabled = $false
        $buttonInstallFeatures.Text = "Windows-Features bereits installiert"
        Write-Host "Features sind bereits installiert, Button wurde deaktiviert."
    } else {
        Write-Host "Features sind NICHT vollständig installiert, Button bleibt aktiv."
    }
    
    # Echtzeitschutz überprüfen
    if (Test-RealtimeProtectionDisabled) {
        $buttonDisableRealtime.Enabled = $false
        $buttonDisableRealtime.Text = "Echtzeitschutz bereits deaktiviert"
    }
    
    # VC++ Redistributables prüfen
    if (Test-VCRedistInstalled -DisplayNamePattern "*Visual C++ 2012*") {
        $buttonInstallVC.Enabled = $false
        $buttonInstallVC.Text = "Visual C++ 2012 Redistributable bereits installiert"
    }
    
    if (Test-VCRedistInstalled -DisplayNamePattern "*Visual C++ 2013*") {
        $buttonInstallVC2013.Enabled = $false
        $buttonInstallVC2013.Text = "Visual C++ 2013 Redistributable bereits installiert"
    }
    
    # URL Rewrite Modul prüfen
    if (Test-URLRewriteInstalled) {
        $buttonInstallURLRewrite.Enabled = $false
        $buttonInstallURLRewrite.Text = "IIS URL Rewrite Modul bereits installiert"
    }
    
    # Wenn es sich um eine Ausführung nach dem Neustart handelt, den Benutzer informieren
    if ($isPostReboot) {
        [System.Windows.Forms.MessageBox]::Show(
            "Das Skript wurde nach dem Neustart automatisch gestartet.", 
            "Nach Neustart", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }

    $statusInfoLabel.Text = "Status-Überprüfung abgeschlossen. Du kannst fortfahren."
    $form.Activate()
})

# Starte das Formular
[void] $form.ShowDialog()
