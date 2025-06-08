<#
.SYNOPSIS
  IIS Inetpub Folder ACL Manager
.DESCRIPTION
  The tool is intended to help you with your dailiy business.
  This tool is based on Microsoft's original PowerShell script for managing IIS inetpub folder permissions and has been enhanced with a user-friendly graphical interface.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-06-07

  Copyright (c) Jörn Walter. All rights reserved.
  Web: https://www.der-windows-papst.de
#>


# Admin Funktion
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Admin Funktion privilegierte Rechte
if (-not (Test-Admin)) {
    # If not, restart the script with administrative privileges
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo
    $newProcess.UseShellExecute = $true
    $newProcess.FileName = "PowerShell"
    $newProcess.Verb = "runas"
    $newProcess.Arguments = "-NoProfile -Windowstyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Globale Variablen
$script:logEntries = @()
$script:currentStep = 0
$script:totalSteps = 6
$script:logFilePath = ""

# Log-Datei initialisieren
function Initialize-LogFile {
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $logFileName = "InetpubACL_Log_$timestamp.txt"
        
        # Versuche verschiedene Pfade für die Log-Datei
        $possiblePaths = @(
            (Join-Path $env:TEMP $logFileName),
            (Join-Path $env:USERPROFILE "Desktop\$logFileName"),
            (Join-Path (Get-Location).Path $logFileName)
        )
        
        foreach ($path in $possiblePaths) {
            try {
                $testContent = "=== IIS Inetpub ACL Manager Log-Datei ===`r`nErstellt: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`nBenutzer: $env:USERNAME`r`nComputer: $env:COMPUTERNAME`r`n" + "="*50 + "`r`n"
                Set-Content -Path $path -Value $testContent -Encoding UTF8 -ErrorAction Stop
                $script:logFilePath = $path
                return $path
            } catch {
                continue
            }
        }
        
        # Fallback: Nur GUI-Logging
        Write-Warning "Konnte keine Log-Datei erstellen. Nur GUI-Logging aktiv."
        return $null
        
    } catch {
        Write-Warning "Fehler beim Initialisieren der Log-Datei: $($_.Exception.Message)"
        return $null
    }
}

# Log-Eintrag in Datei schreiben
function Write-LogToFile {
    param($LogEntry)
    
    if ($script:logFilePath -and (Test-Path $script:logFilePath)) {
        try {
            Add-Content -Path $script:logFilePath -Value $LogEntry -Encoding UTF8
        } catch {
            # Stille Behandlung - GUI-Logging läuft weiter
        }
    }
}

# Hauptfunktionen aus dem ursprünglichen Skript
function Test-HasUserData {
    param($inetpubPath, $dhaFolder)
    
    if (-not (Test-Path -Path $inetpubPath)) {
        return $False
    }

    $subFolders = Get-ChildItem -Path $inetpubPath | Where-Object { $_.PSIsContainer -and !($_.Attributes -match "ReparsePoint") }

    if ($null -ne $subfolders -and $null -eq $subfolders.Count -and $subfolders.Name -ne $dhaFolder) {
        return $True
    }

    if ($subFolders.Count -gt 1 -or ($subFolders.Count -eq 1 -and $subFolders[0].Name -ne $dhaFolder)) {
        return $True
    }

    return $False
}

function Set-FolderAcl {
    param (
        [Parameter(Mandatory=$true)]
        [String] $folderPath,
        [Parameter(Mandatory=$true)]
        [String] $sddl,
        [Parameter(Mandatory=$false)]
        [switch] $Quiet
    )

    if (Test-Path $folderPath) {
        $folder = Get-Item $folderPath -ErrorAction SilentlyContinue
        if (($folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            Add-LogEntry "WARNUNG" "${folderPath} ist ein Reparse Point und wird ignoriert."
            return $false
        }
    } else {
        Add-LogEntry "FEHLER" "${folderPath} nicht gefunden."
        return $false
    }

    # Aktuelle ACLs anzeigen
    if (-not $Quiet) {
        $currentAcl = Get-Acl $folderPath
        Add-LogEntry "INFO" "Aktuelle ACL für ${folderPath} wird gelesen..."
        Update-AclDisplay $folderPath $currentAcl "Aktuelle Berechtigung"
    }

    # ACL-Objekt mit SDDL-String erstellen
    try {
        $acl = New-Object -TypeName System.Security.AccessControl.DirectorySecurity
        $acl.SetSecurityDescriptorSddlForm($sddl)
        Add-LogEntry "SUCCESS" "ACL-Objekt erfolgreich erstellt für ${folderPath}"
    } catch {
        Add-LogEntry "FEHLER" "Fehler beim Erstellen des ACL-Objekts: $($_.Exception.Message)"
        return $false
    }

    try {
        Set-Acl -Path $folderPath -AclObject $acl -ErrorAction Stop
        Add-LogEntry "SUCCESS" "ACL erfolgreich gesetzt für ${folderPath}"
    } catch {
        if ($_.Exception -is [System.InvalidOperationException]) {
            Add-LogEntry "WARNUNG" "Set-Acl fehlgeschlagen. Versuche mit Built-in Administrator als Owner..."
            $sddlNew = $sddl -replace 'O:.+G:.+(D:.*)', 'O:BAG:BA$1'
            try {
                $acl = New-Object -TypeName System.Security.AccessControl.DirectorySecurity
                $acl.SetSecurityDescriptorSddlForm($sddlNew)
                Set-Acl -Path $folderPath -AclObject $acl -ErrorAction Stop
                Add-LogEntry "SUCCESS" "ACL erfolgreich gesetzt mit Built-in Administrator als Owner"
            } catch {
                Add-LogEntry "FEHLER" "Auch zweiter Versuch fehlgeschlagen: $($_.Exception.Message)"
                return $false
            }
        } else {
            Add-LogEntry "FEHLER" "Set-Acl fehlgeschlagen: $($_.Exception.Message)"
            return $false
        }
    }

    # Neue ACLs anzeigen
    if (-not $Quiet) {
        $newAcl = Get-Acl $folderPath
        Add-LogEntry "INFO" "Neue ACL für ${folderPath} wird angezeigt..."
        Update-AclDisplay $folderPath $newAcl "Neue Berechtigung"
    }

    return $true
}

function Add-LogEntry {
    param($Type, $Message)
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    $script:logEntries += $logEntry
    
    # In Datei schreiben
    Write-LogToFile $logEntry
    
    # GUI Update nur wenn Form verfügbar ist
    if ($script:form -and $script:form.IsHandleCreated -and -not $script:form.IsDisposed) {
        try {
            $script:form.Invoke([Action]{
                $script:logTextBox.AppendText("$logEntry`r`n")
                $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length
                $script:logTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            })
        } catch {
            # Fallback: Console-Ausgabe wenn GUI nicht verfügbar
            Write-Host $logEntry
        }
    } else {
        Write-Host $logEntry
    }
    
    Start-Sleep -Milliseconds 50
}

function Update-ProgressBar {
    param($Step, $Status)
    
    $script:currentStep = $Step
    $percentage = [math]::Round(($Step / $script:totalSteps) * 100)
    
    # GUI Update nur wenn Form verfügbar ist
    if ($script:form -and $script:form.IsHandleCreated -and -not $script:form.IsDisposed) {
        try {
            $script:form.Invoke([Action]{
                $script:progressBar.Value = $percentage
                $script:statusLabel.Text = $Status
                $script:stepLabel.Text = "Schritt $Step von $($script:totalSteps)"
                [System.Windows.Forms.Application]::DoEvents()
            })
        } catch {
            Write-Host "Schritt $Step $Status ($percentage%)"
        }
    } else {
        Write-Host "Schritt $Step $Status ($percentage%)"
    }
}

function Update-AclDisplay {
    param($Path, $Acl, $Title)
    
    $aclInfo = @"
=== $Title für: $Path ===
Owner: $($Acl.Owner)
Group: $($Acl.Group)

Zugriffsregeln:
"@
    
    foreach ($access in $Acl.Access) {
        $aclInfo += "`r`n- $($access.IdentityReference): $($access.FileSystemRights) ($($access.AccessControlType))"
        if ($access.IsInherited) {
            $aclInfo += " [Vererbt]"
        }
    }
    
    $aclInfo += "`r`n" + "="*60 + "`r`n"
    
    # In Log-Datei schreiben
    Write-LogToFile $aclInfo
    
    # GUI Update nur wenn Form verfügbar ist
    if ($script:form -and $script:form.IsHandleCreated -and -not $script:form.IsDisposed) {
        try {
            $script:form.Invoke([Action]{
                $script:aclTextBox.AppendText("$aclInfo`r`n")
                $script:aclTextBox.SelectionStart = $script:aclTextBox.Text.Length
                $script:aclTextBox.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            })
        } catch {
            Write-Host $aclInfo
        }
    } else {
        Write-Host $aclInfo
    }
}

function Test-AdminRights {
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    return $IsAdmin
}

function Convert-SddlToReadable {
    param([string]$sddl)
    
    try {
        $acl = New-Object System.Security.AccessControl.DirectorySecurity
        $acl.SetSecurityDescriptorSddlForm($sddl)
        return $acl
    } catch {
        return $null
    }
}

function Compare-AclPermissions {
    param(
        [System.Security.AccessControl.DirectorySecurity]$CurrentAcl,
        [System.Security.AccessControl.DirectorySecurity]$TargetAcl
    )
    
    $differences = @()
    $matches = @()
    
    # Vergleiche Owner
    if ($CurrentAcl.Owner -ne $TargetAcl.Owner) {
        $differences += "Owner: Aktuell='$($CurrentAcl.Owner)' vs Ziel='$($TargetAcl.Owner)'"
    } else {
        $matches += "Owner: $($CurrentAcl.Owner) ✓"
    }
    
    # Vergleiche Group
    if ($CurrentAcl.Group -ne $TargetAcl.Group) {
        $differences += "Group: Aktuell='$($CurrentAcl.Group)' vs Ziel='$($TargetAcl.Group)'"
    } else {
        $matches += "Group: $($CurrentAcl.Group) ✓"
    }
    
    # Sammle alle einzigartigen Identitäten
    $allIdentities = @()
    $allIdentities += $CurrentAcl.Access | ForEach-Object { $_.IdentityReference.Value }
    $allIdentities += $TargetAcl.Access | ForEach-Object { $_.IdentityReference.Value }
    $allIdentities = $allIdentities | Sort-Object -Unique
    
    foreach ($identity in $allIdentities) {
        $currentRules = $CurrentAcl.Access | Where-Object { $_.IdentityReference.Value -eq $identity }
        $targetRules = $TargetAcl.Access | Where-Object { $_.IdentityReference.Value -eq $identity }
        
        if ($currentRules -and $targetRules) {
            # Beide haben Regeln - vergleiche Details
            $currentRights = ($currentRules | ForEach-Object { $_.FileSystemRights.ToString() }) -join ", "
            $targetRights = ($targetRules | ForEach-Object { $_.FileSystemRights.ToString() }) -join ", "
            
            if ($currentRights -eq $targetRights) {
                $matches += "${identity}: $currentRights ✓"
            } else {
                $differences += "${identity}: Aktuell='$currentRights' vs Ziel='$targetRights'"
            }
        } elseif ($currentRules -and -not $targetRules) {
            $currentRights = ($currentRules | ForEach-Object { $_.FileSystemRights.ToString() }) -join ", "
            $differences += "${identity}: Wird entfernt (Aktuell: $currentRights)"
        } elseif (-not $currentRules -and $targetRules) {
            $targetRights = ($targetRules | ForEach-Object { $_.FileSystemRights.ToString() }) -join ", "
            $differences += "${identity}: Wird hinzugefügt (Neu: $targetRights)"
        }
    }
    
    return @{
        Differences = $differences
        Matches = $matches
        IsIdentical = ($differences.Count -eq 0)
    }
}

function Get-FolderAnalysis {
    param(
        [string]$FolderPath,
        [string]$TargetSddl,
        [string]$FolderName
    )
    
    $analysis = @{
        Path = $FolderPath
        Name = $FolderName
        Exists = $false
        IsAccessible = $false
        CurrentAcl = $null
        TargetAcl = $null
        Comparison = $null
        IsEmpty = $false
        SubFolders = @()
        HasReparsePoint = $false
        ErrorMessage = ""
    }
    
    try {
        if (Test-Path -Path $FolderPath) {
            $analysis.Exists = $true
            
            # Teste Zugriff
            try {
                $folder = Get-Item -Path $FolderPath -ErrorAction Stop
                $analysis.IsAccessible = $true
                
                # Prüfe Reparse Point
                if (($folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    $analysis.HasReparsePoint = $true
                    $analysis.ErrorMessage = "Ordner ist ein Reparse Point"
                }
                
                # Hole aktuelle ACL
                $analysis.CurrentAcl = Get-Acl -Path $FolderPath -ErrorAction Stop
                
                # Erstelle Ziel-ACL
                $analysis.TargetAcl = Convert-SddlToReadable -sddl $TargetSddl
                
                # Vergleiche Berechtigungen
                if ($analysis.TargetAcl) {
                    $analysis.Comparison = Compare-AclPermissions -CurrentAcl $analysis.CurrentAcl -TargetAcl $analysis.TargetAcl
                }
                
                # Prüfe Inhalt
                $subItems = Get-ChildItem -Path $FolderPath -ErrorAction SilentlyContinue
                $analysis.SubFolders = $subItems | Where-Object { $_.PSIsContainer } | ForEach-Object { $_.Name }
                $analysis.IsEmpty = ($subItems.Count -eq 0)
                
            } catch {
                $analysis.ErrorMessage = "Zugriff verweigert: $($_.Exception.Message)"
            }
        }
    } catch {
        $analysis.ErrorMessage = "Fehler bei der Analyse: $($_.Exception.Message)"
    }
    
    return $analysis
}

# Die korrigierte Show-InitialAnalysis Funktion
function Show-InitialAnalysis {
    Add-LogEntry "INFO" "=== SYSTEM-ANALYSE BEIM START ==="
    
    # Definiere Pfade und SDDL - NUR für inetpub!
    $systemDrive = $env:SystemDrive
    $inetpubPath = Join-Path -Path $systemDrive\ -ChildPath "inetpub"
    $dhaPath = Join-Path -Path $inetpubPath\ -ChildPath "DeviceHealthAttestation"
    $dhabinPath = Join-Path -Path $dhaPath\ -ChildPath 'bin'
    $sddlInetpub = "O:SYG:SYD:P(A;CIOI;GA;;;S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464)(A;CIOI;GA;;;SY)(A;CIOI;GA;;;BA)(A;CIOI;GRGX;;;BU)(A;CIOI;GA;;;CO)"
    
    # Analysiere NUR den Hauptordner inetpub
    $inetpubAnalysis = Get-FolderAnalysis -FolderPath $inetpubPath -TargetSddl $sddlInetpub -FolderName "inetpub"
    
    # Prüfe optionale Unterordner nur auf Existenz (ohne Berechtigungsanalyse)
    $dhaExists = Test-Path -Path $dhaPath
    $dhabinExists = Test-Path -Path $dhabinPath
    
    # Zeige Hauptanalyse für inetpub
    Add-LogEntry "INFO" "--- ANALYSE: $($inetpubAnalysis.Name) ---"
    Add-LogEntry "INFO" "Pfad: $($inetpubAnalysis.Path)"
    
    if ($inetpubAnalysis.Exists) {
        if ($inetpubAnalysis.IsAccessible) {
            Add-LogEntry "SUCCESS" "✓ Ordner existiert und ist zugänglich"
            
            if ($inetpubAnalysis.HasReparsePoint) {
                Add-LogEntry "WARNUNG" "⚠ Reparse Point erkannt - Skript wird nicht ausgeführt"
                return $false
            }
            
            if ($inetpubAnalysis.IsEmpty) {
                Add-LogEntry "INFO" "📁 Ordner ist leer"
            } else {
                Add-LogEntry "INFO" "📁 Ordner enthält: $($inetpubAnalysis.SubFolders -join ', ')"
            }
            
            # Zeige Berechtigungsvergleich NUR für inetpub
            if ($inetpubAnalysis.Comparison) {
                if ($inetpubAnalysis.Comparison.IsIdentical) {
                    Add-LogEntry "SUCCESS" "✓ Berechtigungen sind bereits korrekt gesetzt!"
                    Add-LogEntry "INFO" "Übereinstimmende Berechtigungen:"
                    foreach ($match in $inetpubAnalysis.Comparison.Matches) {
                        Add-LogEntry "SUCCESS" "  $match"
                    }
                } else {
                    Add-LogEntry "WARNUNG" "⚠ Berechtigungen für inetpub weichen ab - Änderungen erforderlich"
                    
                    if ($inetpubAnalysis.Comparison.Matches.Count -gt 0) {
                        Add-LogEntry "INFO" "Korrekte Berechtigungen:"
                        foreach ($match in $inetpubAnalysis.Comparison.Matches) {
                            Add-LogEntry "SUCCESS" "  $match"
                        }
                    }
                    
                    Add-LogEntry "WARNUNG" "Abweichende Berechtigungen:"
                    foreach ($diff in $inetpubAnalysis.Comparison.Differences) {
                        Add-LogEntry "WARNUNG" "  ❌ $diff"
                    }
                }
                
                # Zeige detaillierte ACL-Informationen
                Update-AclDisplay -Path $inetpubAnalysis.Path -Acl $inetpubAnalysis.CurrentAcl -Title "Aktuelle Berechtigung (Ist-Zustand)"
                Update-AclDisplay -Path $inetpubAnalysis.Path -Acl $inetpubAnalysis.TargetAcl -Title "Ziel-Berechtigung (Soll-Zustand)"
            }
        } else {
            Add-LogEntry "FEHLER" "❌ Ordner existiert, aber Zugriff verweigert"
            Add-LogEntry "FEHLER" "Fehler: $($inetpubAnalysis.ErrorMessage)"
        }
    } else {
        Add-LogEntry "INFO" "📁 inetpub-Ordner existiert nicht - wird bei Ausführung erstellt"
        # Zeige Ziel-Berechtigungen
        if ($inetpubAnalysis.TargetAcl) {
            Update-AclDisplay -Path $inetpubAnalysis.Path -Acl $inetpubAnalysis.TargetAcl -Title "Ziel-Berechtigung (wird gesetzt)"
        }
    }
    
    # Zeige Status der optionalen Unterordner (nur Info, keine Berechtigungsanalyse)
    Add-LogEntry "INFO" "--- OPTIONALE UNTERORDNER ---"
    if ($dhaExists) {
        Add-LogEntry "INFO" "📁 DeviceHealthAttestation-Ordner vorhanden (wird mitbearbeitet)"
    } else {
        Add-LogEntry "INFO" "📁 DeviceHealthAttestation-Ordner nicht vorhanden (normal, nur bei DHA Service benötigt)"
    }
    
    if ($dhabinExists) {
        Add-LogEntry "INFO" "📁 DeviceHealthAttestation/bin-Ordner vorhanden (wird mitbearbeitet)"
    } else {
        Add-LogEntry "INFO" "📁 DeviceHealthAttestation/bin-Ordner nicht vorhanden (normal, nur bei DHA Service benötigt)"
    }
    
    Add-LogEntry "INFO" "----------------------------------------"
    
    # Gesamtbewertung - KORRIGIERT: Bezieht sich nur auf inetpub
    Add-LogEntry "INFO" "=== GESAMTBEWERTUNG ==="
    
    if ($inetpubAnalysis.HasReparsePoint) {
        Add-LogEntry "FEHLER" "❌ KRITISCH: Reparse Point erkannt - Skript kann nicht ausgeführt werden"
        return $false
    }
    
    if ($inetpubAnalysis.Exists) {
        if ($inetpubAnalysis.IsAccessible) {
            if ($inetpubAnalysis.Comparison -and $inetpubAnalysis.Comparison.IsIdentical) {
                Add-LogEntry "SUCCESS" "✅ PERFEKT: inetpub-Ordner hat bereits die korrekten Berechtigungen"
                Add-LogEntry "INFO" "ℹ️ Skript-Ausführung ist optional (keine Änderungen erforderlich)"
            } else {
                Add-LogEntry "WARNUNG" "⚠ AKTION ERFORDERLICH: inetpub-Ordner hat abweichende Berechtigungen"
                Add-LogEntry "INFO" "ℹ️ Skript-Ausführung wird Berechtigungen korrigieren"
            }
        } else {
            Add-LogEntry "FEHLER" "❌ PROBLEM: Zugriff auf inetpub-Ordner verweigert"
        }
    } else {
        Add-LogEntry "INFO" "ℹ️ inetpub-Ordner existiert nicht - Skript wird ihn mit korrekten Berechtigungen erstellen"
    }
    
    Add-LogEntry "INFO" "=== ANALYSE ABGESCHLOSSEN ==="
    return $true
}

function Start-InetpubAclProcess {
    param([bool]$WhatIf = $false)
    
    # Variablen initialisieren
    $systemDrive = $env:SystemDrive
    $inetpubFldr = "inetpub"
    $inetpubPath = Join-Path -Path $systemDrive\ -ChildPath $inetpubFldr
    $dhaFolder = "DeviceHealthAttestation"
    $dhaPath = Join-Path -Path $inetpubPath\ -ChildPath $dhaFolder
    $dhabinPath = Join-Path -Path $dhaPath\ -ChildPath 'bin'
    $sddlInetpub = "O:SYG:SYD:P(A;CIOI;GA;;;S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464)(A;CIOI;GA;;;SY)(A;CIOI;GA;;;BA)(A;CIOI;GRGX;;;BU)(A;CIOI;GA;;;CO)"
    
    try {
        # Schritt 1: Administrator-Rechte prüfen
        Update-ProgressBar 1 "Prüfe Administrator-Rechte..."
        Add-LogEntry "INFO" "Starte Berechtigungsprüfung..."
        
        if (-not (Test-AdminRights)) {
            Add-LogEntry "FEHLER" "Das Skript benötigt Administrator-Rechte!"
            Update-ProgressBar 1 "FEHLER: Keine Administrator-Rechte"
            return
        }
        Add-LogEntry "SUCCESS" "Administrator-Rechte bestätigt"
        
        # Schritt 2: Inetpub-Ordner prüfen
        Update-ProgressBar 2 "Prüfe inetpub-Ordner Status..."
        Add-LogEntry "INFO" "Prüfe Existenz von: $inetpubPath"
        
        $oldInetpubAcl = $null
        $folderExists = Test-Path -Path $inetpubPath
        
        if ($folderExists) {
            Add-LogEntry "INFO" "inetpub-Ordner existiert bereits"
            
            $folder = Get-Item -Path $inetpubPath
            if (($folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                Add-LogEntry "FEHLER" "inetpub-Ordner ist ein Reparse Point - wird nicht unterstützt"
                Update-ProgressBar 2 "FEHLER: Reparse Point erkannt"
                return
            }
            
            if (Test-HasUserData $inetpubPath $dhaFolder) {
                Add-LogEntry "FEHLER" "inetpub-Ordner ist nicht leer - keine Aktion wird durchgeführt"
                Update-ProgressBar 2 "FEHLER: Ordner nicht leer"
                return
            }
            
            $oldInetpubAcl = Get-Acl $inetpubPath
            Add-LogEntry "SUCCESS" "inetpub-Ordner ist bereit für Berechtigungsänderung"
        } else {
            Add-LogEntry "INFO" "inetpub-Ordner existiert nicht - wird erstellt"
        }
        
        # Schritt 3: Ordner erstellen (falls notwendig)
        Update-ProgressBar 3 "Erstelle Ordner falls notwendig..."
        
        if (-not $folderExists) {
            if ($WhatIf) {
                Add-LogEntry "WHATIF" "Würde Ordner erstellen: $inetpubPath"
            } else {
                try {
                    New-Item -Path $systemDrive\ -Name $inetpubFldr -Type "Directory" -ErrorAction Stop | Out-Null
                    Add-LogEntry "SUCCESS" "inetpub-Ordner erfolgreich erstellt: $inetpubPath"
                } catch {
                    Add-LogEntry "FEHLER" "Fehler beim Erstellen des Ordners: $($_.Exception.Message)"
                    Update-ProgressBar 3 "FEHLER beim Erstellen"
                    return
                }
            }
        } else {
            Add-LogEntry "INFO" "Ordner bereits vorhanden - Erstellung übersprungen"
        }
        
        # Schritt 4: Hauptordner-Berechtigungen setzen
        Update-ProgressBar 4 "Setze Berechtigungen für inetpub-Ordner..."
        
        if ($WhatIf) {
            Add-LogEntry "WHATIF" "Würde Berechtigungen setzen für: $inetpubPath"
        } else {
            $success = Set-FolderAcl -folderPath $inetpubPath -sddl $sddlInetpub -Quiet
            if (-not $success) {
                Update-ProgressBar 4 "FEHLER beim Setzen der Berechtigungen"
                return
            }
        }
        
        # Schritt 5: DeviceHealthAttestation-Ordner prüfen und bearbeiten
        Update-ProgressBar 5 "Prüfe DeviceHealthAttestation-Ordner..."
        
        if (Test-Path -Path $dhaPath) {
            Add-LogEntry "INFO" "DeviceHealthAttestation-Ordner gefunden"
            if ($WhatIf) {
                Add-LogEntry "WHATIF" "Würde Berechtigungen setzen für: $dhaPath"
            } else {
                Set-FolderAcl -folderPath $dhaPath -sddl $sddlInetpub | Out-Null
            }
        } else {
            Add-LogEntry "INFO" "DeviceHealthAttestation-Ordner nicht vorhanden (optional - nur bei DHA Service)"
        }
        
        if (Test-Path -Path $dhabinPath) {
            Add-LogEntry "INFO" "DeviceHealthAttestation/bin-Ordner gefunden"
            if ($WhatIf) {
                Add-LogEntry "WHATIF" "Würde Berechtigungen setzen für: $dhabinPath"
            } else {
                Set-FolderAcl -folderPath $dhabinPath -sddl $sddlInetpub | Out-Null
            }
        } else {
            Add-LogEntry "INFO" "DeviceHealthAttestation/bin-Ordner nicht vorhanden (optional - nur bei DHA Service)"
        }
        
        # Schritt 6: Abschließende Berechtigungsprüfung
        Update-ProgressBar 6 "Führe abschließende Berechtigungsprüfung durch..."
        
        if (-not $WhatIf -and (Test-Path -Path $inetpubPath)) {
            Add-LogEntry "INFO" "=== BERECHTIGUNGSVERGLEICH ==="
            
            if ($null -ne $oldInetpubAcl) {
                Add-LogEntry "INFO" "Zeige ursprüngliche Berechtigungen:"
                Update-AclDisplay $inetpubPath $oldInetpubAcl "Ursprüngliche Berechtigung"
            }
            
            $finalAcl = Get-Acl $inetpubPath
            Add-LogEntry "INFO" "Zeige finale Berechtigungen:"
            Update-AclDisplay $inetpubPath $finalAcl "Finale Berechtigung"
            
            # Zusätzliche Berechtigungsvalidierung
            Add-LogEntry "INFO" "=== BERECHTIGUNGSVALIDIERUNG ==="
            $requiredAccounts = @("IIS_IUSRS", "SYSTEM", "Administratoren", "Benutzer")
            
            foreach ($account in $requiredAccounts) {
                $hasAccess = $finalAcl.Access | Where-Object { $_.IdentityReference -like "*$account*" }
                if ($hasAccess) {
                    Add-LogEntry "SUCCESS" "✓ $account hat Zugriff"
                } else {
                    Add-LogEntry "WARNUNG" "⚠ $account wurde nicht gefunden"
                }
            }
        }
        
        Update-ProgressBar 6 "Vorgang erfolgreich abgeschlossen!"
        Add-LogEntry "SUCCESS" "=== VORGANG ERFOLGREICH ABGESCHLOSSEN ==="
        
    } catch {
        Add-LogEntry "FEHLER" "Unerwarteter Fehler: $($_.Exception.Message)"
        Update-ProgressBar $script:currentStep "FEHLER aufgetreten"
    }
}

# GUI erstellen - als Script-Variable für globalen Zugriff
$script:form = New-Object System.Windows.Forms.Form
$form = $script:form
$form.Text = "IIS Inetpub Folder ACL Manager"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(1000, 600)

# Hauptpanel
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainPanel.ColumnCount = 2
$mainPanel.RowCount = 4
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 65)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 35)))
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))

# Header-Bereich
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$headerPanel.BackColor = [System.Drawing.Color]::LightBlue

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "IIS Inetpub Folder ACL Manager"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)

$descLabel = New-Object System.Windows.Forms.Label
$descLabel.Text = "Dieses Tool setzt die Sicherheitsberechtigungen für das inetpub-Verzeichnis und zeigt detaillierte Informationen über den Prozess an."
$descLabel.Location = New-Object System.Drawing.Point(10, 40)
$descLabel.Size = New-Object System.Drawing.Size(450, 35)
$descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Fortschrittsbereich
$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 10)
$progressBar.Size = New-Object System.Drawing.Size(1000, 20)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$progressBar.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Bereit zum Start..."
$statusLabel.Location = New-Object System.Drawing.Point(10, 35)
$statusLabel.Size = New-Object System.Drawing.Size(400, 20)
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$stepLabel = New-Object System.Windows.Forms.Label
$stepLabel.Text = "Schritt 0 von 6"
$stepLabel.Location = New-Object System.Drawing.Point(420, 35)
$stepLabel.Size = New-Object System.Drawing.Size(100, 20)

# Als Script-Variablen für globalen Zugriff
$script:progressBar = $progressBar
$script:statusLabel = $statusLabel
$script:stepLabel = $stepLabel

# Log-Bereich
$logGroupBox = New-Object System.Windows.Forms.GroupBox
$logGroupBox.Text = "Prozess-Log"
$logGroupBox.Dock = [System.Windows.Forms.DockStyle]::Fill

$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logTextBox.ReadOnly = $true

# Als Script-Variable für globalen Zugriff
$script:logTextBox = $logTextBox

# ACL-Anzeige-Bereich
$aclGroupBox = New-Object System.Windows.Forms.GroupBox
$aclGroupBox.Text = "Berechtigungen (ACL)"
$aclGroupBox.Dock = [System.Windows.Forms.DockStyle]::Fill

$aclTextBox = New-Object System.Windows.Forms.TextBox
$aclTextBox.Multiline = $true
$aclTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$aclTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$aclTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
$aclTextBox.ReadOnly = $true

# Als Script-Variable für globalen Zugriff
$script:aclTextBox = $aclTextBox

# Button-Bereich - Korrigierte Positionierung
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

# Info-Button für Programm-Informationen
$infoButton = New-Object System.Windows.Forms.Button
$infoButton.Text = "Info/Hilfe"
$infoButton.Location = New-Object System.Drawing.Point(595, 10)
$infoButton.Size = New-Object System.Drawing.Size(80, 30)
$infoButton.BackColor = [System.Drawing.Color]::LightSteelBlue

# Event-Handler für Info-Button
$infoButton.Add_Click({
    $infoText = @"
IIS Inetpub Folder ACL Manager
Version 1.1 - 2025

=== ENTWICKLER ===
© 2025 Jörn Walter
Website: https://www.der-windows-papst.de

=== ÜBER DAS PROGRAMM ===
Dieses Tool basiert auf dem ursprünglichen PowerShell-Skript von Microsoft zur Verwaltung der IIS inetpub-Ordner-Berechtigungen und wurde um eine benutzerfreundliche grafische Oberfläche erweitert.

=== FUNKTIONEN ===
• Automatische Analyse der aktuellen Berechtigungen
• Sicherheitsberechtigungen für inetpub-Ordner setzen
• Simulation (WhatIf) für sicheres Testen
• Detaillierte Protokollierung aller Aktionen
• Vergleich von Ist- und Soll-Berechtigungen
• Automatische Log-Datei-Erstellung

=== VERWENDUNG ===
1. Starte das Programm als Administrator
2. Überprüfe  die automatische Analyse
3. Verwende "Simulation" zum Testen ohne Änderungen
4. Klicke "Ausführen" für die tatsächliche Anwendung
5. Alle Aktionen werden protokolliert und können in der Log-Datei nachverfolgt werden

=== SICHERHEITSHINWEISE ===
• Immer als Administrator ausführen
• Backup der aktuellen Berechtigungen wird empfohlen
• Erst mit Simulation testen, dann ausführen
• Bei Problemen: Log-Datei prüfen

=== TECHNISCHE BASIS ===
Basiert auf Microsoft PowerShell-Skript für IIS inetpub ACL-Management
GUI-Erweiterung: Jörn Walter 2025
"@

    [System.Windows.Forms.MessageBox]::Show(
        $infoText,
        "Programm-Information & Hilfe",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Text = "Ausführen"
$executeButton.Location = New-Object System.Drawing.Point(10, 10)
$executeButton.Size = New-Object System.Drawing.Size(100, 30)
$executeButton.BackColor = [System.Drawing.Color]::LightGreen

$whatIfButton = New-Object System.Windows.Forms.Button
$whatIfButton.Text = "Simulation (WhatIf)"
$whatIfButton.Location = New-Object System.Drawing.Point(120, 10)
$whatIfButton.Size = New-Object System.Drawing.Size(120, 30)
$whatIfButton.BackColor = [System.Drawing.Color]::LightYellow

$openLogButton = New-Object System.Windows.Forms.Button
$openLogButton.Text = "Log-Datei öffnen"
$openLogButton.Location = New-Object System.Drawing.Point(250, 10)
$openLogButton.Size = New-Object System.Drawing.Size(105, 30)
$openLogButton.BackColor = [System.Drawing.Color]::LightGreen

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Log löschen"
$clearButton.Location = New-Object System.Drawing.Point(365, 10)
$clearButton.Size = New-Object System.Drawing.Size(100, 30)
$clearButton.BackColor = [System.Drawing.Color]::LightCoral

$analyzeButton = New-Object System.Windows.Forms.Button
$analyzeButton.Text = "Erneut Analysieren"
$analyzeButton.Location = New-Object System.Drawing.Point(475, 10)
$analyzeButton.Size = New-Object System.Drawing.Size(110, 30)
$analyzeButton.BackColor = [System.Drawing.Color]::LightBlue

# Beenden-Button ganz rechts positionieren
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Beenden"
$exitButton.Size = New-Object System.Drawing.Size(100, 30)
$exitButton.BackColor = [System.Drawing.Color]::LightCoral
# Anchor rechts, damit der Button immer am rechten Rand bleibt
$exitButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# Event-Handler für dynamische Positionierung des Beenden-Buttons
$buttonPanel.Add_Resize({
    $exitButton.Location = New-Object System.Drawing.Point(($buttonPanel.Width - $exitButton.Width - 10), 10)
})

# Initiale Position des Beenden-Buttons (wird beim ersten Resize überschrieben)
$exitButton.Location = New-Object System.Drawing.Point(1080, 10)

# Buttons UND Copyright in das buttonPanel
$buttonPanel.Controls.AddRange(@($executeButton, $whatIfButton, $openLogButton, $clearButton, $analyzeButton, $infoButton, $exitButton))

# Event-Handler - erst nach Button-Erstellung
$executeButton.Add_Click({
    $executeButton.Enabled = $false
    $whatIfButton.Enabled = $false
    
    # Direkte Ausführung im gleichen Thread
    try {
        Start-InetpubAclProcess -WhatIf $false
    } finally {
        $executeButton.Enabled = $true
        $whatIfButton.Enabled = $true
    }
})

$whatIfButton.Add_Click({
    $executeButton.Enabled = $false
    $whatIfButton.Enabled = $false
    
    # Direkte Ausführung im gleichen Thread
    try {
        Start-InetpubAclProcess -WhatIf $true
    } finally {
        $executeButton.Enabled = $true
        $whatIfButton.Enabled = $true
    }
})

# Event-Handler - erst nach Button-Erstellung
$executeButton.Add_Click({
    $executeButton.Enabled = $false
    $whatIfButton.Enabled = $false
    
    # Direkte Ausführung im gleichen Thread
    try {
        Start-InetpubAclProcess -WhatIf $false
    } finally {
        $executeButton.Enabled = $true
        $whatIfButton.Enabled = $true
    }
})

$whatIfButton.Add_Click({
    $executeButton.Enabled = $false
    $whatIfButton.Enabled = $false
    
    # Direkte Ausführung im gleichen Thread
    try {
        Start-InetpubAclProcess -WhatIf $true
    } finally {
        $executeButton.Enabled = $true
        $whatIfButton.Enabled = $true
    }
})

$openLogButton.Add_Click({
    if ($script:logFilePath -and (Test-Path $script:logFilePath)) {
        try {
            Start-Process notepad.exe -ArgumentList $script:logFilePath
            Add-LogEntry "INFO" "Log-Datei geöffnet: $script:logFilePath"
        } catch {
            try {
                # Fallback: Explorer öffnen
                Start-Process explorer.exe -ArgumentList "/select,`"$script:logFilePath`""
                Add-LogEntry "INFO" "Explorer geöffnet für Log-Datei: $script:logFilePath"
            } catch {
                Add-LogEntry "FEHLER" "Konnte Log-Datei nicht öffnen: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Log-Datei befindet sich unter:`n$script:logFilePath",
                    "Log-Datei Pfad",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    } else {
        Add-LogEntry "WARNUNG" "Keine Log-Datei verfügbar"
        [System.Windows.Forms.MessageBox]::Show(
            "Keine Log-Datei verfügbar. Möglicherweise konnte die Datei nicht erstellt werden.",
            "Log-Datei nicht verfügbar",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})

$clearButton.Add_Click({
    $script:logTextBox.Clear()
    $script:aclTextBox.Clear()
    $script:progressBar.Value = 0
    $script:statusLabel.Text = "Bereit zum Start..."
    $script:stepLabel.Text = "Schritt 0 von 6"
    $script:logEntries = @()
    $script:currentStep = 0
    
    # Neue Log-Datei erstellen
    $script:logFilePath = Initialize-LogFile
    if ($script:logFilePath) {
        Add-LogEntry "INFO" "Log gelöscht und neue Log-Datei erstellt: $script:logFilePath"
    } else {
        Add-LogEntry "INFO" "Log gelöscht (nur GUI)"
    }
})

$analyzeButton.Add_Click({
    try {
        Add-LogEntry "INFO" "Starte manuelle System-Analyse..."
        $analysisSuccess = Show-InitialAnalysis
        if ($analysisSuccess) {
            Add-LogEntry "SUCCESS" "Manuelle System-Analyse abgeschlossen."
        } else {
            Add-LogEntry "FEHLER" "System-Analyse ergab kritische Probleme."
        }
    } catch {
        Add-LogEntry "FEHLER" "Fehler bei der manuellen Analyse: $($_.Exception.Message)"
    }
})

$exitButton.Add_Click({
    # Abschließenden Log-Eintrag schreiben
    if ($script:logFilePath) {
        $closingMessage = "`r`n=== SESSION BEENDET ===`r`nBeendet: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`nGesamte Log-Einträge: $($script:logEntries.Count)`r`n" + "="*50
        Write-LogToFile $closingMessage
        Add-LogEntry "INFO" "Session beendet. Vollständiges Log verfügbar unter: $script:logFilePath"
    }
    
    $script:form.Close()
})

# Form anzeigen
$form.Controls.Add($mainPanel)

# Alle GUI-Elemente zum MainPanel hinzufügen
$headerPanel.Controls.AddRange(@($titleLabel, $descLabel))
$mainPanel.Controls.Add($headerPanel, 0, 0)
$mainPanel.SetColumnSpan($headerPanel, 2)

$progressPanel.Controls.AddRange(@($progressBar, $statusLabel, $stepLabel))
$mainPanel.Controls.Add($progressPanel, 0, 1)
$mainPanel.SetColumnSpan($progressPanel, 2)

$logGroupBox.Controls.Add($logTextBox)
$mainPanel.Controls.Add($logGroupBox, 0, 2)
$mainPanel.SetRowSpan($logGroupBox, 1)

$aclGroupBox.Controls.Add($aclTextBox)
$mainPanel.Controls.Add($aclGroupBox, 1, 2)
$mainPanel.SetRowSpan($aclGroupBox, 1)

$buttonPanel.Controls.AddRange(@($executeButton, $whatIfButton, $openLogButton, $clearButton, $analyzeButton, $exitButton))
$mainPanel.Controls.Add($buttonPanel, 0, 3)
$mainPanel.SetColumnSpan($buttonPanel, 2)

# Form vollständig initialisieren bevor Events verwendet werden
$form.Add_Shown({
    # Log-Datei initialisieren
    $script:logFilePath = Initialize-LogFile
    if ($script:logFilePath) {
        Add-LogEntry "SUCCESS" "Log-Datei erstellt: $script:logFilePath"
    } else {
        Add-LogEntry "WARNUNG" "Log-Datei konnte nicht erstellt werden - nur GUI-Logging aktiv"
    }
    
    # Prüfe Administrator-Rechte beim Start
    if (-not (Test-AdminRights)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Dieses Programm benötigt Administrator-Rechte. Bitte starte die PowerShell als Administrator und führe das Skript erneut aus.",
            "Administrator-Rechte erforderlich",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    } else {
        Add-LogEntry "INFO" "GUI erfolgreich gestartet - Administrator-Rechte bestätigt"
        Add-LogEntry "INFO" "Starte automatische System-Analyse..."
        
        # Führe initiale Analyse durch
        try {
            $analysisSuccess = Show-InitialAnalysis
            if ($analysisSuccess) {
                Add-LogEntry "SUCCESS" "System-Analyse abgeschlossen. Bereit für Skript-Ausführung."
                Add-LogEntry "INFO" "Alle Logs werden auch in die Datei geschrieben: $script:logFilePath"
            } else {
                Add-LogEntry "FEHLER" "System-Analyse ergab kritische Probleme. Bitte prüfe die Meldungen."
            }
        } catch {
            Add-LogEntry "FEHLER" "Fehler bei der System-Analyse: $($_.Exception.Message)"
        }
    }
})

$form.ShowDialog()
