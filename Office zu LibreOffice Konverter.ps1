<#
.SYNOPSIS
  Office zu LibreOffice Konverter
.DESCRIPTION
 Voraussetzung: LibreOffice muss installiert sein
.PARAMETER language
.NOTES
  Version:        1.1
  Author:         Jörn Walter
  Creation Date:  2025-08-20
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

# LibreOffice Pfad ermitteln
function Get-LibreOfficePath {
    $paths = @(
        "${env:ProgramFiles}\LibreOffice\program\soffice.exe",
        "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
        "${env:ProgramFiles}\LibreOffice 7\program\soffice.exe",
        "${env:ProgramFiles(x86)}\LibreOffice 7\program\soffice.exe"
    )
    
    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    # Versuche Registry zu durchsuchen
    try {
        $regPath = Get-ItemProperty "HKLM:\SOFTWARE\LibreOffice\UNO\InstallPath" -ErrorAction SilentlyContinue
        if ($regPath -and $regPath.InstallPath) {
            $soffice = Join-Path $regPath.InstallPath "program\soffice.exe"
            if (Test-Path $soffice) {
                return $soffice
            }
        }
    } catch {}
    
    return $null
}

function Update-ButtonStates {
    # "Liste leeren" Button nur aktivieren wenn Dateien vorhanden sind
    $btnClearFiles.IsEnabled = ($script:files.Count -gt 0)
    
    # "Konvertierung starten" Button nur aktivieren wenn ausgewählte Dateien vorhanden sind
    $selectedFiles = $script:files | Where-Object { $_.Selected }
    $btnConvert.IsEnabled = ($selectedFiles.Count -gt 0) -and (-not $script:cancelRequested)
}

# XAML GUI Definition
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office zu LibreOffice Konverter - © 2025 Jörn Walter" 
        Height="870" Width="900"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC"/>
                    <Setter Property="Foreground" Value="#666666"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="White" CornerRadius="5" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="Office zu LibreOffice Konverter" FontSize="20" FontWeight="Bold" Foreground="#333"/>
                <TextBlock Text="Konvertiert DOCX, XLSX und PPTX Dateien in LibreOffice Formate" FontSize="12" Foreground="#666" Margin="0,5,0,0"/>
                <TextBlock FontSize="10" Foreground="#888" Margin="0,3,0,0">
                    <Run Text="Entwickelt von Jörn Walter - Der Windows Papst"/>
                </TextBlock>
            </StackPanel>
        </Border>
        
        <!-- Konvertierungsmodus -->
        <Border Grid.Row="1" Background="White" CornerRadius="5" Padding="10" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="Konvertierungsmodus:" FontWeight="Bold" Margin="0,0,0,5"/>
                <WrapPanel>
                    <RadioButton Name="rbSingleFile" Content="Einzelne Datei(en)" IsChecked="True" Margin="5"/>
                    <RadioButton Name="rbSystemWide" Content="Systemweite Suche" Margin="20,5,5,5"/>
                </WrapPanel>
                
                <!-- Optionen für einzelne Dateien -->
                <StackPanel Name="spSingleFile" Margin="0,10,0,0">
                    <WrapPanel>
                        <Button Name="btnSelectFiles" Content="📁 Dateien auswählen" Width="150"/>
                        <Button Name="btnSelectFolder" Content="📂 Ordner auswählen" Width="150"/>
                    </WrapPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                        <TextBlock Text="Zielordner:" Width="80"/>
                        <TextBox Name="txtTargetFolder" Width="400" Margin="5,0"/>
                        <Button Name="btnBrowseTarget" Content="Durchsuchen..." Width="100"/>
                    </StackPanel>
                </StackPanel>
                
                <!-- Optionen für systemweite Suche -->
                <StackPanel Name="spSystemWide" Visibility="Collapsed" Margin="0,10,0,0">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Startordner:" Width="80"/>
                        <TextBox Name="txtSearchPath" Width="400" Margin="5,0" Text="C:\"/>
                        <Button Name="btnBrowseSearch" Content="Durchsuchen..." Width="100"/>
                    </StackPanel>
                    <CheckBox Name="chkRecursive" Content="Unterordner einbeziehen" IsChecked="True" Margin="85,5,0,0"/>
                    <TextBlock Text="Bei systemweiter Konvertierung werden die Dateien im gleichen Ordner wie die Originale gespeichert." 
                               Foreground="#666" FontStyle="Italic" Margin="85,5,0,0" TextWrapping="Wrap"/>
                </StackPanel>
                
                <!-- Allgemeine Aktionsbuttons - IMMER sichtbar -->
                <WrapPanel Margin="0,15,0,0" HorizontalAlignment="Right">
                    <Button Name="btnClearFiles" Content="❌ Liste leeren" Width="120" IsEnabled="False"/>
                    <Button Name="btnSearch" Content="🔍 Dateien suchen" Width="150" Visibility="Collapsed"/>
                </WrapPanel>
            </StackPanel>
        </Border>
        
        <!-- Dateiliste -->
        <Border Grid.Row="2" Background="White" CornerRadius="5" Padding="10" Height="250">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Zu konvertierende Dateien (Klick auf Zeile öffnet Ordner):" FontWeight="Bold" Margin="0,0,0,5"/>
                <DataGrid Grid.Row="1" Name="dgFiles" AutoGenerateColumns="False" CanUserAddRows="False" 
                          GridLinesVisibility="Horizontal" HeadersVisibility="Column" Height="200"
                          RowHeight="25" CanUserResizeRows="False" ScrollViewer.VerticalScrollBarVisibility="Auto">
                    <DataGrid.Columns>
                        <DataGridCheckBoxColumn Header="Auswahl" Binding="{Binding Selected}" Width="60"/>
                        <DataGridTextColumn Header="Dateiname" Binding="{Binding Name}" Width="2*"/>
                        <DataGridTextColumn Header="Typ" Binding="{Binding Extension}" Width="60"/>
                        <DataGridTextColumn Header="Größe" Binding="{Binding Size}" Width="80"/>
                        <DataGridTextColumn Header="Pfad" Binding="{Binding Path}" Width="2*"/>
                        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>
        
        <!-- Konvertierungsoptionen -->
        <Border Grid.Row="3" Background="White" CornerRadius="5" Padding="10" Margin="0,10,0,0">
            <StackPanel>
                <TextBlock Text="Konvertierungsoptionen:" FontWeight="Bold" Margin="0,0,0,5"/>
                <WrapPanel>
                    <CheckBox Name="chkConvertDocx" Content="DOCX → ODT" IsChecked="True"/>
                    <CheckBox Name="chkConvertXlsx" Content="XLSX → ODS" IsChecked="True" Margin="20,5,5,5"/>
                    <CheckBox Name="chkConvertPptx" Content="PPTX → ODP" IsChecked="True" Margin="20,5,5,5"/>
                    <CheckBox Name="chkOverwrite" Content="Vorhandene Dateien überschreiben" Margin="40,5,5,5"/>
                </WrapPanel>
            </StackPanel>
        </Border>
        
        <!-- Fortschritt -->
        <StackPanel Grid.Row="4" Margin="0,10,0,0">
            <ProgressBar Name="pbProgress" Height="25" Minimum="0" Maximum="100" IsIndeterminate="False"/>
            <TextBlock Name="txtStatus" Text="Bereit" HorizontalAlignment="Center" Margin="0,5,0,0" TextWrapping="Wrap"/>
        </StackPanel>
        
        <!-- Buttons -->
        <WrapPanel Grid.Row="5" HorizontalAlignment="Center" Margin="0,10,0,0">
            <Button Name="btnConvert" Content="▶ Konvertierung starten" Width="180" Height="35" FontSize="14" FontWeight="Bold"/>
            <Button Name="btnCancel" Content="⏹ Abbrechen" Width="120" Height="35" FontSize="14" IsEnabled="False"/>
        </WrapPanel>
        
        <!-- Copyright Footer -->
        <Border Grid.Row="6" Background="#E0E0E0" CornerRadius="5" Padding="10" Margin="0,10,0,0">
            <StackPanel>
                <TextBlock HorizontalAlignment="Center" FontSize="11">
                    <Run Text="© 2025 Jörn Walter - " Foreground="#555"/>
                    <Hyperlink Name="lnkWebsite" NavigateUri="https://www.der-windows-papst.de" Foreground="#0078D4">
                        <Run Text="https://www.der-windows-papst.de"/>
                    </Hyperlink>
                </TextBlock>
                <TextBlock Text="Office zu LibreOffice Konverter v1.1" HorizontalAlignment="Center" FontSize="10" Foreground="#777" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

# XAML laden
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Controls referenzieren
$rbSingleFile = $window.FindName("rbSingleFile")
$rbSystemWide = $window.FindName("rbSystemWide")
$spSingleFile = $window.FindName("spSingleFile")
$spSystemWide = $window.FindName("spSystemWide")
$btnSelectFiles = $window.FindName("btnSelectFiles")
$btnSelectFolder = $window.FindName("btnSelectFolder")
$btnClearFiles = $window.FindName("btnClearFiles")
$txtTargetFolder = $window.FindName("txtTargetFolder")
$btnBrowseTarget = $window.FindName("btnBrowseTarget")
$txtSearchPath = $window.FindName("txtSearchPath")
$btnBrowseSearch = $window.FindName("btnBrowseSearch")
$chkRecursive = $window.FindName("chkRecursive")
$dgFiles = $window.FindName("dgFiles")
$chkConvertDocx = $window.FindName("chkConvertDocx")
$chkConvertXlsx = $window.FindName("chkConvertXlsx")
$chkConvertPptx = $window.FindName("chkConvertPptx")
$chkOverwrite = $window.FindName("chkOverwrite")
$pbProgress = $window.FindName("pbProgress")
$txtStatus = $window.FindName("txtStatus")
$btnSearch = $window.FindName("btnSearch")
$btnConvert = $window.FindName("btnConvert")
$btnCancel = $window.FindName("btnCancel")
$lnkWebsite = $window.FindName("lnkWebsite")

# Globale Variablen
$script:files = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
$script:cancelRequested = $false
$script:searchRunspace = $null
$script:convertRunspace = $null
$script:timer = $null
$script:libreOfficePath = Get-LibreOfficePath
$script:sharedSearchData = $null
$script:sharedConvertData = $null

# DataGrid Binding
$dgFiles.ItemsSource = $script:files

# Initial UI-Setup
$btnSearch.Visibility = "Collapsed"

# LibreOffice Check
if (-not $script:libreOfficePath) {
    [System.Windows.MessageBox]::Show(
        "LibreOffice wurde nicht gefunden. Bitte installiere LibreOffice, um dieses Tool zu verwenden.",
        "LibreOffice nicht gefunden",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
    $window.Close()
    exit
}

# Timer für Progress Updates
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds(100)

# Hilfsfunktionen
function Get-FileExtensions {
    $extensions = @()
    if ($chkConvertDocx.IsChecked) { $extensions += "*.docx" }
    if ($chkConvertXlsx.IsChecked) { $extensions += "*.xlsx" }
    if ($chkConvertPptx.IsChecked) { $extensions += "*.pptx" }
    return $extensions
}

function Add-FileToList {
    param($FilePath)
    
    $file = Get-Item $FilePath
    $fileObj = [PSCustomObject]@{
        Selected = $true
        Name = $file.Name
        Extension = $file.Extension
        Size = "{0:N2} KB" -f ($file.Length / 1KB)
        Path = $file.DirectoryName
        FullPath = $file.FullName
        Status = "Bereit"
    }
    
    # UI-Update im Dispatcher
    $window.Dispatcher.Invoke([System.Action]{
        $script:files.Add($fileObj)
        Update-ButtonStates  # Button-Status aktualisieren
    })
}

function Generate-HTMLReport {
    param(
        [array]$ConvertedFiles,
        [int]$SuccessCount,
        [int]$FailCount,
        [string]$Duration,
        [datetime]$StartTime
    )
    
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    
    # Desktop-Pfad ermitteln und Bericht dort speichern
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $reportFileName = "LibreOffice_Konvertierung_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath = Join-Path $desktopPath $reportFileName
    
    # Sicherstellen, dass Desktop-Pfad existiert (sollte immer der Fall sein)
    if (-not (Test-Path $desktopPath)) {
        Write-Warning "Desktop-Pfad nicht gefunden, verwende Fallback-Pfad"
        $reportPath = Join-Path $env:USERPROFILE "Desktop\$reportFileName"
        
        # Wenn auch das nicht funktioniert, verwende Temp als letzten Ausweg
        if (-not (Test-Path (Split-Path $reportPath))) {
            $reportPath = Join-Path $env:TEMP $reportFileName
        }
    }
    
    # Statistiken berechnen
    $totalFiles = $ConvertedFiles.Count
    $successRate = if ($totalFiles -gt 0) { [math]::Round(($SuccessCount / $totalFiles) * 100, 1) } else { 0 }
    
    # Gruppierung nach Status - sicherstellen dass Arrays zurückgegeben werden
    $successFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "✓ Erfolgreich" })
    $failedFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "✗ Fehler" })
    $existingFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "⚠ Bereits vorhanden" })
    $cancelledFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "Abgebrochen" })
    
    # Gruppierung nach Dateityp
    $docxCount = 0
    $xlsxCount = 0
    $pptxCount = 0
    
    foreach ($file in $ConvertedFiles) {
        switch ($file.Extension.ToLower()) {
            ".docx" { $docxCount++ }
            ".xlsx" { $xlsxCount++ }
            ".pptx" { $pptxCount++ }
        }
    }
    
    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LibreOffice Konvertierungsbericht - $timestamp</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f0f4f8;
            color: #333;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        header {
            background: #0078D4;
            color: white;
            padding: 40px 0;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,120,212,0.3);
            border-radius: 10px;
            margin-bottom: 30px;
        }
        
        h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .subtitle {
            font-size: 1.1em;
            opacity: 0.95;
        }
        
        .save-location {
            background: #e3f2fd;
            color: #1565c0;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            border-left: 4px solid #0078D4;
        }
        
        .save-location strong {
            color: #0d47a1;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        
        .info-card {
            background: white;
            border-radius: 10px;
            padding: 25px;
            box-shadow: 0 2px 15px rgba(0,0,0,0.08);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        
        .info-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 5px 25px rgba(0,0,0,0.15);
        }
        
        .info-card h3 {
            color: #0078D4;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .info-card .value {
            font-size: 2.5em;
            font-weight: bold;
            color: #005a9e;
        }
        
        .info-card.success .value {
            color: #107c10;
        }
        
        .info-card.error .value {
            color: #d13438;
        }
        
        .info-card.warning .value {
            color: #ff8c00;
        }
        
        .progress-bar {
            width: 100%;
            height: 30px;
            background-color: #e0e0e0;
            border-radius: 15px;
            overflow: hidden;
            margin: 20px 0;
            box-shadow: inset 0 2px 5px rgba(0,0,0,0.1);
        }
        
        .progress-fill {
            height: 100%;
            background: #0078D4;
            border-radius: 15px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            transition: width 1s ease;
        }
        
        .table-container {
            background: white;
            border-radius: 10px;
            padding: 20px;
            margin: 30px 0;
            box-shadow: 0 2px 15px rgba(0,0,0,0.08);
            overflow-x: auto;
        }
        
        .table-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #0078D4;
        }
        
        .table-header h2 {
            color: #005a9e;
            font-size: 1.5em;
        }
        
        .filter-buttons {
            display: flex;
            gap: 10px;
        }
        
        .filter-btn {
            padding: 8px 16px;
            border: 2px solid #0078D4;
            background: white;
            color: #0078D4;
            border-radius: 5px;
            cursor: pointer;
            transition: all 0.3s ease;
            font-weight: 600;
        }
        
        .filter-btn:hover, .filter-btn.active {
            background: #0078D4;
            color: white;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
        }
        
        th {
            background: #f8f9fa;
            color: #005a9e;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #0078D4;
            position: sticky;
            top: 0;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        tr:hover {
            background-color: #f0f8ff;
        }
        
        .status-badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
        }
        
        .status-success {
            background: #dff6dd;
            color: #107c10;
        }
        
        .status-error {
            background: #fde7e9;
            color: #d13438;
        }
        
        .status-existing {
            background: #e3f2fd;
            color: #1565c0;
        }
        
        .status-cancelled {
            background: #fff4e6;
            color: #ff8c00;
        }
        
        .file-path {
            font-size: 0.9em;
            color: #666;
            max-width: 400px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        
        .footer {
            text-align: center;
            padding: 30px 0;
            margin-top: 50px;
            border-top: 2px solid #e0e0e0;
            color: #666;
        }
        
        .footer a {
            color: #0078D4;
            text-decoration: none;
            font-weight: 600;
        }
        
        .footer a:hover {
            text-decoration: underline;
        }
        
        .summary-section {
            background: white;
            border-radius: 10px;
            padding: 25px;
            margin: 30px 0;
            box-shadow: 0 2px 15px rgba(0,0,0,0.08);
        }
        
        .chart-container {
            display: flex;
            justify-content: space-around;
            align-items: center;
            margin: 20px 0;
        }
        
        .pie-chart {
            width: 200px;
            height: 200px;
            position: relative;
        }
        
        @media print {
            body {
                background: white;
            }
            .filter-buttons {
                display: none;
            }
            header {
                background: #0078D4;
                print-color-adjust: exact;
                -webkit-print-color-adjust: exact;
            }
        }
        
        @media (max-width: 768px) {
            .info-grid {
                grid-template-columns: 1fr;
            }
            
            table {
                font-size: 0.9em;
            }
            
            .file-path {
                max-width: 200px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 LibreOffice Konvertierungsbericht</h1>
            <div class="subtitle">Erstellt am $timestamp</div>
        </header>
        
        <!-- Info Box über Speicherort -->
        <div class="save-location">
            <strong>💾 Speicherort:</strong> Dieser Bericht wurde auf dem Desktop gespeichert: <code>$reportPath</code>
        </div>
        
        <div class="info-grid">
            <div class="info-card">
                <h3>Gesamtanzahl Dateien</h3>
                <div class="value">$totalFiles</div>
            </div>
            
            <div class="info-card success">
                <h3>Erfolgreich konvertiert</h3>
                <div class="value">$SuccessCount</div>
            </div>
            
            <div class="info-card error">
                <h3>Fehlgeschlagen</h3>
                <div class="value">$FailCount</div>
            </div>
            
            <div class="info-card">
                <h3>Konvertierungsdauer</h3>
                <div class="value">$Duration</div>
            </div>
        </div>
        
        <div class="summary-section">
            <h2 style="color: #005a9e; margin-bottom: 20px;">Erfolgsquote</h2>
            <div class="progress-bar">
                <div class="progress-fill" style="width: $($successRate)%;">
                    $($successRate)%
                </div>
            </div>
            
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-top: 30px;">
                <div style="text-align: center;">
                    <div style="font-size: 2em; color: #0078D4; font-weight: bold;">$docxCount</div>
                    <div style="color: #666; margin-top: 5px;">DOCX → ODT</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 2em; color: #0078D4; font-weight: bold;">$xlsxCount</div>
                    <div style="color: #666; margin-top: 5px;">XLSX → ODS</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 2em; color: #0078D4; font-weight: bold;">$pptxCount</div>
                    <div style="color: #666; margin-top: 5px;">PPTX → ODP</div>
                </div>
            </div>
        </div>
        
        <div class="table-container">
            <div class="table-header">
                <h2>📁 Detaillierte Dateiliste</h2>
                <div class="filter-buttons">
                    <button class="filter-btn active" onclick="filterTable('all')">Alle</button>
                    <button class="filter-btn" onclick="filterTable('success')">Erfolgreich</button>
                    <button class="filter-btn" onclick="filterTable('existing')">Bereits vorhanden</button>
                    <button class="filter-btn" onclick="filterTable('error')">Fehler</button>
                </div>
            </div>
            
            <table id="fileTable">
                <thead>
                    <tr>
                        <th>Dateiname</th>
                        <th>Typ</th>
                        <th>Größe</th>
                        <th>Pfad</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
"@

    foreach ($file in $ConvertedFiles) {
        $statusClass = switch ($file.Status) {
            "✓ Erfolgreich" { "success" }
            "✗ Fehler" { "error" }
            "⚠ Bereits vorhanden" { "existing" }
            "Abgebrochen" { "cancelled" }
            default { "cancelled" }
        }
        
        $statusText = $file.Status -replace '[✓✗⚠]', ''
        
        $html += @"
                    <tr data-status="$statusClass">
                        <td><strong>$($file.Name)</strong></td>
                        <td>$($file.Extension)</td>
                        <td>$($file.Size)</td>
                        <td><span class="file-path" title="$($file.Path)">$($file.Path)</span></td>
                        <td><span class="status-badge status-$statusClass">$statusText</span></td>
                    </tr>
"@
    }

    $html += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>Office zu LibreOffice Konverter</strong></p>
            <p>© 2025 Jörn Walter - <a href="https://www.der-windows-papst.de" target="_blank">www.der-windows-papst.de</a></p>
            <p style="margin-top: 10px; font-size: 0.9em;">Bericht generiert am $timestamp</p>
            <p style="margin-top: 5px; font-size: 0.8em; color: #888;">Gespeichert auf dem Desktop: $reportFileName</p>
        </div>
    </div>
    
    <script>
        function filterTable(status) {
            const rows = document.querySelectorAll('#fileTable tbody tr');
            const buttons = document.querySelectorAll('.filter-btn');
            
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            
            rows.forEach(row => {
                if (status === 'all') {
                    row.style.display = '';
                } else {
                    row.style.display = row.dataset.status === status ? '' : 'none';
                }
            });
        }
    </script>
</body>
</html>
"@

    # HTML-Datei auf dem Desktop speichern
    try {
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Bericht erfolgreich gespeichert: $reportPath"
    }
    catch {
        Write-Warning "Fehler beim Speichern des Berichts: $_"
        # Fallback: Temp-Ordner verwenden
        $reportPath = Join-Path $env:TEMP $reportFileName
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Bericht im Temp-Ordner gespeichert: $reportPath"
    }
    
    return $reportPath
}

function Convert-File {
    param(
        [string]$InputFile,
        [string]$OutputDir
    )
    
    if (-not (Test-Path $InputFile)) {
        return @{ Success = $false; Status = "✗ Fehler"; Message = "Eingabedatei nicht gefunden" }
    }
    
    if (-not (Test-Path $OutputDir)) {
        try {
            New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        } catch {
            return @{ Success = $false; Status = "✗ Fehler"; Message = "Ausgabeordner konnte nicht erstellt werden" }
        }
    }
    
    $inputItem = Get-Item $InputFile
    $extension = $inputItem.Extension.ToLower()
    
    # Bestimme Output-Format
    $outputFormat = switch ($extension) {
        ".docx" { "odt" }
        ".xlsx" { "ods" }
        ".pptx" { "odp" }
        default { 
            return @{ Success = $false; Status = "✗ Fehler"; Message = "Nicht unterstütztes Format: $extension" }
        }
    }
    
    # Output-Dateiname
    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name) + ".$outputFormat"
    $outputPath = Join-Path $OutputDir $outputName
    
    # Prüfe ob Datei existiert
    if ((Test-Path $outputPath) -and -not $chkOverwrite.IsChecked) {
        return @{ Success = $false; Status = "⚠ Bereits vorhanden"; Message = "Datei existiert bereits" }
    }
    
    # LibreOffice Konvertierung
    $arguments = @(
        "--headless",
        "--convert-to", $outputFormat,
        "--outdir", "`"$OutputDir`"",
        "`"$InputFile`""
    )
    
    try {
        $process = Start-Process -FilePath $script:libreOfficePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich" }
        } else {
            return @{ Success = $false; Status = "✗ Fehler"; Message = "LibreOffice Konvertierung fehlgeschlagen (Exit Code: $($process.ExitCode))" }
        }
    } catch {
        return @{ Success = $false; Status = "✗ Fehler"; Message = "Fehler beim Starten von LibreOffice: $($_.Exception.Message)" }
    }
}

function Stop-BackgroundJob {
    if ($script:timer) {
        $script:timer.Stop()
        # WICHTIG: Alle Event-Handler vom Timer entfernen
        $script:timer.Remove_Tick
    }
    
    if ($script:searchRunspace) {
        try {
            $script:searchRunspace.Close()
            $script:searchRunspace.Dispose()
            $script:searchRunspace = $null
        } catch {}
    }
    
    if ($script:convertRunspace) {
        try {
            $script:convertRunspace.Close()
            $script:convertRunspace.Dispose()
            $script:convertRunspace = $null
        } catch {}
    }
    
    # Shared Data zurücksetzen
    $script:sharedSearchData = $null
    $script:sharedConvertData = $null
}

function Complete-Reset {
    # Timer komplett stoppen und alle Events entfernen
    if ($script:timer) {
        $script:timer.Stop()
        $script:timer.Remove_Tick
    }
    
    # Runspaces schließen
    if ($script:searchRunspace) {
        try {
            $script:searchRunspace.Close()
            $script:searchRunspace.Dispose()
            $script:searchRunspace = $null
        } catch {}
    }
    
    if ($script:convertRunspace) {
        try {
            $script:convertRunspace.Close()
            $script:convertRunspace.Dispose()
            $script:convertRunspace = $null
        } catch {}
    }
    
    # Shared Data komplett zurücksetzen
    $script:sharedSearchData = $null
    $script:sharedConvertData = $null
    
    # UI zurücksetzen
    $pbProgress.IsIndeterminate = $false
    $pbProgress.Value = 0
    $btnSearch.IsEnabled = $true
    $btnConvert.IsEnabled = $true
    $btnCancel.IsEnabled = $false
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $script:cancelRequested = $false
    $txtStatus.Text = "Bereit"
    Update-ButtonStates
}

function Complete-Reset-WithTable {
    # Erst Complete-Reset ausführen
    Complete-Reset
    
    # Methode 1: Clear der ObservableCollection
    $script:files.Clear()
    
    # Methode 2: Neue leere ObservableCollection erstellen falls Clear nicht funktioniert
    if ($script:files.Count -gt 0) {
        $script:files = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $dgFiles.ItemsSource = $script:files
    }
    
    # Methode 3: Erzwinge UI-Update im Dispatcher
    $window.Dispatcher.Invoke([System.Action]{
        $dgFiles.Items.Refresh()
        $dgFiles.UpdateLayout()
    }, [System.Windows.Threading.DispatcherPriority]::Render)
    
    # Methode 4: DataGrid komplett neu binden
    $dgFiles.ItemsSource = $null
    $dgFiles.ItemsSource = $script:files
    
    Update-ButtonStates
}

function Reset-UI {
    $pbProgress.IsIndeterminate = $false
    $pbProgress.Value = 0
    $btnSearch.IsEnabled = $true
    $btnConvert.IsEnabled = $true
    $btnCancel.IsEnabled = $false
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $script:cancelRequested = $false
    $txtStatus.Text = "Bereit"
    Update-ButtonStates
}

# Event Handler
$rbSingleFile.Add_Checked({
    $spSingleFile.Visibility = "Visible"
    $spSystemWide.Visibility = "Collapsed"
    $btnSearch.Visibility = "Collapsed"
})

$rbSystemWide.Add_Checked({
    $spSingleFile.Visibility = "Collapsed"
    $spSystemWide.Visibility = "Visible"
    $btnSearch.Visibility = "Visible"
})

$btnSelectFiles.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Multiselect = $true
    $dialog.Filter = "Office Dateien|*.docx;*.xlsx;*.pptx|Alle Dateien|*.*"
    
    if ($dialog.ShowDialog()) {
        foreach ($file in $dialog.FileNames) {
            Add-FileToList -FilePath $file
        }
        Update-ButtonStates
    }
})

$btnSelectFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Wähle einen Ordner mit Office-Dateien aus"
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $extensions = Get-FileExtensions
        foreach ($ext in $extensions) {
            $files = Get-ChildItem -Path $dialog.SelectedPath -Filter $ext -File
            foreach ($file in $files) {
                Add-FileToList -FilePath $file.FullName
            }
        }
        Update-ButtonStates
    }
})

$btnClearFiles.Add_Click({
    # Vollständige Bereinigung mit Tabelle
    Complete-Reset-WithTable
    
    # Zusätzliche Bereinigungsversuche falls immer noch nicht leer
    if ($script:files.Count -gt 0 -or $dgFiles.Items.Count -gt 0) {
        # Alternative 1: Alle Items einzeln entfernen
        while ($script:files.Count -gt 0) {
            $script:files.RemoveAt(0)
        }
        
        # Alternative 2: DataGrid direkt leeren
        $dgFiles.Items.Clear()
        
        # Alternative 3: Neue Collection mit Force-Update
        $script:files = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $dgFiles.ItemsSource = $script:files
        
        # Force UI Update
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Button-Status aktualisieren
    Update-ButtonStates
    $txtStatus.Text = "Liste geleert"
})

$btnBrowseTarget.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Wähle den Zielordner für konvertierte Dateien"
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTargetFolder.Text = $dialog.SelectedPath
    }
})

$btnBrowseSearch.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Wähle den Startordner für die Suche aus"
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSearchPath.Text = $dialog.SelectedPath
    }
})

# Verbesserter Such-Event Handler mit Runspace
$btnSearch.Add_Click({
    if (-not $rbSystemWide.IsChecked) { return }
    
    $searchPath = $txtSearchPath.Text
    if ([string]::IsNullOrWhiteSpace($searchPath) -or -not (Test-Path $searchPath)) {
        [System.Windows.MessageBox]::Show("Bitte gib einen gültigen Suchpfad an.", "Ungültiger Pfad", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $extensions = Get-FileExtensions
    if ($extensions.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wähle mindestens einen Dateityp aus.", "Kein Dateityp", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Complete-Reset-WithTable
    
    Start-Sleep -Milliseconds 200
    
    # UI für Suche vorbereiten
    $script:files.Clear()
    $script:cancelRequested = $false
    $btnSearch.IsEnabled = $false
    $btnConvert.IsEnabled = $false
    $btnCancel.IsEnabled = $true
    $pbProgress.IsIndeterminate = $true
    $txtStatus.Text = "🔍 Suche läuft..."
    
    # Shared Variables für Runspace
    $script:sharedSearchData = [hashtable]::Synchronized(@{
        SearchPath = $searchPath
        Extensions = $extensions
        Recursive = $chkRecursive.IsChecked
        CancelRequested = $false
        FoundFiles = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Status = "Suche läuft..."
        Completed = $false
        Error = $null
    })
    
    # Runspace erstellen
    $script:searchRunspace = [runspacefactory]::CreateRunspace()
    $script:searchRunspace.Open()
    $script:searchRunspace.SessionStateProxy.SetVariable('sharedSearchData', $script:sharedSearchData)
    
    # PowerShell-Instanz für Suche
    $searchScript = {
        param($sharedData)
        
        try {
            $startTime = Get-Date
            $totalFound = 0
            
            foreach ($ext in $sharedData.Extensions) {
                if ($sharedData.CancelRequested) { break }
                
                $sharedData.Status = "Suche nach $ext Dateien..."
                
                if ($sharedData.Recursive) {
                    $foundFiles = Get-ChildItem -Path $sharedData.SearchPath -Filter $ext -File -Recurse -ErrorAction SilentlyContinue
                } else {
                    $foundFiles = Get-ChildItem -Path $sharedData.SearchPath -Filter $ext -File -ErrorAction SilentlyContinue
                }
                
                if ($foundFiles) {
                    foreach ($file in $foundFiles) {
                        if ($sharedData.CancelRequested) { break }
                        
                        $fileInfo = @{
                            Name = $file.Name
                            Extension = $file.Extension
                            Size = "{0:N2} KB" -f ($file.Length / 1KB)
                            Path = $file.DirectoryName
                            FullPath = $file.FullName
                        }
                        
                        $sharedData.FoundFiles.Add($fileInfo) | Out-Null
                        $totalFound++
                        
                        if ($totalFound % 10 -eq 0) {
                            $elapsed = (Get-Date) - $startTime
                            $sharedData.Status = "Gefunden: $totalFound Datei(en) | Zeit: $($elapsed.ToString('mm\:ss'))"
                            Start-Sleep -Milliseconds 10
                        }
                    }
                }
            }
            
            $elapsed = (Get-Date) - $startTime
            if ($sharedData.CancelRequested) {
                $sharedData.Status = "Suche abgebrochen - Gefunden: $($sharedData.FoundFiles.Count) Datei(en)"
            } else {
                $sharedData.Status = "✓ Suche abgeschlossen: $($sharedData.FoundFiles.Count) Datei(en) (Zeit: $($elapsed.ToString('mm\:ss')))"
            }
            
        } catch {
            $sharedData.Error = $_.Exception.Message
            $sharedData.Status = "❌ Fehler bei der Suche: $($_.Exception.Message)"
        } finally {
            $sharedData.Completed = $true
        }
    }
    
    $powerShell = [powershell]::Create()
    $powerShell.Runspace = $script:searchRunspace
    $powerShell.AddScript($searchScript).AddArgument($script:sharedSearchData) | Out-Null
    $powerShell.BeginInvoke() | Out-Null
    
    # Timer für Search Progress Updates
    if ($script:timer.IsEnabled) {
        $script:timer.Stop()
    }
    $script:timer.Remove_Tick
    
    # Event Handler für Suche definieren
    $searchTimerTick = {
        try {
            if ($script:sharedSearchData -eq $null) { return }
            
            $txtStatus.Text = $script:sharedSearchData.Status
            
            if ($script:sharedSearchData.Completed) {
                $script:timer.Stop()
                
                # Stelle sicher, dass Tabelle wirklich leer ist vor dem Import
                if ($script:files.Count -gt 0) {
                    $script:files.Clear()
                    $dgFiles.Items.Refresh()
                }
                
                # Dateien zur Liste hinzufügen
                $addedCount = 0
                foreach ($fileInfo in $script:sharedSearchData.FoundFiles) {
                    $fileObj = [PSCustomObject]@{
                        Selected = $true
                        Name = $fileInfo.Name
                        Extension = $fileInfo.Extension
                        Size = $fileInfo.Size
                        Path = $fileInfo.Path
                        FullPath = $fileInfo.FullPath
                        Status = "Bereit"
                    }
                    $script:files.Add($fileObj)
                    $addedCount++
                }
                
                # Force DataGrid Update
                $dgFiles.Items.Refresh()
                $window.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
                
                Reset-UI
                
                if ($script:sharedSearchData.Error) {
                    [System.Windows.MessageBox]::Show("Fehler bei der Suche: $($script:sharedSearchData.Error)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                } elseif (-not $script:sharedSearchData.CancelRequested) {
                    if ($script:files.Count -eq 0) {
                        [System.Windows.MessageBox]::Show("Keine Dateien gefunden.", "Suche abgeschlossen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    } else {
                        $docxCount = ($script:files | Where-Object { $_.Extension -eq ".docx" }).Count
                        $xlsxCount = ($script:files | Where-Object { $_.Extension -eq ".xlsx" }).Count
                        $pptxCount = ($script:files | Where-Object { $_.Extension -eq ".pptx" }).Count
                        
                        $summary = "Suche erfolgreich abgeschlossen!`n`n"
                        if ($docxCount -gt 0) { $summary += "📄 DOCX: $docxCount`n" }
                        if ($xlsxCount -gt 0) { $summary += "📊 XLSX: $xlsxCount`n" }
                        if ($pptxCount -gt 0) { $summary += "📽️ PPTX: $pptxCount`n" }
                        $summary += "`n📁 Gesamt: $($script:files.Count) Datei(en)"
                        
                        [System.Windows.MessageBox]::Show($summary, "Suche abgeschlossen", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    }
                }
                
                Stop-BackgroundJob
            }
        } catch {
            $script:timer.Stop()
            Complete-Reset-WithTable  # Bei Fehlern Tabelle leeren
        }
    }
    
    $script:timer.Add_Tick($searchTimerTick)
    $script:timer.Start()
})

# Konvertierungs-Event Handler
$btnConvert.Add_Click({
    if ($script:files.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Keine Dateien in der Liste.", "Keine Dateien", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $selectedFiles = $script:files | Where-Object { $_.Selected }
    if ($selectedFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wähle mindestens eine Datei aus.", "Keine Datei ausgewählt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $targetFolder = ""
    if ($rbSingleFile.IsChecked) {
        $targetFolder = $txtTargetFolder.Text
        if ([string]::IsNullOrWhiteSpace($targetFolder)) {
            [System.Windows.MessageBox]::Show("Bitte gib einen Zielordner an.", "Zielordner fehlt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
    }
    
    # UI für Konvertierung vorbereiten
    Complete-Reset
    
    $script:cancelRequested = $false
    $btnConvert.IsEnabled = $false
    $btnCancel.IsEnabled = $true
    $btnSearch.IsEnabled = $false
    $pbProgress.Value = 0
    
    # Shared Variables für Konvertierung
    $script:sharedConvertData = [hashtable]::Synchronized(@{
        Files = @($selectedFiles)
        TargetFolder = $targetFolder
        IsSystemWide = $rbSystemWide.IsChecked
        OverwriteFiles = $chkOverwrite.IsChecked
        LibreOfficePath = $script:libreOfficePath
        CancelRequested = $false
        CurrentFile = 0
        TotalFiles = $selectedFiles.Count
        SuccessCount = 0
        FailCount = 0
        Status = "Konvertierung läuft..."
        Completed = $false
        Error = $null
        Results = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Duration = ""
        StartTime = (Get-Date)
    })
    
    # Runspace für Konvertierung
    $script:convertRunspace = [runspacefactory]::CreateRunspace()
    $script:convertRunspace.Open()
    $script:convertRunspace.SessionStateProxy.SetVariable('sharedConvertData', $script:sharedConvertData)
    
    $convertScript = {
        param($sharedConvertData)
        
        function Convert-File {
            param(
                [string]$InputFile,
                [string]$OutputDir,
                [bool]$OverwriteFiles,
                [string]$LibreOfficePath
            )
            
            if (-not (Test-Path $InputFile)) { 
                return @{ Success = $false; Status = "✗ Fehler"; Message = "Eingabedatei nicht gefunden" }
            }
            
            if (-not (Test-Path $OutputDir)) {
                try {
                    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
                } catch {
                    return @{ Success = $false; Status = "✗ Fehler"; Message = "Ausgabeordner konnte nicht erstellt werden" }
                }
            }
            
            $inputItem = Get-Item $InputFile
            $extension = $inputItem.Extension.ToLower()
            
            $outputFormat = switch ($extension) {
                ".docx" { "odt" }
                ".xlsx" { "ods" }
                ".pptx" { "odp" }
                default { return @{ Success = $false; Status = "✗ Fehler"; Message = "Nicht unterstütztes Format: $extension" } }
            }
            
            $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name) + ".$outputFormat"
            $outputPath = Join-Path $OutputDir $outputName
            
            if ((Test-Path $outputPath) -and -not $OverwriteFiles) {
                return @{ Success = $false; Status = "⚠ Bereits vorhanden"; Message = "Datei existiert bereits" }
            }
            
            $arguments = @(
                "--headless",
                "--convert-to", $outputFormat,
                "--outdir", "`"$OutputDir`"",
                "`"$InputFile`""
            )
            
            try {
                $process = Start-Process -FilePath $LibreOfficePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                if ($process.ExitCode -eq 0) {
                    return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich" }
                } else {
                    return @{ Success = $false; Status = "✗ Fehler"; Message = "LibreOffice Konvertierung fehlgeschlagen" }
                }
            } catch {
                return @{ Success = $false; Status = "✗ Fehler"; Message = "Fehler beim Starten von LibreOffice" }
            }
        }
        
        try {
            $startTime = $sharedConvertData.StartTime
            
            foreach ($file in $sharedConvertData.Files) {
                if ($sharedConvertData.CancelRequested) {
                    $file.Status = "Abgebrochen"
                    $sharedConvertData.Results.Add($file) | Out-Null
                    break
                }
                
                $sharedConvertData.CurrentFile++
                $sharedConvertData.Status = "Konvertiere $($sharedConvertData.CurrentFile) von $($sharedConvertData.TotalFiles): $($file.Name)"
                
                $outputDir = if ($sharedConvertData.IsSystemWide) { $file.Path } else { $sharedConvertData.TargetFolder }
                
                $result = Convert-File -InputFile $file.FullPath -OutputDir $outputDir -OverwriteFiles $sharedConvertData.OverwriteFiles -LibreOfficePath $sharedConvertData.LibreOfficePath
                
                # Setze Status basierend auf dem detaillierten Ergebnis
                $file.Status = $result.Status
                
                # Zähle nur echte Erfolge und Fehler für die Statistik
                if ($result.Success) {
                    $sharedConvertData.SuccessCount++
                } elseif ($result.Status -eq "✗ Fehler") {
                    $sharedConvertData.FailCount++
                }
                # "Bereits vorhanden" wird weder als Erfolg noch als Fehler gezählt
                
                # Füge die verarbeitete Datei zu den Ergebnissen hinzu
                $sharedConvertData.Results.Add($file) | Out-Null
                
                Start-Sleep -Milliseconds 100
            }
            
            $endTime = Get-Date
            $duration = $endTime - $startTime
            $sharedConvertData.Duration = "{0:mm\:ss}" -f $duration
            $sharedConvertData.StartTime = $startTime
            
            if ($sharedConvertData.CancelRequested) {
                $sharedConvertData.Status = "Konvertierung abgebrochen"
            } else {
                $sharedConvertData.Status = "✓ Konvertierung abgeschlossen: $($sharedConvertData.SuccessCount) erfolgreich, $($sharedConvertData.FailCount) fehlgeschlagen"
            }
            
        } catch {
            $sharedConvertData.Error = $_.Exception.Message
            $sharedConvertData.Status = "❌ Fehler bei der Konvertierung"
        } finally {
            $sharedConvertData.Completed = $true
        }
    }
    
    $convertPowerShell = [powershell]::Create()
    $convertPowerShell.Runspace = $script:convertRunspace
    $convertPowerShell.AddScript($convertScript).AddArgument($script:sharedConvertData) | Out-Null
    $convertPowerShell.BeginInvoke() | Out-Null
    
    # Timer für Konvertierungs-Updates
    if ($script:timer.IsEnabled) {
        $script:timer.Stop()
    }
    
    $script:timer.Remove_Tick
    
    # Event Handler für Konvertierung
    $convertTimerTick = {
        try {
            if ($script:sharedConvertData -eq $null) { 
                return 
            }
            
            $txtStatus.Text = $script:sharedConvertData.Status
            
            if ($script:sharedConvertData.TotalFiles -gt 0) {
                $pbProgress.Value = ($script:sharedConvertData.CurrentFile / $script:sharedConvertData.TotalFiles) * 100
            }
            
            # DataGrid refresh
            $dgFiles.Items.Refresh()
            
            if ($script:sharedConvertData.Completed) {
                $script:timer.Stop()
                
                if ($script:sharedConvertData.Error) {
                    Complete-Reset-WithTable  # Bei Fehlern Tabelle leeren
                    [System.Windows.MessageBox]::Show("Fehler bei der Konvertierung: $($script:sharedConvertData.Error)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                } else {
                    $pbProgress.Value = 100
                    
                    # HTML-Bericht generieren - verwende die verarbeiteten Dateien aus dem Runspace
                    $convertedFiles = @($script:sharedConvertData.Results)
                    
                    if ($convertedFiles.Count -gt 0) {
                        try {
                            $reportPath = Generate-HTMLReport -ConvertedFiles $convertedFiles -SuccessCount $script:sharedConvertData.SuccessCount -FailCount $script:sharedConvertData.FailCount -Duration $script:sharedConvertData.Duration -StartTime $script:sharedConvertData.StartTime
                        } catch {
                            $reportPath = $null
                        }
                    } else {
                        $reportPath = $null
                    }
                    
                    if ($script:sharedConvertData.CancelRequested) {
                        # Bei Abbruch: Frage den Benutzer ob Tabelle geleert werden soll
                        if ($reportPath) {
                            $result = [System.Windows.MessageBox]::Show(
                                "Konvertierung wurde abgebrochen.`n`nBis zum Abbruch verarbeitet:`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)`n`nMöchtest du den HTML-Bericht öffnen?",
                                "Konvertierung abgebrochen",
                                [System.Windows.MessageBoxButton]::YesNo,
                                [System.Windows.MessageBoxImage]::Warning
                            )
                        } else {
                            $result = [System.Windows.MessageBox]::Show(
                                "Konvertierung wurde abgebrochen.`n`nBis zum Abbruch verarbeitet:`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)",
                                "Konvertierung abgebrochen",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Warning
                            )
                        }
                        
                        # Nach Abbruch: Tabelle leeren für sauberen Zustand
                        Complete-Reset-WithTable
                        
                    } else {
                        # Bei erfolgreichem Abschluss: Tabelle NICHT leeren (Benutzer soll Ergebnisse sehen)
                        if ($reportPath) {
                            $result = [System.Windows.MessageBox]::Show(
                                "Konvertierung abgeschlossen!`n`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)`nDauer: $($script:sharedConvertData.Duration)`n`nMöchtest du den detaillierten HTML-Bericht öffnen?",
                                "Konvertierung abgeschlossen",
                                [System.Windows.MessageBoxButton]::YesNo,
                                [System.Windows.MessageBoxImage]::Information
                            )
                        } else {
                            $result = [System.Windows.MessageBox]::Show(
                                "Konvertierung abgeschlossen!`n`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)`nDauer: $($script:sharedConvertData.Duration)",
                                "Konvertierung abgeschlossen",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Information
                            )
                        }
                        
                        # Bei erfolgreichem Abschluss: Nur Jobs bereinigen, Tabelle behalten (damit Benutzer Ergebnisse sehen kann)
                        Complete-Reset
                    }
                    
                    if ($reportPath -and $result -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process $reportPath
                    }
                }
            }
        } catch {
            $script:timer.Stop()
            Complete-Reset-WithTable  # Bei Fehlern Tabelle leeren
        }
    }
    
    $script:timer.Add_Tick($convertTimerTick)
    $script:timer.Start()
})

# Cancel-Event Handler
$btnCancel.Add_Click({
    $script:cancelRequested = $true
    
    # Cancel-Signal an Runspaces senden
    if ($script:sharedSearchData) {
        try {
            $script:sharedSearchData.CancelRequested = $true
        } catch {}
    }
    
    if ($script:sharedConvertData) {
        try {
            $script:sharedConvertData.CancelRequested = $true
        } catch {}
    }
    
    $btnCancel.IsEnabled = $false
    $txtStatus.Text = "Abbruch wird durchgeführt..."
    
    # Timeout für erzwungenen Stop mit vollständiger Bereinigung inkl. Tabelle
    $timeoutTimer = New-Object System.Windows.Threading.DispatcherTimer
    $timeoutTimer.Interval = [TimeSpan]::FromSeconds(3)
    $timeoutTimer.Add_Tick({
        $timeoutTimer.Stop()
        Complete-Reset-WithTable
        $txtStatus.Text = "Abbruch abgeschlossen"
    })
    $timeoutTimer.Start()
})

# DataGrid Zeilen-Klick Event Handler für Ordner öffnen
$dgFiles.Add_MouseDoubleClick({
    $selectedItem = $dgFiles.SelectedItem
    if ($selectedItem -and $selectedItem.Path) {
        try {
            # Öffne den Windows Explorer mit dem Ordner der ausgewählten Datei
            Start-Process "explorer.exe" -ArgumentList "`"$($selectedItem.Path)`""
        } catch {
            [System.Windows.MessageBox]::Show(
                "Ordner konnte nicht geöffnet werden: $($selectedItem.Path)",
                "Fehler beim Öffnen",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
        }
    }
})

# Hyperlink Event Handler
$lnkWebsite.Add_RequestNavigate({
    param($sender, $e)
    Start-Process $e.Uri.AbsoluteUri
    $e.Handled = $true
})

# Window Closing Event
$window.Add_Closing({
    Complete-Reset-WithTable
})

# Initial Button-Status setzen
Update-ButtonStates

# Window anzeigen
$window.ShowDialog() | Out-Null