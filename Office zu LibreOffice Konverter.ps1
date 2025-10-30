<#
.SYNOPSIS
  Office zu LibreOffice Konverter - Parallel Edition
.DESCRIPTION
 Voraussetzung: LibreOffice muss installiert sein
 Konvertiert Office-Dateien (DOCX, XLSX, PPTX) in LibreOffice-Formate (ODT, ODS, ODP)
 mit paralleler Verarbeitung für maximale Performance
.PARAMETER language
.NOTES
  Version:        1.4
  Author:         Jörn Walter
  Creation Date:  2025-08-20
  Update:         2025-10-30 - Parallele Konvertierung und Live-Dashboard
  Update:         2025-08-24 - Lösch- und Verschiebefunktion für Quelldateien
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

# XAML GUI Definition - Komplett neu und bereinigt
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office zu LibreOffice Konverter v1.3 - © 2025 Jörn Walter" 
        Height="1030" Width="900" MinHeight="800" MaxHeight="1200"
        WindowStartupLocation="CenterScreen" SizeToContent="Manual" ResizeMode="CanResize"
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
    
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
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
                    <Run Text="Entwickelt von Jörn Walter - Der Windows Papst (v1.4 - Parallel Edition)"/>
                </TextBlock>
            </StackPanel>
        </Border>

        <!-- Dashboard / Statistiken -->
        <Border Grid.Row="1" Name="borderDashboard" Background="White" CornerRadius="5" Padding="10" Margin="0,0,0,10" Visibility="Collapsed">
            <StackPanel>
                <TextBlock Text="📊 Live-Statistiken" FontWeight="Bold" FontSize="14" Margin="0,0,0,10" Foreground="#0078D4"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Verarbeitete Dateien -->
                    <Border Grid.Column="0" Background="#E3F2FD" CornerRadius="5" Padding="10" Margin="0,0,5,0">
                        <StackPanel>
                            <TextBlock Text="Verarbeitet" FontSize="10" Foreground="#1565C0" FontWeight="SemiBold"/>
                            <TextBlock Name="txtDashProcessed" Text="0 / 0" FontSize="18" FontWeight="Bold" Foreground="#0078D4" Margin="0,3,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Geschwindigkeit -->
                    <Border Grid.Column="1" Background="#E8F5E9" CornerRadius="5" Padding="10" Margin="5,0,5,0">
                        <StackPanel>
                            <TextBlock Text="Geschwindigkeit" FontSize="10" Foreground="#2E7D32" FontWeight="SemiBold"/>
                            <TextBlock Name="txtDashSpeed" Text="0.0 Dateien/Sek" FontSize="18" FontWeight="Bold" Foreground="#107C10" Margin="0,3,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Verbleibende Zeit -->
                    <Border Grid.Column="2" Background="#FFF3E0" CornerRadius="5" Padding="10" Margin="5,0,5,0">
                        <StackPanel>
                            <TextBlock Text="Verbleibend" FontSize="10" Foreground="#E65100" FontWeight="SemiBold"/>
                            <TextBlock Name="txtDashRemaining" Text="--:--" FontSize="18" FontWeight="Bold" Foreground="#FF8C00" Margin="0,3,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Parallele Jobs -->
                    <Border Grid.Column="3" Background="#F3E5F5" CornerRadius="5" Padding="10" Margin="5,0,0,0">
                        <StackPanel>
                            <TextBlock Text="Aktive Jobs" FontSize="10" Foreground="#6A1B9A" FontWeight="SemiBold"/>
                            <TextBlock Name="txtDashActiveJobs" Text="0 / 4" FontSize="18" FontWeight="Bold" Foreground="#7B1FA2" Margin="0,3,0,0"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- Fortschrittsbalken pro Dateityp -->
                <StackPanel Margin="0,10,0,0">
                    <Grid Margin="0,0,0,5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="60"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="DOCX" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="#0078D4"/>
                        <ProgressBar Grid.Column="1" Name="pbDocx" Height="18" Minimum="0" Maximum="100" Margin="5,0"/>
                        <TextBlock Grid.Column="2" Name="txtDocxCount" Text="0/0" FontSize="10" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                    </Grid>
                    <Grid Margin="0,0,0,5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="60"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="XLSX" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="#107C10"/>
                        <ProgressBar Grid.Column="1" Name="pbXlsx" Height="18" Minimum="0" Maximum="100" Margin="5,0"/>
                        <TextBlock Grid.Column="2" Name="txtXlsxCount" Text="0/0" FontSize="10" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                    </Grid>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="60"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="50"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="PPTX" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="#D13438"/>
                        <ProgressBar Grid.Column="1" Name="pbPptx" Height="18" Minimum="0" Maximum="100" Margin="5,0"/>
                        <TextBlock Grid.Column="2" Name="txtPptxCount" Text="0/0" FontSize="10" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                    </Grid>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- Konvertierungsmodus -->
        <Border Grid.Row="2" Background="White" CornerRadius="5" Padding="10" Margin="0,0,0,10">
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
                
                <!-- Allgemeine Aktionsbuttons -->
                <WrapPanel Margin="0,15,0,0" HorizontalAlignment="Right">
                    <Button Name="btnClearFiles" Content="❌ Liste leeren" Width="120" IsEnabled="False"/>
                    <Button Name="btnSearch" Content="🔍 Dateien suchen" Width="150" Visibility="Collapsed"/>
                </WrapPanel>
            </StackPanel>
        </Border>

        <!-- Dateiliste -->
        <Border Grid.Row="3" Background="White" CornerRadius="5" Padding="10" Height="200">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Zu konvertierende Dateien (Doppelklick öffnet Ordner):" FontWeight="Bold" Margin="0,0,0,5"/>
                <DataGrid Grid.Row="1" Name="dgFiles" AutoGenerateColumns="False" CanUserAddRows="False" 
                          GridLinesVisibility="Horizontal" HeadersVisibility="Column" Height="160"
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
        <Border Grid.Row="4" Background="White" CornerRadius="5" Padding="10" Margin="0,10,0,0">
            <StackPanel>
                <TextBlock Text="Konvertierungsoptionen:" FontWeight="Bold" Margin="0,0,0,5"/>
                <WrapPanel>
                    <CheckBox Name="chkConvertDocx" Content="DOCX → ODT" IsChecked="True"/>
                    <CheckBox Name="chkConvertXlsx" Content="XLSX → ODS" IsChecked="True" Margin="20,5,5,5"/>
                    <CheckBox Name="chkConvertPptx" Content="PPTX → ODP" IsChecked="True" Margin="20,5,5,5"/>
                </WrapPanel>
                <WrapPanel Margin="0,5,0,0">
                    <CheckBox Name="chkOverwrite" Content="Vorhandene Dateien überschreiben" Margin="5"/>
                </WrapPanel>

                <!-- Parallele Konvertierung -->
                <Border Background="#F0F8FF" BorderBrush="#0078D4" BorderThickness="1" CornerRadius="3" Padding="8" Margin="0,10,0,0">
                    <StackPanel>
                        <TextBlock Text="⚡ Parallele Konvertierung" FontWeight="SemiBold" Foreground="#0078D4" Margin="0,0,0,5"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Maximale parallele Jobs:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <Slider Grid.Column="1" Name="sliderMaxJobs" Minimum="1" Maximum="16" Value="4"
                                    TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <TextBlock Grid.Column="2" Name="txtMaxJobs" Text="4" FontWeight="Bold"
                                       VerticalAlignment="Center" Width="25" Foreground="#0078D4"/>
                        </Grid>
                        <TextBlock Text="💡 Höhere Werte = schneller, aber mehr CPU/RAM-Last"
                                   FontSize="10" Foreground="#666" FontStyle="Italic" Margin="0,3,0,0"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>

        <!-- Quelldateien-Optionen -->
        <Border Grid.Row="5" Background="White" CornerRadius="5" Padding="10" Margin="0,10,0,0">
            <StackPanel>
                <TextBlock Text="Quelldateien nach erfolgreicher Konvertierung:" FontWeight="Bold" Margin="0,0,0,5"/>
                
                <WrapPanel Margin="0,5,0,0">
                    <RadioButton Name="rbKeepSource" Content="📄 Behalten (keine Aktion)" IsChecked="True" Margin="5"/>
                    <RadioButton Name="rbDeleteSource" Content="🗑️ Löschen (unwiderruflich)" 
                                 Foreground="#D13438" FontWeight="Bold" Margin="20,5,5,5"/>
                    <RadioButton Name="rbMoveSource" Content="📦 Verschieben (sicher)" 
                                 Foreground="#0078D4" FontWeight="Bold" Margin="20,5,5,5"/>
                </WrapPanel>
                
                <!-- Verschieben-Optionen -->
                <StackPanel Name="spMoveOptions" Visibility="Collapsed" Margin="0,10,0,0">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Zielordner:" Width="90"/>
                        <TextBox Name="txtMoveFolder" Width="380" Margin="5,0"/>
                        <Button Name="btnBrowseMove" Content="Durchsuchen..." Width="100"/>
                    </StackPanel>
                    <CheckBox Name="chkCreateSubfolders" Content="Unterordner nach Dateityp erstellen (DOCX, XLSX, PPTX)" 
                              IsChecked="True" Margin="95,5,0,0"/>
                    <CheckBox Name="chkAddTimestamp" Content="Zeitstempel zu Ordnernamen hinzufügen" 
                              IsChecked="True" Margin="95,5,0,0"/>
                </StackPanel>
                
                <!-- Warnungen -->
                <Border Name="deleteWarning" Background="#FFF4E6" BorderBrush="#FF8C00" BorderThickness="1" 
                        CornerRadius="3" Padding="8" Margin="0,5,0,0" Visibility="Collapsed">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="⚠️" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#B8860B" FontWeight="SemiBold">
                            <Run Text="ACHTUNG: Die Quelldateien werden unwiderruflich gelöscht! "/>
                            <Run Text="Stelle sicher, dass du Backups hast." FontWeight="Bold"/>
                        </TextBlock>
                    </StackPanel>
                </Border>
                
                <Border Name="moveInfo" Background="#E3F2FD" BorderBrush="#0078D4" BorderThickness="1" 
                        CornerRadius="3" Padding="8" Margin="0,5,0,0" Visibility="Collapsed">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="💡" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock TextWrapping="Wrap" Foreground="#1565C0" FontWeight="SemiBold">
                            <Run Text="INFO: Quelldateien werden sicher verschoben. "/>
                            <Run Text="Du kannst sie jederzeit wiederherstellen." FontWeight="Bold"/>
                        </TextBlock>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>

        <!-- Fortschritt -->
        <StackPanel Grid.Row="6" Margin="0,5,0,0">
            <ProgressBar Name="pbProgress" Height="20" Minimum="0" Maximum="100" IsIndeterminate="False"/>
            <TextBlock Name="txtStatus" Text="Bereit" HorizontalAlignment="Center" Margin="0,3,0,0" TextWrapping="Wrap"/>
        </StackPanel>

        <!-- Buttons -->
        <WrapPanel Grid.Row="7" HorizontalAlignment="Center" Margin="0,5,0,0">
            <Button Name="btnConvert" Content="▶ Konvertierung starten" Width="180" Height="32" FontSize="14" FontWeight="Bold"/>
            <Button Name="btnCancel" Content="⏹ Abbrechen" Width="120" Height="32" FontSize="14" IsEnabled="False"/>
        </WrapPanel>

        <!-- Copyright Footer -->
        <Border Grid.Row="8" Background="#E0E0E0" CornerRadius="5" Padding="8" Margin="0,5,0,0">
            <StackPanel>
                <TextBlock HorizontalAlignment="Center" FontSize="10">
                    <Run Text="© 2025 Jörn Walter - " Foreground="#555"/>
                    <Hyperlink Name="lnkWebsite" NavigateUri="https://www.der-windows-papst.de" Foreground="#0078D4">
                        <Run Text="www.der-windows-papst.de"/>
                    </Hyperlink>
                    <Run Text=" | Office zu LibreOffice Konverter v1.4 - Parallel Edition" Foreground="#777"/>
                </TextBlock>
            </StackPanel>
        </Border>
    </Grid>
    </ScrollViewer>
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
$sliderMaxJobs = $window.FindName("sliderMaxJobs")
$txtMaxJobs = $window.FindName("txtMaxJobs")

# Dashboard Controls
$borderDashboard = $window.FindName("borderDashboard")
$txtDashProcessed = $window.FindName("txtDashProcessed")
$txtDashSpeed = $window.FindName("txtDashSpeed")
$txtDashRemaining = $window.FindName("txtDashRemaining")
$txtDashActiveJobs = $window.FindName("txtDashActiveJobs")
$pbDocx = $window.FindName("pbDocx")
$pbXlsx = $window.FindName("pbXlsx")
$pbPptx = $window.FindName("pbPptx")
$txtDocxCount = $window.FindName("txtDocxCount")
$txtXlsxCount = $window.FindName("txtXlsxCount")
$txtPptxCount = $window.FindName("txtPptxCount")

# Neue Controls für Quelldateien-Optionen
$rbKeepSource = $window.FindName("rbKeepSource")
$rbDeleteSource = $window.FindName("rbDeleteSource")
$rbMoveSource = $window.FindName("rbMoveSource")
$spMoveOptions = $window.FindName("spMoveOptions")
$txtMoveFolder = $window.FindName("txtMoveFolder")
$btnBrowseMove = $window.FindName("btnBrowseMove")
$chkCreateSubfolders = $window.FindName("chkCreateSubfolders")
$chkAddTimestamp = $window.FindName("chkAddTimestamp")
$deleteWarning = $window.FindName("deleteWarning")
$moveInfo = $window.FindName("moveInfo")

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
$script:runspacePool = $null
$script:activeJobs = @()
$script:timer = $null
$script:libreOfficePath = Get-LibreOfficePath
$script:sharedSearchData = $null
$script:sharedConvertData = $null
$script:dashboardStartTime = $null
$script:dashboardProcessed = 0
$script:dashboardTotal = 0

# DataGrid Binding
$dgFiles.ItemsSource = $script:files

# Initial UI-Setup
$btnSearch.Visibility = "Collapsed"

# Standard-Pfad für Verschieben setzen
$defaultMoveFolder = Join-Path ([Environment]::GetFolderPath("Desktop")) "Office_Originaldateien"
$txtMoveFolder.Text = $defaultMoveFolder

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

# Event Handler für Quelldateien-Optionen
$rbKeepSource.Add_Checked({
    $spMoveOptions.Visibility = "Collapsed"
    $deleteWarning.Visibility = "Collapsed"
    $moveInfo.Visibility = "Collapsed"
})

$rbDeleteSource.Add_Checked({
    $spMoveOptions.Visibility = "Collapsed"
    $deleteWarning.Visibility = "Visible"
    $moveInfo.Visibility = "Collapsed"
})

$rbMoveSource.Add_Checked({
    $spMoveOptions.Visibility = "Visible"
    $deleteWarning.Visibility = "Collapsed"
    $moveInfo.Visibility = "Visible"
})

# Slider Event Handler
$sliderMaxJobs.Add_ValueChanged({
    $txtMaxJobs.Text = [int]$sliderMaxJobs.Value
})

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
        Extension = $file.Extension.ToLower()  # Normalisiere Extension
        Size = "{0:N2} KB" -f ($file.Length / 1KB)
        Path = $file.DirectoryName
        FullPath = $file.FullName
        Status = "Bereit"
        SourceDeleted = $false
        SourceMoved = $false
        MovedTo = ""
    }
    
    # UI-Update im Dispatcher
    $window.Dispatcher.Invoke([System.Action]{
        $script:files.Add($fileObj)
        Update-ButtonStates
    })
}

function Generate-HTMLReport {
    param(
        [array]$ConvertedFiles,
        [int]$SuccessCount,
        [int]$FailCount,
        [int]$DeletedCount,
        [int]$MovedCount,
        [string]$Duration,
        [datetime]$StartTime
    )
    
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    
    # Desktop-Pfad ermitteln und Bericht dort speichern
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $reportFileName = "LibreOffice_Konvertierung_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $reportPath = Join-Path $desktopPath $reportFileName
    
    # Sicherstellen, dass Desktop-Pfad existiert
    if (-not (Test-Path $desktopPath)) {
        Write-Warning "Desktop-Pfad nicht gefunden, verwende Fallback-Pfad"
        $reportPath = Join-Path $env:USERPROFILE "Desktop\$reportFileName"
        
        if (-not (Test-Path (Split-Path $reportPath))) {
            $reportPath = Join-Path $env:TEMP $reportFileName
        }
    }
    
    # Statistiken berechnen
    $totalFiles = $ConvertedFiles.Count
    $successRate = if ($totalFiles -gt 0) { [math]::Round(($SuccessCount / $totalFiles) * 100, 1) } else { 0 }
    
    # Gruppierung nach Status
    $successFiles = @($ConvertedFiles | Where-Object { $_.Status -match "✓ Erfolgreich" })
    $failedFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "✗ Fehler" })
    $existingFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "⚠ Bereits vorhanden" })
    $cancelledFiles = @($ConvertedFiles | Where-Object { $_.Status -eq "Abgebrochen" })
    
    # Gruppierung nach Dateityp
    $docxCount = ($ConvertedFiles | Where-Object { $_.Extension.ToLower() -eq ".docx" }).Count
    $xlsxCount = ($ConvertedFiles | Where-Object { $_.Extension.ToLower() -eq ".xlsx" }).Count
    $pptxCount = ($ConvertedFiles | Where-Object { $_.Extension.ToLower() -eq ".pptx" }).Count
    
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
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
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
            font-size: 2.2em;
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
        
        .info-card.deleted .value {
            color: #8B4513;
        }
        
        .info-card.moved .value {
            color: #0078D4;
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
            gap: 8px;
            flex-wrap: wrap;
        }
        
        .filter-btn {
            padding: 6px 12px;
            border: 2px solid #0078D4;
            background: white;
            color: #0078D4;
            border-radius: 5px;
            cursor: pointer;
            transition: all 0.3s ease;
            font-weight: 600;
            font-size: 0.85em;
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
        
        .status-success-deleted {
            background: #e8f5e8;
            color: #0d5f0d;
            border: 1px solid #107c10;
        }
        
        .status-success-moved {
            background: #e3f2fd;
            color: #0078D4;
            border: 1px solid #0078D4;
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
            max-width: 250px;
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
        
        @media (max-width: 768px) {
            .info-grid {
                grid-template-columns: 1fr;
            }
            
            table {
                font-size: 0.9em;
            }
            
            .file-path {
                max-width: 150px;
            }
            
            .filter-buttons {
                justify-content: center;
            }
            
            .filter-btn {
                font-size: 0.8em;
                padding: 4px 8px;
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
            
            <div class="info-card deleted">
                <h3>Quelldateien gelöscht</h3>
                <div class="value">$DeletedCount</div>
            </div>
            
            <div class="info-card moved">
                <h3>Quelldateien verschoben</h3>
                <div class="value">$MovedCount</div>
            </div>
            
            <div class="info-card">
                <h3>Konvertierungsdauer</h3>
                <div class="value" style="font-size: 1.6em;">$Duration</div>
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
                    <button class="filter-btn" onclick="filterTable('deleted')">Mit Löschung</button>
                    <button class="filter-btn" onclick="filterTable('moved')">Mit Verschiebung</button>
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
        $statusClass = switch -Regex ($file.Status) {
            "✓ Erfolgreich.*verschoben" { "success-moved" }
            "✓ Erfolgreich.*gelöscht" { "success-deleted" }
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
            <p><strong>Office zu LibreOffice Konverter v1.4 - Parallel Edition</strong></p>
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
                } else if (status === 'deleted') {
                    row.style.display = row.dataset.status === 'success-deleted' ? '' : 'none';
                } else if (status === 'moved') {
                    row.style.display = row.dataset.status === 'success-moved' ? '' : 'none';
                } else {
                    row.style.display = row.dataset.status === status ? '' : 'none';
                }
            });
        }
    </script>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Bericht erfolgreich gespeichert: $reportPath"
    }
    catch {
        Write-Warning "Fehler beim Speichern des Berichts: $_"
        $reportPath = Join-Path $env:TEMP $reportFileName
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Bericht im Temp-Ordner gespeichert: $reportPath"
    }
    
    return $reportPath
}

# Erweiterte Convert-File Funktion mit Verschiebefunktion
function Convert-File {
    param(
        [string]$InputFile,
        [string]$OutputDir,
        [string]$SourceAction = "keep",  # "keep", "delete", "move"
        [string]$MoveFolder = "",
        [bool]$CreateSubfolders = $false
    )
    
    if (-not (Test-Path $InputFile)) {
        return @{ Success = $false; Status = "✗ Fehler"; Message = "Eingabedatei nicht gefunden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
    }
    
    if (-not (Test-Path $OutputDir)) {
        try {
            New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        } catch {
            return @{ Success = $false; Status = "✗ Fehler"; Message = "Ausgabeordner konnte nicht erstellt werden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
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
            return @{ Success = $false; Status = "✗ Fehler"; Message = "Nicht unterstütztes Format: $extension"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
        }
    }
    
    # Output-Dateiname
    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name) + ".$outputFormat"
    $outputPath = Join-Path $OutputDir $outputName
    
    # Prüfe ob Datei existiert
    if ((Test-Path $outputPath) -and -not $chkOverwrite.IsChecked) {
        return @{ Success = $false; Status = "⚠ Bereits vorhanden"; Message = "Datei existiert bereits"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
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
            # Konvertierung erfolgreich - handle Quelldatei basierend auf Action
            if ($SourceAction -eq "delete") {
                try {
                    if (Test-Path $outputPath) {
                        Remove-Item -Path $InputFile -Force
                        return @{ Success = $true; Status = "✓ Erfolgreich + Quelle gelöscht"; Message = "Konvertierung erfolgreich, Quelldatei gelöscht"; SourceDeleted = $true; SourceMoved = $false; MovedTo = "" }
                    } else {
                        return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich (Quelldatei nicht gelöscht - Zieldatei nicht gefunden)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                    }
                } catch {
                    return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich, aber Quelldatei konnte nicht gelöscht werden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                }
            }
            elseif ($SourceAction -eq "move" -and -not [string]::IsNullOrWhiteSpace($MoveFolder)) {
                try {
                    if (Test-Path $outputPath) {
                        # MoveFolder ist bereits vorbereitet
                        $finalMoveFolder = $MoveFolder
                        
                        # Unterordner nach Dateityp erstellen wenn gewünscht
                        if ($CreateSubfolders) {
                            $subfolderName = switch ($extension) {
                                ".docx" { "DOCX_Dateien" }
                                ".xlsx" { "XLSX_Dateien" }
                                ".pptx" { "PPTX_Dateien" }
                            }
                            $finalMoveFolder = Join-Path $MoveFolder $subfolderName
                        }
                        
                        # Zielordner erstellen falls nicht vorhanden
                        if (-not (Test-Path $finalMoveFolder)) {
                            New-Item -Path $finalMoveFolder -ItemType Directory -Force | Out-Null
                        }
                        
                        # Datei verschieben
                        $movePath = Join-Path $finalMoveFolder $inputItem.Name
                        
                        # Falls Datei bereits im Zielordner existiert, umbenennen
                        if (Test-Path $movePath) {
                            $counter = 1
                            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name)
                            $fileExtension = $inputItem.Extension
                            do {
                                $newName = "$baseName`_$counter$fileExtension"
                                $movePath = Join-Path $finalMoveFolder $newName
                                $counter++
                            } while (Test-Path $movePath)
                        }
                        
                        Move-Item -Path $InputFile -Destination $movePath -Force
                        return @{ Success = $true; Status = "✓ Erfolgreich + Quelle verschoben"; Message = "Konvertierung erfolgreich, Quelldatei verschoben"; SourceDeleted = $false; SourceMoved = $true; MovedTo = $MoveFolder }
                    } else {
                        return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich (Quelldatei nicht verschoben - Zieldatei nicht gefunden)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                    }
                } catch {
                    return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich, aber Quelldatei konnte nicht verschoben werden: $($_.Exception.Message)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                }
            }
            else {
                # Standard: Quelldatei behalten
                return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
            }
        } else {
            return @{ Success = $false; Status = "✗ Fehler"; Message = "LibreOffice Konvertierung fehlgeschlagen (Exit Code: $($process.ExitCode))"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
        }
    } catch {
        return @{ Success = $false; Status = "✗ Fehler"; Message = "Fehler beim Starten von LibreOffice: $($_.Exception.Message)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
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

    # Cleanup parallele Jobs
    if ($script:activeJobs) {
        foreach ($job in $script:activeJobs) {
            try {
                if ($job.PowerShell) {
                    $job.PowerShell.Stop()
                    $job.PowerShell.Dispose()
                }
            } catch {}
        }
        $script:activeJobs = @()
    }

    # Runspace Pool cleanup
    if ($script:runspacePool) {
        try {
            $script:runspacePool.Close()
            $script:runspacePool.Dispose()
            $script:runspacePool = $null
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

    # Cleanup parallele Jobs
    if ($script:activeJobs) {
        foreach ($job in $script:activeJobs) {
            try {
                if ($job.PowerShell) {
                    $job.PowerShell.Stop()
                    $job.PowerShell.Dispose()
                }
            } catch {}
        }
        $script:activeJobs = @()
    }

    # Runspace Pool cleanup
    if ($script:runspacePool) {
        try {
            $script:runspacePool.Close()
            $script:runspacePool.Dispose()
            $script:runspacePool = $null
        } catch {}
    }

    # Shared Data komplett zurücksetzen
    $script:sharedSearchData = $null
    $script:sharedConvertData = $null

    # Dashboard ausblenden
    $borderDashboard.Visibility = "Collapsed"

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
    Complete-Reset-WithTable
    
    if ($script:files.Count -gt 0 -or $dgFiles.Items.Count -gt 0) {
        while ($script:files.Count -gt 0) {
            $script:files.RemoveAt(0)
        }
        
        $dgFiles.Items.Clear()
        
        $script:files = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $dgFiles.ItemsSource = $script:files
        
        [System.Windows.Forms.Application]::DoEvents()
    }
    
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

# Neuer Event Handler für Verschieben-Ordner
$btnBrowseMove.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Wähle den Ordner für verschobene Quelldateien"
    $dialog.SelectedPath = $txtMoveFolder.Text
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtMoveFolder.Text = $dialog.SelectedPath
    }
})

# Search Event Handler - Mit Debug-Ausgaben
$btnSearch.Add_Click({
    if (-not $rbSystemWide.IsChecked) { return }
    
    $searchPath = $txtSearchPath.Text
    if ([string]::IsNullOrWhiteSpace($searchPath) -or -not (Test-Path $searchPath)) {
        [System.Windows.MessageBox]::Show("Bitte gib einen gültigen Suchpfad an.", "Ungültiger Pfad", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Debug: Überprüfe Checkbox-Zustände
    Write-Host "=== DEBUG: Checkbox-Zustände ==="
    Write-Host "DOCX Checkbox: $($chkConvertDocx.IsChecked)"
    Write-Host "XLSX Checkbox: $($chkConvertXlsx.IsChecked)" 
    Write-Host "PPTX Checkbox: $($chkConvertPptx.IsChecked)"
    
    $extensions = Get-FileExtensions
    if ($extensions.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wähle mindestens einen Dateityp aus.", "Kein Dateityp", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Debug: Zeige gewählte Extensions
    $extensionsList = $extensions -join ", "
    Write-Host "=== DEBUG: Suchende Extensions ==="
    Write-Host "Extensions Array: $extensionsList"
    Write-Host "Extensions Count: $($extensions.Count)"
    
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
                
                try {
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
                                Extension = $file.Extension.ToLower()  # Normalisiere Extension
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
                } catch {
                    # Ignoriere Fehler bei einzelnen Ordnern und setze Suche fort
                    continue
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
                        SourceDeleted = $false
                        SourceMoved = $false
                        MovedTo = ""
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
                        # Debug: Zeige alle Extensions in der Kollektion
                        Write-Host "=== DEBUG: Alle gefundenen Dateien ==="
                        foreach ($file in $script:files) {
                            Write-Host "Datei: $($file.Name) | Extension: '$($file.Extension)'"
                        }
                        
                        $docxCount = ($script:files | Where-Object { $_.Extension -eq ".docx" }).Count
                        $xlsxCount = ($script:files | Where-Object { $_.Extension -eq ".xlsx" }).Count
                        $pptxCount = ($script:files | Where-Object { $_.Extension -eq ".pptx" }).Count
                        
                        Write-Host "=== DEBUG: Gezählte Dateien ==="
                        Write-Host "DOCX Count: $docxCount"
                        Write-Host "XLSX Count: $xlsxCount"
                        Write-Host "PPTX Count: $pptxCount"
                        Write-Host "Total Count: $($script:files.Count)"
                        
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
            Complete-Reset-WithTable
        }
    }
    
    $script:timer.Add_Tick($searchTimerTick)
    $script:timer.Start()
})

# Convert Event Handler mit Verschiebefunktion
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
    
    # Bestimme Quelldateien-Aktion
    $sourceAction = "keep"
    $moveFolder = ""
    $createSubfolders = $false
    $addTimestamp = $false
    
    if ($rbDeleteSource.IsChecked) {
        $sourceAction = "delete"
        $result = [System.Windows.MessageBox]::Show(
            "⚠️ ACHTUNG: Quelldateien löschen ist aktiviert!`n`n" +
            "Die Originaldateien werden nach erfolgreicher Konvertierung unwiderruflich gelöscht.`n" +
            "Stelle sicher, dass du Backups hast oder diese Funktion wirklich verwenden möchtest.`n`n" +
            "Möchtest du mit der Konvertierung fortfahren?",
            "Warnung: Quelldateien werden gelöscht",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($result -eq [System.Windows.MessageBoxResult]::No) { return }
    }
    elseif ($rbMoveSource.IsChecked) {
        $sourceAction = "move"
        $moveFolder = $txtMoveFolder.Text
        $createSubfolders = $chkCreateSubfolders.IsChecked
        $addTimestamp = $chkAddTimestamp.IsChecked
        
        if ([string]::IsNullOrWhiteSpace($moveFolder)) {
            [System.Windows.MessageBox]::Show("Bitte gib einen Zielordner für das Verschieben an.", "Zielordner fehlt", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $result = [System.Windows.MessageBox]::Show(
            "📦 Quelldateien verschieben ist aktiviert.`n`n" +
            "Die Originaldateien werden nach erfolgreicher Konvertierung verschoben nach:`n" +
            "$moveFolder`n`n" +
            "Unterordner nach Dateityp: $($chkCreateSubfolders.IsChecked)`n" +
            "Zeitstempel hinzufügen: $($chkAddTimestamp.IsChecked)`n`n" +
            "Möchtest du mit der Konvertierung fortfahren?",
            "Info: Quelldateien werden verschoben",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($result -eq [System.Windows.MessageBoxResult]::No) { return }
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
    
    # Einmaligen Zeitstempel für alle Dateien generieren
    $globalTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Finalen Verschiebe-Ordner einmalig bestimmen
    $finalMoveFolder = $moveFolder
    if ($sourceAction -eq "move" -and $addTimestamp -and -not [string]::IsNullOrWhiteSpace($moveFolder)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($moveFolder)
        $parentDir = [System.IO.Path]::GetDirectoryName($moveFolder)
        if ([string]::IsNullOrEmpty($parentDir)) { $parentDir = [System.IO.Path]::GetDirectoryName($moveFolder) }
        $finalMoveFolder = Join-Path $parentDir "$baseName`_$globalTimestamp"
    }

    # Max parallele Jobs aus Slider
    $maxParallelJobs = [int]$sliderMaxJobs.Value

    # Dashboard initialisieren
    $script:dashboardStartTime = Get-Date
    $script:dashboardProcessed = 0
    $script:dashboardTotal = $selectedFiles.Count
    $borderDashboard.Visibility = "Visible"

    # Statistiken nach Dateityp
    $docxFiles = @($selectedFiles | Where-Object { $_.Extension -eq ".docx" })
    $xlsxFiles = @($selectedFiles | Where-Object { $_.Extension -eq ".xlsx" })
    $pptxFiles = @($selectedFiles | Where-Object { $_.Extension -eq ".pptx" })

    # Shared Variables für Konvertierung (thread-safe)
    $script:sharedConvertData = [hashtable]::Synchronized(@{
        Files = @($selectedFiles)
        TargetFolder = $targetFolder
        IsSystemWide = $rbSystemWide.IsChecked
        OverwriteFiles = $chkOverwrite.IsChecked
        SourceAction = $sourceAction
        MoveFolder = $finalMoveFolder
        CreateSubfolders = $createSubfolders
        AddTimestamp = $false  # Bereits verarbeitet
        GlobalTimestamp = $globalTimestamp
        LibreOfficePath = $script:libreOfficePath
        CancelRequested = $false
        CurrentFile = 0
        TotalFiles = $selectedFiles.Count
        SuccessCount = 0
        FailCount = 0
        DeletedCount = 0
        MovedCount = 0
        Status = "Konvertierung läuft..."
        Completed = $false
        Error = $null
        Results = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Duration = ""
        StartTime = (Get-Date)
        # Erweiterte Statistiken für Dashboard
        DocxTotal = $docxFiles.Count
        DocxProcessed = 0
        XlsxTotal = $xlsxFiles.Count
        XlsxProcessed = 0
        PptxTotal = $pptxFiles.Count
        PptxProcessed = 0
        ActiveJobs = 0
        MaxJobs = $maxParallelJobs
    })
    
    # Runspace Pool für parallele Konvertierung erstellen
    $script:runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxParallelJobs)
    $script:runspacePool.Open()

    # Script für einzelne Datei-Konvertierung (läuft parallel)
    $singleFileConvertScript = {
        param($file, $sharedData)

        function Convert-SingleFile {
            param(
                [string]$InputFile,
                [string]$OutputDir,
                [bool]$OverwriteFiles,
                [string]$SourceAction,
                [string]$MoveFolder,
                [bool]$CreateSubfolders,
                [string]$LibreOfficePath
            )
            
            if (-not (Test-Path $InputFile)) { 
                return @{ Success = $false; Status = "✗ Fehler"; Message = "Eingabedatei nicht gefunden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
            }
            
            if (-not (Test-Path $OutputDir)) {
                try {
                    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
                } catch {
                    return @{ Success = $false; Status = "✗ Fehler"; Message = "Ausgabeordner konnte nicht erstellt werden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                }
            }
            
            $inputItem = Get-Item $InputFile
            $extension = $inputItem.Extension.ToLower()
            
            $outputFormat = switch ($extension) {
                ".docx" { "odt" }
                ".xlsx" { "ods" }
                ".pptx" { "odp" }
                default { return @{ Success = $false; Status = "✗ Fehler"; Message = "Nicht unterstütztes Format: $extension"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" } }
            }
            
            $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name) + ".$outputFormat"
            $outputPath = Join-Path $OutputDir $outputName
            
            if ((Test-Path $outputPath) -and -not $OverwriteFiles) {
                return @{ Success = $false; Status = "⚠ Bereits vorhanden"; Message = "Datei existiert bereits"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
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
                    # Konvertierung erfolgreich - handle Quelldatei basierend auf Action
                    if ($SourceAction -eq "delete") {
                        try {
                            if (Test-Path $outputPath) {
                                Remove-Item -Path $InputFile -Force
                                return @{ Success = $true; Status = "✓ Erfolgreich + Quelle gelöscht"; Message = "Konvertierung erfolgreich, Quelldatei gelöscht"; SourceDeleted = $true; SourceMoved = $false; MovedTo = "" }
                            } else {
                                return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich (Quelldatei nicht gelöscht - Zieldatei nicht gefunden)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                            }
                        } catch {
                            return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich, aber Quelldatei konnte nicht gelöscht werden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                        }
                    }
                    elseif ($SourceAction -eq "move" -and -not [string]::IsNullOrWhiteSpace($MoveFolder)) {
                        try {
                            if (Test-Path $outputPath) {
                                # MoveFolder ist bereits vorbereitet
                                $finalMoveFolder = $MoveFolder
                                
                                # Unterordner nach Dateityp erstellen wenn gewünscht
                                if ($CreateSubfolders) {
                                    $subfolderName = switch ($extension) {
                                        ".docx" { "DOCX_Dateien" }
                                        ".xlsx" { "XLSX_Dateien" }
                                        ".pptx" { "PPTX_Dateien" }
                                    }
                                    $finalMoveFolder = Join-Path $MoveFolder $subfolderName
                                }
                                
                                # Zielordner erstellen falls nicht vorhanden
                                if (-not (Test-Path $finalMoveFolder)) {
                                    New-Item -Path $finalMoveFolder -ItemType Directory -Force | Out-Null
                                }
                                
                                # Datei verschieben
                                $movePath = Join-Path $finalMoveFolder $inputItem.Name
                                
                                # Falls Datei bereits im Zielordner existiert, umbenennen
                                if (Test-Path $movePath) {
                                    $counter = 1
                                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputItem.Name)
                                    $fileExtension = $inputItem.Extension
                                    do {
                                        $newName = "$baseName`_$counter$fileExtension"
                                        $movePath = Join-Path $finalMoveFolder $newName
                                        $counter++
                                    } while (Test-Path $movePath)
                                }
                                
                                Move-Item -Path $InputFile -Destination $movePath -Force
                                return @{ Success = $true; Status = "✓ Erfolgreich + Quelle verschoben"; Message = "Konvertierung erfolgreich, Quelldatei verschoben"; SourceDeleted = $false; SourceMoved = $true; MovedTo = $MoveFolder }
                            } else {
                                return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich (Quelldatei nicht verschoben - Zieldatei nicht gefunden)"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                            }
                        } catch {
                            return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich, aber Quelldatei konnte nicht verschoben werden"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                        }
                    }
                    else {
                        # Standard: Quelldatei behalten
                        return @{ Success = $true; Status = "✓ Erfolgreich"; Message = "Konvertierung erfolgreich"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                    }
                } else {
                    return @{ Success = $false; Status = "✗ Fehler"; Message = "LibreOffice Konvertierung fehlgeschlagen"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
                }
            } catch {
                return @{ Success = $false; Status = "✗ Fehler"; Message = "Fehler beim Starten von LibreOffice"; SourceDeleted = $false; SourceMoved = $false; MovedTo = "" }
            }
        }
        
        try {
            # Abbruch-Prüfung
            if ($sharedData.CancelRequested) {
                $file.Status = "Abgebrochen"
                $file.SourceDeleted = $false
                $file.SourceMoved = $false
                $file.MovedTo = ""
                $sharedData.Results.Add($file) | Out-Null
                return
            }

            # Erhöhe ActiveJobs counter
            $sharedData.ActiveJobs++

            # Output-Verzeichnis bestimmen
            $outputDir = if ($sharedData.IsSystemWide) { $file.Path } else { $sharedData.TargetFolder }

            # Konvertiere die Datei
            $result = Convert-SingleFile -InputFile $file.FullPath -OutputDir $outputDir -OverwriteFiles $sharedData.OverwriteFiles -SourceAction $sharedData.SourceAction -MoveFolder $sharedData.MoveFolder -CreateSubfolders $sharedData.CreateSubfolders -LibreOfficePath $sharedData.LibreOfficePath

            # Setze Status und Flags
            $file.Status = $result.Status
            $file.SourceDeleted = $result.SourceDeleted
            $file.SourceMoved = $result.SourceMoved
            $file.MovedTo = $result.MovedTo

            # Thread-sicher Zähler erhöhen
            if ($result.Success) {
                $sharedData.SuccessCount++
            } elseif ($result.Status -eq "✗ Fehler") {
                $sharedData.FailCount++
            }

            if ($result.SourceDeleted) {
                $sharedData.DeletedCount++
            }

            if ($result.SourceMoved) {
                $sharedData.MovedCount++
            }

            # Statistiken nach Dateityp aktualisieren
            $extension = $file.Extension.ToLower()
            if ($extension -eq ".docx") {
                $sharedData.DocxProcessed++
            } elseif ($extension -eq ".xlsx") {
                $sharedData.XlsxProcessed++
            } elseif ($extension -eq ".pptx") {
                $sharedData.PptxProcessed++
            }

            # Gesamtzähler
            $sharedData.CurrentFile++
            $sharedData.Results.Add($file) | Out-Null

        } catch {
            $file.Status = "✗ Fehler: $($_.Exception.Message)"
            $sharedData.FailCount++
            $sharedData.CurrentFile++
            $sharedData.Results.Add($file) | Out-Null
        } finally {
            # Reduziere ActiveJobs counter
            $sharedData.ActiveJobs--
        }
    }

    # Jobs für alle Dateien starten (parallel)
    $jobs = @()
    foreach ($file in $selectedFiles) {
        $powerShell = [powershell]::Create()
        $powerShell.RunspacePool = $script:runspacePool
        $powerShell.AddScript($singleFileConvertScript).AddArgument($file).AddArgument($script:sharedConvertData) | Out-Null

        $jobs += [PSCustomObject]@{
            PowerShell = $powerShell
            Handle = $powerShell.BeginInvoke()
            File = $file
        }
    }

    # Speichere Jobs für Timer-Zugriff
    $script:activeJobs = $jobs

    # Timer für Konvertierungs-Updates
    if ($script:timer.IsEnabled) {
        $script:timer.Stop()
    }
    
    $script:timer.Remove_Tick
    
    # Event Handler für Konvertierung mit Dashboard-Updates
    $convertTimerTick = {
        try {
            if ($script:sharedConvertData -eq $null) {
                return
            }

            # Dashboard-Updates
            $processed = $script:sharedConvertData.CurrentFile
            $total = $script:sharedConvertData.TotalFiles
            $txtDashProcessed.Text = "$processed / $total"

            # Geschwindigkeit berechnen
            $elapsed = (Get-Date) - $script:dashboardStartTime
            if ($elapsed.TotalSeconds -gt 0 -and $processed -gt 0) {
                $speed = $processed / $elapsed.TotalSeconds
                $txtDashSpeed.Text = "$($speed.ToString('0.0')) Dateien/Sek"

                # Verbleibende Zeit
                $remaining = $total - $processed
                if ($speed -gt 0) {
                    $remainingSeconds = $remaining / $speed
                    $txtDashRemaining.Text = [TimeSpan]::FromSeconds($remainingSeconds).ToString('mm\:ss')
                }
            }

            # Aktive Jobs
            $activeJobsCount = $script:sharedConvertData.ActiveJobs
            $txtDashActiveJobs.Text = "$activeJobsCount / $($script:sharedConvertData.MaxJobs)"

            # Fortschritt pro Dateityp
            if ($script:sharedConvertData.DocxTotal -gt 0) {
                $pbDocx.Value = ($script:sharedConvertData.DocxProcessed / $script:sharedConvertData.DocxTotal) * 100
                $txtDocxCount.Text = "$($script:sharedConvertData.DocxProcessed)/$($script:sharedConvertData.DocxTotal)"
            }
            if ($script:sharedConvertData.XlsxTotal -gt 0) {
                $pbXlsx.Value = ($script:sharedConvertData.XlsxProcessed / $script:sharedConvertData.XlsxTotal) * 100
                $txtXlsxCount.Text = "$($script:sharedConvertData.XlsxProcessed)/$($script:sharedConvertData.XlsxTotal)"
            }
            if ($script:sharedConvertData.PptxTotal -gt 0) {
                $pbPptx.Value = ($script:sharedConvertData.PptxProcessed / $script:sharedConvertData.PptxTotal) * 100
                $txtPptxCount.Text = "$($script:sharedConvertData.PptxProcessed)/$($script:sharedConvertData.PptxTotal)"
            }

            # Gesamt-Fortschritt
            if ($total -gt 0) {
                $pbProgress.Value = ($processed / $total) * 100
            }

            $txtStatus.Text = "Konvertiere $processed von $total Dateien... (Erfolg: $($script:sharedConvertData.SuccessCount), Fehler: $($script:sharedConvertData.FailCount))"

            # DataGrid refresh
            $dgFiles.Items.Refresh()

            # Prüfe ob alle Jobs fertig sind
            $allJobsCompleted = $true
            foreach ($job in $script:activeJobs) {
                if (-not $job.Handle.IsCompleted) {
                    $allJobsCompleted = $false
                    break
                }
            }

            if ($allJobsCompleted) {
                $script:timer.Stop()

                # Cleanup: PowerShell-Objekte beenden
                foreach ($job in $script:activeJobs) {
                    try {
                        $job.PowerShell.EndInvoke($job.Handle)
                        $job.PowerShell.Dispose()
                    } catch {}
                }

                # Runspace Pool schließen
                if ($script:runspacePool) {
                    $script:runspacePool.Close()
                    $script:runspacePool.Dispose()
                    $script:runspacePool = $null
                }

                # Berechne Gesamtdauer
                $endTime = Get-Date
                $duration = $endTime - $script:sharedConvertData.StartTime
                $script:sharedConvertData.Duration = "{0:mm\:ss}" -f $duration
                
                if ($script:sharedConvertData.Error) {
                    Complete-Reset-WithTable
                    [System.Windows.MessageBox]::Show("Fehler bei der Konvertierung: $($script:sharedConvertData.Error)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                } else {
                    $pbProgress.Value = 100
                    
                    # HTML-Bericht generieren
                    $convertedFiles = @($script:sharedConvertData.Results)
                    
                    if ($convertedFiles.Count -gt 0) {
                        try {
                            $reportPath = Generate-HTMLReport -ConvertedFiles $convertedFiles -SuccessCount $script:sharedConvertData.SuccessCount -FailCount $script:sharedConvertData.FailCount -DeletedCount $script:sharedConvertData.DeletedCount -MovedCount $script:sharedConvertData.MovedCount -Duration $script:sharedConvertData.Duration -StartTime $script:sharedConvertData.StartTime
                        } catch {
                            $reportPath = $null
                        }
                    } else {
                        $reportPath = $null
                    }
                    
                    if ($script:sharedConvertData.CancelRequested) {
                        if ($reportPath) {
                            $message = "Konvertierung wurde abgebrochen.`n`nBis zum Abbruch verarbeitet:`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)"
                            if ($script:sharedConvertData.DeletedCount -gt 0) {
                                $message += "`nQuelldateien gelöscht: $($script:sharedConvertData.DeletedCount)"
                            }
                            if ($script:sharedConvertData.MovedCount -gt 0) {
                                $message += "`nQuelldateien verschoben: $($script:sharedConvertData.MovedCount)"
                            }
                            $message += "`n`nMöchtest du den HTML-Bericht öffnen?"
                            
                            $result = [System.Windows.MessageBox]::Show(
                                $message,
                                "Konvertierung abgebrochen",
                                [System.Windows.MessageBoxButton]::YesNo,
                                [System.Windows.MessageBoxImage]::Warning
                            )
                        } else {
                            $message = "Konvertierung wurde abgebrochen.`n`nBis zum Abbruch verarbeitet:`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)"
                            if ($script:sharedConvertData.DeletedCount -gt 0) {
                                $message += "`nQuelldateien gelöscht: $($script:sharedConvertData.DeletedCount)"
                            }
                            if ($script:sharedConvertData.MovedCount -gt 0) {
                                $message += "`nQuelldateien verschoben: $($script:sharedConvertData.MovedCount)"
                            }
                            
                            $result = [System.Windows.MessageBox]::Show(
                                $message,
                                "Konvertierung abgebrochen",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Warning
                            )
                        }
                        
                        Complete-Reset-WithTable
                        
                    } else {
                        if ($reportPath) {
                            $message = "Konvertierung abgeschlossen!`n`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)"
                            if ($script:sharedConvertData.DeletedCount -gt 0) {
                                $message += "`nQuelldateien gelöscht: $($script:sharedConvertData.DeletedCount)"
                            }
                            if ($script:sharedConvertData.MovedCount -gt 0) {
                                $message += "`nQuelldateien verschoben: $($script:sharedConvertData.MovedCount)"
                            }
                            $message += "`nDauer: $($script:sharedConvertData.Duration)`n`nMöchtest du den detaillierten HTML-Bericht öffnen?"
                            
                            $result = [System.Windows.MessageBox]::Show(
                                $message,
                                "Konvertierung abgeschlossen",
                                [System.Windows.MessageBoxButton]::YesNo,
                                [System.Windows.MessageBoxImage]::Information
                            )
                        } else {
                            $message = "Konvertierung abgeschlossen!`n`nErfolgreich: $($script:sharedConvertData.SuccessCount)`nFehlgeschlagen: $($script:sharedConvertData.FailCount)"
                            if ($script:sharedConvertData.DeletedCount -gt 0) {
                                $message += "`nQuelldateien gelöscht: $($script:sharedConvertData.DeletedCount)"
                            }
                            if ($script:sharedConvertData.MovedCount -gt 0) {
                                $message += "`nQuelldateien verschoben: $($script:sharedConvertData.MovedCount)"
                            }
                            $message += "`nDauer: $($script:sharedConvertData.Duration)"
                            
                            $result = [System.Windows.MessageBox]::Show(
                                $message,
                                "Konvertierung abgeschlossen",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Information
                            )
                        }
                        
                        Complete-Reset
                    }
                    
                    if ($reportPath -and $result -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process $reportPath
                    }
                }
            }
        } catch {
            $script:timer.Stop()
            Complete-Reset-WithTable
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
