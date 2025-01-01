<#
.SYNOPSIS
  Windows Event and String Searcher
.DESCRIPTION
  The tool is intended to help you with your dailiy business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.0
  Author:         Jörn Walter
  Creation Date:  2024-12-29
  Purpose/Change: Initial script development

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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Abrufen aller Protokolle unterhalb von "Anwendungs- und Dienstprotokolle"
function Get-AllLogs {
    param (
        [bool]$classicLog
    )
    $logs = Get-WinEvent -ListLog * | Where-Object { $_.IsClassicLog -eq $classicLog }
    return $logs.LogName
}

# Funktion zum Aktualisieren der ComboBox für die Protokollauswahl
function Update-LogComboBox {
    param (
        [System.Windows.Forms.ComboBox]$comboBox,
        [bool]$classicLog
    )
    $comboBox.Items.Clear()
    $comboBox.Items.AddRange((Get-AllLogs -classicLog $classicLog))
}

# Erstellen des Hauptformulars
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Event and String Searcher"
$form.Size = New-Object System.Drawing.Size(700, 590)  # Höhe angepasst, um Platz für das Copyright-Label zu schaffen
$form.StartPosition = "CenterScreen"

# Erstellen des TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(680, 500)
$form.Controls.Add($tabControl)

# Erstellen des ersten Tabs
$tabPage1 = New-Object System.Windows.Forms.TabPage
$tabPage1.Text = "String Suche"
$tabControl.Controls.Add($tabPage1)

# Erstellen der RadioButtons für die Protokollauswahl im ersten Tab
$radioButtonClassic1 = New-Object System.Windows.Forms.RadioButton
$radioButtonClassic1.Location = New-Object System.Drawing.Point(10, 10)
$radioButtonClassic1.AutoSize = $true
$radioButtonClassic1.Text = "Klassische Protokolle"
$radioButtonClassic1.Checked = $true  # Klassische Protokolle vorausgewählt
$tabPage1.Controls.Add($radioButtonClassic1)

$radioButtonModern1 = New-Object System.Windows.Forms.RadioButton
$radioButtonModern1.Location = New-Object System.Drawing.Point(160, 10)
$radioButtonModern1.AutoSize = $true
$radioButtonModern1.Text = "Moderne Protokolle"
$tabPage1.Controls.Add($radioButtonModern1)

# Erstellen der ComboBox für die Protokollauswahl im ersten Tab
$logComboBox1 = New-Object System.Windows.Forms.ComboBox
$logComboBox1.Location = New-Object System.Drawing.Point(10, 40)
$logComboBox1.Size = New-Object System.Drawing.Size(500, 20)
$logComboBox1.Items.AddRange((Get-AllLogs -classicLog $true))  # Klassische Protokolle laden
$tabPage1.Controls.Add($logComboBox1)

# Event-Handler für die RadioButtons im ersten Tab
$radioButtonClassic1.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox1 -classicLog $true
})

$radioButtonModern1.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox1 -classicLog $false
})

# Erstellen des TextBox für die Eingabe des Suchstrings im ersten Tab
$searchTextBox1 = New-Object System.Windows.Forms.TextBox
$searchTextBox1.Location = New-Object System.Drawing.Point(10, 70)
$searchTextBox1.Size = New-Object System.Drawing.Size(500, 20)
$tabPage1.Controls.Add($searchTextBox1)

# Erstellen der ComboBox für die Auswahl der Anzahl der Ergebnisse im ersten Tab
$resultCountComboBox1 = New-Object System.Windows.Forms.ComboBox
$resultCountComboBox1.Location = New-Object System.Drawing.Point(520, 70)
$resultCountComboBox1.Size = New-Object System.Drawing.Size(100, 20)
$resultCountComboBox1.Items.AddRange(@(10, 20, 50, 100, "Alles"))
$resultCountComboBox1.SelectedIndex = 0
$tabPage1.Controls.Add($resultCountComboBox1)

# Erstellen des Buttons zum Auslösen der Suche im ersten Tab
$searchButton1 = New-Object System.Windows.Forms.Button
$searchButton1.Location = New-Object System.Drawing.Point(520, 100)
$searchButton1.Size = New-Object System.Drawing.Size(100, 25)
$searchButton1.Text = "Suchen"
$tabPage1.Controls.Add($searchButton1)

# Erstellen des TextBox für die Ausgabe der Ergebnisse im ersten Tab
$outputTextBox1 = New-Object System.Windows.Forms.TextBox
$outputTextBox1.Location = New-Object System.Drawing.Point(10, 130)
$outputTextBox1.Size = New-Object System.Drawing.Size(650, 270)
$outputTextBox1.Multiline = $true
$outputTextBox1.ReadOnly = $true
$outputTextBox1.ScrollBars = "Vertical"
$tabPage1.Controls.Add($outputTextBox1)

# Erstellen des Buttons zum Exportieren der Ergebnisse im ersten Tab
$exportButton1 = New-Object System.Windows.Forms.Button
$exportButton1.Location = New-Object System.Drawing.Point(10, 410)
$exportButton1.Size = New-Object System.Drawing.Size(100, 25)
$exportButton1.Text = "Exportieren"
$tabPage1.Controls.Add($exportButton1)

# Erstellen des zweiten Tabs
$tabPage2 = New-Object System.Windows.Forms.TabPage
$tabPage2.Text = "Ereignis-ID Suche"
$tabControl.Controls.Add($tabPage2)

# Erstellen der RadioButtons für die Protokollauswahl im zweiten Tab
$radioButtonClassic2 = New-Object System.Windows.Forms.RadioButton
$radioButtonClassic2.Location = New-Object System.Drawing.Point(10, 10)
$radioButtonClassic2.AutoSize = $true
$radioButtonClassic2.Text = "Klassische Protokolle"
$radioButtonClassic2.Checked = $true  # Klassische Protokolle vorausgewählt
$tabPage2.Controls.Add($radioButtonClassic2)

$radioButtonModern2 = New-Object System.Windows.Forms.RadioButton
$radioButtonModern2.Location = New-Object System.Drawing.Point(160, 10)
$radioButtonModern2.AutoSize = $true
$radioButtonModern2.Text = "Moderne Protokolle"
$tabPage2.Controls.Add($radioButtonModern2)

# Erstellen der ComboBox für die Protokollauswahl im zweiten Tab
$logComboBox2 = New-Object System.Windows.Forms.ComboBox
$logComboBox2.Location = New-Object System.Drawing.Point(10, 40)
$logComboBox2.Size = New-Object System.Drawing.Size(500, 20)
$logComboBox2.Items.AddRange((Get-AllLogs -classicLog $true))  # Klassische Protokolle laden
$tabPage2.Controls.Add($logComboBox2)

# Event-Handler für die RadioButtons im zweiten Tab
$radioButtonClassic2.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox2 -classicLog $true
})

$radioButtonModern2.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox2 -classicLog $false
})

# Erstellen des TextBox für die Eingabe der Ereignis-ID im zweiten Tab
$eventIdTextBox = New-Object System.Windows.Forms.TextBox
$eventIdTextBox.Location = New-Object System.Drawing.Point(10, 70)
$eventIdTextBox.Size = New-Object System.Drawing.Size(500, 20)
$tabPage2.Controls.Add($eventIdTextBox)

# Erstellen der ComboBox für die Auswahl der Anzahl der Ergebnisse im zweiten Tab
$resultCountComboBox2 = New-Object System.Windows.Forms.ComboBox
$resultCountComboBox2.Location = New-Object System.Drawing.Point(520, 70)
$resultCountComboBox2.Size = New-Object System.Drawing.Size(100, 20)
$resultCountComboBox2.Items.AddRange(@(10, 20, 50, 100, "Alles"))
$resultCountComboBox2.SelectedIndex = 0
$tabPage2.Controls.Add($resultCountComboBox2)

# Erstellen des Buttons zum Auslösen der Suche im zweiten Tab
$searchButton2 = New-Object System.Windows.Forms.Button
$searchButton2.Location = New-Object System.Drawing.Point(520, 100)
$searchButton2.Size = New-Object System.Drawing.Size(100, 25)
$searchButton2.Text = "Suchen"
$tabPage2.Controls.Add($searchButton2)

# Erstellen des TextBox für die Ausgabe der Ergebnisse im zweiten Tab
$outputTextBox2 = New-Object System.Windows.Forms.TextBox
$outputTextBox2.Location = New-Object System.Drawing.Point(10, 130)
$outputTextBox2.Size = New-Object System.Drawing.Size(650, 270)
$outputTextBox2.Multiline = $true
$outputTextBox2.ReadOnly = $true
$outputTextBox2.ScrollBars = "Vertical"
$tabPage2.Controls.Add($outputTextBox2)

# Erstellen des Buttons zum Exportieren der Ergebnisse im zweiten Tab
$exportButton2 = New-Object System.Windows.Forms.Button
$exportButton2.Location = New-Object System.Drawing.Point(10, 410)
$exportButton2.Size = New-Object System.Drawing.Size(100, 25)
$exportButton2.Text = "Exportieren"
$tabPage2.Controls.Add($exportButton2)

# Erstellen des dritten Tabs
$tabPage3 = New-Object System.Windows.Forms.TabPage
$tabPage3.Text = "Zeitraum Suche"
$tabControl.Controls.Add($tabPage3)

# Erstellen der RadioButtons für die Protokollauswahl im dritten Tab
$radioButtonClassic3 = New-Object System.Windows.Forms.RadioButton
$radioButtonClassic3.Location = New-Object System.Drawing.Point(10, 10)
$radioButtonClassic3.AutoSize = $true
$radioButtonClassic3.Text = "Klassische Protokolle"
$radioButtonClassic3.Checked = $true  # Klassische Protokolle vorausgewählt
$tabPage3.Controls.Add($radioButtonClassic3)

$radioButtonModern3 = New-Object System.Windows.Forms.RadioButton
$radioButtonModern3.Location = New-Object System.Drawing.Point(160, 10)
$radioButtonModern3.AutoSize = $true
$radioButtonModern3.Text = "Moderne Protokolle"
$tabPage3.Controls.Add($radioButtonModern3)

# Erstellen der ComboBox für die Protokollauswahl im dritten Tab
$logComboBox3 = New-Object System.Windows.Forms.ComboBox
$logComboBox3.Location = New-Object System.Drawing.Point(10, 40)
$logComboBox3.Size = New-Object System.Drawing.Size(500, 20)
$logComboBox3.Items.AddRange((Get-AllLogs -classicLog $true))  # Klassische Protokolle laden
$tabPage3.Controls.Add($logComboBox3)

# Event-Handler für die RadioButtons im dritten Tab
$radioButtonClassic3.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox3 -classicLog $true
})

$radioButtonModern3.Add_CheckedChanged({
    Update-LogComboBox -comboBox $logComboBox3 -classicLog $false
})

# Erstellen des DateTimePicker für die Startzeit im dritten Tab
$startDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
$startDateTimePicker.Location = New-Object System.Drawing.Point(10, 70)
$startDateTimePicker.Size = New-Object System.Drawing.Size(200, 20)
$startDateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$startDateTimePicker.CustomFormat = "dd.MM.yyyy HH:mm:ss"
$tabPage3.Controls.Add($startDateTimePicker)

# Erstellen des DateTimePicker für die Endzeit im dritten Tab
$endDateTimePicker = New-Object System.Windows.Forms.DateTimePicker
$endDateTimePicker.Location = New-Object System.Drawing.Point(220, 70)
$endDateTimePicker.Size = New-Object System.Drawing.Size(200, 20)
$endDateTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$endDateTimePicker.CustomFormat = "dd.MM.yyyy HH:mm:ss"
$tabPage3.Controls.Add($endDateTimePicker)

# Erstellen der ComboBox für die Auswahl der Anzahl der Ergebnisse im dritten Tab
$resultCountComboBox3 = New-Object System.Windows.Forms.ComboBox
$resultCountComboBox3.Location = New-Object System.Drawing.Point(430, 70)
$resultCountComboBox3.Size = New-Object System.Drawing.Size(100, 20)
$resultCountComboBox3.Items.AddRange(@(10, 20, 50, 100, "Alles"))
$resultCountComboBox3.SelectedIndex = 0
$tabPage3.Controls.Add($resultCountComboBox3)

# Erstellen der CheckBox zum Umschalten des Datumsformats im dritten Tab
$dateFormatCheckBox = New-Object System.Windows.Forms.CheckBox
$dateFormatCheckBox.Location = New-Object System.Drawing.Point(540, 70)
$dateFormatCheckBox.Size = New-Object System.Drawing.Size(150, 20)
$dateFormatCheckBox.Text = "Englisches Format"
$tabPage3.Controls.Add($dateFormatCheckBox)

# Event-Handler für die CheckBox zum Umschalten des Datumsformats
$dateFormatCheckBox.Add_CheckedChanged({
    if ($dateFormatCheckBox.Checked) {
        $startDateTimePicker.CustomFormat = "MM/dd/yyyy HH:mm:ss"
        $endDateTimePicker.CustomFormat = "MM/dd/yyyy HH:mm:ss"
    } else {
        $startDateTimePicker.CustomFormat = "dd.MM.yyyy HH:mm:ss"
        $endDateTimePicker.CustomFormat = "dd.MM.yyyy HH:mm:ss"
    }
})

# Erstellen des Buttons zum Auslösen der Suche im dritten Tab
$searchButton3 = New-Object System.Windows.Forms.Button
$searchButton3.Location = New-Object System.Drawing.Point(540, 100)
$searchButton3.Size = New-Object System.Drawing.Size(100, 25)
$searchButton3.Text = "Suchen"
$tabPage3.Controls.Add($searchButton3)

# Erstellen des TextBox für die Ausgabe der Ergebnisse im dritten Tab
$outputTextBox3 = New-Object System.Windows.Forms.TextBox
$outputTextBox3.Location = New-Object System.Drawing.Point(10, 130)
$outputTextBox3.Size = New-Object System.Drawing.Size(650, 270)
$outputTextBox3.Multiline = $true
$outputTextBox3.ReadOnly = $true
$outputTextBox3.ScrollBars = "Vertical"
$tabPage3.Controls.Add($outputTextBox3)

# Erstellen des Buttons zum Exportieren der Ergebnisse im dritten Tab
$exportButton3 = New-Object System.Windows.Forms.Button
$exportButton3.Location = New-Object System.Drawing.Point(10, 410)
$exportButton3.Size = New-Object System.Drawing.Size(100, 25)
$exportButton3.Text = "Exportieren"
$tabPage3.Controls.Add($exportButton3)

# Erstellen des Labels für das Copyright
$copyrightLabel = New-Object System.Windows.Forms.Label
$copyrightLabel.Location = New-Object System.Drawing.Point(0, 500)
$copyrightLabel.Size = New-Object System.Drawing.Size(680, 40)
$copyrightLabel.Text = "Copyright 2025 Jörn Walter`nhttps://www.der-windows-papst.de"
$copyrightLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$form.Controls.Add($copyrightLabel)

# Erstellen des vierten Tabs
$tabPage4 = New-Object System.Windows.Forms.TabPage
$tabPage4.Text = "Protokoll und Quelle Suche"
$tabControl.Controls.Add($tabPage4)

# Erstellen der ComboBox für die Protokollauswahl im vierten Tab
$logComboBox4 = New-Object System.Windows.Forms.ComboBox
$logComboBox4.Location = New-Object System.Drawing.Point(10, 10)
$logComboBox4.Size = New-Object System.Drawing.Size(200, 20)
$logComboBox4.Items.AddRange(@("Anwendung", "Sicherheit", "System"))
$tabPage4.Controls.Add($logComboBox4)

# Erstellen der ComboBox für die Quellenauswahl im vierten Tab
$sourceComboBox4 = New-Object System.Windows.Forms.ComboBox
$sourceComboBox4.Location = New-Object System.Drawing.Point(10, 40)
$sourceComboBox4.Size = New-Object System.Drawing.Size(500, 20)
$tabPage4.Controls.Add($sourceComboBox4)

# Erstellen der ComboBox für die Auswahl der Anzahl der Ergebnisse im vierten Tab
$resultCountComboBox4 = New-Object System.Windows.Forms.ComboBox
$resultCountComboBox4.Location = New-Object System.Drawing.Point(520, 40)
$resultCountComboBox4.Size = New-Object System.Drawing.Size(100, 20)
$resultCountComboBox4.Items.AddRange(@(10, 20, 50, 100, "Alles"))
$resultCountComboBox4.SelectedIndex = 0
$tabPage4.Controls.Add($resultCountComboBox4)

# Erstellen des Buttons zum Auslösen der Suche im vierten Tab
$searchButton4 = New-Object System.Windows.Forms.Button
$searchButton4.Location = New-Object System.Drawing.Point(520, 70)
$searchButton4.Size = New-Object System.Drawing.Size(100, 25)
$searchButton4.Text = "Suchen"
$tabPage4.Controls.Add($searchButton4)

# Erstellen des Buttons zum Exportieren der Ergebnisse im vierten Tab
$exportButton4 = New-Object System.Windows.Forms.Button
$exportButton4.Location = New-Object System.Drawing.Point(10, 410)
$exportButton4.Size = New-Object System.Drawing.Size(100, 25)
$exportButton4.Text = "Exportieren"
$tabPage4.Controls.Add($exportButton4)

# Erstellen des TextBox für die Ausgabe der Ergebnisse im vierten Tab
$outputTextBox4 = New-Object System.Windows.Forms.TextBox
$outputTextBox4.Location = New-Object System.Drawing.Point(10, 100)
$outputTextBox4.Size = New-Object System.Drawing.Size(650, 300)
$outputTextBox4.Multiline = $true
$outputTextBox4.ReadOnly = $true
$outputTextBox4.ScrollBars = "Vertical"
$tabPage4.Controls.Add($outputTextBox4)

# Event-Handler für den Such-Button im ersten Tab
$searchButton1.Add_Click({
    $selectedLog = $logComboBox1.SelectedItem
    $searchString = $searchTextBox1.Text
    $resultCount = $resultCountComboBox1.SelectedItem

    if ($selectedLog -and $searchString) {
        $outputTextBox1.Text = "Bitte warten...keine roten Flecken bekommen"
        $outputTextBox1.Refresh()  # Stellt sicher, dass die Nachricht sofort angezeigt wird

        $logName = $selectedLog

        $job = Start-Job -ScriptBlock {
            param($logName, $searchString, $resultCount)
            if ($using:resultCount -eq "Alles") {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.Message -like "*$using:searchString*" }
                $outputArray = @()
                if ($events) {
                    foreach ($event in $events) {
                        $outputArray += "$($event.TimeCreated) - $($event.Message)"
                    }
                }
                if ($outputArray.Count -eq 0) {
                    $outputArray += "Keine Ergebnisse gefunden."
                }
                return $outputArray
            } else {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.Message -like "*$using:searchString*" }
                $output = ""
                $count = [int]$using:resultCount
                if ($events) {
                    $events = $events | Select-Object -First $count
                    foreach ($event in $events) {
                        $output += "$($event.TimeCreated) - $($event.Message)`n`n"
                    }
                }
                if ($output -eq "") {
                    $output = "Keine Ergebnisse gefunden."
                }
                return $output
            }
        } -ArgumentList $logName, $searchString, $resultCount

        $job | Wait-Job
        $output = Receive-Job -Job $job -Keep
        Remove-Job -Job $job

        if ($resultCount -eq "Alles") {
            $outputTextBox1.Text = "Bitte exportieren Sie die Ergebnisse."
            $global:allEvents = $output
        } else {
            $outputTextBox1.Text = $output
        }
    } else {
        $outputTextBox1.Text = "Bitte wählen Sie ein Protokoll und geben Sie einen Suchstring ein."
    }
})

# Event-Handler für den Such-Button im zweiten Tab
$searchButton2.Add_Click({
    $selectedLog = $logComboBox2.SelectedItem
    $eventId = $eventIdTextBox.Text
    $resultCount = $resultCountComboBox2.SelectedItem

    if ($selectedLog -and $eventId) {
        $outputTextBox2.Text = "Bitte warten...keine roten Flecken bekommen"
        $outputTextBox2.Refresh()  # Stellt sicher, dass die Nachricht sofort angezeigt wird

        $logName = $selectedLog

        $job = Start-Job -ScriptBlock {
            param($logName, $eventId, $resultCount)
            if ($using:resultCount -eq "Alles") {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.Id -eq $using:eventId }
                $outputArray = @()
                if ($events) {
                    foreach ($event in $events) {
                        $outputArray += "$($event.TimeCreated) - $($event.Message)"
                    }
                }
                if ($outputArray.Count -eq 0) {
                    $outputArray += "Keine Ergebnisse gefunden."
                }
                return $outputArray
            } else {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.Id -eq $using:eventId }
                $output = ""
                $count = [int]$using:resultCount
                if ($events) {
                    $events = $events | Select-Object -First $count
                    foreach ($event in $events) {
                        $output += "$($event.TimeCreated) - $($event.Message)`n`n"
                    }
                }
                if ($output -eq "") {
                    $output = "Keine Ergebnisse gefunden."
                }
                return $output
            }
        } -ArgumentList $logName, [int]$eventId, $resultCount

        $job | Wait-Job
        $output = Receive-Job -Job $job -Keep
        Remove-Job -Job $job

        if ($resultCount -eq "Alles") {
            $outputTextBox2.Text = "Bitte exportieren Sie die Ergebnisse."
            $global:allEvents = $output
        } else {
            $outputTextBox2.Text = $output
        }
    } else {
        $outputTextBox2.Text = "Bitte wählen Sie ein Protokoll und geben Sie eine Ereignis-ID ein."
    }
})

# Event-Handler für den Such-Button im dritten Tab
$searchButton3.Add_Click({
    $selectedLog = $logComboBox3.SelectedItem
    $startTime = $startDateTimePicker.Value
    $endTime = $endDateTimePicker.Value
    $resultCount = $resultCountComboBox3.SelectedItem

    if ($selectedLog -and $startTime -and $endTime -and $resultCount) {
        $outputTextBox3.Text = "Bitte warten...keine roten Flecken bekommen"
        $outputTextBox3.Refresh()  # Stellt sicher, dass die Nachricht sofort angezeigt wird

        $logName = $selectedLog

        $job = Start-Job -ScriptBlock {
            param($logName, $startTime, $endTime, $resultCount)
            if ($using:resultCount -eq "Alles") {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.TimeCreated -ge $using:startTime -and $_.TimeCreated -le $using:endTime }
                $outputArray = @()
                if ($events) {
                    foreach ($event in $events) {
                        $outputArray += "$($event.TimeCreated) - $($event.Message)"
                    }
                }
                if ($outputArray.Count -eq 0) {
                    $outputArray += "Keine Ergebnisse gefunden."
                }
                return $outputArray
            } else {
                $events = Get-WinEvent -LogName $using:logName | Where-Object { $_.TimeCreated -ge $using:startTime -and $_.TimeCreated -le $using:endTime }
                $output = ""
                $count = [int]$using:resultCount
                if ($events) {
                    $events = $events | Select-Object -First $count
                    foreach ($event in $events) {
                        $output += "$($event.TimeCreated) - $($event.Message)`n`n"
                    }
                }
                if ($output -eq "") {
                    $output = "Keine Ergebnisse gefunden."
                }
                return $output
            }
        } -ArgumentList $logName, $startTime, $endTime, $resultCount

        $job | Wait-Job
        $output = Receive-Job -Job $job -Keep
        Remove-Job -Job $job

        if ($resultCount -eq "Alles") {
            $outputTextBox3.Text = "Bitte exportieren Sie die Ergebnisse."
            $global:allEvents = $output
        } else {
            $outputTextBox3.Text = $output
        }
    } else {
        $outputTextBox3.Text = "Bitte wählen Sie ein Protokoll und geben Sie Start- und Endzeit ein."
    }
})

# Event-Handler für den Export-Button im ersten Tab
$exportButton1.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Ergebnisse exportieren"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($filePath)

        if ($fileExtension -eq ".txt") {
            if ($global:allEvents) {
                $global:allEvents -join "`n" | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $outputTextBox1.Text | Out-File -FilePath $filePath -Encoding utf8
            }
        } elseif ($fileExtension -eq ".html") {
            if ($global:allEvents) {
                $htmlContent = "<html><body><pre>"
                foreach ($event in $global:allEvents) {
                    $htmlContent += "$event`n"
                }
                $htmlContent += "</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $htmlContent = "<html><body><pre>$($outputTextBox1.Text)</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            }
        }
    }
})

# Event-Handler für den Export-Button im zweiten Tab
$exportButton2.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Ergebnisse exportieren"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($filePath)

        if ($fileExtension -eq ".txt") {
            if ($global:allEvents) {
                $global:allEvents -join "`n" | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $outputTextBox2.Text | Out-File -FilePath $filePath -Encoding utf8
            }
        } elseif ($fileExtension -eq ".html") {
            if ($global:allEvents) {
                $htmlContent = "<html><body><pre>"
                foreach ($event in $global:allEvents) {
                    $htmlContent += "$event`n"
                }
                $htmlContent += "</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $htmlContent = "<html><body><pre>$($outputTextBox2.Text)</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            }
        }
    }
})

# Event-Handler für den Export-Button im dritten Tab
$exportButton3.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Ergebnisse exportieren"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($filePath)

        if ($fileExtension -eq ".txt") {
            if ($global:allEvents) {
                $global:allEvents -join "`n" | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $outputTextBox3.Text | Out-File -FilePath $filePath -Encoding utf8
            }
        } elseif ($fileExtension -eq ".html") {
            if ($global:allEvents) {
                $htmlContent = "<html><body><pre>"
                foreach ($event in $global:allEvents) {
                    $htmlContent += "$event`n"
                }
                $htmlContent += "</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $htmlContent = "<html><body><pre>$($outputTextBox3.Text)</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            }
        }
    }
})

# Funktion zum Abrufen aller Quellen eines bestimmten Protokolls
function Get-LogSources {
    param (
        [string]$logName
    )
    $sources = Get-WinEvent -ListLog $logName | Select-Object -ExpandProperty ProviderNames
    return $sources
}

# Funktion zum Aktualisieren der ComboBox für die Quellenauswahl
function Update-SourceComboBox {
    param (
        [System.Windows.Forms.ComboBox]$comboBox,
        [string]$logName
    )
    $comboBox.Items.Clear()
    $comboBox.Items.AddRange((Get-LogSources -logName $logName))
}

# Event-Handler für die Protokollauswahl im vierten Tab
$logComboBox4.Add_SelectedIndexChanged({
    $selectedLog = $logComboBox4.SelectedItem
    if ($selectedLog) {
        # Übersetzen der Protokollnamen in die englischen Namen
        $logNameMapping = @{
            "Anwendung" = "Application"
            "Sicherheit" = "Security"
            "System" = "System"
        }
        $logName = $logNameMapping[$selectedLog]
        Update-SourceComboBox -comboBox $sourceComboBox4 -logName $logName
    }
})

$tabPage4.Controls.Add($outputTextBox4)

# Event-Handler für den Such-Button im vierten Tab
$searchButton4.Add_Click({
    $selectedLog = $logComboBox4.SelectedItem
    $selectedSource = $sourceComboBox4.SelectedItem
    $resultCount = $resultCountComboBox4.SelectedItem

    if ($selectedLog -and $selectedSource) {
        # Übersetzen der Protokollnamen in die englischen Namen
        $logNameMapping = @{
            "Anwendung" = "Application"
            "Sicherheit" = "Security"
            "System" = "System"
        }
        $logName = $logNameMapping[$selectedLog]
        $sourceName = $selectedSource

        $outputTextBox4.Text = "Bitte warten...keine roten Flecken bekommen"
        $outputTextBox4.Refresh()  # Stellt sicher, dass die Nachricht sofort angezeigt wird

        $job = Start-Job -ScriptBlock {
            param($logName, $sourceName, $resultCount)
            if ($using:resultCount -eq "Alles") {
                $events = Get-WinEvent -LogName $using:logName
                $filteredEvents = $events | Where-Object { $_.ProviderName -eq $using:sourceName }
                $outputArray = @()
                if ($filteredEvents) {
                    foreach ($event in $filteredEvents) {
                        $outputArray += "$($event.TimeCreated) - $($event.Message)"
                    }
                }
                if ($outputArray.Count -eq 0) {
                    $outputArray += "Keine Ergebnisse gefunden."
                }
                return $outputArray
            } else {
                $events = Get-WinEvent -LogName $using:logName
                $filteredEvents = $events | Where-Object { $_.ProviderName -eq $using:sourceName }
                $outputArray = @()
                $count = [int]$using:resultCount
                if ($filteredEvents) {
                    $filteredEvents = $filteredEvents | Sort-Object TimeCreated -Descending | Select-Object -First $count
                    foreach ($event in $filteredEvents) {
                        $outputArray += "$($event.TimeCreated) - $($event.Message)"
                    }
                }
                if ($outputArray.Count -eq 0) {
                    $outputArray += "Keine Ergebnisse gefunden."
                }
                return $outputArray
            }
        } -ArgumentList $logName, $sourceName, $resultCount

        $job | Wait-Job
        $output = Receive-Job -Job $job -Keep
        Remove-Job -Job $job

        if ($resultCount -eq "Alles") {
            $outputTextBox4.Text = "Bitte exportieren Sie die Ergebnisse."
            $global:allEvents = $output
        } else {
            $outputTextBox4.Text = $output -join "`n`n"
        }
    } else {
        $outputTextBox4.Text = "Bitte wählen Sie ein Protokoll und eine Quelle aus."
    }
})

# Event-Handler für den Export-Button im vierten Tab
$exportButton4.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Ergebnisse exportieren"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        $fileExtension = [System.IO.Path]::GetExtension($filePath)

        if ($fileExtension -eq ".txt") {
            if ($global:allEvents) {
                $global:allEvents -join "`n" | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $outputTextBox4.Text | Out-File -FilePath $filePath -Encoding utf8
            }
        } elseif ($fileExtension -eq ".html") {
            if ($global:allEvents) {
                $htmlContent = "<html><body><pre>"
                foreach ($event in $global:allEvents) {
                    $htmlContent += "$event`n"
                }
                $htmlContent += "</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            } else {
                $htmlContent = "<html><body><pre>$($outputTextBox4.Text)</pre></body></html>"
                $htmlContent | Out-File -FilePath $filePath -Encoding utf8
            }
        }
    }
})

# Anzeigen des Formulars
$form.ShowDialog()
