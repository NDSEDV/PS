<#
.SYNOPSIS
  SPN Manager
.DESCRIPTION
  The tool is intended to help you with your dailiy business.
.PARAMETER language
    The tool has a German edition but can also be used on English OS systems.
.NOTES
  Version:        1.0
  Author:         Jörn Walter
  Creation Date:  2025-01-04
  Purpose/Change: Initial script development

  Copyright (c) Jörn Walter. All rights reserved.
  Web: https://www.der-windows-papst.de
#>

Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SPN Manager" Height="600" Width="800">
    <Grid>
        <Label Name="DomainLabel" Content="Domain: " HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <Label Name="UserLabel" Content="User: " HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,30,0,0"/>
        <TabControl Margin="10,60,10,10">
            <TabItem Header="Lesen">
                <Grid>
                    <TextBox Name="ServerNameTextBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Height="25" Margin="10"/>
                    <Button Name="ReadButton" Content="Ermitteln" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="25" Margin="220,10,0,0"/>
                    <ListBox Name="SPNListBox" HorizontalAlignment="Left" VerticalAlignment="Top" Width="730" Height="350" Margin="10,50,10,10"/>
                </Grid>
            </TabItem>
            <TabItem Header="Löschen">
                <Grid>
                    <TextBox Name="ServerNameTextBoxDelete" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Height="25" Margin="10"/>
                    <Button Name="ReadButtonDelete" Content="Ermitteln" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="25" Margin="220,10,0,0"/>
                    <ListBox Name="SPNListBoxDelete" HorizontalAlignment="Left" VerticalAlignment="Top" Width="730" Height="350" Margin="10,50,10,10" SelectionMode="Multiple"/>
                    <Button Name="DeleteButton" Content="Löschen" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="25" Margin="10,410,0,0"/>
                </Grid>
            </TabItem>
            <TabItem Header="Hinzufügen">
                <Grid>
                    <Label Content="Servername:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                    <TextBox Name="ServerNameTextBoxAdd" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Height="25" Margin="10,30,0,0"/>
                    <Label Content="SPN:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="220,10,0,0"/>
                    <TextBox Name="SPNTextBoxAdd" HorizontalAlignment="Left" VerticalAlignment="Top" Width="200" Height="25" Margin="220,30,0,0"/>
                    <Button Name="AddButton" Content="Hinzufügen" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="25" Margin="430,30,0,0"/>
                    <ListBox Name="SPNListBoxAdd" HorizontalAlignment="Left" VerticalAlignment="Top" Width="730" Height="350" Margin="10,70,10,10"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Label Content="Copyright 2025 Jörn Walter - https://www.der-windows-papst.de" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="10,0,0,10"/>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Domänenname und Benutzername ermitteln und anzeigen
$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$window.FindName("DomainLabel").Content = "Domain: $($domain.Name)"
$window.FindName("UserLabel").Content = "User: $user"

$window.FindName("ReadButton").Add_Click({
    $serverName = $window.FindName("ServerNameTextBox").Text
    $spns = Get-SPNs -ServerName $serverName
    $window.FindName("SPNListBox").Items.Clear()
    if ($spns -eq $null) {
        $window.FindName("SPNListBox").Items.Add("Server '$serverName' nicht gefunden.")
    } else {
        foreach ($spn in $spns) {
            $window.FindName("SPNListBox").Items.Add($spn)
        }
    }
})

$window.FindName("ReadButtonDelete").Add_Click({
    $serverName = $window.FindName("ServerNameTextBoxDelete").Text
    $spns = Get-SPNs -ServerName $serverName
    $window.FindName("SPNListBoxDelete").Items.Clear()
    if ($spns -eq $null) {
        $window.FindName("SPNListBoxDelete").Items.Add("Server '$serverName' nicht gefunden.")
    } else {
        foreach ($spn in $spns) {
            $window.FindName("SPNListBoxDelete").Items.Add($spn)
        }
    }
})

$window.FindName("DeleteButton").Add_Click({
    $serverName = $window.FindName("ServerNameTextBoxDelete").Text
    $selectedSPNs = $window.FindName("SPNListBoxDelete").SelectedItems
    $currentSPNs = Get-SPNs -ServerName $serverName

    if ($currentSPNs -ne $null) {
        $spnsToRemove = @()
        foreach ($spn in $selectedSPNs) {
            if ($currentSPNs -contains $spn) {
                $spnsToRemove += $spn
            }
        }
        foreach ($spn in $spnsToRemove) {
            Remove-SPN -SPN $spn -ServerName $serverName
            $window.FindName("SPNListBoxDelete").Items.Remove($spn)
        }
    } else {
        $window.FindName("SPNListBoxDelete").Items.Add("Server '$serverName' nicht gefunden.")
    }
})

$window.FindName("AddButton").Add_Click({
    $serverName = $window.FindName("ServerNameTextBoxAdd").Text
    $spn = $window.FindName("SPNTextBoxAdd").Text
    Add-SPN -SPN $spn -ServerName $serverName
    $window.FindName("SPNListBoxAdd").Items.Add("SPN '$spn' hinzugefügt.")
})

function Get-SPNs {
    param (
        [string]$ServerName
    )
    try {
        $spns = Get-ADComputer -Identity $ServerName -Properties ServicePrincipalNames | Select-Object -ExpandProperty ServicePrincipalNames
        return $spns
    } catch {
        return $null
    }
}

function Remove-SPN {
    param (
        [string]$SPN,
        [string]$ServerName
    )
    $currentSPNs = Get-SPNs -ServerName $ServerName
    if ($currentSPNs -contains $SPN) {
        Set-ADComputer -Identity $ServerName -ServicePrincipalNames @{Remove=$SPN}
    }
}

function Add-SPN {
    param (
        [string]$SPN,
        [string]$ServerName
    )
    Set-ADComputer -Identity $ServerName -ServicePrincipalNames @{Add=$SPN}
}

$window.ShowDialog()
