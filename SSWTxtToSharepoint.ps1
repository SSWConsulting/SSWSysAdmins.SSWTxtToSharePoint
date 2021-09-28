<#
.SYNOPSIS
    This PowerShell script copies entries from a .txt file to SharePoint.
.DESCRIPTION
    This PowerShell script copies entries from a .txt file to SharePoint. At SSW.com.au, it is used to copy entries from our Backup Script to a SharePoint list, for easier viewing.
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    Created by Kaique "Kiki" Biancatti for SSW - https://www.ssw.com.au/people/kiki
#>

# Importing the configuration file
$config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1

# Creating variables to determine magic strings and getting them from the configuration file
$LogFile = $config.LogFile
$OriginEmail = $config.OriginEmail
$TargetEmail = $config.TargetEmail
$LogModuleLocation = $config.LogModuleLocation
$SharePointLocation = $config.SharePointLocation
$CredKey = "$PSScriptRoot/$($config.CredKey)"
$CredUser = $config.CredUser
$CredPass = "$PSScriptRoot/$($config.CredPass)"
$SharePointListID = $config.SharePointListID
$UserLogsLocation = $config.UserLogsLocation

# Importing the necessary modules
Import-Module -Name $LogModuleLocation
Import-Module pnp.powershell

# Connecting to PnP Online
$Key = Get-Content $CredKey
$MyCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $CredUser, (Get-Content $CredPass | ConvertTo-SecureString -Key $key)        
Connect-PnPOnline -Url $SharePointLocation -Credentials $MyCredential

<#
.SYNOPSIS
Gets the data from the logfiles and sorts and structures them.

.DESCRIPTION
Gets the data from the logfiles and sorts and structures them.
Uses a new PSCustomObject to create a table with proper properties.

.EXAMPLE
PS> $MainList = Get-LogFileEntries
#>
function Get-LogFileEntries {
    # Creating my array that will be used in a second
    $organizedArray = @()

    # Getting the data from the on-premises log
    $data = get-content $UserLogsLocation

    # Manipulating the data to create a custom object with the correct title, PC and date
    $data | ForEach-Object {
        $items = $_.split(" ")
        $organizedArray += new-object psobject -property @{Date = $items[0] + " " + $items[1]; PC = $items[2]; Title = $items[3]; } }
    $organizedArray | Group-Object "Title", "PC" | foreach {
        $_.Group | Select "Title", "PC", "Date" -Last 1 } | Sort-Object -Property { $_.Date -as [datetime] }
}
<#
.SYNOPSIS
Function to delete all entries from a SharePoint list and add again, based on a list.

.DESCRIPTION
Function to delete all entries from a SharePoint list and add again, based on a list.

.PARAMETER ListID
The Unique IS of the list in SharePoint.

.PARAMETER List
The actual collection of objects that will be added to the list.

.EXAMPLE
PS> Set-SharePointListEntries -List $MainList -ListID $SharePointListID
#>
function Set-SharePointListEntries {
    [CmdletBinding()]
    param (
        [Parameter()]
        $ListID,
        [Parameter()]
        $List
    )
    $items = Get-PnPListItem -List $ListID -PageSize 1000000
    foreach ($item in $items) {
        try {
            Remove-PnPListItem -List $ListID -Identity $item.Id -Force
        }
        catch {
            Write-Host "Error Occurred While Deleting the Item from the SharePoint Online List"
        }
    }

    # Adding 10 hours to the time so SharePoint is happy - there must be a way to do this better but no time to investigate further now
    $HoursToAdd = New-TimeSpan -Hours 10
    $List | foreach {
        $CorrectDate = [Datetime]::ParseExact($_.Date, 'dd/MM/yyyy HH:mm:ss', $null)
        $CorrectDate = $CorrectDate - $HoursToAdd
        $CorrectDate = $CorrectDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        Add-PnPListItem -list $ListID -Values @{"Title" = $_.Title; "DateLastRun" = $CorrectDate; "PCName" = $_.PC }
    }    
}

# Calling all the functions
$MainList = Get-LogFileEntries
Set-SharePointListEntries -List $MainList -ListID $SharePointListID


