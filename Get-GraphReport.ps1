﻿<#
Get-GraphReport

This script retrieves a specified Graph report for Microsoft Office 365.

To run this script you must first setup the LongView-ReportingApp App Registration in your tenant.

Author: Clark B. Lebarge
Company: Long View Systems
Version: 1.0.beta.10152018

#>

param(
[parameter(Mandatory=$true,HelpMessage="The report to download.")]
[validateset('getEmailActivityUserDetail','getEmailActivityCounts','getEmailActivityUserCounts','getEmailAppUsageUserDetail','getEmailAppUsageAppsUserCounts','getEmailAppUsageUserCounts','getEmailAppUsageVersionsUserCounts','getMailboxUsageDetail','getMailboxUsageMailboxCounts','getMailboxUsageQuotaStatusMailboxCounts','getMailboxUsageStorage','getOffice365ActivationsUserDetail','getOffice365ActivationCounts','getOffice365ActivationsUserCounts','getOffice365ActiveUserDetail','getOffice365ActiveUserCounts','getOffice365ServicesUserCounts','getOffice365GroupsActivityDetail','getOffice365GroupsActivityCounts','getOffice365GroupsActivityGroupCounts','getOffice365GroupsActivityStorage','getOffice365GroupsActivityFileCounts','getOneDriveActivityUserDetail','getOneDriveActivityUserCounts','getOneDriveActivityFileCounts','getOneDriveUsageAccountDetail','getOneDriveUsageAccountCounts','getOneDriveUsageFileCounts','getOneDriveUsageStorage','getSharePointActivityUserDetail','getSharePointActivityFileCounts','getSharePointActivityUserCounts','getSharePointActivityPages','getSharePointSiteUsageDetail','getSharePointSiteUsageFileCounts','getSharePointSiteUsageSiteCounts','getSharePointSiteUsageStorage','getSharePointSiteUsagePages','getSkypeForBusinessActivityUserDetail','getSkypeForBusinessActivityCounts','getSkypeForBusinessActivityUserCounts','getSkypeForBusinessDeviceUsageUserDetail','getSkypeForBusinessDeviceUsageDistributionUserCounts','getSkypeForBusinessDeviceUsageUserCounts','getSkypeForBusinessOrganizerActivityCounts','getSkypeForBusinessOrganizerActivityUserCounts','getSkypeForBusinessOrganizerActivityMinuteCounts','getSkypeForBusinessParticipantActivityCounts','getSkypeForBusinessParticipantActivityUserCounts','getSkypeForBusinessParticipantActivityMinuteCounts','getSkypeForBusinessPeerToPeerActivityCounts','getSkypeForBusinessPeerToPeerActivityUserCounts','getSkypeForBusinessPeerToPeerActivityMinuteCounts','getYammerActivityUserDetail','getYammerActivityCounts','getYammerActivityUserCounts','getYammerDeviceUsageUserDetail','getYammerDeviceUsageDistributionUserCounts','getYammerDeviceUsageUserCounts','getYammerGroupsActivityDetail','getYammerGroupsActivityGroupCounts','getYammerGroupsActivityCounts','getTeamsDeviceUsageUserDetail','getTeamsDeviceUsageUserCounts','getTeamsDeviceUsageDistributionUserCounts','getTeamsUserActivityUserDetail','getTeamsUserActivityCounts','getTeamsUserActivityUserCounts')]
[string]$ReportName,
[parameter(Mandatory=$false,HelpMessage="Report period to download. Default is 30 Days.")]
[validateset('7 Days','30 Days','90 Days','180 Days')]
[string]$ReportPeriod="30 Days",
[parameter(Mandatory=$false,HelpMessage="The format of the output file. Excel or CSV. Excel requires Import-Excel module. Default is CSV.")]
[validateset('CSV','XLSX')]
[string]$Format,
[parameter(Mandatory=$false,HelpMessage="Specify an alternate folder for the output file. Default is current folder.")]
[string]$Path
)

#Begin Functions
Function Connect-WebApp
{
#This function connects to an already configured WebApp using stored certificate. The certificate name is intended to match the Tenant ID.
param(
[parameter()]$TenantID,
[parameter()]$AppId,
[parameter()]$resourceAppIdURI
)
    $clientCertificate = Get-ChildItem -Path Cert:CurrentUser\My | where {$_.Subject -match $TenantId}
    $certThumbprint = $clientCertificate.Thumbprint
    $AADApp = Connect-AzureAD -CertificateThumbprint $certThumbprint -TenantId $TenantID -ApplicationId $AppId
    $TenantName = $AADApp.TenantDomain
    $CredPrompt         = "Auto"
    $authority          = "https://login.microsoftonline.com/$TenantName"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
    $certificateCredential = New-Object -TypeName Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate -ArgumentList ($AppId, $clientCertificate)
    $AccessToken        = $authContext.AcquireTokenAsync($resourceAppIdURI,$certificateCredential).Result
    $WebAppHeader = @{
        'Authorization' = $AccessToken.CreateAuthorizationHeader()
        }
return $WebAppHeader
}

try{

#If we're already logged on, this should work.
$AppInfo = Get-AzureADApplication -SearchString "LongView-ReportingApp"
}
catch{


#Login to Azure AD, then get the app.
    #storing this as a global, so it can be found by repeated runs of the app.
    $global:AzureAD = Connect-AzureAD
    $AppInfo = Get-AzureADApplication -SearchString "LongView-ReportingApp"
}


#The connection requires the Application's ID and Redirect URI.
    
    $AppId = $AppInfo.AppId
    $RedirectUri = $AppInfo.ReplyUrls.Item(0)
    $TenantName = $AzureAD.TenantDomain
    $TenantID = $AzureAD.TenantId

#Now lets connect to Graph with our function.
    $resourceAppIdURI   = "https://graph.microsoft.com"
    $WebAppHeader = Connect-WebApp -TenantID $TenantID -AppId $AppId -resourceAppIdURI $resourceAppIdURI

#Build the URI for the Report.
    $Periods = $null
    $Periods = @{
    '7 Days' = 'D7'
    '30 Days' = 'D30'
    '90 Days' = 'D90'
    '180 days' = 'D180'
    }
    $Period = $Periods.$ReportPeriod

    #Some Reports don't have Period, so here's an exeception.
    IF($ReportName -like "getOffice365Activation*")
    {
    $URI = "$resourceAppIdURI/v1.0/Reports/$ReportName"
    }
    ELSE
    {
    $URI = "$resourceAppIdURI/v1.0/Reports/$ReportName(Period='$Period')"
    }
#Get the Report.
    
    #While the report is in CSV format, the manner in which PowerShell downloads the report results in the CSV data being placed in a single object string.
    #Need to split it up to make it usable within PowerShell.
    #Split each separate line, removing empty lines if they occur. 
    $report = (Invoke-RestMethod -Headers $WebAppHeader -Uri $Uri -UseBasicParsing -Method "GET").Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries)
    
    #Create the datatable required to store the converted objects.
    $Header = ($Report.GetValue(0).Split(','))
    $Columns = ($Header | Measure).Count
    $ColumnNames = ($header | Select-Object @{Name='ColumnName';Expression={$_}}).ColumnName.Split(',')
    $ReportDT = New-Object System.Data.DataTable("ReportDT")
    foreach($Column in $ColumnNames)
    {
        $ReportDT.Columns.Add($Column) | Out-Null
    }

    #Add the values to the DataTable and Export to the file format selected.
    $LineCount = 0
    foreach($line in $Report)
    {
        $Row = $ReportDT.NewRow()
        #Skip the header line.
        IF($LineCount -gt 0)
        {
            $linedata = (($Report.GetValue($LineCount).Split(',')) | Select-Object @{Name='LineInfo';Expression={$_}}).LineInfo.Split(',')
            foreach($Column in $ColumnNames)
            {
                $Position = $ReportDT.Columns.Item($Column).Ordinal
                $Row[$Column] = $linedata.GetValue($Position)
            }
            $ReportDT.Rows.Add($Row) | Out-Null
        }
        $LineCount = $LineCount + 1
    }

#Remove the Report Refresh Date Column as it doesn't add anything to the output.
$ReportDT.Columns.Remove('ï»¿Report Refresh Date')

#Output the Report to the format and file path selected.

    $CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
    #Date
    $Datefield = Get-Date -Format {MMddyyyy}

    IF(!$Format){$Format="CSV"}
    IF(!$Path){$Path=$CurrentPath}
            $Incr = $null
            $testpath = $null
            DO{

            $OutFilePath = "$Path\$TenantName-$ReportName-for last -$ReportPeriod-$Incr.$Format"
   
            $Incr = $Incr + 1
            $testpath = Test-Path $OutFilePath

            }
            UNTIL($testpath -eq $false)

    #Exporting
    
    IF($Format -eq "CSV")
    {
        $ReportDT | Export-Csv -Path $OutFilePath -NoTypeInformation
        notepad $OutFilePath
    }
    ELSE
    {
        $ReportDT | Export-Excel -Path $OutFilePath -WorksheetName $ReportName -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter -Show -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors
    }