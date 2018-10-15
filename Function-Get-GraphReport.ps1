

Function Get-GraphReport
{
param(
[parameter(Mandatory=$true,HelpMessage="The report to download.")]
[validateset('getEmailActivityUserDetail','getEmailActivityCounts','getEmailActivityUserCounts','getEmailAppUsageUserDetail','getEmailAppUsageAppsUserCounts','getEmailAppUsageUserCounts','getEmailAppUsageVersionsUserCounts','getMailboxUsageDetail','getMailboxUsageMailboxCounts','getMailboxUsageQuotaStatusMailboxCounts','getMailboxUsageStorage','getOffice365ActivationsUserDetail','getOffice365ActivationCounts','getOffice365ActivationsUserCounts','getOffice365ActiveUserDetail','getOffice365ActiveUserCounts','getOffice365ServicesUserCounts','getOffice365GroupsActivityDetail','getOffice365GroupsActivityCounts','getOffice365GroupsActivityGroupCounts','getOffice365GroupsActivityStorage','getOffice365GroupsActivityFileCounts','getOneDriveActivityUserDetail','getOneDriveActivityUserCounts','getOneDriveActivityFileCounts','getOneDriveUsageAccountDetail','getOneDriveUsageAccountCounts','getOneDriveUsageFileCounts','getOneDriveUsageStorage','getSharePointActivityUserDetail','getSharePointActivityFileCounts','getSharePointActivityUserCounts','getSharePointActivityPages','getSharePointSiteUsageDetail','getSharePointSiteUsageFileCounts','getSharePointSiteUsageSiteCounts','getSharePointSiteUsageStorage','getSharePointSiteUsagePages','getSkypeForBusinessActivityUserDetail','getSkypeForBusinessActivityCounts','getSkypeForBusinessActivityUserCounts','getSkypeForBusinessDeviceUsageUserDetail','getSkypeForBusinessDeviceUsageDistributionUserCounts','getSkypeForBusinessDeviceUsageUserCounts','getSkypeForBusinessOrganizerActivityCounts','getSkypeForBusinessOrganizerActivityUserCounts','getSkypeForBusinessOrganizerActivityMinuteCounts','getSkypeForBusinessParticipantActivityCounts','getSkypeForBusinessParticipantActivityUserCounts','getSkypeForBusinessParticipantActivityMinuteCounts','getSkypeForBusinessPeerToPeerActivityCounts','getSkypeForBusinessPeerToPeerActivityUserCounts','getSkypeForBusinessPeerToPeerActivityMinuteCounts','getYammerActivityUserDetail','getYammerActivityCounts','getYammerActivityUserCounts','getYammerDeviceUsageUserDetail','getYammerDeviceUsageDistributionUserCounts','getYammerDeviceUsageUserCounts','getYammerGroupsActivityDetail','getYammerGroupsActivityGroupCounts','getYammerGroupsActivityCounts')]
[string]$ReportName,
[parameter(Mandatory=$false,HelpMessage="Report period to download. Default is 30 Days.")]
[validateset('7 Days','30 Days','90 Days','180 Days')]
[string]$ReportPeriod="30 Days"

)

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
    $report = (Invoke-RestMethod -Headers $GraphHeader -Uri $Uri -UseBasicParsing -Method "GET").Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries)
    
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

return $ReportDT
}