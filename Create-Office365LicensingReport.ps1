<#
Create-Office365LicensingReport.ps1

This script is based on the Get-UserAndMailboxStatistics script, with the focus on retreiving, analyzing, and generating a report on all license usage in an Office 365 tenancy.
This script does not include functionality for detailed service usage statistics.

This script has the following requirements:

    - Execution is set to RemoteSigned.
        * Requires Admin elevation to perform.
        * Set-ExecutionPolicy -ExecutionPolicy RemoteSigned.
    - Microsoft Online Services, MSOnline module.
        * Requires Admin elevation to perform.
        * Install-Module MSOnline.
        * Note: Currently the Azure AD 2.0 module, which is intended to replace MSOnline does not seem to support cloud service providers. Which is why the older commands are used.
    - Import-Excel
        *Install Module ImportExcel




Create-Office365LicensingReport.ps1
    [-CSPAllDomains] <switch>
    [-CSPCustomerDomain] <string>
    [-ReportType] <listoption> Summary | Full
    [-Path] <string>


    Version: 1.0.beta.10122018
    Author: Clark B. Lebarge
    Company: Long View Systems

#>

#Script Input Parameters
param(
[parameter(ParameterSetName="CSPD",Mandatory=$true,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[string]$CSPCustomerDomain,
[parameter(ParameterSetName="CSPA",Mandatory=$true,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[switch]$CSPAll,
[parameter(ParameterSetName="CSPD",Mandatory=$false,HelpMessage="The folder path for the output Excel file. Current Directory if ommitted.")]
[parameter(ParameterSetName="CSPA",Mandatory=$false,HelpMessage="The folder path for the output Excel file. Current Directory if ommitted.")]
[parameter(ParameterSetName="NoCSP",Mandatory=$false,HelpMessage="The folder path for the output Excel file. Current Directory if ommitted.")]
[string]$Path,
[parameter(ParameterSetName="CSPD",Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[parameter(ParameterSetName="CSPA",Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[parameter(ParameterSetName="NoCSP",Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[string]$ReportType
)

#End Script Parameters

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

return $ReportDT
}

#End Functions

#Script Variables, these are declared and set here to ensure the values are not contaminated from existing PS variables from earlier runs of the script.
#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}
$TenantInfo = $null
$CompanyName = $null
$TenantId = $null
$MSOLUsers = $null
$MSOLUserCount = $null

#End Script Variables

#Script Body

    #Part 1, Login to MSOL, import SKU details, and set scope for CSP usage.
        #This section completes quickly, so no status information is displayed from this part of the script.

        #Import the SKU Details from Excel File
        #This is currently a local file, the data will need to be hosted somewhere or the file referenced included in the final version.
        #I'm thinking an XML file in Azure storage would work.

        $Plans = Import-Excel -Path C:\working\SKUDetails.xlsx

        #Logon to Microsoft Azure AD, with MSOL package.
        #We use MSOL package as the Azure AD 2.0 package still does not support CSP.

        Connect-MsolService

        #Scope for CSP can be all or domain, for normal Office 365 usage scope is always the current tenancy.
        IF($CSPAll)
        {
        $TenantInfo = Get-MsolPartnerContract | select Name,TenantId,DefaultDomainName

        $TenantCount = ($TenantInfo | measure).Count
        }
        ELSE
        {
    
            IF($CSPCustomerDomain)
            {
            $TenantInfo = Get-MsolPartnerContract -DomainName $CSPCustomerDomain | select Name,TenantId,DefaultDomainName
            $TenantCount = 1
            }
            ELSE
            {
            $TenantInfo = Get-MsolCompanyInformation | select @{N='Name';E={$_.DisplayName}}
            $TenantCount = 1
            }
        }

        #This is a hash table of known resellers and their partner IDs for output on the Subscription Details.
        
        $Partners = $null
        $Partners = @{
            'ac07b96d-b94d-4830-90c3-361acdfb17d7' = 'Long View Systems'
            '0c5f32df-56ac-40ff-b99c-e7b918d4ac16' = 'F12.net Inc. - Old'
            'db73bd59-4978-4089-8693-1f5749dc4e57' = 'KPMG Adoxio'
            '813fb400-867f-48d6-8dce-3e3d48e2e0fe' = 'Polycom'
            'f53a30b2-d978-4df8-a063-7f1cd3f44899' = 'UXC Eclipse Canada (CSP)'
        }
    
    #Part 2, Creation of the report for each tenant.
        #This is accomplished in a loop process for CSP usage.
        
        $TenantInfo | ForEach-Object -Begin {$I=0} -Process {
            $Tenant = $_

            #Part 2, Step 1, query Office 365 for user properties and save information on licensing to a variable.

            #Set Variables to tenant info.
            $CompanyName = $Tenant.Name
            $TenantId = $Tenant.TenantId
            $PrimaryDomain = $Tenant.DefaultDomainName

           
            #For CSP usage the tenant ID is needed in the command.
            IF($TenantID)
            {
                $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId -ErrorAction SilentlyContinue | Select DisplayName,BlockCredential,Licenses,UserPrincipalName,LastDirSyncTime
                $MSOLAccountSkus = Get-MsolAccountSku -TenantId $tenantId -ErrorAction SilentlyContinue | where {$_.TargetClass -eq "User"} | select ActiveUnits,ConsumedUnits,LockedOutUnits,SuspendedUnits,WarningUnits,SkuId,SkuPartNumber,@{N="SubscriptionIds";E={$_.SubscriptionIds}}
                $MSOLSubscriptions = Get-MsolSubscription -TenantId $tenantId | where {$_.Status -eq "Enabled"} | select DateCreated,NextLifecycleDate,IsTrial,OwnerObjectId,SkuId,SkuPartNumber,TotalLicenses
                $LVSReportingApp = Get-MsolServicePrincipal -TenantId $TenantId -SearchString "LongView-ReportingApp"
            }
            ELSE
            {
                $MSOLUsers = Get-MsolUser -EnabledFilter All -All -ErrorAction SilentlyContinue | Select DisplayName,BlockCredential,Licenses,UserPrincipalName,LastDirSyncTime
                $MSOLAccountSkus = Get-MsolAccountSku -ErrorAction SilentlyContinue | where {$_.TargetClass -eq "User"} | select ActiveUnits,ConsumedUnits,LockedOutUnits,SuspendedUnits,WarningUnits,SkuId,SkuPartNumber,@{N="SubscriptionIds";E={$_.SubscriptionIds}}
                $MSOLSubscriptions = Get-MsolSubscription | where {$_.Status -eq "Enabled"} | select DateCreated,NextLifecycleDate,IsTrial,OwnerObjectId,SkuId,SkuPartNumber,TotalLicenses
                $LVSReportingApp = Get-MsolServicePrincipal -SearchString "LongView-ReportingApp"
                $TenantId = (Get-MsolAccountSku | Select-Object AccountObjectId -Last 1).AccountObjectId
            }

        #Part 2, Step 2, Create the Summary Report.



            #Set the Output File Name and Location.
            $Incr = $null
            $testpath = $null
            DO{
            IF($Path)
            {
            $OutFilePath = "$Path\$CompanyName-$Datefield$Incr.xlsx"
            }
            ELSE
            {
            $OutFilePath = "$CurrentPath\$CompanyName-$Datefield$Incr.xlsx"
            }

            $Incr = $Incr + 1
            $testpath = Test-Path $OutFilePath

            }
            UNTIL($testpath -eq $false)
            
            #Copy the template XLSX file to the output path file name.
            Copy-Item -Path C:\working\Office365LicensingReport.xlsx -Destination $OutFilePath
            
            #Set the destination file as the Excel Package for the output.
            $Report = Open-ExcelPackage -Path $OutFilePath

            
            #Create Summary Sheet and Subscription Details
                
                #Total Users
                $TotalUsers = ($MSOLUsers | Select-Object UserPrincipalName -Unique | Measure).Count
                #Disabled Users
                $DisabledUsers = ($MSOLUsers | where {$_.BlockCredential -eq $true} | Select-Object UserPrincipalName -Unique | measure).Count
                #Enabled Licensed Users
                $EnabledLicensedUsers = ($MSOLUsers | where {($_.BlockCredential -eq $false) -and ($_.Licenses -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
                #Disabled Licensed Users
                $DisabledLicensedUsers = ($MSOLUsers | where {($_.BlockCredential -eq $true) -and ($_.licenses -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
                #Estimated Monthly Retail Total of all Licensing
                $EstRetailTotal =  ($msolusers | ForEach-Object {$user = $_ ; $Lics = $user.licenses.accountsku.skupartnumber ; IF($Lics){foreach($lic in $lics){($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail}}} | measure -Sum).Sum
                #Estimated Monthly Retail Cost of Enabled Users Licensing
                $EstRetailEnabled = ($msolusers | where {$_.BlockCredential -eq $false} | ForEach-Object {$user = $_ ; $Lics = $user.licenses.accountsku.skupartnumber ; IF($Lics){foreach($lic in $lics){($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail}}} | measure -Sum).Sum
                #Estimated Monthly Retail Cost of Disabled Users Licensing
                $EstRetailDisabled = ($msolusers | where {$_.BlockCredential -eq $true} | ForEach-Object {$user = $_ ; $Lics = $user.licenses.accountsku.skupartnumber ; IF($Lics){foreach($lic in $lics){($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail}}} | measure -Sum).Sum
                $ReportDate = Get-Date -Format MM/dd/yyyy

                $SummarySheet = $Report.Workbook.Worksheets["Summary"]
                Set-ExcelRange -WorkSheet $SummarySheet -Range B3 -Value $CompanyName
                Set-ExcelRange -WorkSheet $SummarySheet -Range B4 -Value $PrimaryDomain
                Set-ExcelRange -WorkSheet $SummarySheet -Range B5 -Value $TotalUsers
                Set-ExcelRange -WorkSheet $SummarySheet -Range B6 -Value $DisabledUsers
                Set-ExcelRange -WorkSheet $SummarySheet -Range B7 -Value $EnabledLicensedUsers
                Set-ExcelRange -WorkSheet $SummarySheet -Range B8 -Value $DisabledLicensedUsers
                Set-ExcelRange -WorkSheet $SummarySheet -Range B9 -Value $EstRetailTotal
                Set-ExcelRange -WorkSheet $SummarySheet -Range B10 -Value $EstRetailEnabled
                Set-ExcelRange -WorkSheet $SummarySheet -Range B11 -Value $EstRetailDisabled
                Set-ExcelRange -WorkSheet $SummarySheet -Range B12 -Value $ReportDate

                #Usage by License
                
                $row = 15
                foreach($sku in $MSOLAccountSkus)
                {
                   $SkuStatus = $Null
                   $SkuStatus = ($Plans | where {$_.SkuStatus -ne "NonBillable" -and $_.SkuPartNumber -eq $sku.SkuPartNumber} | select SkuStatus).SkuStatus
                   IF($SkuStatus)
                   {
                   $LicenseName = ($Plans | where {$_.SkuPartNumber -eq $sku.SkuPartNumber} | select SkuName).SkuName
                   $LicTotal = $sku.ActiveUnits
                   $LicUsed = $sku.ConsumedUnits
                   $LicAvail = $LicTotal - $LicUsed
                   $LicEstRet = $LicTotal * ($Plans | where {$_.SkuPartNumber -eq $sku.SkuPartNumber} | select SkuRetail).SkuRetail
                   $LicEstUnused = ($LicTotal - $LicUsed - ($msolusers | where {$_.BlockCredential -eq $true -and $_.Licenses.AccountSku.SkuPartNumber -eq $sku.SkuPartNumber} | measure -sum -ErrorAction SilentlyContinue).Sum) * ($Plans | where {$_.SkuPartNumber -eq $sku.SkuPartNumber} | select SkuRetail).SkuRetail
                   
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("A" + $row) -Value $LicenseName
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("B" + $row) -Value $LicTotal
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("C" + $row) -Value $LicUsed
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("D" + $row) -Value $LicAvail
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("E" + $row) -Value $LicEstRet
                   Set-ExcelRange -WorkSheet $SummarySheet -Range ("F" + $row) -Value $LicEstUnused

                   $Row = $row + 1
                   }
                }

                #Subscription Summary
                $SubscriptionSheet = $Report.Workbook.Worksheets["Subscription_Summary"]
                $row = 2
                foreach($sub in $MSOLSubscriptions)
                {
                    $SkuStatus = $Null
                    $SkuStatus = ($Plans | where {$_.SkuStatus -ne "NonBillable" -and $_.SkuPartNumber -eq $sub.SkuPartNumber} | select SkuStatus).SkuStatus
                    IF($SkuStatus)
                    {
                        [string]$PartnerID = $sub.OwnerObjectId
                        IF($PartnerID)
                        {
                        $PurchasedFrom = $Partners.$PartnerID
                        }
                        ELSE
                        {
                        $PurchasedFrom = ""
                        }

                        $LicenseName = ($Plans | where {$_.SkuPartNumber -eq $sub.SkuPartNumber} | select SkuName).SkuName
                        $DateCreated = $sub.DateCreated.ToString("MM/dd/yyyy")
                        $NextLifecycleDate = $sub.NextLifecycleDate.ToString("MM/dd/yyyy")
                        $IsTrial = $sub.IsTrial
                        $TotalLicenses = $sub.TotalLicenses
                        $EstUnitCost = ($Plans | where {$_.SkuPartNumber -eq $sub.SkuPartNumber}).SkuRetail
                        $EstTotal = $TotalLicenses * $EstUnitCost

                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("A" + $row) -Value $LicenseName
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("B" + $row) -Value $DateCreated
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("C" + $row) -Value $NextLifecycleDate
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("D" + $row) -Value $IsTrial
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("E" + $row) -Value $TotalLicenses
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("F" + $row) -Value $EstUnitCost
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("G" + $row) -Value $EstTotal
                        Set-ExcelRange -WorkSheet $SubscriptionSheet -Range ("H" + $row) -Value $PurchasedFrom

                        $row = $row + 1
                    }
                }
                

            #Logic check to determine if a full report is requested.
            IF($ReportType -eq "Summary")
            {
                #Save the report.
                Export-Excel -ExcelPackage $Report
            }
            ELSE
            {
                #Connecting to the LVS Reporting App, if available and download Office Active User Details Report for last 180 days.
                #We only use this for getting activity times, so we also do the calculation for last active time in this block.
                IF($LVSReportingApp)
                {
                $ResourceAppIDURI = "https://graph.microsoft.com"
                $WebAppHeader = Connect-WebApp -TenantID $TenantId -AppId $LVSReportingApp.AppPrincipalId -resourceAppIdURI $ResourceAppIDURI
                $GraphDetails = Get-GraphReport -ReportName getOffice365ActiveUserDetail -ReportPeriod '180 Days'
                $ProdActivations = Get-GraphReport -ReportName getOffice365ActivationCounts -ReportPeriod '180 Days'

                $ProdActivationSheet = $Report.Workbook.Worksheets["Product_Activations"]

                $row = 2
                foreach($Product in $ProdActivations)
                {
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("A" + $row) -Value $Product.'Product Type'
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("B" + $row) -Value $Product.'Windows'
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("C" + $row) -Value $Product.'Windows 10 Mobile'
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("D" + $row) -Value $Product.'Android'
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("E" + $row) -Value $Product.'iOS'
                Set-ExcelRange -WorkSheet $ProdActivationSheet -Range ("F" + $row) -Value $Product.'Mac'

                $row = $row + 1
                }

                $ActivityStatus = $null
                $ActivityStatus = @{}

                foreach($line in $GraphDetails)
                    {
                    $UPN = $line.'User Principal Name'
                    $LastDate = ($line.'Exchange Last Activity Date',$line.'OneDrive Last Activity Date',$line.'SharePoint Last Activity Date',$line.'Skype For Business Last Activity Date',$line.'Teams Last Activity Date',$line.'Yammer Last Activity Date') | Sort-Object | Select-Object -Last 1
                    
                    $ActivityStatus.Add($UPN,$LastDate)
                    }

                }


                #Creating the Users detail sheet.
                $UsersSheet = $Report.Workbook.Worksheets["User_Details"]

                $MSOLUsers | ForEach-Object -Begin {$row=2; $If=0} -Process {
                    $user = $_

                    $DisplayName = $user.DisplayName
                    $UserPrincipalName = $user.UserPrincipalName
                    IF($user.BlockCredential -eq $False)
                    {
                    $LoginStatus = "Enabled"
                    }
                    ELSE
                    {
                    $LoginStatus = "Disabled"
                    }

                    $LastActivityDate = $ActivityStatus.$UserPrincipalName
                    #Licenses Break Out
                        [string]$Licenses = $null
                        [int32]$EstRetail = $null
                        $Lics = $user.Licenses.AccountSku

                        foreach($lic in $Lics)
                        {
                        $thislic = ($Plans | where {$_.SkuPartNumber -eq $lic.SkuPartNumber}).SkuName
                        $Licenses = ($Licenses + " " + $thislic)
                        $thisretail = ($Plans | where {$_.SkuPartNumber -eq $lic.SkuPartNumber}).SkuRetail
                        $EstRetail = $EstRetail + $thisretail
                        }
                    
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("A" + $row) -Value $DisplayName
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("B" + $row) -Value $UserPrincipalName
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("C" + $row) -Value $LoginStatus
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("D" + $row) -Value $Licenses
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("E" + $row) -Value $EstRetail
                    Set-ExcelRange -WorkSheet $UsersSheet -Range ("F" + $row) -Value $LastActivityDate

                    $row = $row + 1
                    #Progress Bar
                    $If = $If+1
                    Write-Progress -Activity "Creating All User Details for $CompanyName" -Status "Progress:" -PercentComplete ($If/($MSOLUsers | measure).Count*100)
                }

            #Save the report.
            Export-Excel -ExcelPackage $Report
         
        #Hey look, a progress bar.
        $I = $I+1
        Write-Progress -Activity "Creating Office 365 Licensing Reports" -Status "Progress:" -PercentComplete ($I/$TenantCount*100)
        }
    }
write-host "Job's done, Boss."
