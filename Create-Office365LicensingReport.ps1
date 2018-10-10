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


    Version: 1.0.beta.10092018
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
            }
            ELSE
            {
                $MSOLUsers = Get-MsolUser -EnabledFilter All -All -ErrorAction SilentlyContinue | Select DisplayName,BlockCredential,Licenses,UserPrincipalName,LastDirSyncTime
                $MSOLAccountSkus = Get-MsolAccountSku -ErrorAction SilentlyContinue | where {$_.TargetClass -eq "User"} | select ActiveUnits,ConsumedUnits,LockedOutUnits,SuspendedUnits,WarningUnits,SkuId,SkuPartNumber,@{N="SubscriptionIds";E={$_.SubscriptionIds}}
                $MSOLSubscriptions = Get-MsolSubscription | where {$_.Status -eq "Enabled"} | select DateCreated,NextLifecycleDate,IsTrial,OwnerObjectId,SkuId,SkuPartNumber,TotalLicenses
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

                #Usage by License
                
                $row = 14
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

                #Subscription Details
                $SubscriptionSheet = $Report.Workbook.Worksheets["Subscription_Details"]
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
                #Doing nothing!
            }
            ELSE
            {

#This is basic code for the Last Activity Time gathered from Microsoft Graph. Currently only functional in WFP for testing.
###########################################################################################################################################
                IF($CompanyName -like "*Western Forest*")
                {
                $resourceAppIdURI = "https://graph.microsoft.com"
                $ClientID         = "b1f930e5-f5ab-4ff7-ab5f-3174877483ea"   #AKA Application ID
                $TenantName       = "westernforest0.onmicrosoft.com"             #Your Tenant Name
                $CredPrompt       = "Auto"                                   #Auto, Always, Never, RefreshSession
                $redirectUri      = "https://clark-lvs-test7.westernforest.com"                #Your Application's Redirect URI
                $Uri              = "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserDetail(Period='d180')" #The query you want to issue to Invoke a REST command with
                $Method           = "GET"
    
                  if (!$CredPrompt){$CredPrompt = 'Auto'}
                    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
                    if ($AadModule -eq $null) {$AadModule = Get-Module -Name "AzureADPreview" -ListAvailable}
                    if ($AadModule -eq $null) {write-host "AzureAD Powershell module is not installed. The module can be installed by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt. Stopping." -f Yellow;exit}
                    if ($AadModule.count -gt 1) {
                        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
                        $aadModule      = $AadModule | ? { $_.version -eq $Latest_Version.version }
                        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
                        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
                        }
                    else {
                        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
                        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
                        }
                    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
                    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
                    $authority          = "https://login.microsoftonline.com/$TenantName"
                    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
                    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
                    $AccessToken        = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters).Result

                    $Header = @{
                        'Authorization' = $AccessToken.CreateAuthorizationHeader()
                        }

                $ActivityReport = (Invoke-RestMethod -Headers $Header -Uri $Uri -UseBasicParsing -Method $Method).Split([Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries)

                $ActivityStatus = $null
                $ActivityStatus = @{}

                foreach($line in $ActivityReport)
                {
                [string]$UPNfromGraph = ($line.split(',')).GetValue(1)
                        IF(($line.split(',')).GetValue(11))
                        {
                        [datetime]$ExchangeLA = ($line.split(',')).GetValue(11)
                        $LastActivityDate = $ExchangeLA.ToShortDateString()

                        }
                        IF(($line.split(',')).GetValue(12))
                        {
                        [datetime]$OneDriveLA = ($line.split(',')).GetValue(12)
                            IF($OneDriveLA -ge $LastActivityDate)
                            {
                            $LastActivityDate = $OneDriveLA.ToShortDateString()
                            }
                                    }
                        IF(($line.split(',')).GetValue(13))
                        {
                        [datetime]$SharePointLA = ($line.split(',')).GetValue(13)
                                    IF($SharePointLA -ge $LastActivityDate)
                            {
                            $LastActivityDate = $SharePointLA.ToShortDateString()
                            }
                        }
                        IF(($line.split(',')).GetValue(14))
                        {
                        [datetime]$SkypeLA = ($line.split(',')).GetValue(14)
                                    IF($SkypeLA -ge $LastActivityDate)
                            {
                            $LastActivityDate = $SkypeLA.ToShortDateString()
                            }
                        }
                        IF(($line.split(',')).GetValue(15))
                        {
                        [datetime]$YammerLA = ($line.split(',')).GetValue(15)
                                    IF($YammerLA -ge $LastActivityDate)
                            {
                            $LastActivityDate = $YammerLA.ToShortDateString()
                            }
                        }
                        IF(($line.split(',')).GetValue(16))
                        {
                        [datetime]$TeamsLA = ($line.split(',')).GetValue(16)
                                    IF($TeamsLA -ge $LastActivityDate)
                            {
                            $LastActivityDate = $TeamsLA.ToShortDateString()
                            }
                        }
                    $ActivityStatus.Add($UPNfromGraph,$LastActivityDate)
                }
                }
###########################################################################################################################################
#End of Last Activity Test Code.
                
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
