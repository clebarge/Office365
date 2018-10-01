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


    Version: 1.0.beta.9282018
    Author: Clark B. Lebarge
    Company: Long View Systems

#>

param(
[parameter(ParameterSetName="CSPD",Mandatory=$true,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[string]$CSPCustomerDomain,
[parameter(ParameterSetName="CSPA",Mandatory=$true,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[switch]$CSPAll,
[parameter(Mandatory=$false,HelpMessage="The folder path for the output Excel file. Current Directory if ommitted.")]
[string]$Path,
[parameter(Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[string]$ReportType
)

Function Create-TempFile
{

param(
[parameter()]$MSOLUsers
)

foreach($user in $MSOLUsers)
{
    #Set Login Status from Office 365.
    IF($user.BlockCredential -eq $false)
    {
    $LoginStatus = "Enabled"
    }
    ELSE
    {
    $LoginStatus = "Disabled"
    }

    #Set License Readable Name

    $Lics = $user.licenses.accountsku.skupartnumber
    Foreach($lic in $lics)
    {
    $SkuName = $null
    $SkuType = $null
    $SkuRetail = $null
    $SkuStatus = $null

    $SkuName = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuName
    $SkuType = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuType
    $SkuRetail = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail
    $SkuStatus = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuStatus
    
 
    
    #Output the raw statistics CSV file to the temp directory.

    New-Object -TypeName PSObject -Property @{
                        LoginStatus = $LoginStatus
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        SkuName = $SkuName
                        SkuType = $SkuType
                        SkuRetail = $SkuRetail
                        SkuStatus = $SkuStatus
                        } | Export-Csv -NoTypeInformation -Path $TempFilePath -Append
    }
}
}

#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}

#Import the SKU Details from Excel File
$Plans = Import-Excel -Path C:\working\SKUDetails.xlsx



#Get and save credentials and connect to MSOL.

IF($CSPCustomerDomain -or $CSPAll)
{
$Credentials = Get-Credential -Message "Please provide CSP login credentials for Office 365 and Exchange Online."
Connect-MsolService -Credential $Credentials
}
ELSE
{
Connect-MSOLService
}



#If the CSPAll switch is set, get list of all customer tenant ids.
IF($CSPAll)
{
$TenantInfo = Get-MsolPartnerContract | select Name,TenantId
}
ELSE
{
    
    IF($CSPCustomerDomain)
    {
    $TenantInfo = Get-MsolPartnerContract -DomainName $CSPCustomerDomain | select Name,TenantId
    }
    ELSE
    {
    $TenantInfo = Get-MsolCompanyInformation | select @{N='Name';E={$_.DisplayName}}
    }
}

Foreach($Tenant in $TenantInfo)
{
$CompanyName = $Tenant.Name
$TenantId = $Tenant.TenantId
Write-Host "Gathering information for Company: $CompanyName."

#Set the Temp File Name and Location.
    $Incr = 1
    $testpath = $null
    DO{
    
    $TempFilePath = "$Env:TEMP\$TenantId$Datefield$Incr.csv"
    $Incr = $Incr + 1
    $testpath = Test-Path $TempFilePath

    }
    UNTIL($testpath -eq $false) 

#Step 1
Write-Host "Executing Step 1: Query for user details from Office 365. Please wait."

IF($TenantID)
{
    $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
}
ELSE
{
    $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
}

#Create the Temp File.

    #Bit of logic to deal with empty customers when running CSP.
    $MSOLUserCount = ($MSOLUsers | measure).Count
    IF($MSOLUserCount -lt 5)
    {
    Write-Host "Under 5 users, tenancy appears unused. Skipping."
    Continue
    }

Create-TempFile -MSOLUsers $MSOLUsers

#Step 2
Write-Host "Step 2: Creating Report."

$Tempfile = Import-Csv -Path $TempFilePath

#Set the Output File Name and Location.
    $Incr = 1
    $testpath = $null
    DO{
    IF($Path)
    {
    $OutFilePath = "$Path\$CompanyName$Datefield$Incr.xlsx"
    }
    ELSE
    {
    $OutFilePath = "$CurrentPath\$CompanyName$Datefield$Incr.xlsx"
    }

    $Incr = $Incr + 1
    $testpath = Test-Path $OutFilePath

    }
    UNTIL($testpath -eq $false) 


#Create Summary Sheet
Write-Host "Creating Summary Sheet."
    #Total Users
    $TotalUsers = ($Tempfile | Select-Object UserPrincipalName -Unique | Measure).Count
    #Disabled Users
    $DisabledUsers = ($Tempfile | where {$_.LoginStatus -eq "Disabled"} | Select-Object UserPrincipalName -Unique | measure).Count
    #Enabled Licensed Users
    $EnabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
    #Disabled Licensed Users
    $DisabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
    #Estimated Monthly Retail Cost of Enabled Users Licensing
    $EstRetailEnabled = (($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -like "*")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum
    #Estimated Monthly Retail Cost of Disabled Users Licensing
    $EstRetailDisabled = (($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -like "*")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum

#Export Summary Sheet
    New-Object -TypeName PSObject -Property @{
        TotalUsers = $TotalUsers
        DisabledUsers = $DisabledUsers
        EnabledLicensedUsers = $EnabledLicensedUsers
        DisabledLicensedUsers = $DisabledLicensedUsers
        EstRetailEnabled = $EstRetailEnabled
        EstRetailDisabled = $EstRetailDisabled
    } | export-excel -Path $OutFilePath -WorksheetName "Summary"

#Create All Users Sheet
$AllUsersSheet = New-Object System.Data.DataTable
$AllUsersSheet.Columns.Add("DisplayName","string")
$AllUsersSheet.Columns.Add("UserPrincipalName","string")
$AllUsersSheet.PrimaryKey = $AllUsersSheet.Columns[1]
$AllUsersSheet.Columns.Add("Licenses","string")
$AllUsersSheet.Columns.Add("LoginStatus","string")
$AllUsersSheet.Columns.Add("EstimatedRetailCostUSD","int32")


Write-Host "Creating Full Report"

$Tempfile | ForEach-Object {
    $line = $_
    #Check if UPN already in AllUsersSheet

    $testline = $AllUsersSheet | where {$_.UserPrincipalName -eq $line.UserPrincipalName}
    IF(-not $testline)
    {
    $r = $AllUsersSheet.NewRow()
    $r.DisplayName = $line.DisplayName
    $r.UserPrincipalName = $line.userprincipalname
    $r.Licenses = $line.SkuName
    $r.LoginStatus = $line.LoginStatus
    $r.EstimatedRetailCostUSD = [int32]$line.SkuRetail
    $AllUsersSheet.Rows.Add($r)
    }
    ELSE
    {
    $AllUsersSheet | where {$_.UserPrincipalName -eq $line.UserPrincipalName} | foreach {
        $_.Licenses = ($_.Licenses + ";" + $line.SkuName)
        $_.EstimatedRetailCostUSD = $_.EstimatedRetailCostUSD + $line.SkuRetail
        }
    }
} 

$AllUsersSheet | Export-Excel -Path $OutFilePath -WorksheetName "All Users"

#Create the Disabled Users with Licenses Sheet

$DisUserWithLicenseSheet = $AllUsersSheet | Where {$_.LoginStatus -eq "Disabled"}

$DisUserWithLicenseSheet | Export-Excel -Path $OutFilePath -WorksheetName "Disabled Users"


#Clean Up Temp File
Remove-Item $TempFilePath -Confirm:$false

$TenantId = $null
$MSOLUsers = $null
}



write-host "Job's done, Boss."
