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
[string]$CSPAll,
[parameter(Mandatory=$true,HelpMessage="The file and folder path for the output Excel file.")]
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

#Set the Temp File Name and Location.
    $Incr = 1
    DO{
    
    $TempFilePath = "$Env:TEMP\$Datefield$Incr.csv"
    $Incr = $Incr + 1
    $testpath = Test-Path $TempFilePath

    }
    UNTIL($testpath -eq $false) 

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





Write-Host "Executing Step 1: Query for user details from Office 365. Please wait."

#Compile list of Office 365 users.
IF($CSPCustomerDomain -or $CSPAll)
{
    #Get the Tenant ID
    $TenantId = (Get-MsolPartnerContract -DomainName $CSPCustomerDomain).TenantId

    $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select BlockCredential,Licenses,UserPrincipalName
}
ELSE
{
    $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select BlockCredential,Licenses,UserPrincipalName
}

#Create the Temp File.
Create-TempFile -MSOLUsers $MSOLUsers

write-host "Job's done, Boss."
