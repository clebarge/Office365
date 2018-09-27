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
    [-UserPrincipalName] <string>
    [-MSOLGroup] <string>
    [-All] <switch>
    [-CSPAllDomains] <switch>
    [-CSPCustomerDomain] <string>
    [-ReportType] <listoption> Summary | Full
    [-Path] <string>


    Version: 1.0.beta.9262018
    Author: Clark B. Lebarge
    Company: Long View Systems

#>

param(
[parameter(ParameterSetName="UPN",Mandatory=$true,HelpMessage="Get Statistics for a specific user.")]
[parameter(ParameterSetName="CSPDu",Mandatory=$true,HelpMessage="Get Statistics for a specific user.")]
[parameter(ParameterSetName="CSPAu",Mandatory=$true,HelpMessage="Get Statistics for a specific user.")]
[string]$UserPrincipalName,
[parameter(ParameterSetName="MSOLGroup",Mandatory=$true,HelpMessage="Get statistics for a specific group of users.")]
[parameter(ParameterSetName="CSPDg",Mandatory=$true,HelpMessage="Get statistics for a specific group of users.")]
[parameter(ParameterSetName="CSPAg",Mandatory=$true,HelpMessage="Get statistics for a specific group of users.")]
[string]$MSOLGroup,
[parameter(ParameterSetName="All",Mandatory=$true,HelpMessage="Get statistics for all users.")]
[parameter(ParameterSetName="CSPDa",Mandatory=$true,HelpMessage="Get statistics for all users.")]
[parameter(ParameterSetName="CSPAa",Mandatory=$true,HelpMessage="Get statistics for all users.")]
[switch]$All,
[parameter(ParameterSetName="CSPDa",Mandatory=$false,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[parameter(ParameterSetName="CSPDg",Mandatory=$false,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[parameter(ParameterSetName="CSPDu",Mandatory=$false,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[string]$CSPCustomerDomain,
[parameter(ParameterSetName="CSPAa",Mandatory=$false,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[parameter(ParameterSetName="CSPAg",Mandatory=$false,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[parameter(ParameterSetName="CSPAu",Mandatory=$false,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[string]$CSPAll,
[parameter(Mandatory=$false,HelpMessage="The file and folder path for the output Excel file.")]
[string]$Path,
[parameter(Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[string]$ReportType
)

#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}

#Hashtable of all possible plans which include an Exchange Online mailbox license.
#May need occasional updating.
$Plans = $null
$Plans = @{
    #Enterprise SKUs
    'STANDARDPACK'                       =    'Office 365 Enterprise E1'
    'ENTERPRISEPACK'                     =    'Office 365 Enterprise E3'
    'ENTERPRISEPACKWITHSCAL'             =    'Office 365 Enterprise E4'
    'ENTERPRISEPREMIUM'                  =    'Office 365 Enterprise E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'       =    'Office 365 Enterprise E5 without PSTN Conferencing'
    
    #Frontline Worker (Kiosk)
    'DESKLESSPACK'                       =    'Office 365 F1'
    
    #Microsoft 365
    'SPE_F1'                             =    'Microsoft 365 F1'
    'SPE_E3'                             =    'Microsoft 365 E3'
    'SPE_E5'                             =    'Microsoft 365 E5'
    
    #Small and Medium Business
    'O365_BUSINESS_PREMIUM'              =    'Office 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'           =    'Office 365 Business Essentials'
    
    #Education SKUs
    'STANDARDWOFFPACK_FACULTY'           =    'Office 365 A1 for Faculty'
    'STANDARDWOFFPACK_STUDENT'           =    'Office 365 A1 for Students'
    'STANDARDWOFFPACK_IW_FACULTY'        =    'Office 365 A1 Plus for Faculty'
    'STANDARDWOFFPACK_IW_STUDENT'        =    'Office 365 A1 Plus for Students'
    'ENTERPRISEPACK_FACULTY'             =    'Office 365 A3 for Faculty'
    'ENTERPRISEPACK_STUDENT'             =    'Office 365 A3 for Students'
    'ENTERPRISEPREMIUM_FACULTY'          =    'Office 365 A5 for Faculty'
    'ENTERPRISEPREMIUM_STUDENT'          =    'Office 365 A5 for Students'
    
    #Standalone SKUs
    'EXCHANGEDESKLESS'                   =    'Exchange Online Kiosk'
    'EXCHANGESTANDARD'                   =    'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                 =    'Exchange Online (Plan 2)'
    'EXCHANGEARCHIVE'                    =    'Exchange Online Archiving for Exchange Server'
    'EXCHANGEARCHIVE_ADDON'              =    'Exchange Archive for Exchange Online'
}

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

#Compile list of Office 365 users.
IF($CSPCustomerDomain -or $CSPAll)
{
    #Get the Tenant ID
    $TenantId = (Get-MsolPartnerContract -DomainName $CSPCustomerDomain).TenantId
    
    IF($All)
    {
        $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select BlockCredential,Licenses,UserPrincipalName
    }

    IF($UserPrincipalName)
    {
    $MSOLUsers = Get-MsolUser -UserPrincipalName $UserPrincipalName -TenantId $TenantId | Select BlockCredential,Licenses,UserPrincipalName
    }

    IF($MSOLGroup)
    {
    $MSOLGroupID = (Get-MsolGroup -TenantId $TenantId -SearchString "$MSOLGroup" | select ObjectID).ObjectID
    $MSOLUsers= Get-MsolGroupMember -TenantId $TenantId -GroupObjectId $MSOLGroupID | foreach 
            {
                Get-MsolUser -EnabledFilter All -TenantId $TenantId -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
            }
    }
}
ELSE
{
    IF($All)
    {
        $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select BlockCredential,Licenses,UserPrincipalName
    }

    IF($UserPrincipalName)
    {
    $MSOLUsers = Get-MsolUser -UserPrincipalName $UserPrincipalName | Select BlockCredential,Licenses,UserPrincipalName
    }

    IF($MSOLGroup)
    {
    $MSOLGroupID = (Get-MsolGroup -SearchString "$MSOLGroup" | select ObjectID).ObjectID
    $MSOLUsers= Get-MsolGroupMember -GroupObjectId $MSOLGroupID | foreach 
            {
                Get-MsolUser -EnabledFilter All -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
            }
    }
}

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
    $License = $null
    $Lics = $user.licenses.accountsku.skupartnumber
    Foreach($lic in $lics)
    {
    IF($Plans.Item($lic)){
    $License = $Plans.Item($lic)
    }
    }

    #Output the statistics, and save as a report.
    IF(-Not $Path)
    {
    $Path = "$CurrentPath\Office365Licensing$Datefield.csv"
    }
    $NewLine = New-Object -TypeName PSObject -Property @{
                        LoginStatus = $LoginStatus
                        UserPrincipalName = $user.UserPrincipalName
                        License = $License
                        }
    $NewLine | ft
    $NewLine | Export-Csv -NoTypeInformation -Path $Path -Append
}

write-host "Job's done, Boss."
