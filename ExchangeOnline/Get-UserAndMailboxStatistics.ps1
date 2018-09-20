<#
This script will retrieve user, licensing, and mailbox statistics for Office 365 reporting on the user account in AD, the mailbox in Exchange Online, and the license assigned in Office 365.
The script only reports on licensing applicable to Exchange Online usage, which reflects the primary license on the user.
While possible, you wouldn't really see a person assigned an E3, have EXO disabled and then assigned a Exchange P2.
So this is a good bet to be the primary/only license assigned to the user.

This script does not include the commands to connect to Office 365, you'll first need to connect to MSOnline and Exchange Online.
The module MSOnline needs to be installed and loaded, and you will need to connect to the EXO remote powershell.
To install: install-module MSOnline

Get-UserAndMailboxStatistics
    [-UserPrincipalName] <string>
    [-MSOLGroup] <string>
    [-All] <switch>
    [-NoAD] <switch>
    [-CSPCustomerDomain] <string>
    [-Path] <string>

Version: 2.0.beta.9182018
Author: Clark B. Lebarge

#>

param(
[parameter(ParameterSetName="UPN",Mandatory=$true,HelpMessage="Get Statistics for a specific user.")][string]$UserPrincipalName,
[parameter(ParameterSetName="MSOLGroup",Mandatory=$true,HelpMessage="Get statistics for a specific group of users.")][string]$MSOLGroup,
[parameter(ParameterSetName="All",Mandatory=$true,HelpMessage="Get statistics for all users.")][switch]$All,
[parameter(Mandatory=$false,HelpMessage="Do not gather AD information.")][switch]$NoAD,
[parameter(Mandatory=$false,HelpMessage="The customer domain name, required for Cloud Service Partners.")][string]$CSPCustomerDomain,
[parameter(Mandatory=$false,HelpMessage="The name and path for the output CSV file.")][string]$Path
)

#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}

#Hashtable of all possible plans which include an Exchange Online mailbox license.
#May need occasional updating.
$Plans = $null
$Plans = @{}
    #Enterprise Plans
    $Plans.Add("STANDARDPACK","Office 365 Enterprise E1")
    $Plans.Add("ENTERPRISEPACK","Office 365 Enterprise E3")
    $Plans.Add("ENTERPRISEPACKWITHSCAL","Office 365 Enterprise E4")
    $Plans.Add("ENTERPRISEPREMIUM","Office 365 Enterprise E5")
    $Plans.Add("ENTERPRISEPREMIUM_NOPSTNCONF","Office 365 Enterprise E5 without PSTN Conferencing")
    #Frontline Worker (Kiosk)
    $Plans.Add("DESKLESSPACK","Office 365 F1")
    #Microsoft 365
    $Plans.Add("SPE_F1","Microsoft 365 F1")
    $Plans.Add("SPE_E3","Microsoft 365 E3")
    $Plans.Add("SPE_E5","Microsoft 365 E5")
    #Small and Medium Business
    $Plans.Add("O365_BUSINESS_PREMIUM","Office 365 Business Premium")
    $Plans.Add("O365_BUSINESS_ESSENTIALS","Office 365 Business Essentials")
    #Education Plans
    $Plans.Add("STANDARDWOFFPACK_FACULTY","Office 365 A1 for Faculty")
    $Plans.Add("STANDARDWOFFPACK_STUDENT","Office 365 A1 for Students")
    $Plans.Add("STANDARDWOFFPACK_IW_FACULTY","Office 365 A1 Plus for Faculty")
    $Plans.Add("STANDARDWOFFPACK_IW_STUDENT","Office 365 A1 Plus for Students")
    $Plans.Add("ENTERPRISEPACK_FACULTY","Office 365 A3 for Faculty")
    $Plans.Add("ENTERPRISEPACK_STUDENT","Office 365 A3 for Students")
    $Plans.Add("ENTERPRISEPREMIUM_FACULTY","Office 365 A5 for Faculty")
    $Plans.Add("ENTERPRISEPREMIUM_STUDENT","Office 365 A5 for Students")
    #Standalone Plans
    $Plans.Add("EXCHANGEDESKLESS","Exchange Online Kiosk")
    $Plans.Add("EXCHANGESTANDARD","Exchange Online (Plan 1)")
    $Plans.Add("EXCHANGEENTERPRISE","Exchange Online (Plan 2)")
    $Plans.Add("EXCHANGEARCHIVE","Exchange Online Archiving for Exchange Server")
    $Plans.Add("EXCHANGEARCHIVE_ADDON","Exchange Archive for Exchange Online")

#Get and save credentials and connect to MSOL.
$Credentials = Get-Credential -Message "Please provide login credentials for Office 365 and Exchange Online."
Connect-MsolService -Credential $Credentials

#Compile list of all Office 365 users.
IF($CSPCustomerDomain)
{
    #Login to the customer's Exchange Online endpoint.
    $Session = New-PSSession `
        -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$CSPCustomerDomain" `
        -Credential $Credentials `
        -Authentication Basic `
        -AllowRedirection

   

    #Get the Tenant ID
    $TenantId = (Get-MsolPartnerContract -DomainName $CSPCustomerDomain).TenantId
    
    IF($All)
    {
    $MSOLUsers = Get-MsolUser -All -TenantId $TenantId
    }

    IF($UserPrincipalName)
    {
    $MSOLUsers = Get-MsolUser -UserPrincipalName $UserPrincipalName -TenantId $TenantId
    }

    IF($MSOLGroup)
    {
    $MSOLGroupID = (Get-MsolGroup -TenantId $TenantId -All | where {$_.DisplayName -like "*$MSOLGroup*"}).ObjectId
    $MSOLUsers= Get-MsolGroupMember -GroupObjectId $MSOLGroupID | foreach {Get-MsolUser -UserPrincipalName $_.EmailAddress}
    }
}
ELSE
{
    IF($All)
    {
    $MSOLUsers = Get-MsolUser -All
    }

    IF($UserPrincipalName)
    {
    $MSOLUsers = Get-MsolUser -UserPrincipalName $UserPrincipalName
    }

    IF($MSOLGroup)
    {
    $MSOLGroupID = (Get-MsolGroup -All| where {$_.DisplayName -like "*$MSOLGroup*"}).ObjectId
    $MSOLUsers= Get-MsolGroupMember -GroupObjectId $MSOLGroupID | foreach {Get-MsolUser -UserPrincipalName $_.EmailAddress}
    }
}

#Queries

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
    

    #Get the statistics for the mailbox
    #$MBX = 
    $UPN = $user.userprincipalname
    $command = "Get-MailboxStatistics $UPN"
    $ScriptBlock = [scriptblock]::Create($command)
    $MBX = Invoke-Command -Session $Session -ScriptBlock $scriptblock
    

    #Get Active Directory statistics for the user. Enabled by default.
    IF(-Not $NoAD)
    {
    $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,DisplayName,Department,LastLogonDate | select Mail,samAccountName,DisplayName,Department,LastLogonDate
    }

    #Output the statistics, and save as a report.
    IF(-Not $Path)
    {
    $Path = "$CurrentPath\Office365Licensing$Datefield.csv"
    }
    $NewLine = New-Object -TypeName PSObject -Property @{
                        LoginStatus = $LoginStatus
                        UserPrincipalName = $user.UserPrincipalName
                        LastMailboxLoginTime = $MBX.LastLogonTime
                        LastMailboxLogoffTime = $MBX.LastLogoffTime
                        MailboxSize = $MBX.TotalItemSize
                        UserName = $Details.samAccountName
                        DisplayName = $Details.DisplayName
                        Department = $Details.Department
                        Email = $Details.Mail
                        License = $License
                        LastWindowsLoginTime = $Details.LastLogonDate
                        }
    $NewLine | ft
    $NewLine | Export-Csv -NoTypeInformation -Path $Path -Append
}

Remove-PSSession $Session

<#
get-msoluser -all | ForEach-Object {

    $user = $_
    
    IF($user.Licenses.accountskuid -like "*ENTERPRISEPACK*"){
                    IF($user.BlockCredential -eq $false)
                    {
                    $LoginStatus = "Enabled"
                    }
                    ELSE
                    {
                    $LoginStatus = "Disabled"
                    }
                    $LIC = "Office 365 Enterprise E3"
                    $MBX = Get-MailboxStatistics $user.UserPrincipalName
                    $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,DisplayName,Division,LastLogonDate | select Mail,samAccountName,DisplayName,Division,LastLogonDate
                    New-Object -TypeName PSObject -Property @{
                        LoginStatus = $LoginStatus
                        UserPrincipalName = $user.UserPrincipalName
                        LastMailboxLoginTime = $MBX.LastLogonTime
                        LastMailboxLogoffTime = $MBX.LastLogoffTime
                        MailboxSize = $MBX.TotalItemSize
                        UserName = $Details.samAccountName
                        DisplayName = $Details.DisplayName
                        Division = $Details.Division
                        Email = $Details.Mail
                        License = $LIC
                        LastWindowsLoginTime = $Details.LastLogonDate
                        }
                }
    ELSE{
        IF($user.Licenses.accountskuid -like "*STANDARDPACK*"){
                        IF($user.BlockCredential -eq $false)
                        {
                        $LoginStatus = "Enabled"
                        }
                        ELSE
                        {
                        $LoginStatus = "Disabled"
                        }
                        $LIC = "Office 365 Enterprise E1"
                        $MBX = Get-MailboxStatistics $user.UserPrincipalName
                        $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,UserPrincipalName,DisplayName,Division,LastLogonDate | select Mail,samAccountName,userprincipalname,DisplayName,Division,LastLogonDate
                         New-Object -TypeName PSObject -Property @{
                            LoginStatus = $LoginStatus
                            UserPrincipalName = $user.UserPrincipalName
                            LastMailboxLoginTime = $MBX.LastLogonTime
                            LastMailboxLogoffTime = $MBX.LastLogoffTime
                            MailboxSize = $MBX.TotalItemSize
                            UserName = $Details.samAccountName
                            DisplayName = $Details.DisplayName
                            Division = $Details.Division
                            Email = $Details.Mail
                            License = $LIC
                            LastWindowsLoginTime = $Details.LastLogonDate
                            }
                    }
        ELSE{
            IF($user.Licenses.accountskuid -like "*DESKLESSPACK*"){
                            IF($user.BlockCredential -eq $false)
                            {
                            $LoginStatus = "Enabled"
                            }
                            ELSE
                            {
                            $LoginStatus = "Disabled"
                            }
                            $LIC = "Office 365 Enterprise K1"
                            $MBX = Get-MailboxStatistics $user.UserPrincipalName
                            $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,UserPrincipalName,DisplayName,Division,LastLogonDate | select Mail,samAccountName,userprincipalname,DisplayName,Division,LastLogonDate
                            New-Object -TypeName PSObject -Property @{
                                LoginStatus = $LoginStatus
                                UserPrincipalName = $user.UserPrincipalName
                                LastMailboxLoginTime = $MBX.LastLogonTime
                                LastMailboxLogoffTime = $MBX.LastLogoffTime
                                MailboxSize = $MBX.TotalItemSize
                                UserName = $Details.samAccountName
                                DisplayName = $Details.DisplayName
                                Division = $Details.Division
                                Email = $Details.Mail
                                License = $LIC
                                LastWindowsLoginTime = $Details.LastLogonDate
                                }
                        }
            ELSE{
                    IF($user.Licenses.accountskuid -like "*ENTERPRISEPREMIUM*"){
                                    IF($user.BlockCredential -eq $false)
                                    {
                                    $LoginStatus = "Enabled"
                                    }
                                    ELSE
                                    {
                                    $LoginStatus = "Disabled"
                                    }
                                    $LIC = "Office 365 Enterprise E5"
                                    $MBX = Get-MailboxStatistics $user.UserPrincipalName
                                    $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,UserPrincipalName,DisplayName,Division,LastLogonDate | select Mail,samAccountName,userprincipalname,DisplayName,Division,LastLogonDate
                                    New-Object -TypeName PSObject -Property @{
                                        LoginStatus = $LoginStatus
                                        UserPrincipalName = $user.UserPrincipalName
                                        LastMailboxLoginTime = $MBX.LastLogonTime
                                        LastMailboxLogoffTime = $MBX.LastLogoffTime
                                        MailboxSize = $MBX.TotalItemSize
                                        UserName = $Details.samAccountName
                                        DisplayName = $Details.DisplayName
                                        Division = $Details.Division
                                        Email = $Details.Mail
                                        License = $LIC
                                        LastWindowsLoginTime = $Details.LastLogonDate
                                        }
                                }
                    ELSE{
                        IF($user.Licenses.accountskuid -like "*EXCHANGEENTERPRISE*"){
                                        IF($user.BlockCredential -eq $false)
                                        {
                                        $LoginStatus = "Enabled"
                                        }
                                        ELSE
                                        {
                                        $LoginStatus = "Disabled"
                                        }
                                        $LIC = "Exchange Online (Plan 2)"
                                        $MBX = Get-MailboxStatistics $user.UserPrincipalName
                                        $Details = Get-ADUser -filter {(UserPrincipalName -eq $User.UserPrincipalName)} -Properties Mail,samAccountName,UserPrincipalName,DisplayName,Division,LastLogonDate | select Mail,samAccountName,userprincipalname,DisplayName,Division,LastLogonDate
                                        New-Object -TypeName PSObject -Property @{
                                            LoginStatus = $LoginStatus
                                            UserPrincipalName = $user.UserPrincipalName
                                            LastMailboxLoginTime = $MBX.LastLogonTime
                                            LastMailboxLogoffTime = $MBX.LastLogoffTime
                                            MailboxSize = $MBX.TotalItemSize
                                            UserName = $Details.samAccountName
                                            DisplayName = $Details.DisplayName
                                            Division = $Details.Division
                                            Email = $Details.Mail
                                            License = $LIC
                                            LastWindowsLoginTime = $Details.LastLogonDate
                                            }
                                    }
                }
                }
                        
        }

    }
} | Export-CSV -NoTypeInformation -Path "$CurrentPath\Office365Licensing$Datefield.csv"
#>