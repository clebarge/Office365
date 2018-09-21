<#
This script will retrieve user, licensing, and mailbox statistics for Office 365 reporting on the user account in AD, the mailbox in Exchange Online, and the license assigned in Office 365.
The script only reports on licensing applicable to Exchange Online usage, which reflects the primary license on the user.
While possible, you wouldn't really see a person assigned an E3, have EXO disabled and then assigned a Exchange P2.
So this is a good bet to be the primary/only license assigned to the user.

This script requires the MSOnline module to be installed as well as the Exchange Online remote powershell.

For CSP use, currently the script doesn't work with MFA authentication.

Get-UserAndMailboxStatistics
    [-UserPrincipalName] <string>
    [-MSOLGroup] <string>
    [-All] <switch>
    [-IncludeAD] <switch>
    [-LicenseOnly] <switch>
    [-DisabledOnly] <switch>
    [-CSPCustomerDomain] <string>
    [-Path] <string>

Version: 2.1.9212018
Author: Clark B. Lebarge

#>

param(
[parameter(ParameterSetName="UPN",Mandatory=$true,HelpMessage="Get Statistics for a specific user.")][string]$UserPrincipalName,
[parameter(ParameterSetName="MSOLGroup",Mandatory=$true,HelpMessage="Get statistics for a specific group of users.")][string]$MSOLGroup,
[parameter(ParameterSetName="All",Mandatory=$true,HelpMessage="Get statistics for all users.")][switch]$All,
[parameter(Mandatory=$false,HelpMessage="Gather AD information. Use only if you're internal to the network.")][switch]$IncludeAD,
[parameter(Mandatory=$false,HelpMessage="Skip mailbox statistics, gather license and account status only. For a faster report.")][switch]$LicenseOnly,
[parameter(ParameterSetName="MSOLGroup",Mandatory=$false,HelpMessage="Report on disabled users only. For a faster report.")]
[parameter(ParameterSetName="All",Mandatory=$false,HelpMessage="Report on disabled users only. For a faster report.")]
[switch]$DisabledOnly,
[parameter(Mandatory=$false,HelpMessage="The customer domain name, required for Cloud Service Partners.")][string]$CSPCustomerDomain,
[parameter(Mandatory=$false,HelpMessage="The name and path for the output CSV file.")][string]$Path
)

Function Logon-CSPEXO
{
    #Login to the customer's Exchange Online endpoint.
    $Script:Session = New-PSSession `
        -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$CSPCustomerDomain" `
        -Credential $Credentials `
        -Authentication Basic `
        -AllowRedirection
}

Function Logon-EXO
{
    #Logon to Exchange Online
        $modules = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
        $moduleName =  Join-Path $modules[0].Directory.FullName "Microsoft.Exchange.Management.ExoPowershellModule.dll"
        Import-Module -FullyQualifiedName $moduleName -Force
        $scriptName =  Join-Path $modules[0].Directory.FullName "CreateExoPSSession.ps1"
        . $scriptName
        $null = Connect-EXOPSSession
        $Script:Session = (Get-PSSession | Where-Object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.State -eq 'Opened') })[0]
}

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

IF($CSPCustomerDomain)
{
$Credentials = Get-Credential -Message "Please provide CSP login credentials for Office 365 and Exchange Online."
Connect-MsolService -Credential $Credentials
}
ELSE
{
Connect-MSOLService
}

#Compile list of Office 365 users.
IF($CSPCustomerDomain)
{
    #Get the Tenant ID
    $TenantId = (Get-MsolPartnerContract -DomainName $CSPCustomerDomain).TenantId
    
    IF($All)
    {
        IF($DisabledOnly)
        {
        $MSOLUsers = Get-MsolUser -EnabledFilter DisabledOnly -All -TenantId $TenantId | Select BlockCredential,Licenses,UserPrincipalName
        }
        ELSE
        {
        $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select BlockCredential,Licenses,UserPrincipalName
        }
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
                IF($DisabledOnly)
                {
                Get-MsolUser -EnabledFilter DisabledOnly -TenantId $TenantId -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
                }
                ELSE
                {
                Get-MsolUser -EnabledFilter All -TenantId $TenantId -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
                }
            }
    }
}
ELSE
{
    IF($All)
    {
        IF($DisabledOnly)
        {
        $MSOLUsers = Get-MsolUser -EnabledFilter DisabledOnly -All | Select BlockCredential,Licenses,UserPrincipalName
        }
        ELSE
        {
        $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select BlockCredential,Licenses,UserPrincipalName
        }
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
                IF($DisabledOnly)
                {
                Get-MsolUser -EnabledFilter DisabledOnly -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
                }
                ELSE
                {
                Get-MsolUser -EnabledFilter All -UserPrincipalName $_.EmailAddress | Select BlockCredential,Licenses,UserPrincipalName
                }
            }
    }
}

#Initial Connection to Exchange Online.

IF($CSPCustomerDomain)
{
Logon-CSPEXO

}
ELSE
{
Logon-EXO
}

#Due to the Session time out of 1 hour, Setting a time now when the script will reconnect to Exchange Online to renew the session.
$SessionTime = (get-date).AddMinutes(55)

#Queries
    
#Get the list of mailboxes for the domain, we'll use this to compare to the MSOL list and only query for statistics on existing mailboxes. This should speed up the process.
IF($LicenseOnly -eq $false)
{
$Mailboxes = Invoke-Command -Session $Session -ScriptBlock {Get-Mailbox -ResultSize Unlimited | Select-Object EmailAddresses}


}

foreach($user in $MSOLUsers)
{
    #For large environments with over 2000 mailboxes, the script will need to refresh the Session.
    IF((Get-Date) -ge $SessionTime)
    {
    
        Remove-PSSession $Session
        
        IF($CSPCustomerDomain)
        {
        Logon-CSPEXO
        }
        ELSE
        {
        Logon-EXO
        }

        #Due to the Session time out of 1 hour, Setting a time now when the script will reconnect to Exchange Online to renew the session.
        $SessionTime = (get-date).AddMinutes(55)
    }
    
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
    

    #Get the statistics for the mailbox.
    IF($LicenseOnly -eq $false)
    {
        Foreach($EmailAddress in $Mailboxes)
        {
        $UPN = $user.userprincipalname
        IF($EmailAddress.EmailAddresses -match $UPN)
        {
            $command = "Get-MailboxStatistics $UPN | Select-Object LastLogonTime,LastLogoffTime,TotalItemSize"
            $ScriptBlock = [scriptblock]::Create($command)
            $MBX = Invoke-Command -Session $Session -ScriptBlock $scriptblock
        
            #Breaking out of the loop now that we found a match.
            Break
        }
        ELSE
        {
        $MBX = $null
        }
        }
    }

    #Get Active Directory statistics for the user. Enabled by default.
    IF($IncludeAD)
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

IF($LicenseOnly -eq $false)
{
Remove-PSSession $Session
write-host "Work Complete."
}
ELSE
{
write-host "Job's done, Boss."
}