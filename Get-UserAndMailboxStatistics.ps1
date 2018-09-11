<#
This script will retrieve user, licensing, and mailbox statistics for Office 365 reporting on the user account in AD, the mailbox in Exchange Online, and the license assigned in Office 365.
The script only reports on licensing applicable to Exchange Online usage, which reflects the primary license on the user.
While possible, you wouldn't really see a person assigned an E3, have EXO disabled and then assigned a Exchange P2.
So this is a good bet to be the primary/only license assigned to the user.
No parameters are defined for this script.

This script does not include the commands to connect to Office 365, you'll first need to connect to MSOnline and Exchange Online.
The module MSOnline needs to be installed and loaded, and you will need to connect to the EXO remote powershell.
To install: install-module MSOnline

#>

#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}

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
