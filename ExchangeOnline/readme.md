#ExchangeOnline

Powershell scripts related to management of Exchange Online.

Get-UserAndMailboxStatistics.ps1: This powershell script is meant to be ran from a domain joined system internal to the organization. It connects to Office 365 to retrieve the license for the user's mailbox in Exchange Online, it then retrieves the mailbox statistics from Exchange Online, and finally also queries Active Directory for the windows login statistics. It compiles this information into a CSV file for review. Requires connecting to MSOnline, EXO, and Active Directory through powershell.

In order to accomodate adding in the ability to check CSP customers, the script will prompt for a username and password. For regular Office 365 tenants, login with your global or other admin account. For Cloud Service Providers, the username and password would be your login, not one in the customer's domain.

Get-UserAndMailboxStatistics
    [-UserPrincipalName] <string>
    [-MSOLGroup] <string>
    [-All] <switch>
    [-IncludeAD] <switch>
    [-LicenseOnly] <switch>
    [-DisabledOnly] <switch>
    [-CSPCustomerDomain] <string>
    [-Path] <string>

UserPrincipalName: Gets stats for a specific user only.

MSOLGroup: Gets stats for members of a specific group in Office 365, only checks users that are members, doesn't drill down into nested groups.

All: Gets stats for all users.

IncludeAD: Gets information from Active Directory. AD requires you run the command from within the internal network.

LicenseOnly: Skips the exhaustive mailbox statistics, reducing run time to minutes instead of hours for large organizations.

DisabledOnly: Skips checking enabled users, speeds up the run time. Combine with LicenseOnly for a fast and quick license evaluation of your Office 365 environment. This is only really accurate for tenants with Exchange Online. Doesn't check for Skype or other licensing.

CSPCustomerDomain: For cloud service providers, allows you to query this information for a specific customer by specifying their domain.

Path: The output location for the report, if ommitted this is defaulted to the current directory with an automatic file name including the date.

Example: Get all information for your Exchange Online implementation, creating the report file automatically in the current folder.
Get-UserAndMailboxStatistics.ps1 -All

Example: Get information for a specific user, saving it to a specific location.
Get-UserAndMailboxStatistics.ps1 -UserPrincipalName john.doe@domain.com -Path C:\Working\johndoe.csv

Example: Get information for a cloud service provider customer, skip AD information.
Get-UserAndMailboxStatistics.ps1 -All -NoAD -CSPCustomerDomain customer.onmicrosoft.com -Path C:\Working\customer.csv
