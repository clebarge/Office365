#ExchangeOnline

Powershell scripts related to management of Exchange Online.

Get-UserAndMailboxStatistics.ps1: This powershell script is meant to be ran from a domain joined system internal to the organization. It connects to Office 365 to retrieve the license for the user's mailbox in Exchange Online, it then retrieves the mailbox statistics from Exchange Online, and finally also queries Active Directory for the windows login statistics. It compiles this information into a CSV file for review. Requires connecting to MSOnline, EXO, and Active Directory through powershell.

Get-UserAndMailboxStatistics
    [-UserPrincipalName] <string>
    [-MSOLGroup] <string>
    [-All] <switch>
    [-NoAD] <switch>
    [-CSPCustomerDomain] <string>
    [-Path] <string>
