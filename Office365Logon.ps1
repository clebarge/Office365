#Logon to Office 365 and save the logon information into a global variable.
#This is very useful as it allows you to run this script, then launch and run configuration scripts without needing to include logon in those configuration scripts.
#The script assumes that UPN equals Email Address.
#This logon is intended for management of Office 365 users, licensing, and mailboxes.
#Logon for Azure, SharePoint, or Skype for Business administration requires separate modules to be loaded and are not covered by this script.
$Global:UserCredential = Get-Credential -UserName YourEmailAddress -Message "Please Logon to Office 365 with your Office 365 admin credentials."
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
Connect-MSOLService -credential $usercredential