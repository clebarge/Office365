# Office365
Powershell Scripts and Utilities for Managing Office 365.

Office365Logon.ps1: This script is a basic logon to Office 365 using the MSOnline module, and then connecting to Exchange Online remote powershell. This does not support MFA for login to Exchange Online. To use MFA with Exchange Online requires installation of the Exchange Online powershell environment.

Office365-EXO_with_MFA_logon.ps1: This script uses the EXO powershell module loaded locally to connect to Exchange Online with MFA. It also connects to the MSOnline.
