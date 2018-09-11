<#
Connect to Exchange Online with MFA authentication.
Need to install a module, follow these instructions;
https://technet.microsoft.com/en-us/library/mt775114(v=exchg.160).aspx

Also connects to the MSOL service to manage licensing and users in Office 365.
Requires installing the module with this command:
install-module -name MSOnline
#>


#Load the Modules.
$modules = @(Get-ChildItem -Path "$($env:LOCALAPPDATA)\Apps\2.0" -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
$moduleName =  Join-Path $modules[0].Directory.FullName "Microsoft.Exchange.Management.ExoPowershellModule.dll"
Import-Module -FullyQualifiedName $moduleName -Force
$scriptName =  Join-Path $modules[0].Directory.FullName "CreateExoPSSession.ps1"
. $scriptName
$null = Connect-EXOPSSession
$exchangeOnlineSession = (Get-PSSession | Where-Object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -and ($_.State -eq 'Opened') })[0]
Import-PSSession $exchangeOnlineSession -Prefix exo

Connect-MsolService

Write-Host "Connected to Exchange Online and Office 365 for remote powershell."
Write-Host "Exchange online commands are prefixed with EXO, so Get-MoveRequest needs to be specified as Get-EXOMoveRequest."
Write-Host "This ensures no ambiguity exists when the on-prem Exchange Admin tools are installed on this system. Commands issued without the prefix are always on-prem."