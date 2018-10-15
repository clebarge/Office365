Function Connect-WebApp
{
#This function connects to an already configured WebApp using stored certificate. The certificate name is intended to match the Tenant ID.
param(
[parameter()]$TenantID,
[parameter()]$AppId,
[parameter()]$resourceAppIdURI
)

$certThumbprint = (Get-ChildItem -Path Cert:\LocalMachine\My | where {$_.Subject -match $TenantId}).Thumbprint


$AADApp = Connect-AzureAD -CertificateThumbprint $certThumbprint -TenantId $TenantID -ApplicationId $AppId
$TenantName = $AADApp.TenantDomain
$redirectUri = $AADApp.ReplyUrls.Item(0)
    
    $CredPrompt         = "Auto"
    $authority          = "https://login.microsoftonline.com/$TenantName"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
    $AccessToken        = $authContext.AcquireTokenAsync($resourceAppIdURI, $AppId, $redirectUri, $platformParameters).Result
    $WebAppHeader = @{
        'Authorization' = $AccessToken.CreateAuthorizationHeader()
        }
return $WebAppHeader
}

$TenantId = "413cc1cd-c360-4ebb-a814-7e0c629fc4c1"
$AppId = "7882a4d1-ee1a-41aa-9ed0-f3d498fc3fd6"
$ResourceAppIDURI = "https://graph.microsoft.com"

Connect-WebApp -TenantID $TenantId -AppId $AppId -resourceAppIdURI $ResourceAppIDURI