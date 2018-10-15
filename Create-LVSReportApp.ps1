<#
Create the Certificate for LongView Reporting App.

#>

$AzureAD = Connect-AzureAD

$TenantID = $AzureAD.TenantId

# Define certificate start and end dates

$currentDate =
Get-Date

$endDate =
$currentDate.AddYears(2)

$notAfter =
$endDate.AddYears(2)

# Generate new self-signed certificate from "Run as Administrator" PowerShell session

$dnsName = "https://LongView-ReportingApp"

$certStore =
"Cert:\LocalMachine\My"

$certThumbprint =
(New-SelfSignedCertificate `
-DnsName "$dnsName" `
-CertStoreLocation $CertStore `
-KeyExportPolicy Exportable `
-Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
-NotAfter $notAfter `
-Subject "CN=$TenantID").Thumbprint

$pfxPassword =
Read-Host `
-Prompt "Enter password to protect exported certificate:" `
-AsSecureString

$pfxFilepath =
Read-Host `
-Prompt "Enter full path to export certificate (ex C:\folder\filename.pfx)"


Export-PfxCertificate `
-Cert "$($certStore)\$($certThumbprint)" `
-FilePath $pfxFilepath `
-Password $pfxPassword

Import-PfxCertificate -FilePath $pfxFilepath -Password $pfxPassword -CertStoreLocation "Cert:CurrentUser\My"

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate($pfxFilepath, $pfxPassword)
$keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())

$application = New-AzureADApplication -DisplayName "LongView-ReportingApp" -IdentifierUris $dnsName -ReplyUrls $dnsName
New-AzureADApplicationKeyCredential -ObjectId $application.ObjectId -CustomKeyIdentifier "$dnsName" -Type AsymmetricX509Cert -Usage Verify -Value $keyValue

$sp=New-AzureADServicePrincipal -AppId $application.AppId

Add-AzureADDirectoryRoleMember -ObjectId (Get-AzureADDirectoryRole | where-object {$_.DisplayName -eq "Directory Readers"}).Objectid -RefObjectId $sp.ObjectId