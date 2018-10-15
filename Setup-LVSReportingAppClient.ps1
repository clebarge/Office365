<#
Setup LVS Reporting App Client.

This script requires it run under Admin.
#>

#Set the execution policy to Remote Signed.

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

#Set the PSGallery as a trusted repository

Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

#Install Modules or Update Modules

$ImportExcel = get-module -Name ImportExcel -ListAvailable
IF(!$ImportExcel)
{
    Install-Module ImportExcel
}
ELSE
{
    Update-Module ImportExcel
}

$MSOnline = Get-Module -Name MSOnline -ListAvailable
IF(!$MSOnline)
{
    Install-Module MSOnline
}
ELSE
{
    Update-Module MSOnline
}

$AzureAD = Get-Module -Name AzureAD -ListAvailable
IF(!$AzureAD)
{
    Install-Module AzureAD
}
ELSE
{
    Update-Module AzureAD
}

$AzureRM = Get-Module -Name AzureRM -ListAvailable
IF(!$AzureRM)
{
    Install-Module AzureRM
}
ELSE
{
    Update-Module AzureRM
}