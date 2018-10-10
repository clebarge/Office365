<#
Login to Microsoft Graph.
This script is intended for testing to determine the minimum code required to connect to Microsoft Graph through PowerShell.

Information provided by Microsoft at https://blogs.technet.microsoft.com/cloudlojik/2018/06/29/connecting-to-microsoft-graph-with-a-native-app-using-powershell/ forms
the basis of this work.

About Graph
Graph is actually an Application Programing Interface (API), not a service in itself, as such you do not actually connect to Graph from within Powershell.
Instead, you need to first create an Application, which will interface to Graph on your behalf.
Microsoft has made Graph available through the Azure AD App Registration system, there are several options currently available which may work.
    Using REST API to connect to AzureAD Native App connecting to Graph.
    Using REST API to connect to AzureAD 2.0 Native App connecting to Graph.
    Using Web API to connect to AzureAD Web App connecting to Graph.
    Using Web API to connect to AzureAD 2.0 Web App connecitng to Graph.

    As this is a new project, the focus is on the most current and first to be understood Azure AD 2.0 Native App with REST API.
    Please follow the instructions in document "Login-MSGraph Create Azure AD 2.0 App Registration.docx."

Because Graph is an API, in general if you need to pull information from Office 365 and Azure AD using the web interface Graph Explorer or implementing a connection
through Power BI will be preferable. The most likely reason to use PowerShell is to have scripted automation in response to information in Graph, or perhaps in IF this then That situations.

Login-MSGraph
    [-AppName] <string>

This script only logs on and connects to Graph. It performs no other actions.

#>

#Command Line Parameters
param(
[parameter(Mandatory=$false,HelpMessage="Specify an App Name, if omitted default is LoginGraphAAD2")][string]$AppName="LoginGraphAAD2"

)

#Required Variables
    
    #This variable is used in the creation of the access token for the content in Graph.
    $resourceAppIdURI = "https://graph.microsoft.com"

    #Credential Prompt Behaviour. This value modifies how, the connection to Graph will deal with credentials.
        #Auto - This is the default behaviour, if credentials already exist, it uses the existing credential. It will prompt if more than one credential exists that are valid.
        #Always - You'll always be prompted to pick a login or provide a login.
        #Never - You won't be prompted, however if more than one credential exists it will fail silently.
        #RefreshSession - For when/if your session has expired, usually after 1 hour of inactivity.
    $CredPrompt       = "Auto"                                   #Auto, Always, Never, RefreshSession

###Begin Body###

#The connection requires the Application's ID
    $AppId = (Get-AzureADApplication -SearchString $AppName).AppId

#We need to connect, or to have connected already to the Azure AD 2.0 endpoint.
    
    #This code is a bit of advanced powershell scripting I don't fully understand. Taken from example provided by Microsoft.
    #Essentially, rather than running Connect-AzureAD, then running the access token request, this loads only the modules necessary.
    #If you have already logged in with Connect-AzureAD it may prompt you to select the login again.
    #If you have not logged on, it should prompt you to login to Azure AD.
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    if ($AadModule -eq $null) {$AadModule = Get-Module -Name "AzureADPreview" -ListAvailable}
    if ($AadModule -eq $null) {write-host "AzureAD Powershell module is not installed. The module can be installed by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt. Stopping." -f Yellow;exit}
    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
        $aadModule      = $AadModule | ? { $_.version -eq $Latest_Version.version }
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    else {
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

    #These lines build the Access Token. As this script is intended to be used as test/basis for other scripts the GraphHeader is saved to a Global variable.
    $authority          = "https://login.microsoftonline.com/$TenantName"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
    $AccessToken        = $authContext.AcquireTokenAsync($resourceAppIdURI, $AppId, $redirectUri, $platformParameters).Result
    $Global:GraphHeader = @{
        'Authorization' = $AccessToken.CreateAuthorizationHeader()
        }