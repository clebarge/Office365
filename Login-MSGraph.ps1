<#
Login to Microsoft Graph
This script is intended for testing to determine the minimum code required to connect to Microsoft Graph through PowerShell.

About Graph
Graph is actually an Application Programing Interface (API), not a service in itself, as such you do not actually connect to Graph from within Powershell.
Instead, you need to first create an Application, which will interface to Graph on your behalf.
Microsoft has made Graph available through the Azure AD App Registration system, there are several options currently available which may work.
    Using REST API to connect to AzureAD Native App connecting to Graph.
    Using REST API to connect to AzureAD 2.0 Native App connecting to Graph.
    Using Web API to connect to AzureAD Web App connecting to Graph.
    Using Web API to connect to AzureAD 2.0 Web App connecitng to Graph.

    As this is a new project, the focus is on the most current and first to be understood Azure AD 2.0 Native App with REST API.

Because Graph is an API, in general if you need to pull information from Office 365, Azure, or other service using the web interface Graph Explorer or implementing a connection
through Power BI will be preferable. The most likely reason to use PowerShell is to have scripted automation in response to information in Graph, or perhaps in IF this then That situations.
#>