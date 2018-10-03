<#
Create-Office365LicensingReport.ps1

This script is based on the Get-UserAndMailboxStatistics script, with the focus on retreiving, analyzing, and generating a report on all license usage in an Office 365 tenancy.
This script does not include functionality for detailed service usage statistics.

This script has the following requirements:

    - Execution is set to RemoteSigned.
        * Requires Admin elevation to perform.
        * Set-ExecutionPolicy -ExecutionPolicy RemoteSigned.
    - Microsoft Online Services, MSOnline module.
        * Requires Admin elevation to perform.
        * Install-Module MSOnline.
        * Note: Currently the Azure AD 2.0 module, which is intended to replace MSOnline does not seem to support cloud service providers. Which is why the older commands are used.
    - Import-Excel
        *Install Module ImportExcel




Create-Office365LicensingReport.ps1
    [-CSPAllDomains] <switch>
    [-CSPCustomerDomain] <string>
    [-ReportType] <listoption> Summary | Full
    [-Path] <string>


    Version: 1.0.beta.10012018
    Author: Clark B. Lebarge
    Company: Long View Systems

#>

#Script Input Parameters
param(
[parameter(ParameterSetName="CSPD",Mandatory=$true,HelpMessage="The customer domain name, required for Cloud Service Partners.")]
[string]$CSPCustomerDomain,
[parameter(ParameterSetName="CSPA",Mandatory=$true,HelpMessage="For Cloud Service Partners, runs through all CSP customers.")]
[switch]$CSPAll,
[parameter(Mandatory=$false,HelpMessage="The folder path for the output Excel file. Current Directory if ommitted.")]
[string]$Path,
[parameter(Mandatory=$true,HelpMessage="Do you want a Summary or Full Report?")][validateset('Summary','Full')]
[string]$ReportType
)

#End Script Parameters

#Script Variables, these are declared and set here to ensure the values are not contaminated from existing PS variables from earlier runs of the script.
#Current Path
$CurrentPath=Split-Path $script:MyInvocation.MyCommand.Path
#Date
$Datefield = Get-Date -Format {MMddyyyy}
$TenantInfo = $null
$CompanyName = $null
$TenantId = $null
$MSOLUsers = $null
$MSOLUserCount = $null

#End Script Variables

#Script Body

    #Part 1, Login to MSOL, import SKU details, and set scope for CSP usage.
        #This section completes quickly, so no status information is displayed from this part of the script.

        #Import the SKU Details from Excel File
        #This is currently a local file, the data will need to be hosted somewhere or the file referenced included in the final version.
        #I'm thinking an XML file in Azure storage would work.

        $Plans = Import-Excel -Path C:\working\SKUDetails.xlsx

        #Logon to Microsoft Azure AD, with MSOL package.
        #We use MSOL package as the Azure AD 2.0 package still does not support CSP.

        Connect-MsolService

        #Scope for CSP can be all or domain, for normal Office 365 usage scope is always the current tenancy.
        IF($CSPAll)
        {
        $TenantInfo = Get-MsolPartnerContract | select Name,TenantId,DefaultDomainName

        $TenantCount = ($TenantInfo | measure).Count
        }
        ELSE
        {
    
            IF($CSPCustomerDomain)
            {
            $TenantInfo = Get-MsolPartnerContract -DomainName $CSPCustomerDomain | select Name,TenantId,DefaultDomainName
            $TenantCount = 1
            }
            ELSE
            {
            $TenantInfo = Get-MsolCompanyInformation | select @{N='Name';E={$_.DisplayName}}
            $TenantCount = 1
            }
        }
    
    #Part 2, Creation of the report for each tenant.
        #This is accomplished in a loop process for CSP usage.
        
        $TenantInfo | ForEach-Object -Begin {$I=0} -Process {
            $Tenant = $_

            #Part 2, Step 1, query Office 365 for user properties and save information on licensing to a temporary CSV file.

            #Set Variables to tenant info.
            $CompanyName = $Tenant.Name
            $TenantId = $Tenant.TenantId
            $PrimaryDomain = $Tenant.DefaultDomainName

            #Set the Temp File Name and Location.
            $Incr = 1
            $testpath = $null
            DO{
    
            $TempFilePath = "$Env:TEMP\$TenantId$Datefield$Incr.csv"
            $Incr = $Incr + 1
            $testpath = Test-Path $TempFilePath

            }
            UNTIL($testpath -eq $false)
            
            #For CSP usage the tenant ID is needed in the command.
            IF($TenantID)
            {
                $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
            }
            ELSE
            {
                $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
            }
            #Create the Temp File.

            #Bit of logic to deal with empty or test customers when running CSP.
            $MSOLUserCount = ($MSOLUsers | measure).Count

            #We take the raw information from Office 365 and replace license IDs with readable names and export to CSV each line.
                #We do this in a loop too!
                $MSOLUsers | ForEach-Object -Begin {$Ia=0} -Process {
                $user = $_
                #Set Login Status from Office 365.

                IF($user.BlockCredential -eq $false)
                {
                $LoginStatus = "Enabled"
                }
                ELSE
                {
                $LoginStatus = "Disabled"
                }

                #Set License Readable Name

                $Lics = $user.licenses.accountsku.skupartnumber

                IF($Lics)
                {
                    Foreach($lic in $lics)
                    {
                    $SkuName = $null
                    $SkuType = $null
                    $SkuRetail = $null
                    $SkuStatus = $null
                    $SkuPartNumber = $null

                    $SkuName = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuName
                    $SkuType = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuType
                    $SkuRetail = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail
                    $SkuStatus = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuStatus
                    $SkuPartNumber = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuPartNumber
    
 
    
                    #Output the raw statistics CSV file to the temp directory.

                    New-Object -TypeName PSObject -Property @{
                                        LoginStatus = $LoginStatus
                                        DisplayName = $user.DisplayName
                                        UserPrincipalName = $user.UserPrincipalName
                                        SkuName = $SkuName
                                        SkuType = $SkuType
                                        SkuRetail = $SkuRetail
                                        SkuStatus = $SkuStatus
                                        SkuPartNumber = $SkuPartNumber
                                        } | Export-Csv -NoTypeInformation -Path $TempFilePath -Append
                    }
                }
                ELSE
                {
                    $SkuName = "No Licenses"
                    $SkuType = ""
                    $SkuRetail = ""
                    $SkuStatus = ""
                    $SkuPartNumber = ""
                    New-Object -TypeName PSObject -Property @{
                                        LoginStatus = $LoginStatus
                                        DisplayName = $user.DisplayName
                                        UserPrincipalName = $user.UserPrincipalName
                                        SkuName = $SkuName
                                        SkuType = $SkuType
                                        SkuRetail = $SkuRetail
                                        SkuStatus = $SkuStatus
                                        SkuPartNumber = $SkuPartNumber
                                        } | Export-Csv -NoTypeInformation -Path $TempFilePath -Append
                }
                $Ia=$Ia+1
                Write-Progress -Activity "Querying $CompanyName Office 365 for user statitics." -Status "Sub-progress:" -PercentComplete ($Ia/$MSOLUserCount*100)
                } 
        
        #Part 2, Step 2, Create the Summary Report.

            #import in the temp file as a variable.
            $Tempfile = Import-Csv -Path $TempFilePath

            #Set the Output File Name and Location.
            $Incr = $null
            $testpath = $null
            DO{
            IF($Path)
            {
            $OutFilePath = "$Path\$CompanyName-$Datefield$Incr.xlsx"
            }
            ELSE
            {
            $OutFilePath = "$CurrentPath\$CompanyName-$Datefield$Incr.xlsx"
            }

            $Incr = $Incr + 1
            $testpath = Test-Path $OutFilePath

            }
            UNTIL($testpath -eq $false)
            
            #Copy the template XLSX file to the output path file name.
            Copy-Item -Path C:\working\Office365LicensingReport.xlsx -Destination $OutFilePath
            
            #Set the destination file as the Excel Package for the output.
            $Report = Open-ExcelPackage -Path $OutFilePath

            
            #Create Summary Sheet
                
                #Total Users
                $TotalUsers = ($Tempfile | Select-Object UserPrincipalName -Unique | Measure).Count
                #Disabled Users
                $DisabledUsers = ($Tempfile | where {$_.LoginStatus -eq "Disabled"} | Select-Object UserPrincipalName -Unique | measure).Count
                #Enabled Licensed Users
                $EnabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -ne "No Licenses")} | Select-Object UserPrincipalName -Unique | measure).Count
                #Disabled Licensed Users
                $DisabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -ne "No Licenses")} | Select-Object UserPrincipalName -Unique | measure).Count
                #Estimated Monthly Retail Total of all Licensing
                $EstRetailTotal =  (($Tempfile | where {$_.SKuName -ne "No Licenses"} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum
                #Estimated Monthly Retail Cost of Enabled Users Licensing
                $EstRetailEnabled = (($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -ne "No Licenses")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum
                #Estimated Monthly Retail Cost of Disabled Users Licensing
                $EstRetailDisabled = (($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -ne "No Licenses")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum

            $SummarySheet = $Report.Workbook.Worksheets["Summary"]
            Set-ExcelRange -WorkSheet $SummarySheet -Range B3 -Value $CompanyName
            Set-ExcelRange -WorkSheet $SummarySheet -Range B4 -Value $PrimaryDomain
            Set-ExcelRange -WorkSheet $SummarySheet -Range B5 -Value $TotalUsers
            Set-ExcelRange -WorkSheet $SummarySheet -Range B6 -Value $DisabledUsers
            Set-ExcelRange -WorkSheet $SummarySheet -Range B7 -Value $EnabledLicensedUsers
            Set-ExcelRange -WorkSheet $SummarySheet -Range B8 -Value $DisabledLicensedUsers
            Set-ExcelRange -WorkSheet $SummarySheet -Range B9 -Value $EstRetailTotal
            Set-ExcelRange -WorkSheet $SummarySheet -Range B10 -Value $EstRetailEnabled
            Set-ExcelRange -WorkSheet $SummarySheet -Range B11 -Value $EstRetailDisabled

            #Logic check to determine if a full report is requested.
            IF($ReportType -eq "Summary")
            {
                #Doing nothing!
            }
            ELSE
            {
                #THIS SECTION NEEDS HELP. TOO MANY LOOPS, NEED TO DO GOODER LOGIC!
                
                #We need to combine all licenses for each unique user to create readable reporting.
                $CombinedLicenses = @{}
                $Tempfile | ForEach-Object {
                    $line = $_
                    
                    $UPN = $line.UserPrincipalName

                    IF(-not $CombinedLicenses.ContainsKey($UPN))
                    {
                    $CombinedLicenses.Add($line.UserPrincipalName,$line.SkuName)
                    }
                    ELSE
                    {
                    $SkuLine = ($CombinedLicenses.item($line.UserPrincipalName) + ";" + $Line.SkuName)
                    $CombinedLicenses.Remove($line.UserPrincipalName)
                    $CombinedLicenses.add($line.UserPrincipalName,$SkuLine)
                    }
                }

                #And also for the RetailUSD of the licenses.
                $CombinedRetail = @{}
                $Tempfile | ForEach-Object {
                    $line = $_
                    
                    $UPN = $line.UserPrincipalName
                    [int32]$SkuRetail = $line.SkuRetail

                    IF(-not $CombinedRetail.ContainsKey($UPN))
                    {
                    $CombinedRetail.Add($UPN,$SkuRetail)
                    }
                    ELSE
                    {
                    $SkuLine = $CombinedRetail.item($UPN) + $SkuRetail
                    $CombinedRetail.Remove($UPN)
                    $CombinedRetail.add($UPN,$SkuLine)
                    }
                }

                            
            #Update the All_Users sheet.
                $All_UsersSheet = $Report.Workbook.Worksheets["All_Users"]
                $Row = 2
                Foreach($line in $CombinedLicenses.Keys)
                {
                $UPN = $line
                $Licenses = $CombinedLicenses.$UPN
                $DisplayName = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).DisplayName
                $LoginStatus = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).LoginStatus
                $EstRetailUSD = $CombinedRetail.Item($UPN)
                
                Set-ExcelRange -WorkSheet $All_UsersSheet -Value $DisplayName -Range ("A" + $Row)
                Set-ExcelRange -WorkSheet $All_UsersSheet -Value $UPN -Range ("B" + $Row)
                Set-ExcelRange -WorkSheet $All_UsersSheet -Value $LoginStatus -Range ("C" + $Row)
                Set-ExcelRange -WorkSheet $All_UsersSheet -Value $Licenses -Range ("D" + $Row)
                Set-ExcelRange -WorkSheet $All_UsersSheet -Value $EstRetailUSD -Range ("E" + $Row)

                $Row = $Row + 1

                }
            #Update the Licensed_Enabled sheet.
                $Licensed_EnabledSheet = $Report.Workbook.Worksheets["Licensed_Enabled"]
                $Row = 2
                Foreach($line in $CombinedLicenses.Keys)
                {
                $UPN = $line
                $Licenses = $CombinedLicenses.$UPN
                $DisplayName = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).DisplayName
                $EstRetailUSD = $CombinedRetail.Item($UPN)
                $LoginStatus = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).LoginStatus

                IF($LoginStatus -eq "Enabled" -and $Licenses -ne "No Licenses")
                {
                Set-ExcelRange -WorkSheet $Licensed_EnabledSheet -Value $DisplayName -Range ("A" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_EnabledSheet -Value $UPN -Range ("B" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_EnabledSheet -Value $LoginStatus -Range ("C" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_EnabledSheet -Value $Licenses -Range ("D" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_EnabledSheet -Value $EstRetailUSD -Range ("E" + $Row)

                $Row = $Row + 1
                }
                }

            #Update the Licensed_Disabled sheet.
                $Licensed_DisabledSheet = $Report.Workbook.Worksheets["Licensed_Disabled"]
                $Row = 2
                Foreach($line in $CombinedLicenses.Keys)
                {
                $UPN = $line
                $Licenses = $CombinedLicenses.$UPN
                $DisplayName = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).DisplayName
                $EstRetailUSD = $CombinedRetail.Item($UPN)
                $LoginStatus = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).LoginStatus
                IF($LoginStatus -eq "Disabled" -and $Licenses -ne "No Licensing")
                {
                Set-ExcelRange -WorkSheet $Licensed_DisabledSheet -Value $DisplayName -Range ("A" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_DisabledSheet -Value $UPN -Range ("B" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_DisabledSheet -Value $LoginStatus -Range ("C" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_DisabledSheet -Value $Licenses -Range ("D" + $Row)
                Set-ExcelRange -WorkSheet $Licensed_DisabledSheet -Value $EstRetailUSD -Range ("E" + $Row)

                $Row = $Row + 1
                }
                }

            #Update the Unlicensed_Enabled sheet.
                $UnLicensed_EnabledSheet = $Report.Workbook.Worksheets["UnLicensed_Enabled"]
                $Row = 2
                Foreach($line in $CombinedLicenses.Keys)
                {
                $UPN = $line
                $Licenses = $CombinedLicenses.$UPN
                $DisplayName = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).DisplayName
                $EstRetailUSD = $CombinedRetail.Item($UPN)
                $LoginStatus = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).LoginStatus
                IF($LoginStatus -eq "Enabled" -and $Licenses -eq "No Licenses")
                {
                Set-ExcelRange -WorkSheet $UnLicensed_EnabledSheet -Value $DisplayName -Range ("A" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_EnabledSheet -Value $UPN -Range ("B" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_EnabledSheet -Value $LoginStatus -Range ("C" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_EnabledSheet -Value $Licenses -Range ("D" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_EnabledSheet -Value $EstRetailUSD -Range ("E" + $Row)

                $Row = $Row + 1
                }
                }

            #Update the Unlicensed_Disabled sheet.
                $UnLicensed_DisabledSheet = $Report.Workbook.Worksheets["UnLicensed_Disabled"]
                $Row = 2
                Foreach($line in $CombinedLicenses.Keys)
                {
                $UPN = $line
                $Licenses = $CombinedLicenses.$UPN
                $DisplayName = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).DisplayName
                $EstRetailUSD = $CombinedRetail.Item($UPN)
                $LoginStatus = ($Tempfile | where {$_.UserPrincipalName -eq $UPN}).LoginStatus
                IF($LoginStatus -eq "Disabled" -and $Licenses -eq "No Licenses")
                {
                Set-ExcelRange -WorkSheet $UnLicensed_DisabledSheet -Value $DisplayName -Range ("A" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_DisabledSheet -Value $UPN -Range ("B" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_DisabledSheet -Value $LoginStatus -Range ("C" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_DisabledSheet -Value $Licenses -Range ("D" + $Row)
                Set-ExcelRange -WorkSheet $UnLicensed_DisabledSheet -Value $EstRetailUSD -Range ("E" + $Row)

                $Row = $Row + 1
                }
                }

            }
            #Save the report.
            Export-Excel -ExcelPackage $Report
        
        #Clean Up Temp File
        Remove-Item $TempFilePath -Confirm:$false
         
        #Hey look, a progress bar.
        $I = $I+1
        Write-Progress -Activity "Creating Office 365 Licensing Reports" -Status "Progress:" -PercentComplete ($I/$TenantCount*100)
        }

<#
Foreach($Tenant in $TenantInfo)
{
$CompanyName = $Tenant.Name
$TenantId = $Tenant.TenantId
Write-Host "Gathering information for Company: $CompanyName."

#Set the Temp File Name and Location.
    $Incr = 1
    $testpath = $null
    DO{
    
    $TempFilePath = "$Env:TEMP\$TenantId$Datefield$Incr.csv"
    $Incr = $Incr + 1
    $testpath = Test-Path $TempFilePath

    }
    UNTIL($testpath -eq $false) 

#Step 1
Write-Host "Executing Step 1: Query for user details from Office 365. Please wait."

IF($TenantID)
{
    $MSOLUsers = Get-MsolUser -EnabledFilter All -All -TenantId $TenantId | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
}
ELSE
{
    $MSOLUsers = Get-MsolUser -EnabledFilter All -All | Select DisplayName,BlockCredential,Licenses,UserPrincipalName
}

#Create the Temp File.

    #Bit of logic to deal with empty customers when running CSP.
    $MSOLUserCount = ($MSOLUsers | measure).Count
    IF($MSOLUserCount -lt 5)
    {
    Write-Host "Under 5 users, tenancy appears unused. Skipping."
    Continue
    }

foreach($user in $MSOLUsers)
{
    #Set Login Status from Office 365.
    IF($user.BlockCredential -eq $false)
    {
    $LoginStatus = "Enabled"
    }
    ELSE
    {
    $LoginStatus = "Disabled"
    }

    #Set License Readable Name

    $Lics = $user.licenses.accountsku.skupartnumber
    Foreach($lic in $lics)
    {
    $SkuName = $null
    $SkuType = $null
    $SkuRetail = $null
    $SkuStatus = $null

    $SkuName = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuName
    $SkuType = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuType
    $SkuRetail = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuRetail
    $SkuStatus = ($Plans | Where {$_.SkuPartNumber -eq $Lic}).SkuStatus
    
 
    
    #Output the raw statistics CSV file to the temp directory.

    New-Object -TypeName PSObject -Property @{
                        LoginStatus = $LoginStatus
                        DisplayName = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        SkuName = $SkuName
                        SkuType = $SkuType
                        SkuRetail = $SkuRetail
                        SkuStatus = $SkuStatus
                        } | Export-Csv -NoTypeInformation -Path $TempFilePath -Append
    }
}

#Step 2
Write-Host "Step 2: Creating Report."

$Tempfile = Import-Csv -Path $TempFilePath

#Set the Output File Name and Location.
    $Incr = 1
    $testpath = $null
    DO{
    IF($Path)
    {
    $OutFilePath = "$Path\$CompanyName$Datefield$Incr.xlsx"
    }
    ELSE
    {
    $OutFilePath = "$CurrentPath\$CompanyName$Datefield$Incr.xlsx"
    }

    $Incr = $Incr + 1
    $testpath = Test-Path $OutFilePath

    }
    UNTIL($testpath -eq $false) 


#Create Summary Sheet
Write-Host "Creating Summary Sheet."
    #Total Users
    $TotalUsers = ($Tempfile | Select-Object UserPrincipalName -Unique | Measure).Count
    #Disabled Users
    $DisabledUsers = ($Tempfile | where {$_.LoginStatus -eq "Disabled"} | Select-Object UserPrincipalName -Unique | measure).Count
    #Enabled Licensed Users
    $EnabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
    #Disabled Licensed Users
    $DisabledLicensedUsers = ($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -like "*")} | Select-Object UserPrincipalName -Unique | measure).Count
    #Estimated Monthly Retail Cost of Enabled Users Licensing
    $EstRetailEnabled = (($Tempfile | where {($_.LoginStatus -eq "Enabled") -and ($_.SKuName -like "*")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum
    #Estimated Monthly Retail Cost of Disabled Users Licensing
    $EstRetailDisabled = (($Tempfile | where {($_.LoginStatus -eq "Disabled") -and ($_.SKuName -like "*")} | Select-Object SkuRetail).SkuRetail | Measure -sum).Sum

#Export Summary Sheet
    New-Object -TypeName PSObject -Property @{
        TotalUsers = $TotalUsers
        DisabledUsers = $DisabledUsers
        EnabledLicensedUsers = $EnabledLicensedUsers
        DisabledLicensedUsers = $DisabledLicensedUsers
        EstRetailEnabled = $EstRetailEnabled
        EstRetailDisabled = $EstRetailDisabled
    } | export-excel -Path $OutFilePath -WorksheetName "Summary"

#Create All Users Sheet
$AllUsersSheet = New-Object System.Data.DataTable
$AllUsersSheet.Columns.Add("DisplayName","string")
$AllUsersSheet.Columns.Add("UserPrincipalName","string")
$AllUsersSheet.PrimaryKey = $AllUsersSheet.Columns[1]
$AllUsersSheet.Columns.Add("Licenses","string")
$AllUsersSheet.Columns.Add("LoginStatus","string")
$AllUsersSheet.Columns.Add("EstimatedRetailCostUSD","int32")


Write-Host "Creating Full Report"

$Tempfile | ForEach-Object {
    $line = $_
    #Check if UPN already in AllUsersSheet

    $testline = $AllUsersSheet | where {$_.UserPrincipalName -eq $line.UserPrincipalName}
    IF(-not $testline)
    {
    $r = $AllUsersSheet.NewRow()
    $r.DisplayName = $line.DisplayName
    $r.UserPrincipalName = $line.userprincipalname
    $r.Licenses = $line.SkuName
    $r.LoginStatus = $line.LoginStatus
    $r.EstimatedRetailCostUSD = [int32]$line.SkuRetail
    $AllUsersSheet.Rows.Add($r)
    }
    ELSE
    {
    $AllUsersSheet | where {$_.UserPrincipalName -eq $line.UserPrincipalName} | foreach {
        $_.Licenses = ($_.Licenses + ";" + $line.SkuName)
        $_.EstimatedRetailCostUSD = $_.EstimatedRetailCostUSD + $line.SkuRetail
        }
    }
} 

$AllUsersSheet | Export-Excel -Path $OutFilePath -WorksheetName "All Users"

#Create the Disabled Users with Licenses Sheet

$DisUserWithLicenseSheet = $AllUsersSheet | Where {$_.LoginStatus -eq "Disabled"}

$DisUserWithLicenseSheet | Export-Excel -Path $OutFilePath -WorksheetName "Disabled Users"


#Clean Up Temp File
Remove-Item $TempFilePath -Confirm:$false

$TenantId = $null
$MSOLUsers = $null
}
#>


write-host "Job's done, Boss."
