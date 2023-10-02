﻿#.SYNOPSIS
 # <Windows_11_Compliance_Dashboard_&_Report_Email_Subscription Using PowerShell>
#.DESCRIPTION
 # <Windows_11_Compliance_Dashboard_&_Report_Email_Subscription Using PowerShell>
#.Demo
#<YouTube video link-->https://www.youtube.com/@ChanderManiPandey
#.INPUTS
 # <Provide all required inforamtion in User Input Section-line No 96-105 & 142-145>
#.OUTPUTS
 # <You will get Windows_11_Compliance_Dashboard_&_Report_Email_Subscription + report in CSV>
#.NOTES
 <# Version:       1.0
  Author:          Chander Mani Pandey
  Creation Date:   28 Sep 2023
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
  Reddit:-         https://www.reddit.com/u/ChanderManiPandey 
 #>

cls
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' 
$error.clear() ## this is the clear error history 
#=====================================================User Input Section Start=============================================================================================
$From          = "cmpdemouser@gmail.com"
$To            = "cmpdemouser@gmail.com"
$CC            = "cmpdemouser@gmail.com"
$SmtpServer    = "smtp.gmail.com"
$Port          = '587'
$Priority      = "Normal"
$Subject     = ""
$tenant = “IntuneLab1.onmicrosoft.com”
$clientId = “10851d04-aea6-4ab2-87b5-e1c3f38baa”
$clientSecret = “S8t8Q~RG.PAzRVFnbm~rgue2VYAJY0Fpv1Ddt_”
$WorkingFolder = "C:\TEMP\Windows11Report"
$FinalReportPath = "$WorkingFolder\Windows11_Compliance_Report.csv"
#=====================================================User Input Section End===============================================================================================
$startTime = Get-Date
Write-Host "===============================Phase-1 (Exporting Intune Device Dump) =========================================================(Started)" -ForegroundColor Green
$futureDate = Get-Date -Year 2025 -Month 10 -Day 14 -Hour 00 -Minute 0 -Second 0
$currentDate = Get-Date
$daysLeft = ($futureDate - $currentDate).Days
#Creating working Folder
$null = New-Item -ItemType Directory -Path $WorkingFolder -Force -ErrorAction SilentlyContinue
$Path ="$WorkingFolder\IntuneDump\"
$MGIModule = Get-module -Name "Microsoft.Graph.Intune" -ListAvailable
Write-Host "Checking Microsoft.Graph.Intune is Installed or Not"
    If ($MGIModule -eq $null) 
    {
        Write-Host "Microsoft.Graph.Intune module is not Installed"
        Write-Host "Installing Microsoft.Graph.Intune module"
        Install-Module -Name Microsoft.Graph.Intune -Force
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force
    }
    ELSE 
    {   Write-Host "Microsoft.Graph.Intune is Installed"
        Write-Host "Importing Microsoft.Graph.Intune module"
        Import-Module Microsoft.Graph.Intune -Force
    }
$tenant = $tenant
$authority = “https://login.windows.net/$tenant”
$clientId = $clientId  
$clientSecret = $clientSecret
Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $ClientSecret -Quiet
Update-MSGraphEnvironment -SchemaVersion "Beta" -Quiet -InformationAction SilentlyContinue
#============Create Request Body==========================================================================================================================================================
$postBody = @{
 'reportName' = "DevicesWithInventory"
 'filter' = "(DeviceType eq '1') "
 'select' =  ("DeviceId"),("SerialNumber"), ("DeviceName") ,("ownerType") ,("OSVersion"),("UPN"),("LastContact"),("JoinType"),("Manufacturer"),("Model"),("ManagementAgent"),("SkuFamily"),("StorageTotal") ,("StorageFree")
  }
#=========== MakeRequest ==================================================================================================================================================================
$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody
Write-Host "Export Job initiated for $ReportName Report "
#====================================Checking Report Ready status==========================================================================================================================
do{ 
$exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Start-sleep -second 2
    Write-Host -NoNewline '...........'
  } while ($exportJob.status -eq 'inprogress')
  Write-Host 'Report is in Ready(Completed) status for Downloading' -ForegroundColor White
  If ($exportJob.status -eq 'completed') 
  { $fileName = (Split-path -Path $exportJob.url -Leaf).split('?')[0]
  Write-host "Export Job completed.......  Writing File $fileName to Disk........" -ForegroundColor White
  Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile $fileName
# Check if the directory exists before attempting to remove items
if (Test-Path -Path $Path -PathType Container) {
    Remove-Item –Path "$Path\*" -Include *.csv
    Write-Host "CSV files in $Path removed successfully."
} else {
    Write-Host "Directory $Path does not exist, no items to remove."
}
  Expand-Archive -Path $fileName -DestinationPath $Path 
  $FileName = Get-ChildItem -Path $Path* -Include *.csv | Where {! $_.PSIsContainer } | Select Name,FullName
  $DevicesInfos = import-csv -Path $FileName.fullName
} 
Write-Host "===============================Phase-1 (Exporting Intune Device Dump) =======================================================(Completed)" -ForegroundColor Green
Write-Host "" 
Write-Host "===============================Phase-2 (Genrating Windows 11 compliance Dashboard) ============================================(Started)" -ForegroundColor Green
Write-Host "Genrating Windows 11 compliance Dashboard"  -ForegroundColor White
# Define the end dates for each windows version
$Version_22H2 = "14-Oct-25";$Version_21H2 = "11-Jun-24";$Version_21H1 = "13-Dec-22";$Version_20H2 = "9-May-23";$Version_2004 = "14-Dec-21";$Version_1909 = "10-May-22";$Version_1903 = "8-Dec-20"
$Version_1809 = "11-May-21";$Version_1803 = "11-May-21";$Version_1709 = "13-Oct-20";$Version_1703 = "8-Oct-19";$Version_1607 = "9-Apr-19";$Version_1511 = "10-Oct-17";$Version_1507 = "9-May-17" ; $Version_7601 = "Win7"
#Invoke-Item -path $FileName.FullName
$Win10_1507 = 0;$Win10_1511 = 0;$Win10_1607 = 0;$Win10_1703 = 0;$Win10_1709 = 0;$Win10_1803 = 0;$Win10_1809 = 0;$Win10_1903 = 0;$Win10_1909 = 0;$Win10_2004 = 0;$Win10_20H2 = 0
$Win10_21H1 = 0;$Win10_21H2 = 0;$Win10_22H2 = 0;$Win11 = 0;$Win7 = 0;$No_OS_version = 0;$Unknown_OS_Version = 0
$totalDevices = $DevicesInfos.Count

$progress = 0
$Win11Info = @()
foreach ($DevicesInfo in $DevicesInfos) 
{ 
  $Win11InfoHSProps = [ordered] @{
    DeviceID        = $DevicesInfo."Device ID"
    Serialnumber    = $DevicesInfo."Serial number"
    Devicename      = $DevicesInfo."Device name"
    Ownership       = $DevicesInfo."Ownership"
    OSversion       = $DevicesInfo."OS version"
    PrimaryuserUPN  = $DevicesInfo."Primary user UPN"
    Lastcheckin     = $DevicesInfo."Last check-in"
    JoinType        = $DevicesInfo."JoinType"
    Manufacturer    = $DevicesInfo."Manufacturer"
    Model           = $DevicesInfo."Model"
    Managedby       = $DevicesInfo."Managed by"
    SkuFamily       = $DevicesInfo."SkuFamily"
    Totalstorage    = $DevicesInfo."Total storage"
    Freestorage     = $DevicesInfo."Free storage"
    Build           = $OSVersion = $DevicesInfo."OS version".Split(".")[2]
    Build_OS        = If ($OSVersion -eq '10240') {'Win10-1507';$Win10_1507++ } 
                        ElseIf ($OSVersion -eq '10586') {'Win10-1511';$Win10_1511++} ElseIf ($OSVersion -eq '14393') {'Win10-1607';$Win10_1607++} 
                        ElseIf ($OSVersion -eq '15063') {'Win10-1703';$Win10_1703++} ElseIf ($OSVersion -eq '16299') {'Win10-1709';$Win10_1709++}
                        ElseIf ($OSVersion -eq '17134') {'Win10-1803';$Win10_1803++} ElseIf ($OSVersion -eq '17763') {'Win10-1809';$Win10_1809++} 
                        ElseIf ($OSVersion -eq '18362') {'Win10-1903';$Win10_1903++} ElseIf ($OSVersion -eq '18363') {'Win10-1909';$Win10_1909++}
                        ElseIf ($OSVersion -eq '19041') {'Win10-2004';$Win10_2004++} ElseIf ($OSVersion -eq '19042') {'Win10-20H2';$Win10_20H2++} 
                        ElseIf ($OSVersion -eq '19043') {'Win10-21H1';$Win10_21H1++} ElseIf ($OSVersion -eq '19044') {'Win10-21H2';$Win10_21H2++} 
                        ElseIf ($OSVersion -eq '19045') {'Win10-22H2';$Win10_22H2++} ElseIf ($OSVersion -eq '7601') {'Win7';$Win7++} 
                        ElseIf ($OSVersion -ge '22000') {'Win11';$Win11++} ElseIf ($OSVersion -eq '0') {'0.0.0.0';$Unknown_OS_Version++} 
                        ElseIf ($OSVersion -eq $null) {'No OS version';$No_OS_version++} Else { 'Unknown_OS_Version';$Unknown_OS_Version++}
    EOL_Date       =  $EOL
  }
  # Populate the EOL_Date property based on OSVersion
  switch ($OSVersion) {'10240' { $Win11InfoHSProps.EOL_Date = $Version_1507 }'10586' { $Win11InfoHSProps.EOL_Date = $Version_1511 }'14393' { $Win11InfoHSProps.EOL_Date = $Version_1607 }
    '15063' { $Win11InfoHSProps.EOL_Date = $Version_1703 }'16299' { $Win11InfoHSProps.EOL_Date = $Version_1709 }'17134' { $Win11InfoHSProps.EOL_Date = $Version_1803 }
    '17763' { $Win11InfoHSProps.EOL_Date = $Version_1809 }'18362' { $Win11InfoHSProps.EOL_Date = $Version_1903 }'18363' { $Win11InfoHSProps.EOL_Date = $Version_1909 }
    '19041' { $Win11InfoHSProps.EOL_Date = $Version_2004 }'19042' { $Win11InfoHSProps.EOL_Date = $Version_2004 }'19043' { $Win11InfoHSProps.EOL_Date = $Version_2004 }
    '19044' { $Win11InfoHSProps.EOL_Date = $Version_2004 }'19045' { $Win11InfoHSProps.EOL_Date = $Version_22H2 }'Win7' { $Win11InfoHSProps.EOL_Date = $Version_7601 }
    {$_ -ge '22000'} { $Win11InfoHSProps.EOL_Date = 'AlreadyOnWin11' }'0' { $Win11InfoHSProps.EOL_Date = "VersionMissing" }'7601' { $Win11InfoHSProps.EOL_Date = $Version_7601 }
    default { $Win11InfoHSProps.EOL_Date = "VersionMissing" }
  }
  $Win11InfoHSobject = New-Object -Type PSObject -Property $Win11InfoHSProps
  $Win11Info += $Win11InfoHSobject
  $progress++
  $percentComplete = [String]::Format("{0:0.00}", ($progress / $totalDevices) * 100)
  Write-Progress -Activity "Generating Windows11 Compliance Report" -Status "Progress: $percentComplete% Complete" -PercentComplete $percentComplete
}


$FinalReport = $Win11Info | Select-Object Devicename,DeviceID,Serialnumber,Ownership,OSversion,PrimaryuserUPN,Lastcheckin,JoinType,Manufacturer,Model,Managedby,SkuFamily,Totalstorage,Freestorage,Build,Build_OS,EOL_Date
$FinalReport | Export-Csv -path $FinalReportPath -NoTypeInformation
$TotalDevice = $DevicesInfos.Count
$TotalWin11Devcies= $Win11 
$TotalNonwin11Device= $DevicesInfos.Count - $Win11 
$Win11Compliance = ($TotalWin11Devcies / $TotalDevice)*100
$Win11Compliance = [math]::Round($Win11Compliance, 2)
$NoInfo = $No_OS_version + $Unknown_OS_Version
$PerShareWin11 = [math]::Round(($Win11 / $TotalDevice) * 100, 2)
$PerShareWin10_22H2 = [math]::Round(($Win10_22H2 / $TotalDevice) * 100, 2);$PerShareWin10_21H2= [math]::Round(($Win10_21H2/ $TotalDevice) * 100, 2)
$PerShareWin10_21H1= [math]::Round(($Win10_21H1 / $TotalDevice) * 100, 2);$PerShareWin10_20H2 = [math]::Round(($Win10_20H2 / $TotalDevice) * 100, 2)
$PerShareWin10_2004 = [math]::Round(($Win10_2004 / $TotalDevice) * 100, 2);$PerShareWin10_1909 = [math]::Round(($Win10_1909 / $TotalDevice) * 100, 2)
$PerShareWin10_1903 = [math]::Round(($Win10_1903 / $TotalDevice) * 100, 2);$PerShareWin10_1809 = [math]::Round(($Win10_1809 / $TotalDevice) * 100, 2)
$PerShareWin10_1803 = [math]::Round(($Win10_1803 / $TotalDevice) * 100, 2);$PerShareWin10_1709 = [math]::Round(($Win10_1709 / $TotalDevice) * 100, 2)
$PerShareWin10_1703 = [math]::Round(($Win10_1703 / $TotalDevice) * 100, 2);$PerShareWin10_1607 = [math]::Round(($Win10_1607 / $TotalDevice) * 100, 2)
$PerShareWin10_1511 = [math]::Round(($Win10_1511 / $TotalDevice) * 100, 2);$PerShareWin10_1507 = [math]::Round(($Win10_1507 / $TotalDevice) * 100, 2)
$PerShareWin7  = [math]::Round(($Win7 / $TotalDevice) * 100, 2);$PerShareNoInfo = [math]::Round(($NoInfo / $TotalDevice) * 100, 2)
$NonWin11 = [math]::Round((($DevicesInfos.Count-$TotalWin11Devcies)/$DevicesInfos.Count) * 100, 2);
# Get the current date
$CurrentDate = Get-Date
# Calculate the number of days left for each version and store in a variable
$DaysLeft_Version_22H2 = (Get-Date $Version_22H2) - $CurrentDate;$DaysLeft_Version_21H2 = (Get-Date $Version_21H2) - $CurrentDate;$DaysLeft_Version_21H1 = (Get-Date $Version_21H1) - $CurrentDate
$DaysLeft_Version_20H2 = (Get-Date $Version_20H2) - $CurrentDate;$DaysLeft_Version_2004 = (Get-Date $Version_2004) - $CurrentDate;$DaysLeft_Version_1909 = (Get-Date $Version_1909) - $CurrentDate
$DaysLeft_Version_1903 = (Get-Date $Version_1903) - $CurrentDate;$DaysLeft_Version_1809 = (Get-Date $Version_1809) - $CurrentDate;$DaysLeft_Version_1803 = (Get-Date $Version_1803) - $CurrentDate
$DaysLeft_Version_1709 = (Get-Date $Version_1709) - $CurrentDate;$DaysLeft_Version_1703 = (Get-Date $Version_1703) - $CurrentDate;$DaysLeft_Version_1607 = (Get-Date $Version_1607) - $CurrentDate
$DaysLeft_Version_1511 = (Get-Date $Version_1511) - $CurrentDate;$DaysLeft_Version_1507 = (Get-Date $Version_1507) - $CurrentDate
# Convert negative days to "AlreadyEol"
$DaysLeft_Version_22H2 = If ($DaysLeft_Version_22H2.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_22H2.Days };$DaysLeft_Version_21H2 = If ($DaysLeft_Version_21H2.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_21H2.Days }
$DaysLeft_Version_21H1 = If ($DaysLeft_Version_21H1.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_21H1.Days };$DaysLeft_Version_20H2 = If ($DaysLeft_Version_20H2.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_20H2.Days }
$DaysLeft_Version_2004 = If ($DaysLeft_Version_2004.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_2004.Days };$DaysLeft_Version_1909 = If ($DaysLeft_Version_1909.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1909.Days }
$DaysLeft_Version_1903 = If ($DaysLeft_Version_1903.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1903.Days };$DaysLeft_Version_1809 = If ($DaysLeft_Version_1809.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1809.Days }
$DaysLeft_Version_1803 = If ($DaysLeft_Version_1803.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1803.Days };$DaysLeft_Version_1709 = If ($DaysLeft_Version_1709.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1709.Days }
$DaysLeft_Version_1703 = If ($DaysLeft_Version_1703.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1703.Days };$DaysLeft_Version_1607 = If ($DaysLeft_Version_1607.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1607.Days }
$DaysLeft_Version_1511 = If ($DaysLeft_Version_1511.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1511.Days };$DaysLeft_Version_1507 = If ($DaysLeft_Version_1507.TotalDays -lt 0) { "AlreadyEol" } Else { $DaysLeft_Version_1507.Days }

Write-Host ""
Write-host "Over All Status" -ForegroundColor Yellow
Write-Host ""
Write-host "Total Devices:-          "$DevicesInfos.Count 
Write-host "Total Win11 Devices:-    "$TotalWin11Devcies" ("$PerShareWin11"%)"
Write-host "Total Non Win11 Device:- "$TotalNonwin11Device" ("$NonWin11"%)"
Write-host "Win11 % Share:-            "$PerShareWin11"%" 
Write-Host ""
Write-Host "Operating System Wise Percentage (%) Share" -ForegroundColor Yellow
Write-Host ""
Write-Host "Device On Windows 11:-     "$Win11" ("$PerShareWin11"%)"
Write-Host "Device On Windows 10 22H2:-"$Win10_22H2" ("$PerShareWin10_22H2"%)"
Write-Host "Device On Windows 10 21H2:-"$Win10_21H2" ("$PerShareWin10_21H2"%)"
Write-Host "Device On Windows 10 21H1:-"$Win10_21H1" ("$PerShareWin10_21H1"%)"
Write-Host "Device On Windows 10 20H2:-"$Win10_20H2" ("$PerShareWin10_20H2"%)"
Write-Host "Device On Windows 10 2004:-"$Win10_2004" ("$PerShareWin10_2004"%)"
Write-Host "Device On Windows 10 1909:-"$Win10_1909" ("$PerShareWin10_1909"%)"
Write-Host "Device On Windows 10 1903:-"$Win10_1903" ("$PerShareWin10_1903"%)"
Write-Host "Device On Windows 10 1809:-"$Win10_1809" ("$PerShareWin10_1809"%)"
Write-Host "Device On Windows 10 1803:-"$Win10_1803" ("$PerShareWin10_1803"%)"
Write-Host "Device On Windows 10 1709:-"$Win10_1709" ("$PerShareWin10_1709"%)"
Write-Host "Device On Windows 10 1703:-"$Win10_1703" ("$PerShareWin10_1703"%)"
Write-Host "Device On Windows 10 1607:-"$Win10_1607" ("$PerShareWin10_1607"%)"
Write-Host "Device On Windows 10 1511:-"$Win10_1511" ("$PerShareWin10_1511"%)"
Write-Host "Device On Windows 10 1507:-"$Win10_1507" ("$perShareWin10_1507"%)"
Write-Host "Device On Win7:-           "$Win7" ("$PerShareWin7"%)"
Write-Host "Device On NoOSInfo:-       "$NoInfo" ("$PerShareNoInfo"%)"

Write-Host "===============================Phase-2 (Genrating Windows 11 compliance Dashboard) ==========================================(Completed)" -ForegroundColor Green
Write-Host ""
Write-Host "===============================Phase-3 (Sending E-mail(Windows Patching Compliance Report) ====================================(Started)" -ForegroundColor Green
Write-Host "Sending E-mail Notification"  -ForegroundColor White
$NotificationBody = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows 11 Status</title>
    <style>
        table {
            border-collapse: collapse;
            width: 60%;
        }
        th, td {
            border: 1px solid black;
            text-align: left;
            padding: 0px;
        }
        .black-bg {
            background-color: black;
            color: white;
        }
        .lightgray-bg {
            background-color: lightgray;
        }
        .lightcoral-bg {
            background-color: lightcoral;
        }
        .lightblue-bg {
            background-color: lightblue;
        }
        .lightgreen-bg {
            background-color: lightgreen;
        }
   </style>
</head>
<body>
    <p>Hello All,</p>
    <p>Windows 10 will reach end of support on October 14, 2025. The current version, 22H2, will be the final version of Windows 10, and all editions will remain in support with monthly security update releases through that date. Existing LTSC releases will continue to receive updates beyond that date based on their specific lifecycles.</p>
    <h2 style="color: #FF5733;">Countdown to Windows 10 EOL: October 14, 2025</h2>
    <table>
        <tr>
            <tr>
            <td style="border: 1px solid #dddddd; padding: 8px; background-color: black; color: white;">Days Remaining</td>
            <td style="border: 1px solid #dddddd; padding: 8px; background-color: black; color: white; text-align: center;">$daysLeft</td>
</tr>
</tr>
        </table>
<p><strong><u>Over All Status</u></strong></p>
<table>
    <tr style="background-color: lightgray;">
        <td><strong>Total Devices</strong></td>
        <td style="text-align: center;">$TotalDevice</td>
    </tr>
    <tr style="background-color: lightgreen;">
        <td>Device On Win11</td>
        <td style="text-align: center;">$Win11</td>
    </tr>
    <tr style="background-color: lightcoral;">
        <td>Devices below Win11</td>
        <td style="text-align: center;">$TotalNonwin11Device</td>
    </tr>
    <tr style="background-color: lightblue;">
        <td><strong>Win11 Compliance (%)</strong></td>
        <td style="text-align: center;">$Win11Compliance%</td>
    </tr>
</table>
<p><strong><u>Operating System Wise Percentage (%) Share</u></strong></p>
       <table>
           <tr class="lightblue-bg">
            <td style="text-align: center;"><strong>OperatingSystem</strong></td>
            <td style="text-align: center;"><strong>Device Count</strong></td>
            <td style="text-align: center;"><strong>EOL Date</strong></td>
            <td style="text-align: center;"><strong>OS % Share</strong></td>
            <td style="text-align: center;"><strong>Duration left</strong></td>
        </tr>
        <tr class="lightgreen-bg">
            <td style="text-align: center;">Windows 11</td>
            <td style="text-align: center;">$Win11</td>
            <td style="text-align: center;">AlreadyOnWin11</td>
            <td style="text-align: center;">$PerShareWin11</td>
            <td style="text-align: center;">AlreadyOnWin11</td>
        </tr>
        <tr class="lightcoral-bg">
            <td style="text-align: center;">Windows 10 22H2</td>
            <td style="text-align: center;">$Win10_22H2</td>
            <td style="text-align: center;">$Version_22H2</td>
            <td style="text-align: center;">$PerShareWin10_22H2</td>
            <td style="text-align: center;">$DaysLeft_Version_22H2 Days</td>
        </tr>
        <tr class="lightcoral-bg">
            <td style="text-align: center;">Windows 10 21H2</td>
            <td style="text-align: center;">$Win10_21H2</td>
            <td style="text-align: center;">$Version_21H2</td>
            <td style="text-align: center;">$PerShareWin10_21H2</td>
            <td style="text-align: center;">$DaysLeft_Version_21h2 Days</td>
        </tr>
        <tr class="lightgray-bg">
            <td style="text-align: center;">Windows 10 21H1</td>
            <td style="text-align: center;">$Win10_21H1</td>
            <td style="text-align: center;">$Version_21H1</td>
            <td style="text-align: center;">$PerShareWin10_21H1</td>
            <td style="text-align: center;">$DaysLeft_Version_21h1</td>
        </tr>
        <tr class="lightgray-bg">
            <td style="text-align: center;">Windows 10 20H2</td>
            <td style="text-align: center;">$Win10_20H2</td>
            <td style="text-align: center;">$Version_20h2</td>
            <td style="text-align: center;">$PerShareWin10_20H2</td>
            <td style="text-align: center;">$DaysLeft_Version_20h2</td>
        </tr>
        <tr class="lightgray-bg">
            <td style="text-align: center;">Windows 10 2004</td>
            <td style="text-align: center;">$Win10_2004</td>
            <td style="text-align: center;">$Version_2004</td>
            <td style="text-align: center;">$PerShareWin10_2004</td>
            <td style="text-align: center;">$DaysLeft_Version_2004</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1909</td>
    <td style="text-align: center;">$Win10_1909</td>
    <td style="text-align: center;">$Version_1909</td>
    <td style="text-align: center;">$PerShareWin10_1909</td>
    <td style="text-align: center;">$DaysLeft_Version_1909</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1903</td>
    <td style="text-align: center;">$Win10_1903</td>
    <td style="text-align: center;">$Version_1903</td>
    <td style="text-align: center;">$PerShareWin10_1903</td>
    <td style="text-align: center;">$DaysLeft_Version_1903</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1809</td>
    <td style="text-align: center;">$Win10_1809</td>
    <td style="text-align: center;">$Version_1809</td>
    <td style="text-align: center;">$PerShareWin10_1809</td>
    <td style="text-align: center;">$DaysLeft_Version_1809</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1803</td>
    <td style="text-align: center;">$Win10_1803</td>
    <td style="text-align: center;">$Version_1803</td>
    <td style="text-align: center;">$PerShareWin10_1803</td>
    <td style="text-align: center;">$DaysLeft_Version_1803</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1709</td>
    <td style="text-align: center;">$Win10_1709</td>
    <td style="text-align: center;">$Version_1709</td>
    <td style="text-align: center;">$PerShareWin10_1709</td>
    <td style="text-align: center;">$DaysLeft_Version_1709</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1703</td>
    <td style="text-align: center;">$Win10_1703</td>
    <td style="text-align: center;">$Version_1703</td>
    <td style="text-align: center;">$PerShareWin10_1703</td>
    <td style="text-align: center;">$DaysLeft_Version_1703</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1607</td>
    <td style="text-align: center;">$Win10_1607</td>
    <td style="text-align: center;">$Version_1607</td>
    <td style="text-align: center;">$PerShareWin10_1607</td>
    <td style="text-align: center;">$DaysLeft_Version_1607</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1511</td>
    <td style="text-align: center;">$Win10_1511</td>
    <td style="text-align: center;">$Version_1511</td>
    <td style="text-align: center;">$PerShareWin10_1511</td>
    <td style="text-align: center;">$DaysLeft_Version_1511</td>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">Windows 10 1507</td>
    <td style="text-align: center;">$Win10_1507</td>
    <td style="text-align: center;">$Version_1507</td>
    <td style="text-align: center;">$PerShareWin10_1507</td>
    <td style="text-align: center;">$DaysLeft_Version_1507</td>
</tr>
</tr>
<tr class="lightgray-bg">
    <td style="text-align: center;">NoOSInfo</td>
    <td style="text-align: center;">$NoInfo</td>
    <td style="text-align: center;">NoInfo</td>
    <td style="text-align: center;">$PerShareNoInfo</td>
    <td style="text-align: center;">NoInfo</td>
</tr>
<tr class="lightblue-bg">
    <td style="text-align: center;"><strong>Total Devices</strong></td>
    <td style="text-align: center;"><strong>$TotalDevice</strong></td>
    <td style="text-align: center;"><strong> </strong></td>
    <td style="text-align: center;"><strong> </strong></td>
    <td style="text-align: center;"><strong> </strong></td>
</tr>
   </table>
    <p>Regards,<br> Patch Management Team</p>
</body>
</html>
"@

#Email Params
$TD = (Get-Date).ToString("MM/dd/yyyy")
$Subject     = "Windows 11 Compliance Dashboard & Report : as of $TD"  
Send-MailMessage -From $From -to $To -CC $CC -Subject $Subject  -Body $NotificationBody -BodyAsHtml  -SmtpServer $SMTPServer -Port $Port -Attachments $FinalReportPath
Write-Host "Mail successfully sent to $to,$cc" -ForegroundColor White
Write-Host "Windows 11 compliance Dashboard Report is avaialbe at this location:-" $FinalReportPath -ForegroundColor White
Write-Host "===============================Phase-3 E-mailSent (Windows Patching Compliance Report)=======================================(Completed)" -ForegroundColor Green
$endTime = Get-Date
$duration = $endTime - $startTime
Write-Host ""
Write-Host "Time duation to successfully excute this script is:- $duration" -ForegroundColor White
Invoke-Item -Path $FinalReportPath