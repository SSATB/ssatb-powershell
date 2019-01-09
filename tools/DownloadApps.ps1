###
# DownloadApps.ps1
#
# Script to download School Applications files
# for the school specified by the input params.
#
# options:
#
# -LastRunDate       '1/1/1970 12:30PM' Defaults to the Academic Year Start (August-01)
# -ExportZipArchive   '' - A Zip File or a Directory where to Put the Zip file in 
# -AcademicYear       0  - Filter Apps for a Given year if Necessary
# -SchoolCode             - SSATB SchoolCode

Param(
    [Parameter ()]
    [DateTime] $LastRunDate = '1/1/1970 12:00:00 AM',
    [Parameter()]
    [String]$ExportZipArchive = "",
    [Parameter()]
    [String]$ClientId = "",
    [Parameter()]
    [String]$ClientSecret = "",
    [Parameter()]
    [String]$SchoolCode = "",
    [Parameter()]
    [int]$AcademicYear = 0
)

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
}
if (-not ([Net.ServicePointManager]::SecurityProtocol).ToString().Contains([Net.SecurityProtocolType]::Tls12)) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol.toString() + ', ' + [Net.SecurityProtocolType]::Tls12
}
# Verify Required Client Parameters are present
if ($ClientId -eq "" -Or $ClientSecret -eq "" -Or $SchoolCode -eq "") {
    Write-Error "Missing Required Authentication Parameters"
    exit
}
    
Write-Debug "DownloadApps called with initParams LastRunDate=$LastRunDate and ExportZipArchive=$ExportZipArchive" 

if ($ExportZipArchive -eq "") {
    $ExportZipArchive = Get-Location
}
$TimeNowEST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'Eastern Standard Time')
$ExportFolder = $ExportZipArchive
if ($ExportZipArchive.ToLower().endswith(".zip") ) {
    $ExportFolder = Split-Path -Path $ExportZipArchive -Parent
}
else {
    $ExportFolder = $ExportZipArchive.TrimEnd("\");
    $ExportZipArchive = $ExportFolder + "\apps_${SchoolCode}_" + $TimeNowEST.ToString("yyyy-MM-dd HH_mm_ss") + ".zip"
}
if (!(Test-Path -Path $ExportZipArchive -IsValid)) {
    Write-Error "Path $ExportZipArchive is not valid. Please fix and try again."
    exit
}

###
# Initialization Parameters
$apiRoot = 'https://api.ssat.org'

##
# Get an oAuth accessToken for our session 
# and store it as a Dictionary to be used 
# in subsequent requests

$urlToken = "$apiRoot/oauth/token"
$authRequestBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = $SchoolCode
}

$timeNowLocal = Get-Date
$authResponse = Invoke-RestMethod -Method Post -Uri $urlToken -Body $authRequestBody 
Write-Debug "Authentication Response: $authResponse" 

#TODO: do a bit of checking to ensure the reponse was valid

$accessTokenHeader = @{
    "Authorization" = ("Bearer", $authResponse.access_token -join " ")
    "Accept"        = "application/json"
}
$tokenExpiryDate = $timeNowLocal.AddSeconds($authResponse.expires_in - 200)
$currentDir = Get-Location 
$metaFile = "${currentDir}\${SchoolCode}_meta.dat"
##
# If Last Run Date is not supplied Set the Last Run Date as the Current Year Academic Start
if ($LastRunDate -eq '1/1/1970 12:00:00 AM') {
    if (!([System.IO.File]::Exists($metaFile))) {  
        [DateTime]$acadmicYearStart = "8/1/" + $TimeNowEST.Year.ToString();
        if ($TimeNowEST.Month -lt 8) {
            $acadmicYearStart = "8/1/" + ($TimeNowEST.Year - 1).ToString();
        }
        $LastRunDate = $acadmicYearStart
    }
    else {
        $content = Get-Content -path $metaFile
        $LastRunDate = [datetime]::Parse($content)
    }
}


##
# Now get the master index of all available 
# students. 

$urlAllFolios = "$apiRoot/sao/StudentApplication/Folios/$schoolCode"
$urlAllFolios += "?LastUpdateDateRangeStart=" + $LastRunDate.ToString("yyyy-MM-dd HH:mm:ss") + "&LastUpdateDateRangeEnd=" + $TimeNowEST.ToString("yyyy-MM-dd HH:mm:ss")

if ($AcademicYear -ne 0) {
    $urlAllFolios += "&ApplicationSessionYear=$AcademicYear"
}
Write-Output  "Request All Apps Since : $LastRunDate" 
$allFoliosResponse = Invoke-RestMethod -Method Get -Uri $urlAllFolios -Headers $accessTokenHeader

$count = $allFoliosResponse.MemberSchoolStudents.Count
Write-Output "Student Count :  $count"
if ($count -eq 0) {
    Set-Content -Path $metaFile -Value $TimeNowEST
    Write-Output  "No Changed Apps. Exiting" 
    exit;
}

#TODO: do a bit of checking to ensure the reponse was valid


##
# For each student in the response
# get their scores and then request each
# score report individually exporting the file to 
# the output directory

$urlStudentReport = "$apiRoot/sao/StudentApplication/$schoolCode"
$tempFolderName = $TimeNowEST.ToString("yyyy-MM-dd HH_mm_ss");
$tempFolder = "${ExportFolder}\${tempFolderName}"
New-Item -Path $ExportFolder -Name $tempFolderName -ItemType "directory"
$indexFile = "${tempFolder}\index.txt"
Add-Content $indexFile """File Name""`t""SAO ID""`t""Document Type""" 
#loop over each student in the response
foreach ($student in $allFoliosResponse.MemberSchoolStudents) {
    $timeNowLocal = Get-Date
    if ($timeNowLocal -ge $tokenExpiryDate) {
        $authResponse = Invoke-RestMethod -Method Post -Uri $urlToken -Body $authRequestBody 
        Write-Debug "Authentication Response: $authResponse" 

        #TODO: do a bit of checking to ensure the reponse was valid

        $accessTokenHeader = @{
            "Authorization" = ("Bearer", $authResponse.access_token -join " ")
            "Accept"        = "application/json"
        }
        $tokenExpiryDate = $timeNowLocal.AddSeconds($authResponse.expires_in - 200)
    }

    $firstName = $student.FirstName
    $middleName = $student.MiddleName
    $lastName = $student.LastName
    $folioId = $student.FolioId
    $appYear = $student.ApplicationSessionYear
    
    Write-Output "Request App Data for $firstName $middleName $lastName (FolioId=$folioId)"
    $appUrl = "$urlStudentReport/$folioId"
    $appUrl += "?ApplicationSessionYear=$appYear"
    $folioResponse = Invoke-RestMethod -Method Get -Uri $appUrl -Headers $accessTokenHeader
    $components = $folioResponse.Application.Components
    $pattern = '[^a-zA-Z0-9]'
    foreach ($component in $components) { 
        $componentId = $component.ComponentId
        $componentName = $component.ComponentName
        if ($component.CompletionDate) {
            $pdfUrl = "$apiRoot/sao/StudentApplication/ComponentWithAttachments/PDF/$schoolCode/$folioId/$componentId"
            $pdfUrl += "?ApplicationSessionYear=$appYear" 
            $componentFileName = $componentName -replace $pattern, '_' 
            $fileId = "${folioId}_${appYear}_${componentFileName}.pdf"
            $fileName = "${tempFolder}\\${fileId}";
            Invoke-RestMethod -Uri $pdfUrl -Headers $accessTokenHeader -OutFile $fileName
            Write-Output "Saved $fileName"
            Add-Content $indexFile """${fileId}""`t""$folioId$appYear""`t""$componentName""" 

        }
    }
}
Compress-Archive -Path "${tempFolder}\*" -CompressionLevel Fastest -DestinationPath $ExportZipArchive
Set-Content -Path $metaFile -Value $TimeNowEST
Remove-Item –path ${tempFolder} –recurse
