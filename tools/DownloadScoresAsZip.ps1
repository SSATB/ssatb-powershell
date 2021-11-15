###
# DownloadScoresAsZip.ps1
#
# Script to download score report files as Zip file
# for the schools using SLATE.
#
# options:
#
# -LastRunDate       '1/1/1970 12:30PM'
# -ExportFolder      '' 
#

Param(
[Parameter ()]
 [DateTime] $LastRunDate = '1/1/1970 12:00:00 AM',
[Parameter()]
 [String]$ExportFolder = "",
[Parameter()]
 [String]$ClientId = "",
[Parameter()]
 [String]$ClientSecret = "",
[Parameter()]
 [String]$SchoolCode = ""
)

###
# Initialization Parameters
$apiRoot = 'https://api.ssat.org'

$ErrorActionPreference ='Stop' # Stop on Errors

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
# if export folder is suplied then use that otherwise fallback to the current folder
if ($ExportFolder -eq "") {
    $ExportFolder = (Get-Location).ToString()
}
$currentDir = Get-Location 
$metaFile = "${currentDir}\${SchoolCode}_scores_meta.dat"

##
# If Last Run Date is not supplied then check to see if the metadata file exists if it does read the last run from meta data file else 
# Set the Last Run Date as the Current Year Academic Start
##
if ($LastRunDate -eq '1/1/1970 12:00:00 AM'){
    [DateTime]$currentDate = Get-Date
    [DateTime]$acadmicYearStart = "8/1/" + $currentDate.Year.ToString();
    if ($currentDate.Month -lt 8){
        $acadmicYearStart = "8/1/" + ($currentDate.Year -1).ToString();
    }
 if (!([System.IO.File]::Exists($metaFile))) {  
    $LastRunDate = $acadmicYearStart
    }
    else {
        $content = Get-Content -path $metaFile
        try{
            $LastRunDate = [datetime]::Parse($content)
        }
        catch
        {
           Write-Output  "Error reading Metadata file. Assuming lastRun as $acadmicYearStart" 
           $LastRunDate = $acadmicYearStart
        }
    }
}


Write-Output "DownloadScores called with initParams LastRunDate=$LastRunDate and ExportFolder=$ExportFolder" 
##
# Get an oAuth accessToken for our session 
# and store it as a Dictionary to be used 
# in subsequent requests

$urlToken = "$apiRoot/oauth/token"
$authRequestBody = @{
    grant_type = "client_credentials"
    client_id = $ClientId
    client_secret= $ClientSecret
    scope= $SchoolCode
}


$authResponse=Invoke-RestMethod -Method Post -Uri $urlToken -Body $authRequestBody 
Write-Debug "Authentication Response: $authResponse" 

#TODO: do a bit of checking to ensure the reponse was valid

$accessTokenHeader = @{
	"Authorization"=("Bearer", $authResponse.access_token -join " ")
}



##
# Now get the master index of all available 
# students. 

$urlAllFolios = "$apiRoot/scores/StudentScoresAll/$schoolCode"
$urlAllFolios += "?LastUpdateDateRangeStart=" + $LastRunDate.ToString("yyyy-MM-dd HH:mm:ss")
Write-Debug  "Request All Scores Since : $LastRunDate" 
$allFoliosResponse=Invoke-RestMethod -Method Get -Uri $urlAllFolios -Headers $accessTokenHeader

$count = $allFoliosResponse.MemberSchoolStudents.Count
Write-Output "Student Count :  $count"
$TimeNowEST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'Eastern Standard Time')

if ($count -eq 0) {
    Set-Content -Path $metaFile -Value $TimeNowEST
    Write-Output  "No New Scores. Exiting" 
    exit;
}


#TODO: do a bit of checking to ensure the reponse was valid


##
# For each student in the response
# get their scores and then request each
# score report individually exporting the file to 
# the output directory

$ExportFolder = $ExportFolder.TrimEnd("\");
$ExportZipArchive = $ExportFolder + "\scores_${SchoolCode}_" + $TimeNowEST.ToString("yyyy-MM-dd HH_mm_ss") + ".zip"
if (!(Test-Path -Path $ExportZipArchive -IsValid)) {
    Write-Error "Path $ExportZipArchive is not valid. Please fix and try again."
    exit
}
$tempFolderName = $TimeNowEST.ToString("yyyy-MM-dd HH_mm_ss");
$tempFolder = "${currentDir}\${tempFolderName}"
New-Item -Path ${currentDir}\ -Name $tempFolderName -ItemType "directory"
$indexFile = "${tempFolder}\index.txt"
Add-Content $indexFile """File Name""`t""Reg ID""`t""Folio ID""`t""First name""`t""Last name""" 

#loop over each student in the response
foreach($student in $allFoliosResponse.MemberSchoolStudents) {
    $firstName =  $student.FirstName
    $middleName = $student.MiddleName
    $lastName =  $student.LastName
    $folioId =  $student.FolioId
    $testRegId = $student.TestRegistrationID;

    Write-Debug "Request Scores for $firstName $middleName $lastName (FolioId=$folioId)"
    ${fileId} = ("${firstName}_${middleName}_${lastName}_${folioId}_${testRegId}" -replace '[^a-zA-Z0-9_]', '') + ".pdf"
    $pdfUrl =  "$apiRoot/scores/StudentScores/pdf/$schoolCode/$folioId/$testRegId" 
    $fileName = "${tempFolder}\${fileId}";
    Invoke-RestMethod -Uri $pdfUrl -Headers $accessTokenHeader -OutFile $fileName
    Write-Debug "Saved $fileName"
    Add-Content $indexFile """${fileId}""`t""$testRegId""`t""$folioId""`t""$firstName""`t""$lastName""" 
 }

Compress-Archive -Path "${tempFolder}\*" -CompressionLevel Fastest -DestinationPath $ExportZipArchive
Set-Content -Path $metaFile -Value $TimeNowEST
Remove-Item –path ${tempFolder} –recurse
