###
# DownloadScores.ps1
#
# Script to download score report files
# for the school specified by the input params.
#
# options:
#
# -LastRunDate       '1/1/1970 12:30PM'
# -ExportFolder      '' 
#
# TODO Date should be generated 
# at each run and stored as a 
# parameter in a file that can
# be read each time this script is run

Param(
[Parameter ()]
[DateTime] $LastRunDate = '1/1/1970 12:00:00 AM',
[Parameter()]
 [String]$ExportFolder = ""
)

If ($PSBoundParameters['Debug']) {
    $DebugPreference = 'Continue'
}

Write-Debug "DownloadScores called with initParams LastRunDate=$LastRunDate and ExportFolder=$ExportFolder" 

###
# Initialization Parameters
$apiRoot = 'https://api.ssat.org'

# The next three are Client Specific : TODO : Save it in a Config File
$clientId = "MY CLIENT ID"
$clientSecret = "MY CLIENT SECRET"
$schoolCode = "MY SCHOOLCODE"


##
# Get an oAuth accessToken for our session 
# and store it as a Dictionary to be used 
# in subsequent requests

$urlToken = "$apiRoot/oauth/token"
$authRequestBody = @{
    grant_type = "client_credentials"
    client_id = $clientId
    client_secret= $clientSecret
    scope= $schoolCode
}


$authResponse=Invoke-RestMethod -Method Post -Uri $urlToken -Body $authRequestBody 
Write-Debug "Authentication Response: $authResponse" 

#TODO: do a bit of checking to ensure the reponse was valid

$accessTokenHeader = @{
	"Authorization"=("Bearer", $authResponse.access_token -join " ")
}

##
# If Last Run Date is not supplied Set the Last Run Date as the Current Year Academic Start
if ($LastRunDate -eq '1/1/1970 12:00:00 AM'){
    [DateTime]$currentDate = Get-Date
    [DateTime]$acadmicYearStart = "8/1/" + $currentDate.Year.ToString();
    if ($currentDate.Month -lt 8){
        $acadmicYearStart = "8/1/" + ($currentDate.Year -1).ToString();
    }
    $LastRunDate = $acadmicYearStart
}


##
# Now get the master index of all available 
# students. 

$urlAllFolios = "$apiRoot/scores/StudentScoresAll/$schoolCode"
$urlAllFolios += "?LastUpdateDateRangeStart=" + $LastRunDate.ToString("yyyy-MM-dd HH:mm:ss")
Write-Debug  "Request All Scores Since : $LastRunDate" 
$allFoliosResponse=Invoke-RestMethod -Method Get -Uri $urlAllFolios -Headers $accessTokenHeader

$count = $allFoliosResponse.MemberSchoolStudents.Count
Write-Debug "Student Count :  $count"


#TODO: do a bit of checking to ensure the reponse was valid


##
# For each student in the response
# get their scores and then request each
# score report individually exporting the file to 
# the output directory

$urlStudentReport = "$apiRoot/scores/StudentScores/PDF/Consolidated/$schoolCode/"

#loop over each student in the response
foreach($student in $allFoliosResponse.MemberSchoolStudents) {
    $firstName =  $student.FirstName
    $middleName = $student.MiddleName
    $lastName =  $student.LastName
    $folioId =  $student.FolioId
    $testRegId = $student.TestRegistrationID;

    Write-Debug "Request Scores for $firstName $middleName $lastName (FolioId=$folioId)"
    
    $pdfUrl =  "$apiRoot/scores/StudentScores/pdf/$schoolCode/$folioId/$testRegId" 
    $fileName = "${ExportFolder}\${firstName}_${middleName}_${lastName}_${folioId}_${testRegId}.pdf";
    Invoke-RestMethod -Uri $pdfUrl -Headers $accessTokenHeader -OutFile $fileName
    Write-Debug "Saved $fileName"
       
 }

