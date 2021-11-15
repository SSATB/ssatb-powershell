EMA Powershell Scripts ReadMe
=======================================

**DownloadApps.ps1**
--------
**Purpose:**

Create a Zip Archive containing:
  1. EMA SAO Application Component PDF Files
  2. Index File containing metadata about the Downloaded Components
--------
**Prerequisites:**
   1. Setup OAuth Credentials for Using SSATB SAO APIs
   2. A Windows 10 PC running atleast Powershell 5.0. The Script creates large zip file archive so it would need atleast 8 GB of RAM
   3. PowerShell Executuon Policy. By Default Powershell Execution Policy is set to "Restricted", which basically disables exectuion of any scripts. This needs to be changed to atleast "__RemoteSigned__" . You can do that be logging in as an Admin and executing `set-executionpolicy remotesigned`
   4. Folder Permissions : This scripts creates TempDirectory in the folder it is executing or in the *ExportZipArchive* folder that is specified. It needs permisions to create and delete files in that folder.
--------
**How to Use:**
   1. Startup Powershell and Navigate to the directory
   
         ` .\DownloadApps.ps1 -ClientId "emaissuedclientid" -ClientSecret "emaissuedclientsecret" -SchoolCode "1111" `
        
        Arguments to the Script:  
                - **ClientId** : OAuth 2 ClientId provided by EMA. This is required.   
                - **ClientSecret** : OAuth 2 Client secret provided by EMA. This is required.  
                - **SchoolCode** : Your SSATB SchoolCode. This is Required.  
                - *AcademicYear* : If supplied filters applications for a particular academic year only. By default retrieves all applications.  
                - *LastRunDate* : '1/1/1970 12:00 AM' If supplied gets filtered Applications that have changed since this date. The default value is start of   
                Academic year (8/1). Please note the application maintains a file by the name __schoolcode_meta.data__ 
                where this value is persisted, so that the script can pick up from where it left of when the next time its run. 
                __PLEASE DO NOT DELETE THIS FILE.__    
                - *ExportZipArchive* : A fully qualified path to a directory or a File where you want the Zip Archive. If this is a directory then a File by the name apps_{schoolcode}_{Date of Run}.zip is Prepaired.

**DownloadScoresAsZip.ps1**
--------
**Purpose:**

Create a Zip Archive containing:
  1. Score Report PDF Files 
  2. Index File containing metadata about the SSAT Scores
--------
**Prerequisites:**
   1. Setup OAuth Credentials for Using SSATB Score APIs
   2. A Windows 10 PC running atleast Powershell 5.0. The Script creates large zip file archive so it would need atleast 8 GB of RAM
   3. PowerShell Executuon Policy. By Default Powershell Execution Policy is set to "Restricted", which basically disables exectuion of any scripts. This needs to be changed to atleast "__RemoteSigned__" . You can do that be logging in as an Admin and executing `set-executionpolicy remotesigned`
   4. Folder Permissions : This scripts creates TempDirectory in the folder it is executing or in the *ExportFolder* folder that is specified. It needs permisions to create and delete files in that folder.
--------
**How to Use:**
   1. Startup Powershell and Navigate to the directory
   
         ` .\DownloadScoresAsZip.ps1 -ClientId "emaissuedclientid" -ClientSecret "emaissuedclientsecret" -SchoolCode "1111" `
        
        Arguments to the Script:  
                - **ClientId** : OAuth 2 ClientId provided by EMA. This is required.   
                - **ClientSecret** : OAuth 2 Client secret provided by EMA. This is required.  
                - **SchoolCode** : Your SSATB SchoolCode. This is Required.  
                - *LastRunDate* : '1/1/1970 12:00 AM' If supplied gets filtered Score Reports that have changed since this date. The default value is start of   
                Academic year (8/1). Please note the application maintains a file by the name __schoolcode_scores_meta.dat__ 
                where this value is persisted, so that the script can pick up from where it left of when the next time its run. 
                __PLEASE DO NOT DELETE THIS FILE.__    
                - *ExportFolder* : A fully qualified path to a directory or a File where you want the Zip Archive.
                
       The App Prepares a Zip Archive be the Name scores_{schoolcode}_{Date of Run}.zip
               
