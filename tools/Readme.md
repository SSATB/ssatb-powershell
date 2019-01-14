EMA Powershell Scripts ReadMe
=======================================

**DownloadApps.ps1**
--------
**Purpose:**

Create a Zip Archive containing:
  1. EMA SAO Application Component Pdf Files
  2. Index File Containing metadata about the Downloaded Components
--------
**Prerequisites:**
   1. Setup OAuth Credentials for Using SSATB SAO APIs
   2. A Windows 10 PC running atleast Powershell 5.0. The Script Creates Lazrge Zip File Archive so it would need atleast 8 GB of RAM
   3. PowerShell Executuon Policy. By Default Powershell ExecutionPolicy is set to "Restricted" , which basically disables Exectuion of any Scripts. This needs to be changed to atleast "__RemoteSigned__" . You can do that be Logging in as an Admin and executing `set-executionpolicy remotesigned`
   4. Folder Permissions : This Scripts Creates TempDirectory in the Folder it is Executing or in the *ExportZipArchive* folder f that is specified. It needs permisions to create and delete files in that folder
--------
**How to Use:**
   1. Startup Powershell and Navigate to the directory
   
         ` .\DownloadApps.ps1 -ClientId "emaissuedclientid" -ClientSecret "emaissuedclientsecret" -SchoolCode "1111" `
        
        Arguments to the Script:  
                - **ClientId** : OAuth 2 ClientId Provided by EMA. This is Required.   
                - **ClientSecret** : OAuth 2 Client Secret Provided by EMA. This is Required.  
                - **SchoolCode** : Your SSATB SchoolCode. This is Required.  
                - *AcademicYear* : If supplied filters Applications for a particular academic year only.By Default retirens all applications.  
                - *LastRunDate* : '1/1/1970 12:30PM' If supplied gets filters Applications that have changed since this Date. The default value is start of   
                Academic year (8/1). Please note the application maintains a file by the name __schoolcode_meta.data__ 
                where this value is persisted, so that the script can pick up from where it left of when the next time its run. 
                __PLEASE DO NOT DELETE THIS FILE.__    
                - *ExportZipArchive* : A Fully Qualified path to a Directory or a File where you want the Zip Archive.If this is a directory then a File by the name apps_{schoolcode}_{Date of Run}.zip is Prepaired.
               
