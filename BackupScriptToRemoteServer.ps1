##########################################################################################################################
#####Begin Variables#####

###Variables that you may need to change
#####IMPORTANT PLEASE READ:  Get store name dynamically from the folder name on the Tag Appliance or put the name manually (no spaces) and match naming convention#####
$VarFacility = Get-ChildItem -Path C:\inbox -Name

#PowerShell drive credentials
$VarDriveCredentialUser = "computer\backupuser"
$VarDriveCredentialPassword = "some password"

#SMTP credentials
$VarSMTPServer = "smtp.somedomain.local"
$VarSMTPCredentialUser = "domain\user"
$VarSMTPCredentialPassword = "somePassword"

#To and From email variables.  Multiple email addresses can be added to the VarToEmail variable but they need to be seperated by commas
$VarToEmail = "someemail@someemail", "someemail2@someemail"
$VarFromEmail = "do-not-reply@testdomain.test"

#Root folder variable on the Tag Appliance, where the production files are found
$VarRootFolder = "C:\Inbox\"
$VarRootFolderWithFacilty = $VarRootFolder + $VarFacility

#Root folder variable on the Tag Appliance, where the staging files go
$VarRootStagingFolder = "C:\Outbox\"

#Main backup root path on the corporate server
$VarCorporateServer = "server name"
$VarCorporateServerBaseRootPath ="\\" + $VarCorporateServer + "\Backup" #The backup folder is in the Inbox on the corporate server
$VarCorporateServerRootPath = $VarCorporateServerBaseRootPath + "\" + $VarFacility

#Zip filenames
$VarDatabaseFile = "\SQL_Backups.zip"
$VarOPCFile = "\ConfigFiles_Backups.zip"
$VarWebsiteFile = "\Web_Backups.zip"

#################################################################################################################################
###Variables that you don't need to change
#Success email variables
$VarSuccessfulEmailSubject = "Successful Backup: " + $VarFacility
$VarSuccessfulEmailBody = "Files were on the corporate server, and the correct size, for this backup"

#Unsuccess email variables
$VarUnsuccessfulEmailSubject = "Not a Successful Backup: " + $VarFacility
$VarUnsuccessfulFileSizeEmailBody = "Error:  Files were on the corporate server, but file size was incorrect for this backup"
$VarUnsuccessfulFileCountEmailBody = "Error:  File(s) were not copied to the corporate server for this backup"
$VarUnsuccessfulMapDriveEmailBody = "Error:  Could not map drive to corporate server. No files copied"
$VarUnsuccessfulMultipleFoldersEmailBody = "Error:  Multiple store folders exist on the Tag Appliance"
$VarUnsuccessfulCorporateUTCFolderEmailBody = "Error:  Could not create the backup folder on the corporate server. No files copied"

#Corporate Server
$VarCorporateServerDatabasePath = "\Data\SQL\"
$VarCorporateServerOPCPath = "\Data\config\"
$VarCorporateServerWebsitePath = "\Web\"

#Tag Appliance paths
$VarTagApplianceOPCPath = $VarRootFolderWithFacilty + "\Data\config"
$VarTagApplianceWebsitePath = $VarRootFolderWithFacilty + "\Web"
#No path is needed for databases because this script talks directly to SQL

#Tag Appliance root staging path
$VarTagApplianceBackupStagingRootPath = "_StagingBackups\"

#Tag Appliance database staging paths
$VarTagApplianceBackupStagingDatabaseRootPath = $VarRootStagingFolder + $VarTagApplianceBackupStagingRootPath + "Data\SQL\"
$VarTagApplianceBackupStagingDeleteDatabaseZipPath = $VarTagApplianceBackupStagingDatabaseRootPath + "*.*"
$VarTagApplianceBackupStagingDatabasePath = $VarTagApplianceBackupStagingDatabaseRootPath + "DatabaseBackup\"
$VarTagApplianceBackupStagingDeleteDatabasePath = $VarTagApplianceBackupStagingDatabasePath + "*.*"

#Zips on the Tag Appliance
$VarTagApplianceBackupStagingDatabaseZip = $VarRootStagingFolder + $VarTagApplianceBackupStagingRootPath + "Data\SQL" + $VarDatabaseFile
$VarTagApplianceBackupStagingOPCZip = $VarRootStagingFolder + $VarTagApplianceBackupStagingRootPath + "Data\config" + $VarOPCFile
$VarTagApplianceBackupStagingWebsiteZip = $VarRootStagingFolder + $VarTagApplianceBackupStagingRootPath + "Web" + $VarWebsiteFile
######################################################################################################################################
#####End Variables#####

#Create credentials for SMTP
$VarSMTPPassword = ConvertTo-SecureString –String $VarSMTPCredentialPassword –AsPlainText -Force
$VarSMTPCredential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $VarSMTPCredentialUser, $VarSMTPPassword

#Create credentials for user to make PowerShell drives
$VarDriveCredentialPasswordSecured = ConvertTo-SecureString –String $VarDriveCredentialPassword –AsPlainText -Force
$VarDriveCredential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $VarDriveCredentialUser, $VarDriveCredentialPasswordSecured

#Detect if facilty name consists of multiple folders
If($VarFacility -is [system.array]) #This variable becomes an array when more than one folder exists
    {        
        Write-Host "Multiple store folders exist on the Tag Appliance" -ForegroundColor Red
        Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulMultipleFoldersEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
        Clear-Variable -Name Var*
        Exit
    }

###Detect store folder on corporate server or create it
#Create temporary PowerShell drive on the Tag Appliance
$VarTestPathPowerShellDrive = New-PSDrive -Credential $VarDriveCredential -Root $VarCorporateServerBaseRootPath -Name “TestPathDrive” -PSProvider FileSystem

#Detect if PowerShell drive was created successfully
If (!$VarTestPathPowerShellDrive) #If variable is null
    {
        Write-Host "Could not create Test Path Drive" -ForegroundColor Red
        Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulMapDriveEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
        Clear-Variable -Name Var*
        Exit
    }

#Create test folder path
$VarTestPathCorporateServerRootPath = $VarTestPathPowerShellDrive.Root + "\" + $VarFacility

#Check to see if store folder has been created on the corporate server, if not then this will create it
If (!(Test-Path -Path $VarTestPathCorporateServerRootPath))
    {        
        Write-Host "Store folder does not exist on corporate server for this store.  This script will create it now" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $VarTestPathCorporateServerRootPath
    }
Else
    {
        "Store folder already exists on the corporate server for this store"
    }

#Remove temporary PowerShell drive
Remove-PSDrive -Name TestPathDrive

###Get current date in UTC format and convert to a string
$VarToday = Get-Date
$VarTodayUTC = $Vartoday.ToUniversalTime()
$VarFolderName = $VarTodayUTC.ToString("yyyy-MM-dd-hh-mm-ss")

#Detect if the folder on the corporate server was created successfully
If (!$VarFolderName) #If variable is null
    {
        Write-Host "Could not generate the corporate server folder based on UTC" -ForegroundColor Red        
        Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulCorporateUTCFolderEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
        Clear-Variable -Name Var*
        Exit
    }

###Backup databases on Tag Appliance
Import-Module “sqlps” #This module is included in SQL Management Studio
cd SQLSERVER:\SQL\localhost\default\Databases
    foreach($Vardatabase in (Get-ChildItem))
        {
            $VardbName = $Vardatabase.Name
            $VarDatabaseBackupFile = $VarTagApplianceBackupStagingDatabasePath + $VardbName + ".bak"
            Backup-SqlDatabase -Database $VardbName -BackupFile $VarDatabaseBackupFile
        }
cd c:\

###Create temporary PowerShell drive on the Tag Appliance
$VarPowerShellDrive = New-PSDrive -Credential $VarDriveCredential -Root $VarCorporateServerRootPath -Name “BackupDrive” -PSProvider FileSystem

#Detect if PowerShell drive was created successfully
If (!$VarPowerShellDrive) #If variable is null
    {
        Write-Host "Could not create Backup Drive" -ForegroundColor Red 
        Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulMapDriveEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
        Clear-Variable -Name Var*
        Exit
    }

###Generate folder paths based on UTC
$VarCorporateServerDatabaseBackupLocation = $VarPowerShellDrive.Root + $VarCorporateServerDatabasePath + $VarFolderName
$VarCorporateServerOPCBackupLocation = $VarPowerShellDrive.Root + $VarCorporateServerOPCPath + $VarFolderName
$VarCorporateServerWebsiteBackupLocation = $VarPowerShellDrive.Root + $VarCorporateServerWebsitePath + $VarFolderName

#Create folders on corporate server based on UTC
New-Item -ItemType Directory -Path "$VarCorporateServerDatabaseBackupLocation", "$VarCorporateServerOPCBackupLocation", "$VarCorporateServerWebsiteBackupLocation"

###Zip up files on the Tag Appliance
Add-Type -Assembly "System.IO.Compression.FileSystem" ;
[System.IO.Compression.ZipFile]::CreateFromDirectory($VarTagApplianceBackupStagingDatabasePath, $VarTagApplianceBackupStagingDatabaseZip) ;

Add-Type -Assembly "System.IO.Compression.FileSystem" ;
[System.IO.Compression.ZipFile]::CreateFromDirectory($VarTagApplianceOPCPath, $VarTagApplianceBackupStagingOPCZip) ;

Add-Type -Assembly "System.IO.Compression.FileSystem" ;
[System.IO.Compression.ZipFile]::CreateFromDirectory($VarTagApplianceWebsitePath, $VarTagApplianceBackupStagingWebsiteZip) ;

###Copy zip files to corporate server
Copy-Item -Path $VarTagApplianceBackupStagingDatabaseZip -Destination $VarCorporateServerDatabaseBackupLocation -Recurse
Copy-Item -Path $VarTagApplianceBackupStagingOPCZip -Destination $VarCorporateServerOPCBackupLocation -Recurse
Copy-Item -Path $VarTagApplianceBackupStagingWebsiteZip -Destination $VarCorporateServerWebsiteBackupLocation -Recurse

###Remove old folders on corporate server
#Generate folder paths
$VarDatabaseBackupFolders = $VarPowerShellDrive.Root + $VarCorporateServerDatabasePath
$VarOPCBackupFolders = $VarPowerShellDrive.Root + $VarCorporateServerOPCPath
$VarWebsiteBackupFolders = $VarPowerShellDrive.Root + $VarCorporateServerWebsitePath

#Remove folders that are older than 5 days
Get-Childitem -Path $VarDatabaseBackupFolders | Where {$_.CreationTime -lt (Get-Date).adddays(-5)} | Remove-Item -Recurse
Get-Childitem -Path $VarOPCBackupFolders | Where {$_.CreationTime -lt (Get-Date).adddays(-5)} | Remove-Item -Recurse
Get-Childitem -Path $VarWebsiteBackupFolders | Where {$_.CreationTime -lt (Get-Date).adddays(-5)} | Remove-Item -Recurse

###Check to see if files were correctly copied, if yes then send success email, but if not send unsuccessful email
#Generate folder paths
$VarCorporateServerBackupStagingDatabaseZip = $VarCorporateServerDatabaseBackupLocation + $VarDatabaseFile
$VarCorporateServerBackupStagingOPCZip = $VarCorporateServerOPCBackupLocation + $VarOPCFile
$VarCorporateServerBackupStagingWebsiteZip = $VarCorporateServerWebsiteBackupLocation + $VarWebsiteFile

############For Testing Only############
#Start-Sleep -s 100 ##########
########################################

###Database file sizes
#Database file size on Tag Appliance
$VarSourceDatabaseItem = Get-ItemProperty -Path $VarTagApplianceBackupStagingDatabaseZip
$VarSourceDatabaseItem.Length
#Database file size on corporate server
$VarDestinationDatabaseItem = Get-ItemProperty -Path $VarCorporateServerBackupStagingDatabaseZip
$VarDestinationDatabaseItem.Length

#OPC file sizes
#OPC file size on Tag Appliance
$VarSourceOPCItem = Get-ItemProperty -Path $VarTagApplianceBackupStagingOPCZip
$VarSourceOPCItem.Length
#OPC file size on corporate server
$VarDestinationOPCItem = Get-ItemProperty -Path $VarCorporateServerBackupStagingOPCZip
$VarDestinationOPCItem.Length

#Website file sizes
#Website file size on Tag Appliance
$VarSourceWebsiteItem = Get-ItemProperty -Path $VarTagApplianceBackupStagingWebsiteZip
$VarSourceWebsiteItem.Length
#Website file size on corporate server
$VarDestinationWebsiteItem = Get-ItemProperty -Path $VarCorporateServerBackupStagingWebsiteZip
$VarDestinationWebsiteItem.Length

#Count of files on corporate server
$VarContentCopiedDatabase = Get-ChildItem -Path $VarCorporateServerDatabaseBackupLocation
$VarContentCopiedOPC = Get-ChildItem -Path $VarCorporateServerOPCBackupLocation
$VarContentCopiedWebsite = Get-ChildItem -Path $VarCorporateServerWebsiteBackupLocation

$VarCountOfFiles = 0,0,0
$VarCountOfFiles[0] = $VarContentCopiedDatabase.Count
$VarCountOfFiles[1] = $VarContentCopiedOPC.Count
$VarCountOfFiles[2] = $VarContentCopiedWebsite.Count

#First it checks that files were copied to the corporate server, if not send unsuccess email
If (($VarCountOfFiles[0] -gt 0) -and ($VarCountOfFiles[1] -gt 0) -and ($VarCountOfFiles[2] -gt 0))
    {
        #Second it checks that the file sizes are correct on the corporate server, if not send unsuccess email
        If (($VarSourceDatabaseItem.Length -eq $VarDestinationDatabaseItem.Length) -and ($VarSourceOPCItem.Length -eq $VarDestinationOPCItem.Length) -and ($VarSourceWebsiteItem.Length -eq $VarDestinationWebsiteItem.Length))
            {
                Write-Host "Success" -ForegroundColor Green
                Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarSuccessfulEmailSubject -Body $VarSuccessfulEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
            }
        Else
            {                
                Write-Host "Unsuccess on File Size" -ForegroundColor Red
                Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulFileSizeEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
            }
    }
Else
    {        
        Write-Host "Unsuccess on File Count" -ForegroundColor Red
        Send-MailMessage -to $VarToEmail -From $VarFromEmail -Subject $VarUnsuccessfulEmailSubject -Body $VarUnsuccessfulFileCountEmailBody -Smtpserver $VarSMTPServer -Credential $VarSMTPCredential
    }

###Remove local backup files from Tag Appliance
Remove-Item -Path $VarTagApplianceBackupStagingDeleteDatabaseZipPath, $VarTagApplianceBackupStagingDeleteDatabasePath, $VarTagApplianceBackupStagingOPCZip, $VarTagApplianceBackupStagingWebsiteZip

###Remove temporary PowerShell drive
Remove-PSDrive -Name BackupDrive

#Clear all user defined variables
Clear-Variable -Name Var*