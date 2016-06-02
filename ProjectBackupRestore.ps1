#Initial variables
$restoretemplatename = "RestoreTemplate"
$server = "<enter server name>"
#Setup user to connect to SharePoint
$username = "<account>"
#Get credential from encoded file
$file = "\\$server\InBox\ProjectBackups\file.txt"
$key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$password = Get-Content $file | ConvertTo-SecureString -key $key
$credential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
$Global:test = $null
$Global:job = $null

#Functions
function perComp ($cmdjob)
{
    $i = 0                  
    Do
    {            
        Write-Progress -Activity "Percentage completed" -status $cmdjob.state -PercentComplete $i                
        Start-Sleep -s 2
        If ($i -lt 100)
        {
            $i++                                       
        }
        If ($cmdjob.state -eq "completed")
        {
            Write-Progress -Activity "Percentage completed" -status $cmdjob.state -PercentComplete 100
            Start-Sleep -s 2                                      
        }
    } while ($cmdjob.state -ne "completed")             
    Write-Progress -Activity "Complete" -Complete
}

function checkSite ($checksite)
{
    Write-Host "Checking to see if site exists..."  -ForegroundColor Green
    $Global:test = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($site=$checksite) Add-PSSnapin Microsoft.SharePoint.Powershell; 
        Get-SPWeb($site) -ErrorAction SilentlyContinue} -ArgumentList $checksite -AsJob

    perComp $test
    return
}

function restoreSite ($completepathonclient, $restorepathonserver, $restoresite, $rootsite)
{
    Write-Host "Restoring site..."  -ForegroundColor Green
                Start-Sleep -s 3
                
                Copy-Item $completepathonclient -Destination \\$server\InBox\ProjectBackups\Restore\ -Force
                $Global:job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($site=$restoresite, $restore=$restorepathonserver) Add-PSSnapin Microsoft.SharePoint.Powershell; 
                    Import-SPWeb $site -Path $restore -UpdateVersions Ignore -IncludeUserSecurity -HaltOnError -NoLogFile} -ArgumentList $restoresite,$restorepathonserver -AsJob
                
                perComp $job                
                
                Remove-Item -Path $restorepathonserver
                Start-Sleep -s 3   
                Write-Host "Site restored"  -ForegroundColor Green  
                Start-Sleep -s 3
            
                #Cleanup site
                Write-Host "Starting cleanup of site..."  -ForegroundColor Green
                Start-Sleep -s 3
           
                $Global:job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($root=$rootsite,$site=$restoresite) Add-PSSnapin Microsoft.SharePoint.Powershell; 
                    $web = Get-SPWeb $site;
                    $notebookitem = $web.Navigation.QuickLaunch | where {$_.title -eq "Notebook"};
                    $notebookitemurl = $notebookitem.url;
                    $partialnotebookul = $notebookitem.url;   
                    $result = $notebookitemurl -split "/";                
                    $indexcount = $result.Count;
                    $indexcount = $indexcount - 1;                
                    [System.Collections.ArrayList]$oresult = $result;
                    $oresult.RemoveRange(0,$indexcount);                
                    $partialnotebookul = $notebookitemurl -split "SiteAssets";
                    $indexcount = $partialnotebookul.Count;                
                    $indexcount = $indexcount - 1;
                    [System.Collections.ArrayList]$opartialnotebookul = $partialnotebookul;
                    $opartialnotebookul.RemoveRange(1,$indexcount);                
                    $fixedurl = $root + $opartialnotebookul + "_layouts/15/WopiFrame.aspx?sourcedoc=" + $opartialnotebookul + "SiteAssets/" + $oresult + "&action=default&RootFolder=" + $opartialnotebookul +"SiteAssets%2f" + $oresult;                                      
                    Start-Sleep -s 10;
                    $notebookitem.Delete();
                    Start-Sleep -s 10;
                    $node = New-Object -TypeName Microsoft.SharePoint.Navigation.SPNavigationNode("Notebook", $fixedurl);
                    Start-Sleep -s 10;
                    $web = Get-SPWeb $site;
                    Start-Sleep -s 10;
                    $homeitem = $web.Navigation.QuickLaunch | where {$_.title -eq "Home"};
                    $web.Navigation.QuickLaunch.Add($node, $homeitem);
                    $web.Update()
                    } -ArgumentList $rootsite,$restoresite -AsJob
                
                perComp $job                

                Write-Host "Site cleaned up"  -ForegroundColor Green 
                Write-Host "Go through site and perform testing and clean up additional items as necessary"  -BackgroundColor Red
}

function selectedNothing
{
    Write-Host "Did not select [Y]es or [N]o - so exiting" -BackgroundColor Red
    Exit
}

function serverError
{
    Write-Host "Error running command on the server - so exiting" -BackgroundColor Red
    Exit
}

function selectedNo
{
    Write-Host "Selected [N]o - so exiting" -BackgroundColor Red
    Exit
}
#End functions

 #Initial prompt
"Do want to backup or restore a site?"
$choice = Read-Host "[B]ackup or [R]estore"

If ($choice -eq "B" -or $choice -eq "Backup")
{
    ####Backup steps
    #Get user input for project title and URL
    $project = Read-Host "Enter the project title (ex. 141A1test)"
    $backuppathonclient = Read-Host "Enter the folder path, on your computer, where you want to place the backup file (ex. C:\Outbox)"      
    If ($project -like '*.cmp*')
    {
        $backuppathonserver = "C:\InBox\ProjectBackups\Backup\" + $project
        $fullbackuppathonclient = $backuppathonclient + "\" + $project
    }
    Else
    {
        $backuppathonserver = "C:\InBox\ProjectBackups\Backup\" + $project + ".cmp"
        $fullbackuppathonclient = $backuppathonclient + "\" + $project + ".cmp" 
    }              
    
    If ($backuppathonclient.EndsWith("\"))
    {
        $backuppathonclient = $backuppathonclient.TrimEnd("\")
    }
    $backupsite = Read-Host "Enter the URL of the site you want to backup (ex. http://domain.com/Sites/141A1testbackup)"

    If ((!$project) -or (!$backuppathonclient) -or (!$backupsite))
    {
        Write-Host "You didn't enter anything in one of the prompts. Try again - exiting now..." -BackgroundColor Red
        Exit
    }

    #Processing URL
    $separator = "/"
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $result = $backupsite.Split($separator,$option)                                                                          
    [System.Collections.ArrayList]$oresult = $result                                
    
    #Checking if input URL was valid (ie. not a /year/sites pattern)
    If (!(($result.Count -eq 5) -or ($result.Count -eq 3)))
    {       
       Write-Host "URL was not correct. Try again - exiting now..." -BackgroundColor Red
       Exit 
    }
    
    $verifypath = Test-Path $backuppathonclient
    If ($verifypath -eq $false)
    {
        Write-Host "Backup folder on your computer" $project "does not exist. Create it first and try again - exiting now..." -BackgroundColor Red
        Exit
    }    

    #Check if site exists
    checkSite $backupsite    
          
        If ($test.HasMoreData -eq $false)
        {
            Write-Host "Site to backup does not exist. Try again - exiting now..." -BackgroundColor Red
            Exit
        }
        Else
        {
            #Backup site    
            Write-Host "Backing up site..."  -ForegroundColor Green
    
            Remove-Item -Path \\$server\InBox\ProjectBackups\Backup\*.cmp
            $job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($site=$backupsite, $backup=$backuppathonserver) Add-PSSnapin Microsoft.SharePoint.Powershell; 
                Export-SPWeb $site -Path $backup -IncludeVersions All -HaltOnError -NoLogFile} -ArgumentList $backupsite,$backuppathonserver -AsJob
            
            perComp $job
            
            Copy-Item \\$server\InBox\ProjectBackups\Backup\*.cmp -Destination $backuppathonclient -Force
            Remove-Item -Path \\$server\InBox\ProjectBackups\Backup\*.cmp
    
            #test document to see if site was backed up    
            Start-Sleep -s 3

            $document = Get-ItemProperty -Path $fullbackuppathonclient    
            If (($document.Length) -gt 0)
            { 
                Write-Host "Backed up site"  -ForegroundColor Green
                Write-Host "Are you sure you want to delete this site from SharePoint at URL" $backupsite -BackgroundColor Red
                $backupdecision = Read-Host "[Y]es or [N]o"

                If ($backupdecision -eq "Y" -or $backupdecision -eq "Yes")
                {                    
                    Write-Host "Removing backup site from SharePoint..."  -ForegroundColor Green
                    Start-Sleep -s 3
                    $job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($site=$backupsite) Add-PSSnapin Microsoft.SharePoint.Powershell; 
                        Remove-SPWeb $site -Confirm:$false -Recycle} -ArgumentList $backupsite -AsJob
                        
                    perComp $job
                                         
                    Write-Host "Removed backup site from SharePoint"  -ForegroundColor Green                
                }
                ElseIf ($backupdecision -eq "N" -or $backupdecision -eq "No")
                {                    
                    selectedNo
                }
                Else
                {                    
                    selectedNothing
                }                
            }
            Else
            {
                Write-Host "Backup was unsuccessful. Try again - exiting now..." -BackgroundColor Red
                Exit
            }   
        }    
}
ElseIf ($choice -eq "R" -or $choice -eq "Restore")
{    
    ####Restore steps
    #Get user input for project title and URL
    $project = Read-Host "Enter the backup file name"
    $restorepathonclient = Read-Host "Enter the folder path, on your computer, to the backup file (ex. C:\Outbox)"
    If ($restorepathonclient.EndsWith("\"))
    {
        $restorepathonclient = $restorepathonclient.TrimEnd("\")
    }
    If ($project -like '*.cmp*')
    {
        $completepathonclient = $restorepathonclient + "\" + $project
        $restorepathonserver = "\\$server\InBox\ProjectBackups\Restore\" + $project
    }
    Else
    {
        $completepathonclient = $restorepathonclient + "\" + $project + ".cmp" 
        $restorepathonserver = "\\$server\InBox\ProjectBackups\Restore\" + $project + ".cmp"
    }                  
    
    $restoresite = Read-Host "Enter the URL of the site you want to restore (ex. http://domain.com/Sites/141A1testimport)"    

    If ((!$project) -or (!$restorepathonclient) -or (!$restoresite))
    {
        Write-Host "You didn't enter anything in one of the prompts. Try again - exiting now..." -BackgroundColor Red
        Exit
    }

    #Processing URL to get root site URL and Site Collection URL
    $separator = "/"
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $result = $restoresite.Split($separator,$option)                                                                          
    [System.Collections.ArrayList]$oresult = $result
    
    #Checking if input URL was valid
    If (!(($result.Count -eq 5) -or ($result.Count -eq 3)))
    {       
       Write-Host "URL was not correct. Try again - exiting now..." -BackgroundColor Red
       Exit 
    }
                                    
    $indexcount = $result.Count              
    $oresult.RemoveRange(0,1);             
    $sitecollection = "/" + $oresult[1] + "/" + $oresult[2]
    $indexcount = $oresult.Count;  
    $indexcount = $indexcount - 1; 
    $oresult.RemoveRange(1,$indexcount);    
    $rootsite = $oresult.ToString()
    $rootsite = "http://" + $oresult         
    $sitecollection = $rootsite + $sitecollection   

    $verifypath = Test-Path $completepathonclient
    If ($verifypath -eq $false)
    {
        Write-Host "Backup file" $project "does not exist. Try again - exiting now..." -BackgroundColor Red
        Exit
    }

    #Check if site exists
    checkSite $restoresite     
    
        If ($test.HasMoreData -eq $false)
        {            
            Write-Host "Site" $restoresite "does not exist." -ForegroundColor Green            
            Write-Host "Creating site now..."  -ForegroundColor Green

            #Creating site 
            $job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {param($sc=$sitecollection,$site=$restoresite,$rtn=$restoretemplatename) Add-PSSnapin Microsoft.SharePoint.Powershell; 
                        $scobject = Get-SPWeb($sc); 
                        $template = $scobject.GetAvailableWebTemplates(1033) | Where-Object {$_.Title -eq $rtn}; 
                        $newsite = New-SPWeb $site -Name "IMPORTSITE";
                        $siteobject = Get-SPWeb($site);
                        $siteobject.ApplyWebTemplate($template.Name)} -ArgumentList $sitecollection,$restoresite,$restoretemplatename -AsJob
            
            perComp $job            
            
            #Check to see if Invoke-Command executed sucessfully
            If ($error -ne $null)
            {
                serverError                
            }
                
            #Checking site again to see if it was created sucessfully
            checkSite $restoresite                      
            
            If ($test.HasMoreData -eq $false)
            {
                Write-Host "Site" $restoresite "was not created. Try again - exiting now..." -BackgroundColor Red
                Exit
            }
            Else
            {
                Start-Sleep -s 3
                Write-Host "Site created"  -ForegroundColor Green
                Start-Sleep -s 3
                #Restore site
                restoreSite $completepathonclient $restorepathonserver $restoresite $rootsite                 
            }
        }
        Else
        {
            
            Write-Host "Site" $restoresite "already EXISTS"  -BackgroundColor Red
            Write-Host "Are you sure you want to restore to this URL (it will overwrite all content)" $restoresite -BackgroundColor Red
            $restoredecision = Read-Host "[Y]es or [N]o"


            If ($restoredecision -eq "Y" -or $restoredecision -eq "Yes")
            {
                #Restore site
                restoreSite $completepathonclient $restorepathonserver $restoresite $rootsite                
            }
            ElseIf ($restoredecision -eq "N" -or $restoredecision -eq "No")
            {                
                selectedNo
            }
            Else
            {                
                selectedNothing
            }
        }      
}
Else
{
    Write-Host "Did not select [B]ackup or [R]estore - so exiting" -BackgroundColor Red
    Exit
}