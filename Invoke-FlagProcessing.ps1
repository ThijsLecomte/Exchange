<#
.SYNOPSIS
    This script checks which items where flagged in a previous time (imported through a CSV) and removes all the flags of the mails which are currently flagged, but were not flagged the last time.

.DESCRIPTION
    This script imports a CSV which contains all the emails that were flagged at the moment the CSV was created.
    It gets all the current flagged items and unflags all items that are currently flagged, but are not part of the CSV.

    The user is prompoted to select the CSV. The CSV should always be using the following type of name:
        "C:\Temp\flaggedITems Wednesday, 16 October 2019 - test.collective@domain.com.csv"
    It uses EWS as an API, so this needs to be installed on the computer on which it is executed.
        It can be downloaded on any PC using the following link: https://www.microsoft.com/en-us/download/details.aspx?id=42951
        It is installed by default on all Exchange servers

    It checks one account, the info of the account has derived from the CSVFileName that is used to import.

    You must login with an administrator account that has application impersonation rights in Exchange.

    This scripts creates a log file each time the script is executed. 
    It deleted all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\CustomScript\Invoke-FlagProcessing.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.PARAMETER DLLPathEWS
    Path to DLL of EWS
    Default path of regular computer: C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll
    Default path on Exchange Server: C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll

.PARAMETER EWSURL
    This is the URL (public or internal) through which EWS is made available on the Exchange Server

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\Invoke-FlagProcessing.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\Export-FlaggedMails.ps1 -DLLPathEWS "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -EWSURL "https://outlook.domain.com/EWS/Exchange.asmx"

.NOTES
    File Name  : Invoke-FlagProcessing.ps1  
    Author     : Thijs Lecomte 
    Company    : The Collective Consulting BV
#>

#region Parameters
#Define Parameter LogPath
param (
    $LogPath = "C:\Temp\CustomScripts\Invoke-FlagProcessing.txt",
    $DLLPathEWS = "C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll",
    $EWSURL = "https://outlook.domain.com/EWS/Exchange.asmx",
    $csvImportPath = "C:\temp\ArchiveMigration",
    $csvDonePath = "C:\temp\ArchiveMigration\OK"
)
#endregion

#region variables
$MaxAgeLogFiles = 30

#region Log file creation
#Create Log file
  Try{
    #Create log file based on logPath parameter followed by current date
    $date = Get-Date -Format yyyyMMddTHHmmss
    $date = $date.replace("/","").replace(":","")
    $logpath = $logpath.insert($logpath.IndexOf(".txt")," $date")
    $logpath = $LogPath.Replace(" ","")
    New-Item -Path $LogPath -ItemType File -Force -ErrorAction Stop

    #Delete all log files older than x days (specified in $MaxAgelogFiles variable)
    $limit = (Get-Date).AddDays(-$MaxAgeLogFiles)
    Get-ChildItem -Path $logPath.substring(0,$logpath.LastIndexOf("\")) -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
    
  } catch {
    #Throw error if creation of loge file fails
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup($_.Exception.Message,0,"Creation Of LogFile failed",0x1)
    exit
  }
#endregion

#region functions
#Define Log function
Function Write-Log {
    Param ([string]$logstring)

    $DateLog = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $WriteLine = $DateLog + "|" + $logstring
    try {
        Add-Content -Path $LogPath -Value $WriteLine -ErrorAction Stop
    } catch {
        Start-Sleep -Milliseconds 100
        Write-Log $logstring
    }
    Finally{
        if($logstring -contains "[ERROR]"){
            Write-Host $logstring -ForegroundColor Red
        }
        else{
            Write-Host $logstring 
        }
    }
}

Function Import-EWS($DLLPathEWS){
    Write-Log "[INFO] - Starting Function Import-EWS"
    Try{
        Import-Module $DLLPathEWS
        Write-Log "[INFO] - Imported DLL"
    }
    Catch{
        Write-Log "[ERROR] - Could not import DLL from Path $DLLPathEWS"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Import-EWS"
}

Function Connect-EWS($EWSURL,$MailboxToImpersonate, $AdminCreds){
    Write-Log "[INFO] - Starting Function Connect-EWS"

    try{
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
        
        $service.Credentials = $AdminCreds
        $service.URL = $EWSURL
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate )
        Write-Log "[INFO] - Connected to EWS succesfully"
    }
    catch{
        Write-Log "[ERROR] - Error connecting to EWS to $EWSURL with account $AdminAccount"
        Write-Log "$($_.Exception.Message)"
    }

    Write-Log "[INFO] - Exiting Function Connect-EWS"

    return $service
}

Function Get-FlaggedItems($Service){
    Write-Log "[INFO] - Starting Function Get-FlaggedItems"

    #Get all folders
    try{
        $RootFolder= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$ImpersonatedMailboxName)
        $fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(31000)
        $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
        $folders = $service.FindFolders($RootFolder,$fvFolderView)
        Write-Log "[INFO] - Correctly got all folders"
    }
    catch{
        Write-Log "[ERROR] - Getting all folders"
        Write-Log "$($_.Exception.Message)"
    }


    $flagged = @()

    #Check all folders and search for ToDoSearch Folder (this is the folder with all the tasks)
    Foreach($folder in $folders){
        Write-Log "[INFO] - Checking folder $($folder.DisplayName)"
        
        if($folder.DisplayName -eq "To-Do Search" -or $folder.DisplayName -eq "Zoeken naar taken"){
            Write-Log "[INFO] - Found To Do folder"

            try{
                $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView 300
                $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+Exists(new-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1090, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer))
                $items = $folder.FindItems($filter, $view)
                Write-Log "[INFO] - Got all items"
            }
            catch{
                Write-Log "[ERROR] - Getting all items from folder $($folder.DisplayName)"
                Write-Log "$($_.Exception.Message)"
            }

            #Check all items to see which item is flagged
            Foreach($item in $items){
                if($item.flag.FlagStatus -eq "Flagged"){
                    #Found a flagged item, add it to our collection
                    $flagged += $item
                }
            }
            Break
        }
    }

    Write-Log "[INFO] - Found $($flagged.Count) current flagged items"
    Write-Log "[INFO] - Exiting Function Get-FlaggedItems"

    return $flagged
}

Function Set-ItemCompleted($MailItem){
    Write-Log "[INFO] - Starting Function Set-ItemCompleted for $($MailItem.Subject) received on $($MailItem.DateTimeReceived)"

    Try{
        $MailItem.flag.FlagStatus = "Complete"
        $MailItem.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve);
        Write-Log "[INFO] - Setting ItemCompleted for $($MailItem.Subject) received on $($MailItem.DateTimeReceived) completed succesfully"
    }
    Catch{
        Write-Log "[ERROR] - Setting ItemCompleted for $($MailItem.Subject) received on $($MailItem.DateTimeReceived) failed"
        Write-Log "$($_.Exception.Message)"
    }
    

    Write-Log "[INFO] - Exiting Function Set-ItemCompleted"
}

Function Import-FlaggedItems($csvImportPath){
    Write-Log "[INFO] - Starting Function Import-FlaggedItems($csvImportPath)"
    Try{
        $items = Import-Csv -Path $csvImportPath
        Write-Log "[INFO] - Imported CSV, got $($items.count) items that were flagged pre archive migration"
    }
    Catch{
        Write-Log "[ERROR] - Importing CSV"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Import-FlaggedItems($csvImportPath)"

    return $items
}

Function Invoke-FlagComparision($previousFlags,$currentFlags, $CSVCreated){
    Write-Log "[INFO] - Starting Function Invoke-FlagComparision"
    $CSVCreated = Get-Date $CSVCreated -Hour 0 -Minute 0 -Second 0
    Write-Log "[INFO] - CSV Datetime is $CSVCreated"
    Write-Log "[INFO] - Starting foreach to enumerate over all current flagged items"
    foreach($flag in $currentFlags){
        Write-Log "[INFO] - Checking current flagged item $($flag.Subject) received on $($flag.DateTimeReceived) with InternetMessageID $($flag.InternetMessageId)"
        if($flag.Flag.StartDate -ge (Get-Date $CSVCreated -Format "yyyy/MM/dd HH:mm:ss")){
            Write-Log "[INFO] - Flag was created - $($flag.Flag.StartDate) - after CSV file was generated - $CSVCreated - so it is current"
            $flagCurrent = $true
        }
        else{
            Write-Log "[INFO] - Flag wasn't been created - $($flag.Flag.StartDate) - after CSV file was generated - $CSVCreated"
            $flagCurrent = $false
            foreach($previousFlagged in $previousFlags){
                Write-Log "[INFO] - Comparing item with previous flagged item $($previousFlagged.Subject) received on $($previousFlagged.DateTimeReceived) with InternetMessageID $($previousFlagged.InternetMessageId)"
                if($flag.InternetMessageId -eq $previousFlagged.InternetMessageId){
                    Write-Log "[INFO] - Match found, flag is current"
                    $flagCurrent = $true
                    Break
                }
            }
        }
        

        if($flagCurrent -eq $false){
            Write-Log "[INFO] - No match found, flag is not current. Removing flag"
            Set-ItemCompleted -MailItem $flag
        }
        Write-Log "[INFO] ------- Foreach looped ------------"
    }
    Write-Log "[INFO] - Exiting Function Invoke-FlagComparision"
}

Function Get-CSVPath{
    Write-Log "[INFO] - Starting Function Get-CSVPath"

    Add-Type -AssemblyName System.Windows.Forms

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = 'C:\temp\ArchiveMigration'
        Filter = 'Comme Seperated Files (*.csv)|*.csv'
        Title = 'Please browse to the CSV'
    }

    $DialogResult = $FileBrowser.ShowDialog()

    $csvPath = $FileBrowser.FileName

    Write-Log "File chosen: $csvPath"
    Write-Log "[INFO] - Exiting Function Get-CSVPath"

    return $csvPath
}
#endregion

#region Operational Script
Write-Log "[INFO] - Starting script"

$csvs = Get-ChildItem -Path $csvImportPath -File

Write-Log "[INFO] - Found $($csvs.count) CSVs"

Import-EWS -DLLPathEWS $DLLPathEWS

$psCred = Get-Credential -Message "Exchange Admin creds"
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())

foreach($csv in $csvs){
    Write-Log "[INFO] - Checking $($csv.FullName)"

    if($csv.Name.StartsWith("UsersInfo")){
        Write-Log "[INFO] - Not a valid CSV of flags - $($csv.FullName)"
    }
    else{
        $PreviousFlaggedItems = Import-FlaggedItems -csvImportPath $csv.FullName

        $CSVTimeCreated = (Get-Item $csv.FullName| Select-Object -Property CreationTime).CreationTime

        #Get username dynamicallly from path of CSV file
        $split = $($csv.FullName).Split(" ")[-1]
        $MailboxToSearch = $split.substring(0,$split.indexOf(".csv"))

        $service = Connect-EWS -EWSURL $EWSURL -MailboxToImpersonate $MailboxToSearch -AdminCreds $creds

        $CurrentFlaggedItems = Get-FlaggedItems -Service $service

        Invoke-FlagComparision -previousFlags $PreviousFlaggedItems -currentFlags $CurrentFlaggedItems -CSVCreated $CSVTimeCreated

        Move-Item -Path $csv.FullName -Destination $csvDonePath -force

        Write-Log "[INFO] - Moved CSV to different location. End processing of current CSV"
        Write-Log "-----------------------------------------------------"
    }
}
Write-Log "[INFO] - Stopping script"
#endregion