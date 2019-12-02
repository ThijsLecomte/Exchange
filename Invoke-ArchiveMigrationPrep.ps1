<#
.SYNOPSIS
    This is a script that prepares the mailbox for archive migration.
    It gets all the current flagged items and exports some basic info for the user to a CSV that needs to be sent to Quadrotech.

.DESCRIPTION
    This script uses a CSV file as input (defined in usersCSV) in which all the Emailaddress of the users who are to be migrated are populated.
        This CSV requires the header 'Emailaddress'
        The script prompts the user to select the right CSV
    
        EmailAddress
        John.Doe@domain.com
        Jane.Doe@domain.com

    This scripts the mailboxes and outputs all flagged items into a CSV.

    It also gets some basic information for all mailboxes that are needed by Quadrotech to start the migration.
    This info is exported to a CSV, the path is defined in the ExportPath Variable

    It uses EWS as an API, so this needs to be installed on the computer on which it is executed.
        It can be downloaded on any PC using the following link: https://www.microsoft.com/en-us/download/details.aspx?id=42951
        It is installed by default on all Exchange servers

    You must login with an administrator account that has application impersonation rights in Exchange.

    It outputs to file to the export path defined in the ExportPath variable.

    This scripts creates a log file each time the script is executed. 
    It deleted all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\CustomScript\Export-FlaggedMails.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.PARAMETER DLLPathEWS
    Path to DLL of EWS
    Default path of regular computer: C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll
    Default path on Exchange Server: C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll

.PARAMETER EWSURL
    This is the URL (public or internal) through which EWS is made available on the Exchange Server

.PARAMETER ExportPath
    Path to which CSV should be exported.  

.PARAMETER EVServer
    EV Server

.PARAMETER EVPowershellPath
    Path of the EV Powershell Module

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\Export-FlaggedMails.ps1 -DLLPathEWS "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" -EWSURL "https://outlook.rodekruis.be/EWS/Exchange.asmx" -ExportPath "C:\temp" -EVServer "SERVER1" -EVPowershellPath ".\Symantec.EnterpriseVault.PowerShell.Monitoring.dll"

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\Export-FlaggedMails.ps1 -logPath "C:\Windows\Temp\CustomScripts\CustomLogPath.txt"

.NOTES
    File Name  : Invoke-ArchiveMigrationPrep.ps1  
    Author     : Thijs Lecomte 
    Company    : The Collective Consulting BV
#>

#region Parameters
#Define Parameter LogPath
param (
    $LogPath = "C:\Windows\Temp\CustomScripts\Invoke-ArchiveMigrationPrep.txt",
    $DLLPathEWS = "C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll",
    $EWSURL = "https://outlook.domain.com/EWS/Exchange.asmx",
    $ExportPath = "C:\temp\ArchiveMigration2",
    $EVServer,
    $EVPowershellPath
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

Function Connect-EWS($EWSURL,$MailboxToImpersonate, $AdminCredentials){
    Write-Log "[INFO] - Starting Function Connect-EWS"

    try{
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
        $psCred = $AdminCredentials
        $creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
        $service.Credentials = $creds
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

Function Get-FlaggedItems($Service,$mailboxToSearch){
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

    Write-Log "[INFO] - Found $($flagged.Count) flagged items"
    Write-Log "[INFO] - Exiting Function Get-FlaggedItems"

    return $flagged
}

Function Export-ToCSV($items, $ExportPath, $MailboxToSearch){
    Write-Log "[INFO] - Starting Function Export-ToCSV"

    $date = Get-Date -Format D

    $path = "$ExportPath\ $date - $MailboxToSearch.csv"

    try{
        $items | Export-csv -Path $path -NoTypeInformation
    }
    catch{
        Write-Log "[ERROR] - Exporting to csv to $path"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Export-ToCSV"
}

Function Import-Users ($CSVPath){
    Write-Log "[INFO] - Starting Function Import-Users ($CSVPath)"
    Try{
        $users = Import-Csv -Path $CSVPath
        Write-Log "[INFO] - Succesfully imported users, found $($users.count)"
    }
    Catch{
        Write-Log "[ERROR] - Import of users at path $CSVPath failed"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Import-Users ($CSVPath)"

    return $users
}

Function Connect-OnPremExchange{
    Write-Log "[INFO] - Starting Function Connect-OnPremExchange"
    Try{
        $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell "
        Invoke-Expression $CallEMS
        Write-Log "[INFO] - Connected to on-prem Exchange"
    }
    Catch{
        Write-Log "[ERROR] - Connecting to on-prem Exchange"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Connect-OnPremExchange"
}

Function Get-MigrationInfo($EmailAddress){
    Write-Log "[INFO] - Starting Function Get-MigrationInfo($DisplayName)"
    Try{
        $mailbox = Get-Mailbox -Identity $EmailAddress | Select-Object -Property DisplayName, SamAccountName, PrimarySmtpAddress
        Write-Log "[INFO] - Got MailboxInfo"
    }
    Catch{
        Write-Log "[ERROR] - Getting MailboxInfo"
        Write-Log "$($_.Exception.Message)"
    }
    
    Try{
        $size = Get-MailboxStatistics -Identity $EmailAddress | select-object -Property TotalItemSize
        Write-Log "[INFO] - Got mailboxSize"
    }
    Catch{
        Write-Log "[ERROR] - Getting mailboxSize"
        Write-Log "$($_.Exception.Message)"
    }
    

    $UserStat = [PSCustomObject]@{
        "DisplayName" = $mailbox.DisplayName
        "Alias" = $mailbox.SamAccountName
        "Email" = $mailbox.PrimarySmtpAddress
        "MailboxSize" = $size.TotalItemSize
    }

    Write-Log "[INFO] - Exiting Function Get-MigrationInfo($DisplayName)"

    return $UserStat
}

Function Get-CSVPath{
    Write-Log "[INFO] - Starting Function Get-CSVPath"

    Add-Type -AssemblyName System.Windows.Forms

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = 'C:\scripts\Migration\Archive Migration'
        Filter = 'Comme Seperated Files (*.csv)|*.csv'
        Title = 'Please browse to the CSV'
    }

    $DialogResult = $FileBrowser.ShowDialog()

    $csvPath = $FileBrowser.FileName

    Write-Log "File chosen: $csvPath"
    Write-Log "[INFO] - Exiting Function Get-CSVPath"

    return $csvPath
}


Function Invoke-EVArchiveCheck($DisplayName, $EVServerName, $EVPowerShellPath){
    Write-Log "[INFO] - Starting Function Invoke-EVArchiveCheck"
    try {
        Write-Log "[INFO] - Getting Archive for $DisplayName"
        $archive = Invoke-command -ConfigurationName Microsoft.PowerShell32 -ComputerName $EVServerName -ScriptBlock {param($DisplayName, $EVPowershellPath) Import-Module $EVPowershellPath; Get-EVArchive -ArchiveName "$DisplayName" } -Args $DisplayName, $EVPowershellPath
        Write-Log "[INFO] - Got EV Archive for $DisplayName"
    }
    catch{
        Write-Log "[ERROR] - Getting EV Archive for $DisplayName"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function Invoke-EVArchiveCheck"

    if($archive){
        return $true
    }
    else{
        return $false
    }
}
#endregion

#region Operational Script
Write-Log "[INFO] - Starting script"

$creds = Get-Credential -Message "Please provide Exchange Administrator account that has Application impersonation permissions om mailbox"

$csvPath = Get-CSVPath

if($csvPath -eq ""){
    Write-Log "[ERROR] - You didn't select a file, exiting script. Please select a CSV and try again" -ForeGroundColor Red
    Break
}


Import-EWS -DLLPathEWS $DLLPathEWS

Connect-OnPremExchange

$users = Import-Users -CSVPath $csvPath

$InformationUsers = @()

foreach($user in $users){

    $userinfo = Get-MigrationInfo -EmailAddress $user.EmailAddress

    Write-Log "[INFO] - User email is $($userInfo.Email)"

    $archive = Invoke-EVArchiveCheck -DisplayName $userinfo.DisplayName -EVServerName $EVServer -EVPowerShellPath $EVPowershellPath

    if($archive){
        Write-Log "[INFO] - User has an archive"
        $service = Connect-EWS -EWSURL $EWSURL -MailboxToImpersonate $userinfo.Email -AdminCredentials $creds

        $flaggedItems = Get-FlaggedItems -service $service -mailboxToSearch $userinfo.Email

        if($flaggedItems.Count -ne 0){
            Export-ToCSV -items $flaggedItems -ExportPath $ExportPath -MailboxToSearch $userinfo.Email
        }
        else{
            Write-Log "[INFO] - No flagged items found for $($userinfo.DisplayName), not creating a file."
        }

        $informationUsers += $userinfo

    }
    else{
        Write-Log "[INFO] - User doesn't have an Archive, skipping"
    }
}

$date = Get-Date -Format D

$InformationUsers | Export-Csv -path "$ExportPath\UsersInfo $date.csv" -NoTypeInformation -Force

Write-Log "[INFO] - Stopping script"
#endregion