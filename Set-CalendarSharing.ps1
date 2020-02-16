<#
.SYNOPSIS
    Sets calendar permissions for all users to a specified scope and accessrights.

.DESCRIPTION
    Sets calendar permissions for all users to a specified scope and accessrights.
    It uses a few known languages and currently only supports German, French, Dutch and Chinese. Edit the foldercalendars parameter to add more languages

    Exchange online module needs to be installed: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

    This scripts creates a log file each time the script is executed. 
    It deleted all the logs it created that are older than 30 days. This value is defined in the MaxAgeLogFiles variable.

.PARAMETER LogPath
    This defines the path of the logfile. By default: "C:\Windows\Temp\CustomScript\Set-CalendarSharing.txt"
    You can overwrite this path by calling this script with parameter -logPath (see examples)

.PARAMETER accessrights
    Specifies which accessrights should be applied: Reviewer, Author, Contributor, Editor, None, NonEditingAuthor, Owner, PublishingEditor, PublishingAuthor

.PARAMETER Scope
    Specifies which users should receive the accessrights. Use default to add permissions to all calendars

.PARAMETER FolderCalendars
    Array list that holds all supported languages, add other languages here.
    For example if you have Spanish users, add 'calendario'

.EXAMPLE
    Use the default logpath without the use of the parameter logPath
    ..\Set-CalendarSharing.ps1

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\Set-CalendarSharing.ps1 -logPath "C:\Windows\Temp\CustomScripts\CustomLogPath.txt"

.NOTES
    File Name  : Set-CalendarSharing.ps1  
    Author     : Thijs Lecomte 
    Company    : The Collective Consulting BV
#>

#region Parameters
#Define Parameter LogPath
param (
    $LogPath = "C:\Windows\Temp\CustomScripts\Set-CalendarSharing.txt",
    $accessrights = "Reviewer",
    $Scope = "Default",
    $FolderCalendars = @("Agenda", "Calendar","Calendrier", "Kalender", "日历")
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
        Write-Host $logstring
    }
}

Function Connect-ExchangeOnline{
    Write-Log "[INFO] - Starting Function ConnectoTo-ExchangeOnline"

    Try{
        $CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
        . "$CreateEXOPSSession\CreateExoPSSession.ps1"
        Connect-EXOPSSession
        Write-Log "[INFO] - Connected to Exchange Online"
    }
    Catch{
        Write-Log "[ERROR] - Signing into Exchange Online"
        Write-Log "$($_.Exception.Message)"
    }
    Write-Log "[INFO] - Exiting Function ConnectoTo-ExchangeOnline"
}

Function Get-Users{
    Write-Log "[INFO] - Starting Function Get-Users"
    try{
        $users = Get-Mailbox -RecipientTypeDetails UserMailbox
        Write-Log "[INFO] - Got users"
    }
    catch{
        Write-Log "[ERROR] - Error getting users"
        Write-Log "$($_.Exception.Message)"
    }

    Write-Log "[INFO] - Starting Function Get-Users"

    return $users
}

Function Set-Sharing($users,$accessrights, $scope, $FolderCalendars){
    Write-Log "[INFO] - Starting Function Set-Sharing"
    Write-Log "[INFO] - Set sharing for $($users.count) users, with accessrights $accessrights"
    foreach($user in Get-Mailbox -RecipientTypeDetails UserMailbox) {
        $calendarName = $null
        Write-Log "[INFO] - Checking $($user.Identity)"
        $Calendars = (Get-MailboxFolderStatistics $user.Identity -FolderScope Calendar)
        if($calendars.count -gt 1){
            Write-Log "[INFO] - User got multiple calendars, checking with known calendars"
            foreach($calendar in $calendars){
                #Only continue if correct calendar names hasn't been found
                if($calendarName -eq $null){
                    Write-Log "[INFO] - Checking calendar: $($calendar.Name)"
                    foreach($folder in $FolderCalendars){
                        Write-Log "[INFO] - Checking known calendar: $folder"
                        if($folder -eq $calendar.Name){
                            Write-Log "[INFO] - Match found, saving $folder as CalendarName"
                            $calendarName = $folder
                            break;
                        }
                    }
                }
            }
        }
        else{
            Write-Log "[INFO] - User only has one calendar $($Calendars.Name)"
            $calendarName = $Calendars.Name
        }
        if($calendarName -eq $null){
            Write-Host "[ERROR] - No calendar found, skipping"
        }
        else{
            $cal = $user.Identity+":\$CalendarName"
            Write-Log "[INFO] - Setting permissions for $($cal) with scope $scope to $accessrights"
            
            try{
                Set-MailboxFolderPermission -Identity $cal -User $scope -AccessRights $accessrights
                Write-Log "[INFO] - Set calendar properties"
            }
            catch{
                Write-Log "[ERROR] - Error setting properties"
                Write-Log "$($_.Exception.Message)"
            }
        }
        Write-Host "-------------"
    }

    Write-Log "[INFO] - Starting Function Set-Sharing"
}


#endregion

#region Operational Script
Write-Log "[INFO] - Starting script"
Try{
    Connect-ExchangeOnline

    $users = Get-Users

    Set-Sharing -users $users -accessrights $accessrights -scope $scope -FolderCalendars $FolderCalendars
}
Catch{
    Write-Log "[ERROR] - Signing into Exchange Online"
    Write-Log "$($_.Exception.Message)"
}
Finally{
    #Remove all current PS Sessions
    Get-PSSession | Remove-PSSession
}

Write-Log "[INFO] - Stopping script"
#endregion