# Script to export the user mailbox

Param(
    [Parameter(Mandatory=$True, HelpMessage='Enter the email address that you want to export')]
    $Mailbox,
    [Parameter(Mandatory=$True, HelpMessage='Enter the User Fullname, this full name used to create a mailbox search name as like an example, Rajkumar Ramasamy')]
    $UserFullname,
    [Parameter(Mandatory=$True, HelpMessage='Enter the path where you want to save the PST file. !NO TRAILING BACKSLASH! Example like D:\PSTfile ')]
    $LocalcpExptLocation,     #to set a default for this parameter.
    [Parameter(Mandatory=$True, HelpMessage='Enter mailbox user country location, where this location name create a new folder as a subfolder to save the PST file, Example like D:\PSTfile\Australia')]
    $MailboxUserCountry

)

# Choose a log folder path to capture this PS job script execution logs
$LogFolder = "C:\Users\Public\Documents\O365_PSscriptExecution"
# Remove existing log files from that logfolder
Get-ChildItem -Path $LogFolder -File | Remove-Item

$Date = (Get-Date).ToString("ddMMyyyHHmm")
$Logfilepath = $LogFolder +"\" +$Date +"_Mbx_exportjob.log"
md $LogFolder

# Create a new script execution log file on the target computer
New-Item -Path $Logfilepath -ItemType File
Write-host "Execution logs will capture in the below file"
Write-host "$Logfilepath"
Start-Transcript -Append $Logfilepath #>>Start capture the PS job script execution logs
Set-PSDebug -Trace 1

# Following are the parameter inputs that are used in this job script
Write-Output "# Following are the parameter inputs that are used in this job script"
Write-Output "Mailbox name $Mailbox"
Write-Output "Mailbox user $UserFullname"
Write-output "Mailbox export filepath $LocalcpExptLocation"
Write-Output "Mailbox user accessibility base country location: $MailboxUserCountry"

#Get the Logical drives that matches local diskdrive type with NTFS partition = '3"
Write-Host "#Get the Logical drives that matches local diskdrive type with NTFS partition = '3"
$CheckDriveType = (Get-WmiObject -Class Win32_logicaldisk -Filter "DriveType = 3").DeviceID

#Define the Folder path
$rootdrive = $LocalcpExptLocation.Split("\")
$rootdrive = $rootdrive[0]
#$Logfilepath = $rootdrive +"\O365Logfolder" 
$Exportfilepath = "$LocalcpExptLocation\$MailboxUserCountry"

#*******************************************************************************************************
#*******************************************************************************************************

#create file path, if not exist
if((!(Test-Path $Exportfilepath)-ieq $True) -or (!(Test-Path $LogFolder) -ieq $True))
{
    
    if($CheckDriveType -contains $rootdrive )
    {
        Write-Output "PST file export location drive has a NTSF partition on a local diskDrive $rootdrive"
        md $Exportfilepath
        md $Logfilepath
        Write-output "So, Local copy backup location folder created, path: $Exportfilepath"
    }
    else
    {
        Write-host "Local drive $rootdrive : not exists, please check & provide the correct file location & re-run this job script"
        exit
    }
}
else
{
    Write-Output "pst file path already exists $Exportfilepath"
    Write-Output "log file path already exits $LogFolder"
}


#*******************************************************************************************************
#*******************************************************************************************************
Write-Output "Create a Export job search name"
# Create a search name & Search Export name. You can change this to suit your preference
$SearchName = $userFullname +"_mailbox"
$ExportName = $SearchName +"_Export"
#$SearchName = "Jimmy Hoi_Mailbox"
#$ExportName = $SearchName +"_Export"
Write-Output $SearchName
Write-Output $ExportName

##******************************************************************************************************
##******************************************************************************************************
# To check for existing search & export jobs Function
Function Get-ExistingSearchJobResult()
{
    [CmdletBinding()]
    [OutputType([Int])]
    Param (
    [Parameter(Mandatory = $true)] [string] $SearchName,
    [Parameter(Mandatory = $true)] [string] $ExportName
    )
    #To look for search job state
    try{
        $Search_status =Get-ComplianceSearch -Identity $SearchName -ErrorAction Stop
        "JobFound"
        #$Search_status = 
    }
    Catch {
        "JobNotFound"
        #$Search_status = 
    }
    #To look for search action job state
    Try{
        $SearchAction_Status = Get-ComplianceSearchAction -Identity $ExportName -ErrorAction stop
        "ExportJobFound"
        #$SearchAction_Status = 
    }
     Catch{
        "ExportJobNotFound"
        #$SearchAction_Status = "
    }
}

##******************************************************************************************************
##******************************************************************************************************
# Job Function to start the search job

Function Start-MailboxSearchExport()
{
    <#
    #Connect Compliance Center 
    & Search for mailbox export job & 
    download the exported file to local storage
    #>
    Param (

        [Parameter(Mandatory = $true)] [string] $Mailbox,
        [Parameter(Mandatory = $true)] [string] $SearchName,
        [Parameter(Mandatory = $true)] [string] $ExportName
            
        )
    #Connect Compliance Center
    Write-Host "Connecting to Security & Compliance Center. Enter your admin credentials in the pop-up (pop-under?) window."
    Connect-IPPSSession -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" 
 
    ##Check if the the compliance search name exist
    $ExistingSearchJobResult = Get-ExistingSearchJobResult $SearchName $ExportName
    Switch($ExistingSearchJobResult[0],$ExistingSearchJobResult[1])
    {
        "JobFound" 
        {       
            Write-Warning "Already search job exist in the same name: $SearchName"
            Write-Warning "Kindly review that job in GUI & proceed for this export job"
            do{

                Write-host = "Enter your input Would you like to delete job in GUI or Rename the job to continue" 
                $SearchJob_UserInput = Read-Host -Prompt "TYPE D to quit this script or R to rename "
                $Confirmation = Read-Host -prompt "You have entered $SearchJob_UserInput ,please confirm (y/n)?"
            }

            While ($Confirmation -ne "y")

            If($SearchJob_UserInput -ieq "D")
            {
                exit
            }
            elseif($SearchJob_UserInput -ieq "R")
            {
                $SearchName = $SeachName +"_02"
            }
                                   <#
            ##Wait for user input a). to delete existing search_job or b). rename the search_job
            If(User input is = a)
            ask user to delete this search_job name and confirm
            and continue
            if(user input is = b)
            rename that search_job value and assign to variable
            and continue
            #>
            Start-Sleep -Seconds 10
            $SearchName = "Renamed_SeachName"
            #exit
            #Break
            
        
        }
        "ExportJobFound"
        {
            Write-Warning "Already export job exist in the same name: $ExportName"
            Write-Warning "Kindly review & delete this Export job in GUI & re-run"
            exit
            <#
            ##Wait for user input a). to delete existing Export_job or b). rename the Export_job
            If(User input is = a)
            ask user to delete this Export_job name and confirm
            and continue
            if(user input is = b)
            rename that Export_job value and assign to variable
            and continue
            #>
            #>
                        
            #Break
        }
    }
    $SearchName
    Write-Host "Creating compliance search..."
    #New-ComplianceSearch -Name "$SearchName" -ExchangeLocation "$Mailbox" -Description "EMail ID: $Mailbox" -AllowNotFoundExchangeLocationsEnabled $true #Create a content search, including the the entire contents of the user's email and onedrive. If you didn't provide a OneDrive URL, or it wasn't valid, it will be ignored.
    $Newsearch = New-ComplianceSearch -Name "$SearchName" -ExchangeLocation "$Mailbox" -Description "EMail ID: $Mailbox" -AllowNotFoundExchangeLocationsEnabled $true #Create a content search, including the the entire contents of the user's email and onedrive. If you didn't provide a OneDrive URL, or it wasn't valid, it will be ignored.
    Write-Host "Starting compliance search..."
    Start-ComplianceSearch -Identity $SearchName #Start the search created above
    Write-Host "Waiting for compliance search to complete..."
    for ($SearchStatus;$SearchStatus -notlike "Completed";)
    { #Wait then check if the search is complete, loop until complete
        Start-Sleep -s 10
        $SearchStatus = Get-ComplianceSearch $SearchName | Select-Object -ExpandProperty Status #Get the status of the search
        Write-Host -NoNewline "." # Show some sort of status change in the terminal
    }
    Write-Host "Compliance search is complete!"
    #"Creating export from the search..."
    Write-Host "Creating exportjob from the search..."
    $SearchAction = New-ComplianceSearchAction -SearchName $SearchName -Export -Format FxStream -ExchangeArchiveFormat PerUserPst -Scope BothIndexedAndUnindexedItems -EnableDedupe $true -SharePointArchiveFormat IndividualMessage -IncludeSharePointDocumentVersions $true 
    Start-Sleep -s 5 # wait 5 seconds to give microsoft's side time to create the SearchAction before the next commands try to run against it.

    # Find the Unified Export Tool's location and create a variable for it
    $ExportExe = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)

    # Gather the URL and Token from the export in order to start the download
    # We only need the ContainerURL and SAS Token at a minimum but we're also pulling others to help with tracking the status of the export.
    $ExportName = $SearchName +"_Export"
    $ExportDetails = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details # Get details for the export action
    For ($MonitorSearchAction;$MonitorSearchAction -notlike "Completed")
    {
        #Wait then check if the search is complete, loop until complete
        Start-Sleep -s 10
        $MonitorSearchAction = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details | Select-Object -ExpandProperty status  # Get details for the export action
        Write-Host -NoNewline "." # Show some sort of status change in the terminal
    }    
    $ExportDetails = $ExportDetails.Results.split(";")
    $ExportContainerUrl = $ExportDetails[0].trimStart("Container url: ")
    $ExportSasToken = $ExportDetails[1].trimStart(" SAS token: ")
    $ExportEstSize = ($ExportDetails[18].TrimStart(" Total estimated bytes: ") -as [double])
    $ExportTransferred = ($ExportDetails[20].TrimStart(" Total transferred bytes: ") -as [double])
    $ExportProgress = $ExportDetails[22].TrimStart(" Progress: ").TrimEnd("%")
    $ExportStatus = $ExportDetails[25].TrimStart(" Export status: ")
    
    # Download the exported files from Office 365
    Write-Host "Initiating download"
    Write-Host "Saving export to: " + $Exportfilepath
    $Arguments = "-name ""$SearchName""","-source ""$ExportContainerUrl""","-key ""$ExportSasToken""","-dest ""$Exportfilepath""","-trace true"
    Start-Process -FilePath "$ExportExe" -ArgumentList $Arguments

    while(Get-Process microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue -Verbose){
        $Downloaded = Get-ChildItem $Exportfilepath +"\$ExportName\" -Recurse | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum
        Write-Progress -Id 1 -Activity "Export in Progress" -Status "Complete..." -PercentComplete $ExportProgress
        if ("Completed" -notlike $ExportStatus){Write-Progress -Id 2 -Activity "Download in Progress" -Status "Estimated Complete..." -PercentComplete ($Downloaded/$ExportEstSize*100) -CurrentOperation "$Downloaded/$ExportEstSize bytes downloaded."}
        else {Write-Progress -Id 2 -Activity "Download in Progress" -Status "Complete..." -PercentComplete ($Downloaded/$ExportEstSize*100) -CurrentOperation "$Downloaded/$ExportTransferred bytes downloaded."}
        Start-Sleep 60
        $ExportDetails = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details # Get details for the export action
        $ExportDetails = $ExportDetails.Results.split(";")
        $ExportEstSize = ($ExportDetails[18].TrimStart(" Total estimated bytes: ") -as [double])
        $ExportTransferred = ($ExportDetails[20].TrimStart(" Total transferred bytes: ") -as [double])
        $ExportProgress = $ExportDetails[22].TrimStart(" Progress: ").TrimEnd("%")
        $ExportStatus = $ExportDetails[25].TrimStart(" Export status: ")
        Write-Host -NoNewline " ."
    }
    Write-Host "Download Complete!"
    Start-Sleep -s 60
     
}

#Connect Exchange online admin center
Write-Host "Connecting to Exchange Online. Enter your admin credentials in the pop-up (pop-under?) window."
$Mailbox = "Jimmy.Hoi@eginnovations.com"
#"Check the Mailbox available in O365 tenant"
Function Get-MailboxAvailability
{
    Param (
            [Parameter(Mandatory = $true)] [string] $Mailbox
            
        )
    <#
    This function to check the usermailbox availability      
    #>
    #Write-host "Check the Mailbox available in O365 tenant"
    Connect-ExchangeOnline
    if(($MbxState = Get-EXOMailbox -Identity $Mailbox -ErrorAction SilentlyContinue))
    {
        $MbxState = "Present"
        $MbxState
    }
        else
    {
        $MbxState = "NotPresent"
        $MbxState
        #exit
    }
    Disconnect-ExchangeOnline -Confirm:$false
}
$MbxaccessibilityState = Get-MailboxAvailability $Mailbox

if($MbxaccessibilityState -contains "Present" )
{
    Start-MailboxSearchExport $Mailbox $SearchName
    Write-output "Mailbox Yes $Mailbox"
    
}
elseif($MbxaccessibilityState -contains "NotPresent")
{
    Write-Output "No Mailbox $Mailbox"
    exit
    #"No Mailbox"

}
#Connect Compliance Center & Search for mailbox export job & download the exported file to local storage

Start-Sleep 10
## Copy Downloaded files to NAS drive folder location
Function Copy-TonasFolder
{
    Write-host "Stop the O365 mailbox export tool process, to move the Locally downloaded backup copy to NAS drive"
    Stop-Process -Name microsoft.office.client.discovery.unifiedexporttool -Confirm:$false
    Start-Sleep -s 20
    $NASdrive = "D:"
    $NASdrivefolder = "$NASdrive\$MailboxUserCountry"
    $CopyLogfilepath = $LogFolder +"\" + $date +"Copyprocess.txt"
    If(!(Test-Path $NASdrivefolder)-ieq $True )
    {
        md $NASdrivefolder
        Write-host 'NAS drive folder path created as $NASdrivefolder'
        Copy-Item -Path $Exportfilepath -Destination $NASdrivefolder -Recurse | get-process | Out-File -FilePath "$CopyLogfilepath"
    }
    elseif ((Test-Path $NASdrivefolder)-ieq $True)
    {
        Write-host 'NAS drive folder path exist $NASdrivefolder'
        Copy-Item -Path $Exportfilepath -Destination $NASdrivefolder -Recurse | get-process | Out-File -FilePath "$CopyLogfilepath"
    }

}

Copy-TonasFolder $MailboxUserCountry

Start-Sleep -Seconds 15
Write-Output "Kill process $PID"
Stop-Process $PID