#############################################################################
#  This script is used to read from a csv containing Microsoft 365 groups,
#  Shared mailboxes, SharePoint and OneDrive Sites.
#  The script will import into an existing eDiscovery Case. 
#
# August 20, 2021
#
# Version 1.0
# Author: Habib Mankal
# 
# ##############################################################################

$day = (get-date).day
$month = (get-date).month
$hour = (get-date).hour
$minute = (get-date).minute
$logsdir = "C:\temp"

Start-Transcript -LiteralPath "$logsdir\ImporteDiscoverySites-Unfied-Shared-Mailboxes-$month-$day-$hour-$minute.log"  -NoClobber -Append

# Connect to SCC PowerShell using modern authentication
if (!$SccSession) {
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession
}

$dataImportTypes = @(
    "1- Sharepoint Sites",
    "2- Microsoft 365 Groups",
    "3- Shared Mailboxes",
    "4- User Mailboxes and OneDrive Sites",
    "5- Teams User chats (User mailboxes)",
    "6- Teams Channel conversations (Teams Mailboxes)"
)

$TypeOfeDisovery = @(

    "Core eDiscovery",
    "Advanced eDiscovery"
)

$DataImportSelection = $null
$eDisoveryTypeSelection = $null
$DataImportSelection = ($dataImportTypes | Out-GridView -OutputMode Single);
$eDisoveryTypeSelection = ($TypeOfeDisovery | Out-GridView -OutputMode Single)
    
If ($eDisoveryTypeSelection -eq "Core eDiscovery") {
    $CaseName = (Get-ComplianceCase | Select-Object Name | Select-Object Name | Out-GridView -OutputMode Single).Name;
    $ExistingHold = (Get-caseholdpolicy * | Select-Object Name | Out-GridView -OutputMode Single).Name
    Start-Sleep -Seconds 5
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $holdname = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter a uinque name for a new Legal Hold")
}

else {
    $CaseName = (Get-ComplianceCase -CaseType AdvancedEdiscovery | Select-Object Name | Out-GridView -OutputMode Single).Name;
    $ExistingHold = (Get-caseholdpolicy * | Select-Object Name | Out-GridView -OutputMode Single).Name;
    Start-Sleep -Seconds 5
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $holdname = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter a name for a new Legal Hold")

}

     


<# Get-caseholdpolicy * | Select Name -ExpandProperty Name
    Start-Sleep -Seconds 6
    $holdexists = (get-caseholdpolicy -identity "$holdname" -case "$casename" -erroraction SilentlyContinue).isvalid

    $holdName = $(Write-host "Enter the name of a new hold: " -foregroundColor yellow -backgroundcolor darkgreen -NoNewline; Read-Host)
    if ($holdexists -eq 'True') {
        ""
        write-host "A hold named '$holdname' already exists. Please specify a new hold name." -foregroundColor Yellow
        ""
    }
    #>
""

Try {

    If ($DataImportSelection -eq "1- Sharepoint Sites") {
        #Import Sharepoint Sites.
        do {
            ""
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
            $spsinputfile = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter the name full path of the csv file that contains the Sharepoint Sites to place on hold. eg c:\Temp\AllSpSites.csv: ")

            ""
            $fileexists = test-path -path $spsinputfile
            if ($fileexists -ne 'True') { write-host "$spsinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
             
        $spsurls = $null                  
        [Array] $spsurlsarray = (Import-Csv $spsinputfile).url

        $Spsurlsarray = $spsurlsarray | Where-Object { $_ }
        $spsurls = $spsurlsarray[0..99] -join ","
        $spsurls = $spsurls -split ' *, *'

        
        Write-Host "Please wait as the $holdname legal hold is being created" -ForegroundColor Yellow
        New-CaseHoldPolicy -Name "$holdname" -Case "$casename" -SharePointLocation $spsurls -Enabled $True 
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose | Statusbar
        
        Write-Host "The $holdname legal hold has been successfully created!" -ForegroundColor Green
        

    } #if


    If ($DataImportSelection -eq "2- Microsoft 365 Groups") {
        #Import Microsoft 365 Groups.
        do {
            ""
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
            $m365groupinputfile = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :")

            ""
            $fileexists = test-path -path $m365groupinputfile
            if ($fileexists -ne 'True') { write-host "$m365groupinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
        [Array]  $M365GroupsArray = @(Import-Csv $m365groupinputfile).ExternalDirectoryObjectId
              
        $M365Groups = $null                  
        $M365GroupsArray = $M365GroupsArray | Where-Object { $_ }
        $M365Groups = $M365Groupsarray[0..999] -join ","
        $M365Groups = $M365Groups -split ' *, *'

        Write-Host "Please wait as the $holdname legal hold is being created" -ForegroundColor Yellow                                          
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $M365Groups -Enabled $True -Verbose
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose
        Write-Host "The $holdname legal hold has been successfully created!" -ForegroundColor Green


    } #if


    If ($DataImportSelection -eq "3- Shared Mailboxes") {
        #Import Shared Mailboxes.
        do {
            ""
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
            $SharedMBXinputfile = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :")

            ""
            $fileexists = test-path -path $SharedMBXinputfile
            if ($fileexists -ne 'True') { write-host "$SharedMBXinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
        [Array]  $SharedMBXArray = @(Import-Csv $SharedMBXinputfile).ExternalDirectoryObjectId
              
        $SharedMBX = $null                  
        $SharedMBXArray = $SharedMBXArray | Where-Object { $_ }
        $SharedMBX = $SharedMBXArray[0..999] -join ","
        $SharedMBX = $SharedMBXArray -split ' *, *'
                                              
        Write-Host "Please wait as the $holdname legal hold is being created" -ForegroundColor Yellow
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $SharedMBX -Enabled $True -Verbose
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose
        Write-Host "The $holdname legal hold has been successfully created!" -ForegroundColor Green

    } #if

    If ($DataImportSelection -eq "4- User Mailboxes and OneDrive Sites") {
        #Import User Mailboxes and Onedrive sites
        do {
            ""
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
            $UserMBXODinputfile = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter the name full path of the csv file that contains the users mailbox and onedrive to place on hold. eg c:\Temp\AllM365Groups.csv :")

            ""
            $fileexists = test-path -path $UserMBXODinputfile
            if ($fileexists -ne 'True') { write-host "$UserMBXODinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
        [Array]  $importUserMBXArray = @(Import-Csv $UserMBXODinputfile).PrimarySmtpAddress
              
        $importUserMBX = $null                  
        $importUserMBXArray = $importUserMBXArray | Where-Object { $_ }
        $importUserMBX = $importUserMBXArray[0..999] -join ","
        $importUserMBX = $importUserMBXArray -split ' *, *'

        [Array]  $importUserODArray = @(Import-Csv $UserMBXODinputfile).PersonalURL
              
        $importUserOD = $null                  
        $importUserODArray = $importUserODArray | Where-Object { $_ }
        $importUserOD = $importUserODArray[0..999] -join ","
        $importUserOD = $importUserODArray -split ' *, *'

        Write-Host "Please wait as the $holdname legal hold is being created" -ForegroundColor Yellow
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $importUserMBX -SharePointLocation $importUserOD -Enabled $True -verbose -Force
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose
        Write-Host "The $holdname legal hold has been successfully created!" -ForegroundColor Green

    } #if

    If ($userIDataImportSelectionnput -eq "5- Teams User chats (User mailboxes)") {
        #Import User Teams chat (User mailboxes)
        do {
            ""    
            [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
            $UserMBXinputfile = [string][Microsoft.VisualBasic.Interaction]::InputBox("Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :")

            ""
            $fileexists = test-path -path $UserMBXinputfile
            if ($fileexists -ne 'True') { write-host "$UserMBXinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')

        [Array]  $importUserMBXArray = @(Import-Csv $UserMBXinputfile).PrimarySmtpAddress
              
        $importUserMBX = $null                  
        $importUserMBXArray = $importUserMBXArray | Where-Object { $_ }
        $importUserMBX = $importUserMBXArray[0..999] -join ","
        $importUserMBX = $importUserMBXArray -split ' *, *'
        
        Write-Host "Please wait as the $holdname legal hold is being created" -ForegroundColor Yellow
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $importUserMBX -Enabled $True -verbose -Force
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "(c:c)(ItemClass=IPM.Note.Microsoft.Conversation)(ItemClass=IPM.Note.Microsoft.Missed)(ItemClass=IPM.Note.Microsoft.Conversation.Voice)(ItemClass=IPM.Note.Microsoft.Missed.Voice)(ItemClass=IPM.SkypeTeams.Message)" -Verbose
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose | Statusbar

    } #if

    If ($DataImportSelection -eq "6- Teams Channel conversations (Teams Mailboxes)") {
        #Import Microsoft 365 Groups.
        do {
            ""
            $msTeamsMbxinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline; read-host)
            ""
            $fileexists = test-path -path $m365groupinputfile
            if ($fileexists -ne 'True') { write-host "$m365groupinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
             
                               
        [Array]  $msTeamsMbxArray = @(Import-Csv $msTeamsMbxinputfile).ExternalDirectoryObjectId
              
        $msTeamsMbx = $null                  
        $msTeamsMbxArray = $msTeamsMbxArray | Where-Object { $_ }
        $msTeamsMbx = $MmsTeamsMbxsarray[0..999] -join ","
        $msTeamsMbx = $msTeamsMbx -split ' *, *'
                                              
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $msTeamsMbx -Enabled $True -Verbose
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose

    } #if



} #try

Catch {
    write-host -f Red "`tError:" $_.Exception.Message
} #catch

Stop-Transcript
