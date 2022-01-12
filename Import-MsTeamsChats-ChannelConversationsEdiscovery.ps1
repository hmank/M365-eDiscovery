#############################################################################
#  This script is used to read from a csv containing Microsft user mailboxes (Teams Chats)
#  Teams Mailboxes (Channel conversations)
#  The script will import into an existing eDiscovery Case.
#
# October 13, 2021
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

Start-Transcript -LiteralPath "$logsdir\ImporteDiscoveryTeamsMailboxes-$month-$day-$hour-$minute.log"  -NoClobber -Append

# Connect to SCC PowerShell using modern authentication
if (!$SccSession) {
    Import-Module ExchangeOnlineManagement
    Connect-IPPSSession
}


write-host "***********************************************"
write-host "   Security & Compliance Center PowerShell  " -foregroundColor yellow -backgroundcolor darkgreen
write-host "   Add Teams user chats or Teams Channel conversation mailboxes to an eDiscovery Case   " -foregroundColor yellow -backgroundcolor darkgreen 
write-host "***********************************************"
" "
do {
    write-host "***********************************************"
    write-host "   Please select below the option you wish to add to the eDiscovery Case " -foregroundColor yellow -backgroundcolor darkgreen
    write-host "   [1] - Teams User chats (User mailboxes)" -foregroundColor yellow -backgroundcolor darkgreen 
    write-host "   [2] - Teams Channel conversations (Teams Mailboxes)" -foregroundColor yellow -backgroundcolor darkgreen 

    $userInput = Read-Host "Select Option & Press Enter: " 
    write-host "***********************************************"

    switch ($userInput) {
    
        '1' { "[1]- Teams User chats (User mailboxes)" }
        '2' { "[2]- Teams Channel conversations (Teams Mailboxes)" }

    }#switch
    
}
While (($userInput -ne '1') -and ($userInput -ne '2') )

"" 

# Get other required information
do {
    Get-ComplianceCase -CaseType AdvancedEdiscovery | Select-Object Name, Casetype ; Get-ComplianceCase | Select-Object Name, Casetype
    $casename = $(Write-Host "Enter the name of the existing case: " -foregroundColor yellow -backgroundcolor darkgreen -NoNewline; Read-Host)
    $caseexists = (get-compliancecase -identity "$casename" -erroraction SilentlyContinue).isvalid
    if ($caseexists -ne 'True') {
        ""
        write-host "A case named '$casename' doesn't exist. Please specify the name of an existing case, or create a new case and then re-run the script." -foregroundColor Yellow
        ""
    }
}
While ($caseexists -ne 'True')
""
write-host "***********************************************"

do {
    Get-caseholdpolicy * | Select Name

    $holdName = $(Write-host "Enter the name of a new hold: " -foregroundColor yellow -backgroundcolor darkgreen -NoNewline; Read-Host)
    $holdexists = (get-caseholdpolicy -identity "$holdname" -case "$casename" -erroraction SilentlyContinue).isvalid
    if ($holdexists -eq 'True') {
        ""
        write-host "A hold named '$holdname' already exists. Please specify a new hold name." -foregroundColor Yellow
        ""
    }
}While ($holdexists -eq 'True')
""

Try {

   
    If ($userInput -eq 1) {
        #Import User Teams chat (User mailboxes)
        do {
            ""
            $UserMBXinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline; read-host)
            ""
            $fileexists = test-path -path $UserMBXinputfile
            if ($fileexists -ne 'True') { write-host "$UserMBXinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
        }while ($fileexists -ne 'True')
            
             
                               
        [Array]  $importUserMBXArray = @(Import-Csv $UserMBXinputfile).PrimarySmtpAddress
              
        $importUserMBX = $null                  
        $importUserMBXArray = $importUserMBXArray | Where-Object { $_ }
        $importUserMBX = $importUserMBXArray[0..999] -join ","
        $importUserMBX = $importUserMBXArray -split ' *, *'
                                        
        New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $importUserMBX -Enabled $True -verbose -Force
        New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "(c:c)(ItemClass=IPM.Note.Microsoft.Conversation)(ItemClass=IPM.Note.Microsoft.Missed)(ItemClass=IPM.Note.Microsoft.Conversation.Voice)(ItemClass=IPM.Note.Microsoft.Missed.Voice)(ItemClass=IPM.SkypeTeams.Message)" -Verbose


    } #if

    If ($userInput -eq 2) {
        #Import Microsoft Teams Mailboxes.
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
