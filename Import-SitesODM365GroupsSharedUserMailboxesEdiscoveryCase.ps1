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
if (!$SccSession)
{
  Import-Module ExchangeOnlineManagement
  Connect-IPPSSession
}


write-host "***********************************************"
write-host "   Security & Compliance Center PowerShell  " -foregroundColor yellow -backgroundcolor black
write-host "   Add Sharepoint or OneDrive Sites, User, Microsoft 365 Groups,Shared Mailboxes to eDiscovery Case   " -foregroundColor yellow -backgroundcolor black 
write-host "***********************************************"
" "
do{
write-host "***********************************************"
write-host "   Please select below the option you wish to add to the eDiscovery Case " -foregroundColor yellow -backgroundcolor black
write-host "   [1] - Sharepoint Sites   " -foregroundColor yellow -backgroundcolor black
write-host "   [2] - Microsoft 365 Groups   " -foregroundColor yellow -backgroundcolor black
write-host "   [3] - Shared Mailboxes  " -foregroundColor yellow -backgroundcolor black 
write-host "   [4] - User Mailboxes" -foregroundColor yellow -backgroundcolor black
write-host "   [5] - OneDrive Sites  " -foregroundColor yellow -backgroundcolor black 
write-host "   [6] - Teams User chats (User mailboxes)" -foregroundColor yellow -backgroundcolor black 
write-host "   [7] - Teams Channel conversations (Teams Mailboxes)" -foregroundColor yellow -backgroundcolor black 

$userInput = Read-Host "Select Option & Press Enter: " 
write-host "***********************************************"

switch ($userInput){
    
    '1' {"[1]- Sharepoint Sites Selected"} 
    '2' {"[2]- Microsoft365 Groups Selected"}
    '3' {"[3]- Shared Mailboxes Selected"}
    '4' {"[4]- User Mailboxes"}
    '5' {"[5]- OneDrive Sites"}
    '6' {"[6]- Teams User chats (User mailboxes)" }
    '7' {"[7]- Teams Channel conversations (Teams Mailboxes)" }

    }#switch
    
}
While (($userInput -ne '1') -and ($userInput -ne '2') -and ($userInput -ne '3') -and ($userInput -ne '4') -and ($userInput -ne '5') -and ($userInput -ne '6')  )

"" 

# Get other required information
do{

$casename = $(Write-Host "Enter the name of the existing case: " -foregroundColor yellow -backgroundcolor black -NoNewline; Read-Host)
$caseexists = (get-compliancecase -identity "$casename" -erroraction SilentlyContinue).isvalid
if($caseexists -ne 'True')
{""
write-host "A case named '$casename' doesn't exist. Please specify the name of an existing case, or create a new case and then re-run the script." -foregroundColor Yellow
""}
}
While($caseexists -ne 'True')
""
write-host "***********************************************"

do{

$holdName = $(Write-host "Enter a unique name of a new case hold: " -foregroundColor yellow -backgroundcolor black -NoNewline; Read-Host)
$holdexists=(get-caseholdpolicy -identity "$holdname" -case "$casename" -erroraction SilentlyContinue).isvalid
if($holdexists -eq 'True')
{""
write-host "A hold named '$holdname' already exists. Please specify a new hold name." -foregroundColor Yellow
""}
}While($holdexists -eq 'True')
""

Try{

 If ($userInput -eq 1){   #Import Sharepoint Sites.
        do{
        ""
        $spsinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Sharepoint Sites to place on hold. eg c:\Temp\AllSpSites.csv: " -NoNewline;read-host)
        ""
        $fileexists = test-path -path $spsinputfile
             if($fileexists -ne 'True'){write-host "$spsinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
             $spsurls = $null                  
            [Array] $spsurlsarray = (Import-Csv $spsinputfile).url

             $Spsurlsarray = $spsurlsarray | Where-Object {$_}
             $spsurls = $spsurlsarray[0..99] -join ","
             $spsurls = $spsurls -split ' *, *'


                New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -SharePointLocation $spsurls -Enabled $True -Verbose
                New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose


        } #if


        If ($userInput -eq 2){   #Import Microsoft 365 Groups.
        do{
        ""
        $m365groupinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline;read-host)
        ""
        $fileexists = test-path -path $m365groupinputfile
             if($fileexists -ne 'True'){write-host "$m365groupinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
                               
             [Array]  $M365GroupsArray = @(Import-Csv $m365groupinputfile).ExternalDirectoryObjectId
              
                $M365Groups = $null                  
                $M365GroupsArray = $M365GroupsArray | Where-Object {$_}
                $M365Groups = $M365Groupsarray[0..999] -join ","
                $M365Groups = $M365Groups -split ' *, *'
                                              
              New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $M365Groups -Enabled $True -Verbose
              New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose

        } #if


        If ($userInput -eq 3){   #Import Shared Mailboxes.
        do{
        ""
        $SharedMBXinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline;read-host)
        ""
        $fileexists = test-path -path $SharedMBXinputfile
             if($fileexists -ne 'True'){write-host "$SharedMBXinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
                               
             [Array]  $SharedMBXArray = @(Import-Csv $SharedMBXinputfile).ExternalDirectoryObjectId
              
                $SharedMBX = $null                  
                $SharedMBXArray = $SharedMBXArray | Where-Object {$_}
                $SharedMBX = $SharedMBXArray[0..999] -join ","
                $SharedMBX = $SharedMBXArray -split ' *, *'
                                              
              New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $SharedMBX -Enabled $True -Verbose
              New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose

        } #if

                If ($userInput -eq 4){   #Import User Mailboxes
        do{
        ""
        $UserMBXODinputfile = $(Write-Host "Enter the name full path of the csv file that contains the users mailboxes to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline;read-host)
        ""
        $fileexists = test-path -path $UserMBXODinputfile
             if($fileexists -ne 'True'){write-host "$UserMBXODinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
                               
                [Array]  $importUserMBXArray = @(Import-Csv $UserMBXODinputfile).PrimarySmtpAddress
              
                $importUserMBX = $null                  
                $importUserMBXArray = $importUserMBXArray | Where-Object {$_}
                $importUserMBX = $importUserMBXArray[0..999] -join ","
                $importUserMBX = $importUserMBXArray -split ' *, *'

                                                              
              New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -ExchangeLocation $importUserMBX -Enabled $True -verbose -Force
              New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose



        } #if

     If ($userInput -eq 5) {
          #Import User Onedrive Sites
          do {
               ""
               $UserMBXODinputfile = $(Write-Host "Enter the name full path of the csv file that contains the users OneDrive sites to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline; read-host)
               ""
               $fileexists = test-path -path $UserMBXODinputfile
               if ($fileexists -ne 'True') { write-host "$UserMBXODinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow }
          }while ($fileexists -ne 'True')
            
         
          [Array]  $importUserODArray = @(Import-Csv $UserMBXODinputfile).PersonalURL
              
          $importUserOD = $null                  
          $importUserODArray = $importUserODArray | Where-Object { $_ }
          $importUserOD = $importUserODArray[0..999] -join ","
          $importUserOD = $importUserODArray -split ' *, *'

                                              
          New-CaseHoldPolicy -Name "$holdName" -Case "$casename" -SharePointLocation $importUserOD -Enabled $True -verbose -Force
          New-CaseHoldRule -Name "$holdName" -Policy "$holdName" -ContentMatchQuery "" -Verbose



     } #if





     If ($userInput -eq 6) {
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

     If ($userInput -eq 7) {
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
