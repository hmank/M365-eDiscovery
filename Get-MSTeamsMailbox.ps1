#############################################################################
#  This script is used to read all Microsoft Teams
#  and export them to a csv file to be used in a script to add
#  the mailboxes to an eDiscovery Case.
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
$exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);


Start-Transcript -LiteralPath "$logsdir\GetTeamsMailboxes-$month-$day-$hour-$minute.log"  -NoClobber -Append


Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline



Try {
 
     
            
     Get-UnifiedGroup | Where-Object {$_.ResourceProvisioningOptions -eq "Team"} | Select DisplayName, PrimarySmtpAddress,Identity,ExternalDirectoryObjectId | Export-Csv "$exportCSVPath\MSTeamsMailboxes-$month-$day.csv" -Append -NoTypeInformation
     Get-UnifiedGroup | Where-Object {$_.ResourceProvisioningOptions -eq "Team"}|  Select DisplayName, PrimarySmtpAddress,Identity ,ExternalDirectoryObjectId
        }
        Catch {
            write-host -f Red "`tError:" $_.Exception.Message
        }
 


Stop-Transcript
