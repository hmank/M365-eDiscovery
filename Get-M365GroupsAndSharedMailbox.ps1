#############################################################################
#  This script is used to read all Microsoft 365 group and Shared mailboxes
#  and export them to a csv file to be used in a script to add
#  the mailboxes to an eDiscovery Case.
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
$exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);


Start-Transcript -LiteralPath "$logsdir\GetUnfiedAndSharedMailbox-$month-$day-$hour-$minute.log"  -NoClobber -Append


Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline



Try {
 
     
            
     Get-UnifiedGroup | Select DisplayName, PrimarySmtpAddress,Identity,ExternalDirectoryObjectId | Export-Csv "$exportCSVPath\M365GroupsMailboxes-$month-$day.csv" -Append -NoTypeInformation
     Get-UnifiedGroup | Select DisplayName, PrimarySmtpAddress,Identity ,ExternalDirectoryObjectId
     Get-Mailbox -RecipientTypeDetails sharedmailbox | Select DisplayName, PrimarySmtpAddress,Identity,ExternalDirectoryObjectId | Export-Csv "$exportCSVPath\SharedMailboxes-$month-$day.csv" -Append -NoTypeInformation
     Get-Mailbox -RecipientTypeDetails sharedmailbox | Select DisplayName, PrimarySmtpAddress,Identity,ExternalDirectoryObjectId
        }
        Catch {
            write-host -f Red "`tError:" $_.Exception.Message
        }
 


Stop-Transcript
