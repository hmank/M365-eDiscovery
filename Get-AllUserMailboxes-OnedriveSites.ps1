#############################################################################
#  This script is used to read export all users Primary SMTP Email addresses
#  and Onedrive personal sites and export them to a csv file.
#  
#  
#  Sept 8, 2021
#
# Version 1.0
# Author: Habib Mankal
# 
################################################################################


Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

#Get the user primary SMTP Email Addresses
$UserMBx = Get-Mailbox -ResultSize unlimited -Filter { RecipientTypeDetails -eq 'UserMailbox'} | Select-Object PrimarySmtpAddress
$UserMBxEmail = $UserMBx.PrimarySmtpAddress 

# Get the organization's domain name. We use this to create the SharePoint admin URL and root URL for OneDrive for Business.
""
$mySiteDomain = Read-Host "Enter the domain name for your SharePoint organization. We use this name to connect to SharePoint admin center and for the OneDrive URLs in your organization. For example, 'contoso' in 'https://contoso-admin.sharepoint.com' and 'https://contoso-my.sharepoint.com'"
""

# Connect to PnP Online using modern authentication
Import-Module PnP.PowerShell
Connect-PnPOnline -Url https://$mySiteDomain-admin.sharepoint.com -UseWebLogin

# Load the SharePoint assemblies from the SharePoint Online Management Shell
# To install, go to https://go.microsoft.com/fwlink/p/?LinkId=255251
if (!$SharePointClient -or !$SPRuntime -or !$SPUserProfile)
{
    $SharePointClient = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    $SPRuntime = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    $SPUserProfile = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.UserProfiles")
    if (!$SharePointClient)
    {
        Write-Error "The SharePoint Online Management Shell isn't installed. Please install it from: https://go.microsoft.com/fwlink/p/?LinkId=255251 and then re-run this script."
        return;
    }
}
#Import the list of addresses from the txt file.  Trim any excess spaces and make sure all addresses 
    #in the list are unique.
  [array]$emailAddresses = $UserMBxEmail
  [int]$dupl = $emailAddresses.count
  [array]$emailAddresses = $emailAddresses | select-object -unique
  $dupl -= $emailAddresses.count
#Validate email addresses so the hold creation does not run in to an error.
if($emailaddresses.count -gt 0){
write-host ($emailAddresses).count "addresses were found. There were $dupl duplicate entries in the file." -foregroundColor Yellow
""
            Write-host "Validating the email addresses. Please wait..." -foregroundColor Yellow
            ""
            $finallist =@()
            foreach($emailAddress in $emailAddresses)
            {
            if((get-recipient $emailaddress -erroraction SilentlyContinue).isvalid -eq 'True')
            {$finallist += $emailaddress}
            else {"Unable to find the user $emailaddress"
            [array]$excludedlist += $emailaddress}
            }
            ""
            #Find user's OneDrive account URL using email address
            Write-Host "Getting the URL for each user's OneDrive for Business site." -foregroundColor Yellow 
            ""
            $AdminUrl = "https://$mySiteDomain-admin.sharepoint.com"
            $mySiteUrlRoot = "https://$mySiteDomain-my.sharepoint.com"
            $urls = @()
            foreach($emailAddress in $finallist)
            {
            try
            {
            $url=Get-PnPUserProfileProperty -Account $emailAddress | Select PersonalUrl
            $urls += $url.PersonalUrl
                   Write-Host "- $emailAddress => $url"
                   [array]$ODadded += $url.PersonalUrl

                   $UserMBxODObject = @{
                   PersonalUrl = $url.PersonalUrl
                   PrimarySmtpAddress = $emailAddress
                   }
                   $UserMBxOD = New-Object PSObject -Property $UserMBxODObject    
                   $UserMBxOD | Export-csv  c:\temp\allusers-mbx-od.csv -Append -NoTypeInformation
      
       }catch { 
 Write-Warning "Could not locate OneDrive for $emailAddress"
 [array]$ODExluded += $emailAddress
 Continue }
}
}
