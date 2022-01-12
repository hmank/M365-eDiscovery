#############################################################################
#  This script is used to read all SharePoint sites in a tenant
#  and export them to a csv file to be used in a script to add
#  the sites to an eDiscovery Case.
#
# August 20, 2021
#
# Version 1.0
# Author: Habib Mankal
# CSV Column Headers
# SiteUrl
# ##############################################################################

$day = (get-date).day
$month = (get-date).month
$hour = (get-date).hour
$minute = (get-date).minute
$logsdir = "C:\temp"
$exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);

Start-Transcript -LiteralPath "$logsdir\GetSpSites-$month-$day-$hour-$minute.log"  -NoClobber -Append

$mySiteDomain = Read-Host "Enter the domain name for your SharePoint organization. We use this name to connect to SharePoint admin center`
, ONLY enter the DOMAIN name not the full URL. For example, if your admin
 URL is 'https://yyyzz-admin.sharepoint.com' ENTER only yyyzzz"

Import-Module PnP.PowerShell
Connect-PnPOnline -Url https://$mySiteDomain-admin.sharepoint.com -UseWebLogin


Try {
 
    #Get All Site collections  
    $SitesCollection = Get-PnPTenantSite
 
    #Loop through each site collection
    ForEach ($Site in $SitesCollection) {  
        Write-host -F Green $Site.Url 
        Get-PnPTenantSite -Identity $Site| Select-Object URL | Export-csv "$exportCSVPath\AllSharePointSites-$month-$day.csv" -Append -NoTypeInformation
        
        Try {
            
            #Get Site Collection subsites
            $SubSites = Get-PnPSubWebs -Recurse
            ForEach ($web in $SubSites) {
                Write-host `t $Web.URL 
                                Get-PnPSubWebs -Identity $web | Export-csv "$exportCSVPath\AllSharePointSubSites-$month-$day.csv" -Append -NoTypeInformation
            }
        }
        Catch {
            write-host -f Red "`tError:" $_.Exception.Message
        }
    }

    Get-Content "$exportCSVPath\AllSharePointSites-$month-$day.csv"  | Where-Object { $_ -notmatch "sharepoint.com/search" } | Out-File "$exportCSVPath\AllSharePointSites-$month-$day-1.csv" -Force
    Get-Content "$exportCSVPath\AllSharePointSites-$month-$day-1.csv"  | Where-Object { $_ -notmatch "my.sharepoint.com/" } | Out-File "$exportCSVPath\AllSharePointSites-$month-$day-Final.csv" -Force
    Remove-Item "$exportCSVPath\AllSharePointSites-$month-$day.csv" 
    Remove-Item "$exportCSVPath\AllSharePointSites-$month-$day-1.csv"
}

Catch {
    write-host -f Red "Error:" $_.Exception.Message
}



Stop-Transcript
