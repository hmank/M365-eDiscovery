#############################################################################
#  This script is used to read single CSV file of either Microsoft 365 groups,
#  Shared mailboxes or Sharepoint sites then split the file into smaller files
#  with the appropraite number of rows as per the limits defined and
#  imported into an eDiscovery Case
#
#  Sept 8, 2021
#
# Version 1.0
# Author: Habib Mankal
# Credit : Adam the 32-bit Aardvark
################################################################################

$day = (get-date).day
$month = (get-date).month
$hour = (get-date).hour
$minute = (get-date).minute
$logsdir = "C:\temp"

Start-Transcript -LiteralPath "$logsdir\CSVSplit-Sites-Unfied-Shared-Mailboxes-$month-$day-$hour-$minute.log"  -NoClobber -Append


write-host "***********************************************"
write-host "  Split a single CSV file into multiple files  " -foregroundColor yellow -backgroundcolor darkgreen
write-host "  For use to import into eDiscovery Case   " -foregroundColor yellow -backgroundcolor darkgreen 
write-host "***********************************************"


do{
write-host "***********************************************"
write-host "   Please select below the M365 file type you wish to split " -foregroundColor yellow -backgroundcolor darkgreen
write-host "   [1] - Sharepoint Sites   " -foregroundColor yellow -backgroundcolor darkgreen
write-host "   [2] - Microsoft 365 Groups   " -foregroundColor yellow -backgroundcolor darkgreen
write-host "   [3] - Shared Mailboxes  " -foregroundColor yellow -backgroundcolor darkgreen 
$userInput = Read-Host "Select Option & Press Enter: " 
write-host "***********************************************"

switch ($userInput){
    
    '1' {"[1]- Sharepoint Sites Selected"} 
    '2' {"[2]- Microsoft365 Groups Selected"}
    '3' {"[3]- Shared Mailboxes Selected"}

    }#switch
    
}
While (($userInput -ne '1') -and ($userInput -ne '2') -and ($userInput -ne '3') )

"" 


Try{

 If ($userInput -eq 1){   #Split Sharepoint file.
        do{
        ""
        $spsinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Sharepoint Sites to place on hold. eg c:\Temp\AllSpSites.csv: " -NoNewline;read-host)
        ""
        $fileexists = test-path -path $spsinputfile
             if($fileexists -ne 'True'){write-host "$spsinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
         # variable used to store the path of the source CSV file


            $sourceCSVCount = (import-csv $spsinputfile).count
            $exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);
            $M365CSVFilename = $(Write-Host "Enter the name filename you wish to save: " -NoNewline;read-host);
            # variable used to advance the number of the row from which the export starts
            $startrow = 0 ;

            # counter used in names of resulting CSV files
            $counter = 1 ;

            # setting the while loop to continue as long as the value of the $startrow variable is smaller than the number of rows in your source CSV file
            while ($startrow -lt $sourceCSVCount)
            {

                # import of however many rows you want the resulting CSV to contain starting from the $startrow position and export of the imported content to a new file
                Import-CSV $spsinputfile | select-object -skip $startrow -first 98 | Export-CSV "$exportCSVPath\$M365CSVFilename$($counter).csv" -NoClobber -Verbose;

                # advancing the number of the row from which the export starts
                $startrow += 98 ;

                # incrementing the $counter variable
                $counter++ ;

            }    
           

        } #if


        If ($userInput -eq 2){   #Spliy Microsoft 365 Groups.
        do{
        ""
        $m365groupinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline;read-host)
        ""
        $fileexists = test-path -path $m365groupinputfile
             if($fileexists -ne 'True'){write-host "$m365groupinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
            $sourceCSVCount = (import-csv $m365groupinputfile).count
            $exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);
            $M365CSVFilename = $(Write-Host "Enter the name filename you wish to save: " -NoNewline;read-host);
            # variable used to advance the number of the row from which the export starts
            $startrow = 0 ;

            # counter used in names of resulting CSV files
            $counter = 1 ;

            # setting the while loop to continue as long as the value of the $startrow variable is smaller than the number of rows in your source CSV file
            while ($startrow -lt $sourceCSVCount)
            {

                # import of however many rows you want the resulting CSV to contain starting from the $startrow position and export of the imported content to a new file
                Import-CSV $m365groupinputfile | select-object -skip $startrow -first 998 | Export-CSV "$exportCSVPath\$M365CSVFilename$($counter).csv" -NoClobber -Verbose;

                # advancing the number of the row from which the export starts
                $startrow += 998 ;

                # incrementing the $counter variable
                $counter++ ;

            }  
            

        } #if


        If ($userInput -eq 3){   #Split Shared Mailboxes.
        do{
        ""
        $SharedMBXinputfile = $(Write-Host "Enter the name full path of the csv file that contains the Microsoft 365 Groups to place on hold. eg c:\Temp\AllM365Groups.csv :" -NoNewline;read-host)
        ""
        $fileexists = test-path -path $SharedMBXinputfile
             if($fileexists -ne 'True'){write-host "$SharedMBXinputfile doesn't exist. Please enter a valid path and filename." -foregroundcolor Yellow}
        }while($fileexists -ne 'True')
            
             
            $sourceCSVCount = (import-csv $SharedMBXinputfile).count
            $exportCSVPath = $(Write-Host "Enter the name full path to export the csv files to: " -NoNewline;read-host);
            $M365CSVFilename = $(Write-Host "Enter the name filename you wish to save: " -NoNewline;read-host);
            # variable used to advance the number of the row from which the export starts
            $startrow = 0 ;

            # counter used in names of resulting CSV files
            $counter = 1 ;

            # setting the while loop to continue as long as the value of the $startrow variable is smaller than the number of rows in your source CSV file
            while ($startrow -lt $sourceCSVCount)
            {

                # import of however many rows you want the resulting CSV to contain starting from the $startrow position and export of the imported content to a new file
                Import-CSV $SharedMBXinputfile | select-object -skip $startrow -first 998 | Export-CSV "$exportCSVPath\$M365CSVFilename$($counter).csv" -NoClobber -Verbose;

                # advancing the number of the row from which the export starts
                $startrow += 998 ;

                # incrementing the $counter variable
                $counter++ ;

            }  #while

                                      
        } #if



} #try

Catch {
 write-host -f Red "`tError:" $_.Exception.Message
} #catch

Stop-Transcript