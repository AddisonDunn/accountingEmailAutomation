# Addison Dunn, Michael Vabner 2016
#http://stackoverflow.com/questions/22159170/grab-files-from-most-recently-received-email-in-specific-outlook-folder

#$ErrorActionPreference = 'Stop'
echo Starting

#Get the inbox folder
Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null
$olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]
$oldFolders = $namespace.Folders.Item('michael.vabner@inforeliance.com').Folders
$outlookObj = new-object -comobject outlook.application
$Namespace = $outlookObj.GetNameSpace(“MAPI”)
$inbox = $Namespace.getDefaultFolder($olFolders::olFolderInBox).Items | Sort-Object ReceivedTime -Descending

New-Item c:\users\addison.dunn\documents\temp_folder -type directory -force
$main_filepath = “C:\Users\addison.dunn\Documents\temp_folder”

#$filteredInbox = $inbox.Restrict(“[UnRead] = ‘True'”)
#$filteredInbox

#$inbox.items | Select-Object -Property Subject, ReceivedTime, Importance, SenderName
#$inbox.Attachments

#Function to look through excel file and turn contents of first column into list 
$Excel = New-Object -ComObject Excel.Application 
$Excel.Visible = $true
$Excel.DisplayAlerts = $false
$ExcelWorkBook = $Excel.Workbooks.Open("C:\Users\addison.dunn\Documents\FolderEmailListExceptions.xlsx") 
$ExcelWorkSheet = $Excel.Sheets.item("Sheet1") 
$ExcelWorkSheet.activate() 
$arrBlackListEmails = @()
$i = 1
Do 
{
    $arrBlackListEmails += $ExcelWorkSheet.Cells.Item($i, 1).Value()
    $i = $i + 1
}
Until ($ExcelWorkSheet.Cells.Item($i, 1).Value() -eq $null)
$excel.Quit()

# Loop through emails in inbox
for($i=1; $i -lt $inbox.Count; $i++)
{
    $email = $inbox.Item($i)
    
    # Check if there is an attachment and that the email is not checked
    If((0 -lt $email.Attachments.Count) -And  ( -Not $email.FlagStatus -eq 1))
    {
        # Get email address
        $address = $email.SenderEmailAddress
        # If email address is internal, this if-statement fixes the formatting
        If ($email.SenderEmailType -eq "EX") 
        {
            $address = $email.Sender.GetExchangeUser().PrimarySmtpAddress
        }

        # Get company name from email address
        $match = $address -replace ".*@" -replace ".com.*"
        $arrBlackListNames = @("Theresa Grouge", "Jen Dunlap", "Beverly Goodwin", "John Sankovich", "Sara Mallory", "Jacob Elliot")


        $b = $true
        #Check for exceptions
        If( -Not $arrBlackListNames.Contains($email.SenderName) -And (-Not $arrBlackListEmails.Contains($address))) {
            Foreach ($element in $arrBlackListEmails)
            {
                If ($address -Match $element)
                {
                    $b = $false
                }
            }

            If ($b){

                # Load attachment
                $attachment = $email.Attachments.Item(1)
                $filepath = $main_filepath + "\" + $match
                If (-Not (Test-Path $filepath))
                {
                    New-Item $filepath -type directory -force
                }

                $date = $email.SentOn.ToString("yyyy-MM-dd")

                $attachment = $email.Attachments.Item(1)
                $startingFilename = $attachment.FileName
                echo "Starting filename: $startingFilename"
                #$startingFilename.getType()
                #$s1, $s2 = $startingFilename -split '.'

               # echo "text: $s1"
                #$filename = $s1 + " " + $date + '.' + $s2
                $filename = $date + " " + $startingFilename

                $attachment | %{$_.saveasfile((join-path $filepath ($filename)))}
                echo "Loaded."
                echo " "

                $email.FlagStatus = 1
               
                
                
            }
            
        }
        
        

        #echo "Email subject: " $email.Subject

        #$attachment = $email.Attachments.Item(1)
        #$attachment | %{$_.saveasfile((join-path $filepath $_.filename))}
        #echo "Loaded."
        #echo " "
    }
}

# Make list of exceptions using excel file