# Addison Dunn, Michael Vabner 2016
#http://stackoverflow.com/questions/22159170/grab-files-from-most-recently-received-email-in-specific-outlook-folder

echo Starting

#Get the inbox folder
Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null
$olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]
$outlookObj = new-object -comobject outlook.application
$Namespace = $outlookObj.GetNameSpace(“MAPI”)
$inbox = $Namespace.getDefaultFolder($olFolders::olFolderInBox).Items | Sort-Object ReceivedTime -Descending

New-Item c:\users\addison.dunn\documents\temp_folder -type directory -force
$filepath = “C:\Users\addison.dunn\Documents\temp_folder”

#$filteredInbox = $inbox.Restrict(“[UnRead] = ‘True'”)
#$filteredInbox

#$inbox.items | Select-Object -Property Subject, ReceivedTime, Importance, SenderName
#$inbox.Attachments
#$email = $inbox.Item(4)
#$email.FlagStatus
#If($email.FlagStatus -eq 2)
#{
#    $email.Subject
#    echo "here"
#}
# For-loop
for($i=1; $i -lt $inbox.Count; $i++)
{
    $email = $inbox.Item($i)
    If((0 -lt $email.Attachments.Count) -And  ( -Not $email.FlagStatus -eq 1))
    {
        
        echo "Email subject: " $email.Subject

        $attachment = $email.Attachments.Item(1)
        $attachment | %{$_.saveasfile((join-path $filepath $_.filename))}
        echo "Loaded."
        echo " "
    }
}

# Make list of exceptions using excel file