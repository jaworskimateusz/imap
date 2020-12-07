﻿Add-Type -Path "C:\chilkat\chilkatdotnet48-9.5.0-x64\ChilkatDotNet48.dll"
$imap = New-Object Chilkat.Imap

cls

# Base config
# Connect to an IMAP server.
# Use TLS
$imap.Ssl = $true
$imap.Port = 993
$imap_server = "outlook.office365.com"
$bUid = $false
$fetchUids = $true
$login = "01129802@pw.edu.pl"
$password = ""
$mailbox = "Inbox"
$basePath = -join("C:\Users\", $env:username, "\Studenci")
$basePathOthers = -join("C:\Users\", $env:username, "\Pozostałe")
$notSeenSearch = "NOT SEEN"


# --------- Main logic starts --------------

$success = $imap.Connect($imap_server)
if ($success -ne $true) {
    $($imap.LastErrorText)
    exit
}

# Login
$success = $imap.Login($login, $password)
if ($success -ne $true) {
    $($imap.LastErrorText)
    exit
}

# Select an IMAP mailbox
$success = $imap.SelectMailbox($mailbox)
if ($success -ne $true) {
    $($imap.LastErrorText)
    exit
}

# $numberOfMails = $imap.NumMessages
# $email = $imap.FetchSingle($numberOfMails, $bUid)

# Get the set of unseen message UIDs
$messageSet = $imap.Search($notSeenSearch, $fetchUids)
if ($imap.LastMethodSuccess -eq $false) {
    $($imap.LastErrorText)
    exit
}

# Fetch the unseen emails into a bundle object:
$bundle = $imap.FetchBundle($messageSet)
if ($imap.LastMethodSuccess -eq $false) {

    $($imap.LastErrorText)
    exit
}

$i = 0
$numberOfMails = $bundle.MessageCount
# Loop over unseen emails
while($i -lt $numberOfMails) {

    $email = $bundle.GetEmail($i)
    $isStudent = $email.FromAddress.EndsWith('stud@pw.edu.pl')

    if ($isStudent) {

        $studentDir = $email.FromAddress.Split("@")[0]
        Write-Host $email.EmailDateStr # todo create str variable like: yyyy-MM-dd 

        #creating directory for student
        if ( Test-Path -Path "$basePath\$studentDir" -PathType Container ) {
            "Folder $studentDir already exists in path: $basePath\" 
        } else { 
            New-Item -Path $basePath -Name "\$studentDir" -ItemType "directory"
            "Directory for student: $studentDir was created."
        }

        #creating .html file with email content in student' dir
        $fileTitle =  $email.Subject
        $content = $email.Body
	    Set-Content -Path "$basePath\$studentDir\$fileTitle.html" -Value $content

        #downloading attachments from email
        for ($i = 0; $i -lt $imap.GetMailNumAttach($email); $i++) {
          $imap.FetchAttachment($email, $i, "$basePath\$studentDir")
        }
            
    } else {
        Write-Host "This mail is not from student, sender is $email.FromAddress"
        # TODO set NOT SEEN
    }
    
    $i = $i + 1
}

# Disconnect from the IMAP server.
$success = $imap.Disconnect()

# --------- Main program ends ----------------

