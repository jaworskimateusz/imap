Add-Type -Path "C:\chilkat\chilkatdotnet48-9.5.0-x64\ChilkatDotNet48.dll"
$imap = New-Object Chilkat.Imap

cls

# Base config
# Connect to an IMAP server.
# Use TLS
$imap.Ssl = $true
$imap.Port = 993
$imap_server = "outlook.office365.com"
$bUid = $false
$login = "01129802@pw.edu.pl"
$password = ""
$mailbox = "Inbox"
$basePath = -join("C:\Users\", $env:username, "\Studenci")

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

# --------- Main logic starts --------------

$numberOfMails = $imap.NumMessages
$email = $imap.FetchSingle($numberOfMails, $bUid)
$isStudent = $email.FromAddress.EndsWith('stud@pw.edu.pl')

createDirectory($email)

saveEmailContent($numberOfMails)

downloadEmailAttachments($numberOfMails)


#for ($i = $numberOfMails; $i -gt 0; $i--) {

    # $email = $imap.FetchSingle($numberOfMails, $bUid)
    # $imap.GetMailFlag($email, '\Seen') 
    # todo do while not seen? we must thing and find flag

#}

# Disconnect from the IMAP server.
$success = $imap.Disconnect()

# --------- Main program ends ----------------


# functions used in script

Function createDirectory($email) {
    $studentDir = $email.FromAddress.Split("@")[0]
    Write-Host "$basePath\$studentDir"

    if ( Test-Path -Path "$basePath\$studentDir" -PathType Container ) {
        "Folder $studentDir already exists in path: $basePath\" 
    } else { 
        New-Item -Path $basePath -Name "\$studentDir" -ItemType "directory"
        "Directory for student: $studentDir was created."
    }
}

Function saveEmailContent($mailNumber) {
    $fileTitle =  $email.Subject
    Write-Host $email.Body # todo get body content to variable $content
    Write-Host $email.EmailDateStr # todo create str variable like: yyyy-MM-dd 
}

Function downloadEmailAttachments($mailNumber) {
    $email = $imap.FetchSingle($mailNumber, $bUid)
    $studentDir = $email.FromAddress.Split("@")[0]
    for ($i = 0; $i -lt $imap.GetMailNumAttach($email); $i++) {
       $imap.FetchAttachment($email, $i, "$basePath\$studentDir")
    }
}
