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
$fetchUids = $true
$login = ""
$password = ""
$mailbox = "Inbox"
$basePath = -join("C:\Users\", $env:username, "\Studenci")
$basePathOthers = -join("C:\Users\", $env:username, "\Pozostałe")
$notSeenSearch = "NOT SEEN"

# --------- GUI ----------------------------

$login = Read-Host "Login"

$password_secure = Read-Host 'Enter a Password' -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password_secure)
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)


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
        $emailDate = $email.EmailDate.ToString("dd-MM-yyyy")

        #creating directory for student
        if ( Test-Path -Path "$basePath\$studentDir" -PathType Container ) {
            "Folder $studentDir already exists in path: $basePath\" 
        } else { 
            New-Item -Path $basePath -Name "\$studentDir" -ItemType "directory"
            "Directory for student: $studentDir was created."
        }

        #create directory for student with date
        if ( Test-Path -Path "$basePath\$studentDir\$emailDate" -PathType Container ) {
            "Folder $studentDir\$emailDate already exists in path: $basePath\" 
        } else { 
            New-Item -Path $basePath -Name "\$studentDir\$emailDate" -ItemType "directory"
            "Directory for today' mail: $studentDir\$emailDate was created."
        }

        #creating .html file with email content in student' dir
        $fileTitle =  $email.Subject
        $content = $email.Body.Replace("charset=iso-8859-2","charset=""UTF-8""")
	    Set-Content -Path "$basePath\$studentDir\$emailDate\$fileTitle.html" -Value $content

        #downloading attachments from email
        for ($i = 0; $i -lt $imap.GetMailNumAttach($email); $i++) {
          $imap.FetchAttachment($email, $i, "$basePath\$studentDir\$emailDate")
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

