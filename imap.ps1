Add-Type -Path ".\dll\ChilkatDotNet48.dll" 
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
$basePath = (Get-Content -Path .\config.txt -TotalCount 6)[-1]
#$basePath = -join("C:\Users\", $env:username, "\Studenci")
$seenSearch = "SEEN"
$notSeenSearch = "NOT SEEN"

# --------- GUI ----------------------------

$login = Read-Host "Login"

$password_secure = Read-Host 'Enter a Password' -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password_secure)
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$login_file = (Get-Content -Path .\config.txt -TotalCount 2)[-1]
$password_file = (Get-Content -Path .\config.txt -TotalCount 4)[-1]

if($login -eq ""){
	$login = $login_file	
	$password = $password_file	
	exit
}


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

        #finding lectture name which is in [...]
        $lectureDir = $email.Subject
        $lectureDir = $lectureDir.Substring($lectureDir.IndexOf("[") + 1)
        $lectureDir = $lectureDir.Substring(0, $lectureDir.IndexOf("]"))

        $fileTitle = $email.Subject.Substring($email.Subject.IndexOf("]") + 1)
        $studentDir = $email.FromAddress.Split("@")[0]
        $studentDir = $studentDir.Substring(0, $studentDir.Length - 5)
        $emailDate = $email.EmailDate.ToString("dd-MM-yyyy-HH.mm")

        #creating directory for lecture
        if ( Test-Path -Path "$basePath\$lectureDir" -PathType Container ) {
            Write-Host "Directory $lectureDir already exists." 
        } else { 
            New-Item -Path $basePath -Name "\$lectureDir" -ItemType "directory"
            Write-Host "Directory for lecture: $lectureDir was created."
        }

        #creating directory for student
        if ( Test-Path -Path "$basePath\$lectureDir\$studentDir" -PathType Container ) {
            Write-Host "Directory $lectureDir\$studentDir already exists." 
        } else { 
            New-Item -Path $basePath -Name "\$lectureDir\$studentDir" -ItemType "directory"
            Write-Host "Directory for student: $lectureDir\$studentDir was created."
        }

        #create directory for student with date
        if ( Test-Path -Path "$basePath\$lectureDir\$studentDir\$emailDate" -PathType Container ) {
            Write-Host "Directory $lectureDir\$studentDir\$emailDate already exists."
        } else { 
            New-Item -Path $basePath -Name "\$lectureDir\$studentDir\$emailDate" -ItemType "directory"
            Write-Host "Directory for today' mail: $lectureDir\$studentDir\$emailDate was created."
        }

        #creating .html file with email content in student' dir
        $content = $email.Body.Replace("charset=iso-8859-2","charset=""UTF-8""")
	    Set-Content -Path "$basePath\$lectureDir\$studentDir\$emailDate\$fileTitle.html" -Value $content

        #downloading attachments from email
        for ($i = 0; $i -lt $imap.GetMailNumAttach($email); $i++) {
          $imap.FetchAttachment($email, $i, "$basePath\$lectureDir\$studentDir\$emailDate")
        }
            
    } else {
        Write-Host "This mail is not from student, sender is $email.FromAddress"
        $success = $imap.SetMailFlag($email,"SEEN",0)
        if ($success -ne $true) {
           $($imap.LastErrorText)
              exit
        }
    }
    
    $i = $i + 1
}

# Disconnect from the IMAP server.
$success = $imap.Disconnect()

# --------- Main program ends ----------------

