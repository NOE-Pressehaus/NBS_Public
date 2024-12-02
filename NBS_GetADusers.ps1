# Import the Active Directory module
Import-Module ActiveDirectory

# Directory for the files where they are saved at
$newDirPath = "\\server.domain.local\share\NBS"

# Groupnames where our licensed Users are a member of, you can also use Get-Aduser -Filer * to get all Users.
$groupNames = @(
    "LicenseGroup_freieMitarbeiter",
    "LicenseGroup_Angestellt"
)

# Users to exclude
$excludedUsers = @(
    "t.user@test2.com",
    "t.user@test.com"
    
)

# Create an empty array to store the user objects
$users = @()

# Function to set the institution field from the upn, so if it upn end with @company.at the institution is Company, if it would end with nbs.at it Would be NBS.
# The value you want to set needs to be in the {}
function Get-UserObject {
    param($user, $groupMembership)

    $institution = switch -Regex ($user.Mail) {
        '@Company\.at$' { "Company" }
        '@Comapany1\.at$' { "Company1" }
        '@nbs\.at$' {"NBS"}
        default { "Unknown" }
    }

#sets the fields for the csv
    $userObject = [PSCustomObject] @{
        username = $user.UserPrincipalName.ToLower()
        lastname = $user.Surname
        firstname = $user.GivenName        
        email = $user.Mail.ToLower()
        department = $user.department
        maildisplay = "0"
        auth = "oauth2"
        profile_field_MA_Status = $groupMembership
        cohort1 = "ALLE_MA"
        institution = $institution
    }

    return $userObject
}

# Iterate through the group names and retrieve user objects with specific attributes
foreach ($groupName in $groupNames) {
    $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive |
                    Get-ADUser -Properties Enabled, * |
                    Where-Object { $_.Enabled -eq $true -and $excludedUsers -notcontains $_.UserPrincipalName }

    $groupMembership = if ($groupName -match "freieMitarbeiter) { "Frei" } else { "Fix" }

    foreach ($member in $groupMembers) {
        [array]$users += Get-UserObject -user $member -groupMembership $groupMembership
    }
}

# Export the new user information to a CSV file
$currentDate = (Get-Date).ToString("yyMMdd")
$newCsvPath = "$newDirPath\nbsenabled$currentDate.csv"
$users | Export-Csv -Path $newCsvPath -Encoding UTF8 -NoTypeInformation

# Find the latest "big" file
$latestBigFile = Get-ChildItem -Path $newDirPath -Filter "nbsenabled_*.csv" |
                 Where-Object { $_.Name -notmatch "_deleted" } |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1

# Import the old CSV file
$oldUsers = Import-Csv -Path $latestBigFile.FullName

# Compare the new user list with the old one
$newUsernames = $users.username
$oldUsernames = $oldUsers.username

$deletedUsers = $oldUsers | Where-Object { $newUsernames -notcontains $_.username }

# Create a new object with only the username and deleted fields
$deletedUsers = $deletedUsers | Select-Object @{Name="username"; Expression={$_.username}}, @{Name="deleted"; Expression={"1"}}

# Export the deleted users to a CSV file with the current date
$deletedUsers | Export-Csv -Path "$newDirPath\nbsenabled_deleted_$currentDate.csv" -Encoding UTF8 -NoTypeInformation

# Rename the current "big" file to include the date
Rename-Item -Path $newCsvPath -NewName "nbsenabled_$currentDate.csv"



# Define the email parameters
$emailFrom = "from@sender.com"
$emailTo = "to@recipient.com"
$emailSubject = "NBS Benutzerliste - NBS $currentDate"
$emailBody = "Hallo Test, hier die Automatisierte Benutzerliste für das Unternehmen vom $currentDate. Bitte einpfelgen und um kurze Rückmeldung. Danke!"
$smtpServer = "mailrelay.test.com"

# Paths to the CSV files
$newCsvPath = "$newDirPath\nbsenabled_$currentDate.csv"
$deletedCsvPath = "$newDirPath\nbsenabled_deleted_$currentDate.csv"

# Send the email with the CSV attachments
Send-MailMessage -From $emailFrom -To $emailTo -Subject $emailSubject -Body $emailBody -Attachments $newCsvPath, $deletedCsvPath -SmtpServer $smtpServer -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)

Write-Output "Email sent successfully."
