# Simple ADLDS User Search Script
Import-Module ActiveDirectory

# ADLDS server and search base
$Server = "bezavwseds1001.prg-dc.dhl.com"
$SearchBase = "O=dhl.com"

# Get user input
Write-Host "Enter userids to search (comma-separated)" -ForegroundColor Yellow
$userInput = Read-Host

$userList = $userInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

if (-not $userList) {
    Write-Host "No userids entered."
    exit
}

# Search for each user
$results = @()
foreach ($userid in $userList) {
    try {
        $user = Get-ADObject -Filter "uid -eq '$userid'" -Server $Server -SearchBase $SearchBase -Properties DisplayName,mail,mailhost,uid -ErrorAction Stop
        
        if ($user) {
            $results += [PSCustomObject]@{
                UserID = $user.uid
                DisplayName = $user.DisplayName
                MailHost = $user.mailhost
            }
        }
    }
    catch {
        Write-Host "User '$userid' not found or error occurred: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Display results
if ($results) {
    $results | Format-Table -AutoSize
} else {
    Write-Host "No users found." -ForegroundColor Yellow
}