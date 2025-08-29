# Import the Active Directory module
Import-Module ActiveDirectory

################################################################################################
# AD and ADLDS servers
################################################################################################
# $AdwsLds = "bezavwseds1001.prg-dc.dhl.com"
$AdWsEnabledGc = (Get-ADDomainController -Discover -Service GlobalCatalog,ADWS -SiteName "bru-hub").HostName[0] + ":3268"
$AdWsEnabledDc = (Get-ADDomainController -Discover -Service ADWS        -SiteName "bru-hub").HostName[0]

# --- Let the user choose where to search ---
Write-Host "Select your search scope:" -ForegroundColor Cyan
Write-Host "  1) Entire Directory" -ForegroundColor Green
Write-Host "  2) Domain Controller (AVI-DC)" -ForegroundColor Yellow
# Write-Host "  3) ADLDS/LDAP server" -ForegroundColor DarkBlue

do{
    $choice = Read-Host "Enter 1 or 2"
}while ($choice -notin '1', '2')

if ($choice -eq '1') {
    $Server = $AdWsEnabledGc
    Write-Host "➤ Using Entire Directory: $Server" -ForegroundColor Cyan
}
else{
    $Server = $AdWsEnabledDc
    Write-Host "➤ Using Domain Controller: $Server" -ForegroundColor Yellow
}

# Prompt once for all names (comma-separated)
Write-Host "To make the search faster enter at least first name and last name of the user(s) you want to search for." -ForegroundColor Magenta
$userinput = Read-Host "Enter user names to search (comma-separated)"

# Split the input into an array of names, trimming whitespace
$userList = $userinput -split '\s*,\s*' | Where-Object { $_}

if (-not $userList) {
    Write-Host "No names entered; exiting."
    exit
}

# search base for the AD LDS server
#$Searchbase = "DC=avi-dc,DC=dhl,DC=com" # This is the base DN for the AD

# Create a new array to hold the formatted strings
$found = @()

# Loop through each user in the list
foreach ($name in $userList) {
    # Get the user from Active Directory
        # $u = Get-ADUser -Filter * -Properties Name,Displayname,SAMAccountName,DistinguishedName | Where-Object { $_.DisplayName -like "*$name*" }
        $umatches = Get-ADUser `
            -Server $Server `
            -Filter "displayName -like '$name*'" `
            -Properties DisplayName,SamAccountName,DistinguishedName,employeeType,employeeNumber,Manager,extensionAttribute13,mail
             
    if(-not $umatches) {
        # If the user is not found, output a message
        Write-Host "User $name not found in AD and may be created." -ForegroundColor Red
        continue
    }
    
    # Separate enabled vs disabled users
    $enabled = $umatches | Where-Object { $_.Enabled -eq $true }
    $disabled = $umatches | Where-Object { $_.Enabled -eq $false }

    # Accumulate the enable ones for the final output
    if ($enabled) {        
        $found += $enabled
    }
    
    # Report disabled users
    if ($disabled) {
        Write-Host "following account(s) for '$name' where found but are DISABLED in AD." -ForegroundColor Yellow
        $disabled | ForEach-Object {
            Write-Host " - $($_.DisplayName) ($($_.SAMAccountName))" -ForegroundColor Magenta
        }
    }
    
    
}

# Output the found users
$found |
    Select-Object `
    DisplayName, `
    @{Name = 'UserID'; Expression = { $_.SAMAccountName }}, `
    @{Name = 'DOMAIN'; Expression = { ($_.DistinguishedName -split ',')[4] -replace '^DC='  }}, `
    @{Name = 'GID'; Expression = {$_.employeeNumber}}, `
    @{Name = 'ARP'; Expression = {($_.manager -split ',')[0] -replace '^CN='} }, `
    @{Name = 'License'; Expression = {$_.extensionAttribute13} }, `
    @{Name = 'Email'; Expression = {$_.mail} } `
     | 
    Format-Table -AutoSize -Wrap
