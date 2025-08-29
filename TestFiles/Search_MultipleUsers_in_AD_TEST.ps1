# Import the Active Directory module
Import-Module ActiveDirectory

################################################################################################
# AD and ADLDS servers
################################################################################################

$AdWsEnabledGc = (Get-ADDomainController -Discover -Service GlobalCatalog,ADWS -SiteName "bru-hub").HostName[0] + ":3268"
$AdWsEnabledDc = (Get-ADDomainController -Discover -Service ADWS        -SiteName "bru-hub").HostName[0]

# --- Let the user choose where to search ---
Write-Host "Select your search scope:" -ForegroundColor Cyan
Write-Host "  1) Entire Directory (fast name lookup)" -ForegroundColor Green
Write-Host "  2) Domain Controller (all attributes)" -ForegroundColor Yellow

do {
    $choice = Read-Host "Enter 1 or 2"
} while ($choice -notin '1','2')

if ($choice -eq '1') {
    $PrimaryServer   = $AdWsEnabledGc   # use GC for speed
    $SecondaryServer = $AdWsEnabledDc   # fallback DC
    Write-Host "➤ Using Global Catalog for name lookup: $PrimaryServer" -ForegroundColor Cyan
}
else {
    $PrimaryServer   = $AdWsEnabledDc   # direct DC lookup
    $SecondaryServer = $null
    Write-Host "➤ Using Domain Controller for all lookups: $PrimaryServer" -ForegroundColor Yellow
}

# Prompt once for all names (comma-separated)
Write-Host "Enter at least first and last name(s) for faster search." -ForegroundColor Magenta
$userInput = Read-Host "User display names to search (comma-separated)"

# Split into array
$userList = $userInput -split '\s*,\s*' | Where-Object { $_ }
if (-not $userList) { Write-Host "No names entered; exiting." -ForegroundColor Red; exit }

# Prepare result array
$found = @()

foreach ($name in $userList) {
    # 1) Lookup by GC/DC for basic attributes
    $matches = Get-ADUser `
        -Server    $PrimaryServer `
        -Filter    "displayName -like '$name*'" `
        -Properties DisplayName,SamAccountName,DistinguishedName,Enabled,Manager,targetAddress

    if (-not $matches) {
        Write-Host "User '$name' not found in AD and may be created." -ForegroundColor Red
        continue
    }

    # 2) Pull enabled/disabled
    $enabled  = $matches | Where-Object { $_.Enabled }
    $disabled = $matches | Where-Object { -not $_.Enabled }

    # Report disabled
    if ($disabled) {
        Write-Host "Disabled account(s) for '$name':" -ForegroundColor Yellow
        $disabled | ForEach-Object { Write-Host " - $($_.DisplayName) ($($_.SamAccountName))" -ForegroundColor Magenta }
    }

    # 3) For each enabled user, fetch employeeNumber (GID) from DC if missing
    foreach ($u in $enabled) {
        if (-not $u.employeeNumber) {
            # Only if using GC as primary, fallback to DC
            if ($SecondaryServer) {
                $full = Get-ADUser -Identity $u.SamAccountName `
                                  -Server    $SecondaryServer `
                                  -Properties employeeNumber,extensionAttribute13
                $u.employeeNumber    = $full.employeeNumber
                $u.extensionAttribute13 = $full.extensionAttribute13
            }
        }
        $found += $u
    }
}

# Output the found users
$found | Select-Object `
    @{Name='DisplayName';    Expression={$_.DisplayName}}, `
    @{Name='UserID';         Expression={$_.SamAccountName}}, `
    @{Name='DOMAIN';         Expression={ ($_.DistinguishedName -split ',')[4] -replace '^DC=' }}, `
    @{Name='GID';            Expression={$_.employeeNumber -join ';'}}, `
    @{Name='ARP';            Expression={ ($_.Manager -split ',')[0] -replace '^CN=' }}, `     `
    @{Name='License';        Expression={$_.extensionAttribute13}} `
  | Format-Table -AutoSize -Wrap
