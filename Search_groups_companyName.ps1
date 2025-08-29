# Import the Active Directory module
Import-Module ActiveDirectory

Write-Host "Enter the company name to search for groups: " -ForegroundColor Yellow
$term = Read-Host
if ([string]::IsNullOrWhiteSpace($term)) {
  Write-Warning "No search term provided. Exiting."
  exit 1
}

$gc   = 'prg-dc.dhl.com:3268'   # any GC, include :3268
$base = (Get-ADRootDSE -Server $gc).RootDomainNamingContext
$results = @()

foreach ($word in $term.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)) {
  try {
    Write-Host "Searching for groups containing '$word'..." -ForegroundColor Cyan
    $found = Get-ADGroup -Server $gc -SearchBase $base `
      -LDAPFilter "(&(objectClass=group)(|(cn=*$word*)(name=*$word*)))" -ErrorAction Stop |
      Select-Object Name, DistinguishedName

    if ($found) { $results += $found }
  } catch {
    Write-Warning "Search for '$word' failed: $($_.Exception.Message)"
  }
}

if ($results) {
  Write-Host "Search completed." -ForegroundColor Green
  $results | Sort-Object Name | Format-Table -AutoSize
} else {
  Write-Host "No groups found." -ForegroundColor Yellow
}
