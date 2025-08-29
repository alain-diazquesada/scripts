<#
.SYNOPSIS
  Export managers of Agent/BE0712 users from a chosen domain partition.
.PARAMETER Domain
  Which naming context to use: “AVI” or “PRG”.
#>
param(
  [ValidateSet('AVI','PRG','KUL','PHX')]
  [string]$Domain = 'AVI'
)

# ————— Define per-domain endpoints & bases —————
$DomainConfig = @{
  'AVI' = @{
    GcServer   = (Get-ADDomainController -Discover -Service GlobalCatalog,ADWS -SiteName 'bru-hub').HostName[0] + ':3268'
    DcServer   = (Get-ADDomainController -Discover -Service ADWS        -SiteName 'bru-hub').HostName[0]
    SearchBase = 'DC=avi-dc,DC=dhl,DC=com'
  }
  'PRG' = @{
    # AD LDS / different domain
    GcServer   = 'prg-dc.dhl.com:3268'
    DcServer   = 'prg-dc.dhl.com'
    SearchBase = 'DC=prg-dc,DC=dhl,DC=com'
  }
  
  'KUL' = @{
      GcServer   = 'kul-dc.dhl.com:3268'
      DcServer   = 'kul-dc.dhl.com'
      SearchBase = 'DC=kul-dc,DC=dhl,DC=com'
  }
    'PHX' = @{
        GcServer   = 'phx-dc.dhl.com:3268'
        DcServer   = 'phx-dc.dhl.com'
        SearchBase = 'DC=phx-dc,DC=dhl,DC=com'
    }
  
}

# ————— Pull in the settings for the chosen domain —————
$cfg        = $DomainConfig[$Domain]
$gcServer   = $cfg.GcServer
$dcServer   = $cfg.DcServer
$searchBase = $cfg.SearchBase

Write-Host "Using Domain: $Domain"                  -ForegroundColor Cyan
Write-Host "  GC endpoint: $gcServer"               -ForegroundColor Cyan
Write-Host "  DC endpoint: $dcServer"               -ForegroundColor Cyan
Write-Host "  SearchBase  : $searchBase"            -ForegroundColor Cyan


# ————— Rest of your logic (find BE0712+Agent users, collect manager DNs, resolve & export) —————
if ($Domain -eq 'AVI') {
    $filterString = "dpwncrestcode -eq 'BE0712' -and employeeType -eq 'Agent'"
}
else {
    $filterString = "employeeType -eq 'Agent'"
}
Write-Host "Using filter: $filterString" -ForegroundColor Cyan

# Example: fetch only BE0712+Agent users who have a manager
$qualifiedUsers = Get-ADUser `
    -Server     $dcServer `
    -SearchBase $searchBase `
    -Filter     $filterString `
    -Properties manager |
  Where-Object { $_.manager }

# De-dupe their manager DNs
$managerDNs = $qualifiedUsers | Select-Object -Expand manager -Unique

# Resolve each manager and build your PSCustomObject list…
$results = foreach ($dn in $managerDNs) {

    # Attempt to bind to the manager account
    try {
        $mgr = Get-ADUser -Server $dcServer -Identity $dn `
            -Properties DisplayName,sAMAccountName,Enabled
    }
    catch {
        # Couldn’t find this manager DN—skip it entirely
        continue
    }

    # Emit only for successfully bound manager objects
    [PSCustomObject]@{
        Manager = if ($mgr.DisplayName) { $mgr.DisplayName }
                  else { $mgr.sAMAccountName }
        UserID  = $mgr.sAMAccountName
        Status  = if ($mgr.Enabled) { 'Enabled' } else { 'Disabled' }
    }
}

Write-Host "→ Resolved $($results.Count) manager accounts." -ForegroundColor Cyan

# Export to Excel…
$outFile    = "C:\temp\Managers_BE0712_$dcServer+$(Get-Date -Format yyyyMMdd).xlsx"

if ($results) {
    $results | Export-Excel `
        -Path          $outFile `
        -WorksheetName 'Managers' `
        -TableName     'BE0712Managers' `
        -AutoSize
    Write-Host "✅ Exported $($results.Count) managers to $outFile" -ForegroundColor Green
} else {
    Write-Host "⚠ No managers matched the criteria." -ForegroundColor Yellow
}

