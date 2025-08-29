<#
Report columns:
- userid (sAMAccountName)
- mail AD
- mail EDS (from EDS)
- UPN
- targetaddress
- proxyaddress
- mailhost (from EDS)
- O365License  (raw extensionAttribute13)

Filters:
- employeeType = "Agent"
- manager is jkenrick OR rodemytt (resolved to DN)

Search:
- AD attributes from Global Catalog (forest-wide)
- EDS mail + mailhost from bezavwseds1001.prg-dc.dhl.com (O=dhl.com)

Output:
- XLSX to your Reports path, merged centered title row:
  "All users Managed by Jody Kendrik and Robin Demyttenaere"
#>

#region Servers (your discovery)
$AdWsEnabledGc = (Get-ADDomainController -Discover -Service GlobalCatalog,ADWS -SiteName "bru-hub").HostName[0] + ":3268"
$AdWsEnabledDc = (Get-ADDomainController -Discover -Service ADWS        -SiteName "bru-hub").HostName[0]
$SearchServer   = $AdWsEnabledGc
#endregion

#region Config
$EdsServer   = "bezavwseds1001.prg-dc.dhl.com"
$EdsBaseDN   = "O=dhl.com"

$ManagerSamAccountNames     = @('jkenrick','rodemytt')
$EmployeeTypeValue          = "Agent"
$EnsureEmployeeTypeFromDC   = $true   # verify/filter via DC if GC can't

$OutDir   = "C:\Users\adiazque\DPDHL\EXP GAIT Identity and Access Management - Internal Kitchen - Internal Kitchen\Reports"
$OutFile  = Join-Path $OutDir "ManagedAgents_SMTP_Mailhost_Report.xlsx"
$Sheet    = "Report"
$Title    = "All users Managed by Jody Kendrik and Robin Demyttenaere"

$UserIdAttribute = "sAMAccountName"
#endregion

#region Helpers
function Test-Module {
    param([Parameter(Mandatory)][string]$Name)
    try { Import-Module $Name -ErrorAction Stop; $true } catch { $false }
}

# EDS searcher for mail + mailhost
function New-EdsSearcher {
  param([System.Management.Automation.PSCredential]$Credential)
  $rootPath = "LDAP://$EdsServer/$EdsBaseDN"
  $root = if ($Credential) {
    New-Object System.DirectoryServices.DirectoryEntry($rootPath, $Credential.UserName, $Credential.GetNetworkCredential().Password)
  } else {
    New-Object System.DirectoryServices.DirectoryEntry($rootPath)
  }
  $ds = New-Object System.DirectoryServices.DirectorySearcher($root)
  $ds.PageSize = 1000
  $ds.PropertiesToLoad.Clear()
  @('mail','mailhost') | ForEach-Object { [void]$ds.PropertiesToLoad.Add($_) }
  $ds
}

# Resolve EDS mail + mailhost (try uid=sAM, fallback mail=AD mail)
function Get-EdsMailAndHost {
  param(
    [Parameter(Mandatory)][string]$SamAccountName,
    [string]$AdMail,
    [System.DirectoryServices.DirectorySearcher]$Searcher
  )
  $Searcher.Filter = "(&(objectClass=person)(uid=$SamAccountName))"
  $r = $Searcher.FindOne()
  if (-not $r -and $AdMail) {
    $Searcher.Filter = "(&(objectClass=person)(mail=$AdMail))"
    $r = $Searcher.FindOne()
  }
  if ($r) {
    [pscustomobject]@{
      MailEDS  = ($r.Properties['mail']     | Select-Object -First 1)
      MailHost = ($r.Properties['mailhost'] | Select-Object -First 1)
    }
  } else {
    [pscustomobject]@{ MailEDS = $null; MailHost = $null }
  }
}
#endregion

#region Modules
if (-not (Ensure-Module -Name ActiveDirectory)) {
  Write-Error "ActiveDirectory module not found. Install RSAT or run on a domain-joined admin workstation."
  return
}
$haveImportExcel = Ensure-Module -Name ImportExcel
#endregion

#region Resolve manager DNs via GC
$ManagerDNs = $ManagerSamAccountNames | ForEach-Object {
  $mgr = Get-ADUser -Server $SearchServer -Filter "sAMAccountName -eq '$_'" -Properties DistinguishedName -ErrorAction SilentlyContinue
  if ($mgr) { $mgr.DistinguishedName }
} | Where-Object { $_ }

if (-not $ManagerDNs) { Write-Error "No manager DNs resolved on GC. Aborting."; return }
#endregion

#region Query users (GC first)
$props = @('mail','userPrincipalName','targetAddress','proxyAddresses','extensionAttribute13',
           'manager','sAMAccountName','employeeType','distinguishedName')

# Try filtering on GC including employeeType (works only if in PAS)
$managerOr = ($ManagerDNs | ForEach-Object { "(manager=$_)"} ) -join ''
$ldapFilterWithEmp = "(&(objectClass=user)(!(objectClass=computer))(employeeType=$EmployeeTypeValue)(|$managerOr))"

$users = Get-ADUser -Server $SearchServer -LDAPFilter $ldapFilterWithEmp -Properties $props -ErrorAction SilentlyContinue

# If GC couldn’t apply employeeType, fall back: manager-only from GC, verify employeeType via DC
if ($EnsureEmployeeTypeFromDC -and ($users.Count -eq 0 -or ($users | Where-Object { -not $_.employeeType }).Count -gt 0)) {
  $ldapFilterMgrOnly = "(&(objectClass=user)(!(objectClass=computer))(|$managerOr))"
  $candidates = Get-ADUser -Server $SearchServer -LDAPFilter $ldapFilterMgrOnly -Properties $props -ErrorAction SilentlyContinue

  $users = foreach ($c in $candidates) {
    try {
      $full = Get-ADUser -Server $AdWsEnabledDc -Identity $c.DistinguishedName -Properties $props -ErrorAction Stop
      if ($full.employeeType -eq $EmployeeTypeValue) { $full }
    } catch { }
  }
}

if (-not $users) { Write-Warning "No users matched employeeType='$EmployeeTypeValue' for managers jkenrick/rodemytt."; $users = @() }
#endregion

#region EDS lookup (mail + mailhost)
$EdsCred = $null  # set to Get-Credential if EDS requires it
$EdsSearcher = New-EdsSearcher -Credential $EdsCred
#endregion

#region Build rows (extensionAttribute13 as-is; DC fallback if missing)
$rows = foreach ($u in $users) {
  $sam        = $u.$UserIdAttribute
  $mailAD     = $u.mail
  $eds        = Get-EdsMailAndHost -SamAccountName $sam -AdMail $mailAD -Searcher $EdsSearcher
  $primarySmtpRaw = $u.proxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1
if (-not $primarySmtpRaw) {
    # fallback to any smtp: if primary isn't present
    $primarySmtpRaw = $u.proxyAddresses | Where-Object { $_ -match '^smtp:' } | Select-Object -First 1
}
$primarySmtp = if ($primarySmtpRaw) { $primarySmtpRaw -replace '^(?i)smtp:' } else { '' }

  # Ensure we have extensionAttribute13 even if GC didn't return it
  $ext13 = $u.extensionAttribute13
  if (-not $ext13) {
    try {
      $full = Get-ADUser -Server $AdWsEnabledDc -Identity $u.DistinguishedName -Properties extensionAttribute13 -ErrorAction Stop
      $ext13 = $full.extensionAttribute13
    } catch { }
  }

  [pscustomobject]@{
    userid        = $sam
    'mail AD'     = $mailAD
    'mail EDS'    = $eds.MailEDS
    UPN           = $u.userPrincipalName
    targetaddress = $u.targetAddress
    proxyaddress  = $primarySmtp
    mailhost      = $eds.MailHost
    O365License   = $ext13
  }
}
#endregion

#region Export XLSX (ImportExcel preferred, else COM)
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

if ($haveImportExcel) {
  # Ensure we overwrite the same file (avoid OneDrive conflicted copies / stale tables)
  try {
    Remove-Item -LiteralPath $OutFile -Force -ErrorAction Stop
  } catch {
    Write-Warning "Could not remove existing $OutFile (is it open in Excel?). Attempting to overwrite."
  }

  $null = $rows | Export-Excel -Path $OutFile `
      -WorksheetName $Sheet `
      -TableName 'Users' `
      -AutoSize `
      -BoldTopRow `
      -FreezeTopRow `
      -Title $Title `
      -TitleBold `
      -TitleSize 14
  
}
else {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = $Sheet

  $headers = @('userid','mail AD','mail EDS','UPN','targetaddress','proxyaddress','mailhost','O365License')
  for ($i=0; $i -lt $headers.Count; $i++) {
    $ws.Cells.Item(2, $i+1).Value2 = $headers[$i]
    $ws.Cells.Item(2, $i+1).Font.Bold = $true
  }

  $r = 3
  foreach ($row in $rows) {
    $c = 1
    foreach ($h in $headers) {
      $ws.Cells.Item($r,$c).Value2 = ($row.$h)
      $c++
    }
    $r++
  }

  $lastCol = $headers.Count
  $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$lastCol)).Merge()
  $ws.Cells.Item(1,1).Value2 = $Title
  $ws.Cells.Item(1,1).HorizontalAlignment = -4108  # xlCenter
  $ws.Cells.Item(1,1).Font.Bold = $true
  $ws.Cells.Item(1,1).Font.Size = 14
  $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

  $null = New-Item -ItemType Directory -Force -Path $OutDir -ErrorAction SilentlyContinue
  $wb.SaveAs($OutFile, 51)  # xlsx
  $wb.Close($true)
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
#endregion

Write-Host "Done. Exported $($rows.Count) rows to:`n$OutFile" -ForegroundColor Green
