Import-Module ActiveDirectory
Import-Module ImportExcel

# ─── Parameters ──────────────────────────────────────────────────────────────
$SearchBase = "OU=Users,OU=Express,OU=BE,DC=avi-dc,DC=dhl,DC=com"
$OutputFile = "C:\Users\adiazque\DPDHL\EXP GAIT Identity and Access Management - Internal Kitchen - Internal Kitchen\Reports\Agents_NoEmail.xlsx"

# ─── 1) Pull your data (always as an array) ──────────────────────────────────
$AgentsMissingEmail = @( 
  Get-ADUser `
    -Filter { employeeType -eq "Agent" } `
    -SearchBase $SearchBase `
    -Properties DisplayName,Mail |
  Where-Object { -not $_.Mail }
)

# ─── 2) Export (this wipes Sheet1 completely) ────────────────────────────────
$AgentsMissingEmail |
  Select-Object `
    DisplayName,
    @{Name='UserID';Expression={$_.SamAccountName}} |
  Export-Excel `
    -Path          $OutputFile `
    -WorksheetName 'Sheet1' `
    -TableName     'EmptyEmailUsers' `
    -AutoSize      `
    -ClearSheet    # removes the old Sheet1 so we get a fresh one :contentReference[oaicite:0]{index=0}

# ─── 3) If zero rows, write back just the two headers ────────────────────────
if ($AgentsMissingEmail.Count -eq 0) {
    # this will put “DisplayName” into A1 and “UserID” into B1 :contentReference[oaicite:1]{index=1}
    Set-ExcelRange `
      -Path         $OutputFile `
      -WorksheetName 'Sheet1' `
      -Range        'A1:B1' `
      -Value        @('DisplayName','UserID')
}

Write-Host "Found $($AgentsMissingEmail.Count) Agent accounts without email. Report saved to $OutputFile" `
  -ForegroundColor DarkYellow
