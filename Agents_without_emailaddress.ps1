Import-Module ActiveDirectory

import-module ImportExcel

# Parameters
$SearchBase  = "OU=Users,OU=Express,OU=BE,DC=avi-dc,DC=dhl,DC=com"
$OutputFile  = "C:\Users\adiazque\DPDHL\EXP GAIT Identity and Access Management - Internal Kitchen - Internal Kitchen\Reports\Agents_NoEmail.xlsx"

# Get all users with employeeType = Agent and no email
$AgentsMissingEmail = @(Get-ADUser `
    -Filter { employeeType -eq "Agent" -and Enabled -eq $true } `
    -SearchBase $SearchBase `
    -Properties DisplayName,mail |
  Where-Object { -not $_.mail }
)
# Select the columns you care about
$AgentsMissingEmail |
  Select-Object `
    DisplayName,
    @{Name= 'UserID'; Expression={$_.SamAccountName}} |
    
  Export-Excel `
    -Path $OutputFile `
    -TableName 'EmptyEmailUsers' `
    -WorksheetName 'Sheet1' `
    -AutoSize `
    -ClearSheet

Write-Host "Found $($AgentsMissingEmail.Count) Agent accounts without email. Report saved to $OutputFile" -ForegroundColor DarkYellow
