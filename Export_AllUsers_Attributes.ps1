# Export users with employeeType=Agent from specific OU to Excel with ALL populated attributes
$domain = "avi-dc.dhl.com"
$SearchBase = "OU=Users,OU=Express,OU=BE,DC=avi-dc,DC=dhl,DC=com"

# Get users with employeeType=Agent that are Enabled from the specific OU with ALL properties
Write-Host "Retrieving enabled users with employeeType=Agent..." -ForegroundColor Yellow
$users = Get-ADUser -Server $domain -Filter {employeeType -eq "Agent" -and Enabled -eq $true} -SearchBase $SearchBase -Properties *

if ($users.Count -eq 0) {
    Write-Host "No enabled users found with employeeType=Agent in the specified OU." -ForegroundColor Red
    exit
}

Write-Host "Found $($users.Count) enabled Agent users. Analyzing attributes..." -ForegroundColor Green

# Define attributes to exclude (empty/unwanted attributes)
$excludedAttributes = @(
    'isDeleted', 'HomedirRequired', 'HomeDrive', 'HomePage', 'HomePhone', 'info', 'Initials', 
    'LastKnownParent', 'lastLogoff', 'LockedOut', 'lockoutTime', 'LogonWorkstations', 
    'MNSLogonAccount', 'MobilePhone', 'ModifiedProperties', 'msDS-KeyCredentialLink', 
    'msExchRecipientSoftDeletedStatus', 'msExchTextMessagingState', 'msExchTransportRecipientSettingsFlags', 
    'msExchUserAccountControl', 'msRTCSIP-UserPolicies', 'Organization', 'OtherName', 
    'PasswordNeverExpires', 'PasswordNotRequired', 'personalTitle', 'physicalDeliveryOfficeName', 
    'POBox', 'postalAddress', 'PostalCode', 'PrincipalsAllowedToDelegateToAccount', 'ProfilePath', 
    'ProtectedFromAccidentalDeletion', 'protocolSettings', 'sDRightsEffective', 'ServicePrincipalNames', 
    'SmartcardLogonRequired', 'st', 'State', 'TrustedForDelegation', 'TrustedToAuthForDelegation', 
    'UseDESKeyOnly', 'AccountLockoutTime', 'AccountNotDelegated', 'AddedProperties', 
    'AllowReversiblePasswordEncryption', 'AuthenticationPolicy', 'AuthenticationPolicySilo', 
    'AccountExpirationDate', 'CannotChangePassword', 'Deleted', 'delivContLength'
)

# Get all unique attributes that have values across all users
$allAttributes = @()
foreach ($user in $users) {
    $populatedAttributes = $user.PSObject.Properties | Where-Object {
        $_.Value -ne $null -and 
        $_.Value -ne "" -and 
        $_.Name -notlike "*Properties" -and
        $_.Name -ne "PropertyNames" -and
        $_.Name -ne "PropertyCount" -and
        $_.Name -notin $excludedAttributes
    } | Select-Object -ExpandProperty Name
    
    $allAttributes += $populatedAttributes
}

# Get unique attribute names and sort them
$uniqueAttributes = $allAttributes | Sort-Object -Unique

Write-Host "Found $($uniqueAttributes.Count) unique populated attributes across all users." -ForegroundColor Cyan

# Create export data with all populated attributes
$exportData = @()
foreach ($user in $users) {
    $userRow = [ordered]@{}
    
    foreach ($attribute in $uniqueAttributes) {
        $value = $user.$attribute
        
        # Handle special cases for better readability
        if ($attribute -eq "accountExpires") {
            # Handle accountExpires special case
            if ($value -eq 0 -or $value -eq 9223372036854775807) {
                $userRow[$attribute] = "Never"
            }
            else {
                try {
                    $userRow[$attribute] = [DateTime]::FromFileTime($value).ToString("yyyy-MM-dd HH:mm:ss")
                }
                catch {
                    $userRow[$attribute] = "Never"
                }
            }
        }
        elseif ($attribute -eq "pwdLastSet") {
            if ($value -gt 0) {
                $userRow[$attribute] = [DateTime]::FromFileTime($value).ToString("yyyy-MM-dd HH:mm:ss")
            }
            else {
                $userRow[$attribute] = "Never"
            }
        }
        elseif ($attribute -eq "lastLogon") {
            if ($value -gt 0) {
                $userRow[$attribute] = [DateTime]::FromFileTime($value).ToString("yyyy-MM-dd HH:mm:ss")
            }
            else {
                $userRow[$attribute] = "Unknown"
            }
        }
        elseif ($attribute -eq "lastLogonTimestamp") {
            if ($value -gt 0) {
                $userRow[$attribute] = [DateTime]::FromFileTime($value).ToString("yyyy-MM-dd HH:mm:ss")
            }
            else {
                $userRow[$attribute] = "Unknown"
            }
        }
        elseif ($value -is [DateTime]) {
            $userRow[$attribute] = $value.ToString("yyyy-MM-dd HH:mm:ss")
        }
        elseif ($value -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
            $userRow[$attribute] = ($value | ForEach-Object { $_.ToString() }) -join "; "
        }
        elseif ($value -is [Array]) {
            $userRow[$attribute] = $value -join "; "
        }
        elseif ($value -ne $null -and $value -ne "") {
            $userRow[$attribute] = $value.ToString()
        }
        else {
            $userRow[$attribute] = ""
        }
    }
    
    $exportData += [PSCustomObject]$userRow
}

# Export to Excel file
$outputPath = "C:\temp\Agent_Users_BE_Express_AllAttributes_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').xlsx"
Write-Host "Exporting to Excel..." -ForegroundColor Yellow

$exportData | Export-Excel -Path $outputPath -AutoSize -BoldTopRow -FreezeTopRow -WorksheetName "Agent Users"

Write-Host "Export completed successfully!" -ForegroundColor Green
Write-Host "File saved to: $outputPath" -ForegroundColor Yellow
Write-Host "Total enabled Agent users exported: $($exportData.Count)" -ForegroundColor Cyan
Write-Host "Total attributes exported: $($uniqueAttributes.Count)" -ForegroundColor Cyan

# Display the attributes that will be exported
Write-Host "`nAttributes being exported:" -ForegroundColor Magenta
$uniqueAttributes | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }