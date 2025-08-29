# Interactive AD User Management Script
# Requirements: ActiveDirectory PowerShell module and appropriate permissions

# Import Active Directory module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Active Directory module loaded successfully." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Active Directory module not found or cannot be loaded." -ForegroundColor Red
    Write-Host "Please install RSAT tools or run this script on a domain controller." -ForegroundColor Yellow
    exit 1
}

# Prompt one time for the admin creds that can write "mail" and manage groups
Write-Host "Enter credentials that have rights to update E-mail field and manage security groups:" -ForegroundColor Cyan
Write-Host "(Note: Account needs 'Create Group Objects' permission in the target OU)" -ForegroundColor Gray
$AdminCred = Get-Credential

# Loop to add update more users 
$again = 'Y'

do{

# Function to validate user input
function Get-ValidatedInput {
    param(
        [string]$Prompt,
        [string]$ValidationPattern = ".*",
        [string]$ErrorMessage = "Invalid input. Please try again.",
        [string]$PromptColor = "White"
    )
    
    do {
        Write-Host $Prompt -ForegroundColor $PromptColor -NoNewline
        Write-Host ": " -NoNewline
        $username = Read-Host
        if ($username -match $ValidationPattern -and $username.Trim() -ne "") {
            return $username.Trim()
        }
        Write-Host $ErrorMessage -ForegroundColor Red
    } while ($true)
}

# Function to search for user in AD
function Find-ADUser {
    param([string]$SearchTerm)
    
    $foundUsers = @()
    
    # Method 1: exact by SamAccountName (returns DisplayName by default)
    try {
        $user = Get-ADUser -Identity $SearchTerm -ErrorAction Stop
        Write-Host "Found exact match by SamAccountName: $($user.DisplayName) ($($user.SamAccountName))" -ForegroundColor Green
        return $user
    } catch { }
    
    # Method 2: exact by UPN
    try {
        $user = Get-ADUser -Filter "UserPrincipalName -eq '$SearchTerm'" -ErrorAction Stop
        if ($user) {
            Write-Host "Found exact match by UPN: $($user.DisplayName) ($($user.SamAccountName))" -ForegroundColor Green
            return $user
        }
    } catch { }
    
    # Methods 3–6: fuzzy across DisplayName, GivenName, Surname, Name
    foreach ($prop in 'DisplayName','GivenName','Surname','Name') {
        try {
            $users = Get-ADUser `
                -Filter "$prop -like '*$SearchTerm*'" `
                -Properties DisplayName `
                -ErrorAction SilentlyContinue
            if ($users) { $foundUsers += $users }
        } catch { }
    }
    
    # Dedupe & sort
    $uniqueUsers = @(
        $foundUsers | Sort-Object SamAccountName -Unique
        )
    
    if ($uniqueUsers.Count -eq 0) {
        return $null
    }
    elseif ($uniqueUsers.Count -eq 1) {
        $u = $uniqueUsers[0]
        Write-Host "Found single match: $($u.DisplayName) ($($u.SamAccountName))" -ForegroundColor Green
        return $u
    }
    else {
        Write-Host ""
        Write-Host "Multiple users found matching '$SearchTerm':" -ForegroundColor Yellow
        Write-Host "================================================" -ForegroundColor Yellow
        
        $i = 1
        $displaylist = foreach ($u in $uniqueUsers) {
                $details = Get-ADUser -Identity $u.SamAccountName -Properties Displayname, Department
                [PSCustomObject]@{
                    Index       = $i++
                    DisplayName = $details.DisplayName
                    UserID      = $details.SamAccountName
                    Department  = $details.Department
                    
                }
        }

        # show results in a table format
        $table = $displaylist | Format-Table Index, DisplayName, UserID, Department -AutoSize | Out-String -Stream

        # Write each line in green 
        $table | ForEach-Object {Write-Host $_ -ForegroundColor Green}
        
        Write-Host ""
        $selection = Get-ValidatedInput `
            -Prompt "Select user by number (1-$($displaylist.Count))" `
            -ValidationPattern "^[1-$($displaylist.Count)]$" `
            -ErrorMessage "Please enter a number between 1 and $($displaylist.Count)"
        
        return $uniqueUsers[$selection - 1]
    }
}

# Function to check if security group exists
function Test-SecurityGroup {
    param([string]$GroupName)
    
    try {
        # First try searching in the specific GAIT Groups OU
        $gaitGroupsOU = "OU=Groups,OU=GAIT,OU=Global,DC=avi-dc,DC=dhl,DC=com"
        $group = Get-ADGroup -Identity $GroupName -SearchBase $gaitGroupsOU -ErrorAction Stop
        return $true
    }
    catch {
        try {
            # Fallback: search entire domain
            $group = Get-ADGroup -Identity $GroupName -ErrorAction Stop
            return $true
        }
        catch {
            return $false
        }
    }
}

# Function to create security group with admin credentials
function New-SecurityGroup {
    param(
        [string]$GroupName
    )
    
    try {
        Write-Host "Attempting to create group with admin credentials..." -ForegroundColor Yellow
        Write-Host "Group: $GroupName" -ForegroundColor Yellow
        
        # Use the correct GAIT Groups OU path
        $gaitGroupsOU = "OU=Groups,OU=GAIT,OU=Global,DC=avi-dc,DC=dhl,DC=com"
        Write-Host "Target OU: $gaitGroupsOU" -ForegroundColor Yellow
        
        # Try method 1: Create in the correct GAIT Groups OU
        try {
            New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $gaitGroupsOU -Credential $AdminCred -ErrorAction Stop
            Write-Host "Security group created successfully in GAIT Groups OU: $GroupName" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "Method 1 (GAIT Groups OU) failed: $($_.Exception.Message)" -ForegroundColor Red
            
            # Try method 2: Alternative Groups OU
            Write-Host "Trying alternative Groups OU..." -ForegroundColor Yellow
            $alternativeGroupsOU = "OU=Groups,OU=Global,DC=avi-dc,DC=dhl,DC=com"
            Write-Host "Alternative OU: $alternativeGroupsOU" -ForegroundColor Yellow
            
            try {
                New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $alternativeGroupsOU -Credential $AdminCred -ErrorAction Stop
                Write-Host "Security group created successfully in alternative Groups OU: $GroupName" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "Method 2 (Alternative Groups OU) failed: $($_.Exception.Message)" -ForegroundColor Red
                
                # Try method 3: Default Users container (last resort)
                Write-Host "Trying default Users container..." -ForegroundColor Yellow
                $defaultPath = "CN=Users,DC=avi-dc,DC=dhl,DC=com"
                Write-Host "Default path: $defaultPath" -ForegroundColor Yellow
                
                try {
                    New-ADGroup -Name $GroupName -GroupCategory Security -GroupScope Global -Path $defaultPath -Credential $AdminCred -ErrorAction Stop
                    Write-Host "Security group created successfully in Users container: $GroupName" -ForegroundColor Green
                    return $true
                }
                catch {
                    Write-Host "Method 3 (Users container) failed: $($_.Exception.Message)" -ForegroundColor Red
                    
                    # Method 4: Show current permissions and suggest manual creation
                    Write-Host "All automatic creation methods failed." -ForegroundColor Red
                    Write-Host "Current user context: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -ForegroundColor Yellow
                    Write-Host "Admin credential user: $($AdminCred.UserName)" -ForegroundColor Yellow
                    Write-Host "" -ForegroundColor Yellow
                    Write-Host "SUGGESTION: Please manually create the group '$GroupName' in the GAIT Groups OU and run the script again." -ForegroundColor Cyan
                    Write-Host "OR: Contact your AD administrator to grant group creation permissions." -ForegroundColor Cyan
                    
                    return $false
                }
            }
        }
    }
    catch {
        Write-Host "ERROR: Unexpected error in group creation - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to add user to security group
function Add-UserToSecurityGroup {
    param(
        [string]$UserSamAccountName,
        [string]$GroupName
    )
    
    try {
        # Check if user is already a member
        $isMember = Get-ADGroupMember -Identity $GroupName -ErrorAction Stop -Credential $AdminCred | Where-Object { $_.SamAccountName -eq $UserSamAccountName }
        
        if ($isMember) {
            Write-Host "User is already a member of group: $GroupName" -ForegroundColor Yellow
            return $true
        }
        
        # Add user to group using admin credentials
        Add-ADGroupMember -Identity $GroupName -Members $UserSamAccountName -Credential $AdminCred -ErrorAction Stop
        Write-Host "Successfully added user to group: $GroupName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "ERROR: Failed to add user to group $GroupName - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to update user's email address
function Update-UserEmail {
    param(
      [string]$UserSamAccountName,
      [string]$EmailAddress
    )

    try {
        # 1) Get the user's DN
        $userObj = Get-ADUser -Identity $UserSamAccountName `
                              -Properties DistinguishedName `
                              -ErrorAction Stop
        $dn = $userObj.DistinguishedName

        # 2) Build a proper user-name string for ADSI
        $netCred = $AdminCred.GetNetworkCredential()
        $userName = if ($netCred.Domain) {
            # if domain is set, use DOMAIN\user
            "$($netCred.Domain)\$($netCred.UserName)"
        } else {
            # else assume they entered UPN
            $AdminCred.UserName
        }

        # 3) Bind with credentials in the constructor
        $entry = New-Object System.DirectoryServices.DirectoryEntry(
            "LDAP://$dn",
            $userName,
            $netCred.Password,
            [System.DirectoryServices.AuthenticationTypes]::Secure
        )

        # 4) Update the mail attribute
        $entry.Properties['mail'].Value = $EmailAddress

        # 5) Commit
        $entry.CommitChanges()

        Write-Host "Successfully updated E-mail to: $EmailAddress" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "ERROR: Failed to update E-mail field - $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}


# Main script execution
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  Interactive AD User Management Script" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Get User ID and search for user
$searchTerm = Get-ValidatedInput -Prompt "Enter User Search Term (First Name, Full Name, or User ID)" -PromptColor "DarkYellow"

Write-Host "Searching for user with term: '$searchTerm'..." -ForegroundColor Yellow
$user = Find-ADUser -SearchTerm $searchTerm

if ($null -eq $user) {
    Write-Host "ERROR: User not found in Active Directory." -ForegroundColor Red
    continue
}

Write-Host "User found: $($user.DisplayName) ($($user.SamAccountName))" -ForegroundColor Green
Write-Host ""

# Step 2: Get company name
$company = Get-ValidatedInput -Prompt "Enter Company Name (spaces will be converted to hyphens)" -ValidationPattern "^[a-zA-Z0-9\s_-]+$" -ErrorMessage "Company name should contain only letters, numbers, spaces, hyphens, and underscores."

# Step 2.1a: Normalize separators
$normalized = $company.Trim() `
    -replace '[\s_]+' , '-' `
    -replace '-+'    , '-'

# Step 2.1b: Upper-case it
$companyFormatted = $normalized.ToUpper()

Write-Host "Company name formatted as: $companyFormatted" -ForegroundColor Cyan

# Step 3: Construct security group name and add user
$securityGroupName = "G-GAIT-B2B-$companyFormatted"
Write-Host "Target security group: $securityGroupName" -ForegroundColor Yellow

# Check if security group exists
if (-not (Test-SecurityGroup -GroupName $securityGroupName)) {
    Write-Host "WARNING: Security group '$securityGroupName' does not exist." -ForegroundColor Red
    $createGroup = Read-Host "Would you like to create this group? (y/n)"
    
    if ($createGroup.ToLower() -eq 'y' -or $createGroup.ToLower() -eq 'yes') {
        $groupCreated = New-SecurityGroup -GroupName $securityGroupName
        
        if (-not $groupCreated) {
            Write-Host "Script terminated - Could not create security group." -ForegroundColor Red
            continue
        }
    }
    else {
        Write-Host "Script terminated - Security group does not exist." -ForegroundColor Red
        continue
    }
}

# Add user to security group
$groupResult = Add-UserToSecurityGroup -UserSamAccountName $user.SamAccountName -GroupName $securityGroupName

if (-not $groupResult) {
    Write-Host "Failed to add user to security group. Continuing..." -ForegroundColor Yellow
}

Write-Host ""

# Step 4: Get application access requirements (with loop for multiple apps)
$appGroupResults = @()
$appGroupNames = @()
$appAccesses = @()

do {
    $appAccess = Get-ValidatedInput -Prompt "Enter Application Name for access (or type 'none' if no app access needed)"
    
    if ($appAccess.ToLower() -ne 'none') {
        # Format application name: replace spaces with hyphens and convert to uppercase
        $appFormatted = $appAccess.Trim() -replace '\s+', '-' -replace '_', '-'
        $appFormatted = $appFormatted.ToUpper()
        
        Write-Host "Application name formatted as: $appFormatted" -ForegroundColor Cyan
        
        $appGroupName = "G-GAIT-APP-$appFormatted"
        Write-Host "Target application group: $appGroupName" -ForegroundColor Yellow
        
        $appGroupResult = $true
        
        # Check if application group exists
        if (-not (Test-SecurityGroup -GroupName $appGroupName)) {
            Write-Host "WARNING: Application group '$appGroupName' does not exist." -ForegroundColor Red
            $createAppGroup = Read-Host "Would you like to create this group? (y/n)"
            
            if ($createAppGroup.ToLower() -eq 'y' -or $createAppGroup.ToLower() -eq 'yes') {
                $appGroupResult = New-SecurityGroup -GroupName $appGroupName
            }
            else {
                Write-Host "Skipping application group assignment for $appFormatted..." -ForegroundColor Yellow
                $appGroupResult = $false
            }
        }
        
        # Add user to application group if group exists or was created
        if ($appGroupResult) {
            $appGroupResult = Add-UserToSecurityGroup -UserSamAccountName $user.SamAccountName -GroupName $appGroupName
            
            if (-not $appGroupResult) {
                Write-Host "Failed to add user to application group $appGroupName. Continuing..." -ForegroundColor Yellow
            }
        }
        
        # Store results for summary
        $appGroupResults += $appGroupResult
        $appGroupNames += $appGroupName
        $appAccesses += $appFormatted
        
        # Ask if user wants to add another application
        $addAnotherApp = Read-Host "Do you want to add access to another application? (y/n)"
        $continueLoop = ($addAnotherApp.ToLower() -eq 'y' -or $addAnotherApp.ToLower() -eq 'yes')
    }
    else {
        Write-Host "No application access requested." -ForegroundColor Gray
        $continueLoop = $false
    }
    
} while ($continueLoop)

Write-Host ""

# Step 5: Get external email address and update user
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
$emailAddress = Get-ValidatedInput -Prompt "Enter External Email Address" -ValidationPattern $emailPattern -ErrorMessage "Please enter a valid email address format (e.g., user@domain.com)."

$emailResult = Update-UserEmail -UserSamAccountName $user.SamAccountName -EmailAddress $emailAddress

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "              SUMMARY" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "User: $($user.DisplayName) ($($user.SamAccountName))" -ForegroundColor White
Write-Host "Company: $companyFormatted" -ForegroundColor White
Write-Host "Security Group: $securityGroupName" -ForegroundColor White

if ($appAccesses.Count -gt 0) {
    Write-Host "Application Access:" -ForegroundColor White
    for ($i = 0; $i -lt $appAccesses.Count; $i++) {
        $status = if ($appGroupResults[$i]) { "✓" } else { "✗" }
        Write-Host "  $status $($appAccesses[$i]) → $($appGroupNames[$i])" -ForegroundColor White
    }
}
else {
    Write-Host "Application Access: None requested" -ForegroundColor Gray
}

Write-Host "Email Address: $emailAddress" -ForegroundColor White
Write-Host ""

# Check overall success
$allAppGroupsSuccessful = $appGroupResults.Count -eq 0 -or ($appGroupResults | Where-Object { $_ -eq $false }).Count -eq 0

if ($groupResult -and $allAppGroupsSuccessful -and $emailResult) {
    Write-Host "Script completed successfully!" -ForegroundColor Green
}
else {
    Write-Host "Script completed with some errors. Please review the output above." -ForegroundColor Yellow
}

 # Ask if another user needs to be updated 
 $again = Read-Host "Do you want to update another user? (Y/N)"
 $again = $again.ToUpper().Trim()
} while ($again -eq 'Y')

Write-Host "`nGoodbye!" -ForegroundColor Green
# Press any key to exit
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")