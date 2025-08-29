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

# Prompt one time for the admin creds that can write "mail"
Write-Host "Enter credentials that have rights to update the E-mail field:" -ForegroundColor Cyan
$MailWriterCred = Get-Credential


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
        $displaylist     = foreach ($u in $uniqueUsers) {
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
        Get-ADGroup -Identity $GroupName -ErrorAction Stop
        return $true
    }
    catch {
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
        $isMember = Get-ADGroupMember -Identity $GroupName -ErrorAction Stop | Where-Object { $_.SamAccountName -eq $UserSamAccountName }
        
        if ($isMember) {
            Write-Host "User is already a member of group: $GroupName" -ForegroundColor Yellow
            return $true
        }
        
        # Add user to group
        Add-ADGroupMember -Identity $GroupName -Members $UserSamAccountName -ErrorAction Stop
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
        # 1) Get the user’s DN
        $userObj = Get-ADUser -Identity $UserSamAccountName `
                              -Properties DistinguishedName `
                              -ErrorAction Stop
        $dn = $userObj.DistinguishedName

        # 2) Build a proper user-name string for ADSI
        $netCred = $MailWriterCred.GetNetworkCredential()
        $userName = if ($netCred.Domain) {
            # if domain is set, use DOMAIN\user
            "$($netCred.Domain)\$($netCred.UserName)"
        } else {
            # else assume they entered UPN
            $MailWriterCred.UserName
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
    exit 1
}

Write-Host "User found: $($user.DisplayName) ($($user.SamAccountName))" -ForegroundColor Green
Write-Host ""

# Step 2: Get company name
# ==== New combined Step: Prompt for email, derive company, add to group & update mail ====

# Info text about the what happens whe nentering email address.
Write-Host "Note: Entering the external e-mail address will also add the user to the correspondant company security group in AD and update the external email address in the user properties" -ForegroundColor DarkMagenta

# 1) Ask for the external email address
$emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
$emailAddress = Get-ValidatedInput `
    -Prompt "Enter External Email Address" -ForegroundColor DarkYellow `
    -ValidationPattern $emailPattern `
    -ErrorMessage "Please enter a valid email (e.g. user@domain.com)."

# 2) Extract the company from the domain (between '@' and first '.')
$domain    = $emailAddress.Split('@')[1]
$companyRaw = $domain.Split('.')[0]
Write-Host "Detected company from email: $companyRaw" -ForegroundColor Cyan

# 3) Normalize & uppercase exactly like before
$normalized       = $companyRaw.Trim() `
                       -replace '[\s_]+' , '-' `
                       -replace '-+'    , '-'
$companyFormatted = $normalized.ToUpper()

# 4) Build & process the security group
$securityGroupName = "G-GAIT-B2B-$companyFormatted"
Write-Host "Target security group: $securityGroupName" -ForegroundColor Yellow

if (-not (Test-SecurityGroup -GroupName $securityGroupName)) {
    Write-Host "WARNING: Group '$securityGroupName' not found." -ForegroundColor Red
    if (Read-Host "Create it? (y/n)" -match '^[Yy]') {
        New-ADGroup -Name $securityGroupName `
                    -GroupCategory Security `
                    -GroupScope Global `
                    -Path ("CN=Users," + (Get-ADDomain).DistinguishedName) `
                    -ErrorAction Stop
        Write-Host "Created group: $securityGroupName" -ForegroundColor Green
    }
    else {
        Write-Host "Aborting—group doesn’t exist." -ForegroundColor Red
        exit 1
    }
}

Add-UserToSecurityGroup -UserSamAccountName $user.SamAccountName `
                        -GroupName         $securityGroupName

# 5) Finally, update the user’s mail attribute
Update-UserEmail -UserSamAccountName $user.SamAccountName `
                 -EmailAddress      $emailAddress

Write-Host "Done: user added to $securityGroupName and mail set to $emailAddress" -ForegroundColor Green
# ======================================================================================


<# Testing another alternative, this is the original, just in case i have to fall back to this one
# Step 3: Construct security group name and add user
$securityGroupName = "G-GAIT-B2B-$companyFormatted"
Write-Host "Target security group: $securityGroupName" -ForegroundColor Yellow

# Check if security group exists
if (-not (Test-SecurityGroup -GroupName $securityGroupName)) {
    Write-Host "WARNING: Security group '$securityGroupName' does not exist." -ForegroundColor Red
    $createGroup = Read-Host "Would you like to create this group? (y/n)"
    
    if ($createGroup.ToLower() -eq 'y' -or $createGroup.ToLower() -eq 'yes') {
        try {
            # Create the security group
            $groupPath = "CN=Users," + (Get-ADDomain).DistinguishedName  # You may want to change this path
            New-ADGroup -Name $securityGroupName -GroupCategory Security -GroupScope Global -Path $groupPath -ErrorAction Stop
            Write-Host "Security group created successfully: $securityGroupName" -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR: Failed to create security group - $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "Script terminated - Security group does not exist." -ForegroundColor Red
        exit 1
    }
}

# Add user to security group
$groupResult = Add-UserToSecurityGroup -UserSamAccountName $user.SamAccountName -GroupName $securityGroupName

if (-not $groupResult) {
    Write-Host "Failed to add user to security group. Continuing..." -ForegroundColor Yellow
}

Write-Host ""
#>


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
                try {
                    # Create the application group
                    $groupPath = "CN=Users," + (Get-ADDomain).DistinguishedName  # You may want to change this path
                    New-ADGroup -Name $appGroupName -GroupCategory Security -GroupScope Global -Path $groupPath -ErrorAction Stop
                    Write-Host "Application group created successfully: $appGroupName" -ForegroundColor Green
                }
                catch {
                    Write-Host "ERROR: Failed to create application group - $($_.Exception.Message)" -ForegroundColor Red
                    $appGroupResult = $false
                }
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

<# another alternative is present is step 2 i keep this original in case i have to fall back to it

# Step 5: Get external email address and update user
$emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2}$"
$emailAddress = Get-ValidatedInput -Prompt "Enter External Email Address" -ValidationPattern $emailPattern -ErrorMessage "Please enter a valid email address format (e.g. user@domain.com)."

$emailResult = Update-UserEmail -UserSamAccountName $user.SamAccountName -EmailAddress $emailAddress

#>

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

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")