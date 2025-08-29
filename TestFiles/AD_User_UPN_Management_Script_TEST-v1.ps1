<#
.SYNOPSIS
    Interactive AD User Management: search users, derive+create B2B group, assign, update mail.

.PARAMETER SearchTerm
    First name, full name, or SamAccountName to search for.

.PARAMETER NonInteractive
    Skip all UI (for automation).

.EXAMPLE
    .\AD-Manager.ps1 -SearchTerm "jacob"
.EXAMPLE
    .\AD-Manager.ps1 -SearchTerm "jdoe" -NonInteractive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SearchTerm,

    [switch]$NonInteractive
)

# ─── Load AD module & prompt creds ───────────────────────────────────────────────
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "AD module loaded." -ForegroundColor Green
} catch {
    Write-Host "ERROR: AD module not available." -ForegroundColor Red
    exit 1
}

Write-Host "Enter credentials that can update the mail attribute:" -ForegroundColor Cyan
$MailWriterCred = Get-Credential

# ─── Color variables ─────────────────────────────────────────────────────────────
$ColorHeader    = 'Cyan'
$ColorInfo      = 'Green'
$ColorWarning   = 'Yellow'
$ColorError     = 'Red'
$ColorTableText = 'Magenta'

function Write-Header {
    param($Text)
    Write-Host $Text -ForegroundColor $ColorHeader
}
function Write-Info {
    param($Text)
    Write-Host $Text -ForegroundColor $ColorInfo
}
function Write-WarnUI {
    param($Text)
    Write-Host $Text -ForegroundColor $ColorWarning
}
function Write-ErrorUI {
    param($Text)
    Write-Host $Text -ForegroundColor $ColorError
}

# ─── Input helper ────────────────────────────────────────────────────────────────
function Get-ValidatedInput {
    param(
        [string]$Prompt,
        [string]$ValidationPattern = '.*',
        [string]$ErrorMessage     = 'Invalid input.',
        [string]$PromptColor      = 'White'
    )
    do {
        Write-Host -NoNewline -ForegroundColor $PromptColor "${Prompt}: "
        $userResponse = Read-Host
        if ($userResponse -match $ValidationPattern -and $userResponse.Trim()) {
            return $userResponse.Trim()
        }
        Write-Host $ErrorMessage -ForegroundColor $ColorError
    } while ($true)
}

# ─── 1) Lookup AD user(s) ────────────────────────────────────────────────────────
function Get-UserMatches {
    param([string]$Term)
    $found = @()
    try { return ,(Get-ADUser -Identity $Term -ErrorAction Stop) } catch {}
    try {
        $u = Get-ADUser -Filter "UserPrincipalName -eq '$Term'" -ErrorAction Stop
        if ($u) { return ,$u }
    } catch {}
    foreach ($prop in 'DisplayName','GivenName','Surname','Name') {
        try {
            $users = Get-ADUser -Filter "$prop -like '*$Term*'" -Properties DisplayName
            if ($users) { $found += $users }
        } catch {}
    }
    return $found | Sort-Object SamAccountName -Unique
}

# ─── 2) Show table & pick ────────────────────────────────────────────────────────
function Show-UserSelectionTable {
    param([array]$Users)
    $i = 1
    $table = foreach ($u in $Users) {
        $d = Get-ADUser -Identity $u.SamAccountName -Properties GivenName, DisplayName,Department
        [PSCustomObject]@{
            Index       = $i++
            FullName    = $d.GivenName
            DisplayName = $d.DisplayName
            UserID      = $d.SamAccountName
            Department  = $d.Department
            
        }
    }
    $lines = $table |
        Format-Table Index, FullName, DisplayName, UserID, Department, Title -AutoSize |
        Out-String -Stream
    $lines | ForEach-Object { Write-Host $_ -ForegroundColor $ColorTableText }
    return $table.Count
}

# ─── 3) Get or create B2B group ──────────────────────────────────────────────────
function Get-CompanyGroup {
    param([string]$CompanyRaw)
    $norm      = $CompanyRaw.Trim() -replace '[\s_]+' , '-' -replace '-+' , '-'
    $company   = $norm.ToUpper()
    $groupName = "G-GAIT-B2B-$company"

    if (-not (Get-ADGroup -Identity $groupName -ErrorAction SilentlyContinue)) {
        if ($NonInteractive) { throw "Group $groupName not found." }
        Write-WarnUI "Group '$groupName' not found."
        $ans = Read-Host "Create it? (Y/n)"
        if ($ans.Trim().ToLower() -notin 'n','no') {
            try {
                $path = "CN=Users," + (Get-ADDomain).DistinguishedName
                New-ADGroup -Name $groupName -GroupCategory Security -GroupScope Global `
                  -Path $path -ErrorAction Stop
                Write-Info "Created group $groupName"
            } catch {
                throw "Could not create group: $_"
            }
        } else {
            throw "Aborted: group missing."
        }
    }

    return Get-ADGroup -Identity $groupName
}

# ─── 4) Update mail via ADSI ────────────────────────────────────────────────────
function Update-UserEmail {
    param([string]$SamAccountName, [string]$EmailAddress)
    $u  = Get-ADUser -Identity $SamAccountName -Properties DistinguishedName
    $dn = $u.DistinguishedName

    $net  = $MailWriterCred.GetNetworkCredential()
    $user = if ($net.Domain) { "$($net.Domain)\$($net.UserName)" } else { $MailWriterCred.UserName }

    $entry = New-Object System.DirectoryServices.DirectoryEntry(
        "LDAP://$dn", $user, $net.Password,
        [System.DirectoryServices.AuthenticationTypes]::Secure
    )
    $entry.Properties['mail'].Value = $EmailAddress
    $entry.CommitChanges()
    Write-Info "Email updated to $EmailAddress"
}

# ─── 5) Add to security group ────────────────────────────────────────────────
function Add-UserToSecurityGroup {
    param([string]$UserSamAccountName, [string]$GroupName)
    try {
        $members = Get-ADGroupMember -Identity $GroupName -Recursive
        if (-not ($members | Where-Object { $_.SamAccountName -eq $UserSamAccountName })) {
            Add-ADGroupMember -Identity $GroupName -Members $UserSamAccountName -ErrorAction Stop
            Write-Info "Added user to $GroupName"
        } else {
            Write-WarnUI "User already in $GroupName"
        }
    } catch {
        Write-ErrorUI ("Error adding to {0}: {1} -f $GroupName, $_")
    }
}

# ─── Main flow ───────────────────────────────────────────────────────────────────
try {
    Write-Header "`n=== AD User Manager ===`n"

    # 1) lookup
    $userMatches = Get-UserMatches -Term $SearchTerm
    if (-not $userMatches) {
        Write-ErrorUI "No users found for '$SearchTerm'."
        exit 1
    }

    # 2) select
    if ($userMatches.Count -eq 1 -or $NonInteractive) {
        $chosen = $userMatches[0]
        Write-Info "Selected $($chosen.DisplayName) ($($chosen.SamAccountName))"
    } else {
        Write-WarnUI "Multiple matches:"
        $count = Show-UserSelectionTable -Users $userMatches
        $sel   = Get-ValidatedInput -Prompt "Select user by number (1-$count)" `
                 -ValidationPattern "^[1-$count]$" -ErrorMessage "Enter 1 to $count"
        $chosen = $userMatches[$sel - 1]
    }

    # 3) email → company
    Write-Header "`nNote: your email domain will add user to company group.`n"
    $email      = Get-ValidatedInput -Prompt "Enter External Email" `
                   -ValidationPattern '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$' `
                   -ErrorMessage "Invalid email format"
    $companyRaw = ($email.Split('@')[1]).Split('.')[0]

    # 4) ensure + assign B2B
    $b2bGroup = Get-CompanyGroup -CompanyRaw $companyRaw
    Add-UserToSecurityGroup -UserSamAccountName $chosen.SamAccountName -GroupName $b2bGroup.Name

    # 5) update mail
    Update-UserEmail -SamAccountName $chosen.SamAccountName -EmailAddress $email

    # 6) application‐access loop (unchanged)
    $appGroupResults = @(); $appGroupNames = @(); $appAccesses = @()
    do {
        $appAccess = Get-ValidatedInput -Prompt "Enter Application name (or 'none')" `
                     -ValidationPattern '\S+' -ErrorMessage "Cannot be empty"
        if ($appAccess.ToLower() -eq 'none') { break }

        $appFmt   = ($appAccess.Trim() -replace '\s+','-').ToUpper()
        $appGroup = "G-GAIT-APP-$appFmt"
        Write-Host "Target: $appGroup" -ForegroundColor $ColorWarning

        if (-not (Get-ADGroup -Identity $appGroup -ErrorAction SilentlyContinue)) {
            Write-WarnUI "Missing: $appGroup"
            $c = Read-Host "Create? (Y/n)"
            if ($c.Trim().ToLower() -notin 'n','no') {
                New-ADGroup -Name $appGroup -GroupCategory Security -GroupScope Global `
                  -Path ("CN=Users,"+(Get-ADDomain).DistinguishedName) -ErrorAction Stop
                Write-Info "Created $appGroup"
            } else {
                Write-WarnUI "Skipped $appGroup"
                continue
            }
        }

        Add-UserToSecurityGroup -UserSamAccountName $chosen.SamAccountName -GroupName $appGroup
        $appGroupResults += $true
        $appGroupNames   += $appGroup
        $appAccesses     += $appFmt

        $more = Read-Host "Add another? (Y/n)"
    } while ($more.Trim().ToLower() -in 'y','yes')

    # 7) summary
    Write-Header "`n=== SUMMARY ===`n"
    Write-Host "User:  $($chosen.DisplayName) ($($chosen.SamAccountName))"
    Write-Host "Email: $email"
    Write-Host "Group: $($b2bGroup.Name)"
    if ($appAccesses.Count) {
        Write-Host "`nApp Access:"
        for ($i = 0; $i -lt $appAccesses.Count; $i++) {
            if ($appGroupResults[$i]) { $mark = '✓' } else { $mark = '✗' }
            Write-Host "  $mark $($appAccesses[$i]) → $($appGroupNames[$i])"
        }
    }

    Write-Info "`nAll done!"
}
catch {
    Write-ErrorUI "ERROR: $_"
    exit 1
}
