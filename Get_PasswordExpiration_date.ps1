# Import the Active Directory module 
Import-Module ActiveDirectory

do {
    # Prompt
    Write-Host "`nEnter the userID, first name, or full name (First Last) to check password expiration date" -ForegroundColor Yellow
    $username = Read-Host
    $user     = $null

    # First try by Identity
    try {
        $user = Get-ADUser -Identity $username `
                            -Property PasswordLastSet,PasswordNeverExpires,Enabled `
                            -ErrorAction Stop

        if (-not $user.Enabled) {
            throw "Account '$($user.SamAccountName)' is disabled."
        }

        Write-Host "Found enabled user by ID: $($user.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "$($_.Exception.Message) Trying by name lookup..." -ForegroundColor Yellow

        $nameParts = $username -split '\s+'

        if ($nameParts.Count -eq 2) {
            # Full name (only enabled)
            $firstName = $nameParts[0]
            $lastName  = $nameParts[1]
            try {
                $users = @( Get-ADUser `
                    -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName' -and Enabled -eq 'True'" `
                    -Property PasswordLastSet,PasswordNeverExpires,Department,Title,Office `
                    -ErrorAction Stop )

                if ($users.Count -eq 1) {
                    $user = $users[0]
                    Write-Host "Found user by name: $($user.Name)" -ForegroundColor Green
                }
                elseif ($users.Count -gt 1) {
                    Write-Host "`nMultiple users found with name '$firstName $lastName':" -ForegroundColor Yellow
                    Write-Host "Please choose from the following options:" -ForegroundColor Cyan

                    for ($i=0; $i -lt $users.Count; $i++) {
                        $u      = $users[$i]
                        $dept   = if ($u.Department) { $u.Department } else { "N/A" }
                        $title  = if ($u.Title)      { $u.Title }      else { "N/A" }
                        $office = if ($u.Office)     { $u.Office }     else { "N/A" }
                        Write-Host "[$($i+1)] $($u.Name) (ID: $($u.SamAccountName))" -ForegroundColor White
                        Write-Host "     Dept: $dept | Title: $title | Office: $office" -ForegroundColor Gray
                    }

                    Write-Host "[0] Enter specific User ID instead" -ForegroundColor Yellow
                    do {
                        $choice = Read-Host "`nEnter your choice (0-$($users.Count))"
                        if ($choice -eq '0') {
                            $specific = Read-Host "Enter the specific User ID (SAMAccountName)"
                            try {
                                $user = Get-ADUser -Identity $specific `
                                                  -Property PasswordLastSet,PasswordNeverExpires,Enabled `
                                                  -ErrorAction Stop
                                if (-not $user.Enabled) {
                                    throw "Account '$($user.SamAccountName)' is disabled."
                                }
                                Write-Host "Found enabled user by specific ID: $($user.Name)" -ForegroundColor Green
                                break
                            }
                            catch {
                                Write-Host "$($_.Exception.Message) Try again." -ForegroundColor Red
                                $choice = $null
                            }
                        }
                        elseif ([int]$choice -ge 1 -and [int]$choice -le $users.Count) {
                            $user = $users[[int]$choice - 1]
                            Write-Host "Selected user: $($user.Name)" -ForegroundColor Green
                            break
                        }
                        else {
                            Write-Host "Invalid choice. Enter 0-$($users.Count)." -ForegroundColor Red
                            $choice = $null
                        }
                    } while ($choice -eq $null)
                }
            }
            catch {
                Write-Host "Error searching by name: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        elseif ($nameParts.Count -eq 1) {
            # First-name only (only enabled)
            $firstName = $nameParts[0]
            Write-Host "Searching enabled users by first name: $firstName" -ForegroundColor Yellow
            try {
                $users = @( Get-ADUser `
                    -Filter "GivenName -eq '$firstName' -and Enabled -eq 'True'" `
                    -Property PasswordLastSet,PasswordNeverExpires,Department,Title,Office `
                    -ErrorAction Stop )

                if ($users.Count -eq 0) {
                    Write-Host "No enabled users found with first name '$firstName'." -ForegroundColor Red
                }
                elseif ($users.Count -eq 1) {
                    $user = $users[0]
                    Write-Host "Found user by first name: $($user.Name)" -ForegroundColor Green
                }
                else {
                    Write-Host "`nMultiple users found with first name '$firstName':" -ForegroundColor Yellow
                    Write-Host "Please choose from the following options:" -ForegroundColor Cyan

                    for ($i=0; $i -lt $users.Count; $i++) {
                        $u      = $users[$i]
                        $dept   = if ($u.Department) { $u.Department } else { "N/A" }
                        $title  = if ($u.Title)      { $u.Title }      else { "N/A" }
                        $office = if ($u.Office)     { $u.Office }     else { "N/A" }
                        Write-Host "[$($i+1)] $($u.Name) (ID: $($u.SamAccountName))" -ForegroundColor White
                        Write-Host "     Dept: $dept | Title: $title | Office: $office" -ForegroundColor Gray
                    }

                    Write-Host "[0] Enter specific User ID instead" -ForegroundColor Yellow
                    do {
                        $choice = Read-Host "`nEnter your choice (0-$($users.Count))"
                        if ($choice -eq '0') {
                            $specific = Read-Host "Enter the specific User ID (SAMAccountName)"
                            try {
                                $user = Get-ADUser -Identity $specific `
                                                  -Property PasswordLastSet,PasswordNeverExpires,Enabled `
                                                  -ErrorAction Stop
                                if (-not $user.Enabled) {
                                    throw "Account '$($user.SamAccountName)' is disabled."
                                }
                                Write-Host "Found enabled user by specific ID: $($user.Name)" -ForegroundColor Green
                                break
                            }
                            catch {
                                Write-Host "$($_.Exception.Message) Try again." -ForegroundColor Red
                                $choice = $null
                            }
                        }
                        elseif ([int]$choice -ge 1 -and [int]$choice -le $users.Count) {
                            $user = $users[[int]$choice - 1]
                            Write-Host "Selected user: $($user.Name)" -ForegroundColor Green
                            break
                        }
                        else {
                            Write-Host "Invalid choice. Enter 0-$($users.Count)." -ForegroundColor Red
                            $choice = $null
                        }
                    } while ($choice -eq $null)
                }
            }
            catch {
                Write-Host "Error searching by first name: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        elseif ($nameParts.Count -gt 2) {
            # Multi-part names (only enabled)
            Write-Host "Multiple names detected. Trying different combinations..." -ForegroundColor Yellow
            for ($i=1; $i -lt $nameParts.Count; $i++) {
                $firstName = $nameParts[0..($i-1)] -join ' '
                $lastName  = $nameParts[$i..($nameParts.Count-1)] -join ' '
                try {
                    $users = @( Get-ADUser `
                        -Filter "GivenName -eq '$firstName' -and Surname -eq '$lastName' -and Enabled -eq 'True'" `
                        -Property PasswordLastSet,PasswordNeverExpires,Department,Title,Office `
                        -ErrorAction Stop )

                    if ($users.Count -eq 1) {
                        $user = $users[0]
                        Write-Host "Found user by name combination: $firstName $lastName" -ForegroundColor Green
                        break
                    }
                    elseif ($users.Count -gt 1) {
                        Write-Host "`nMultiple users found with name '$firstName $lastName':" -ForegroundColor Yellow
                        Write-Host "Please choose from the following options:" -ForegroundColor Cyan
                        for ($j=0; $j -lt $users.Count; $j++) {
                            $u      = $users[$j]
                            $dept   = if ($u.Department) { $u.Department } else { "N/A" }
                            $title  = if ($u.Title)      { $u.Title }      else { "N/A" }
                            $office = if ($u.Office)     { $u.Office }     else { "N/A" }
                            Write-Host "[$($j+1)] $($u.Name) (ID: $($u.SamAccountName))" -ForegroundColor White
                            Write-Host "     Dept: $dept | Title: $title | Office: $office" -ForegroundColor Gray
                        }
                        Write-Host "[0] Enter specific User ID instead" -ForegroundColor Yellow
                        do {
                            $choice = Read-Host "`nEnter your choice (0-$($users.Count))"
                            if ($choice -eq '0') {
                                $specific = Read-Host "Enter the specific User ID (SAMAccountName)"
                                try {
                                    $user = Get-ADUser -Identity $specific `
                                                      -Property PasswordLastSet,PasswordNeverExpires,Enabled `
                                                      -ErrorAction Stop
                                    if (-not $user.Enabled) {
                                        throw "Account '$($user.SamAccountName)' is disabled."
                                    }
                                    Write-Host "Found enabled user by specific ID: $($user.Name)" -ForegroundColor Green
                                    break
                                }
                                catch {
                                    Write-Host "$($_.Exception.Message) Try again." -ForegroundColor Red
                                    $choice = $null
                                }
                            }
                            elseif ([int]$choice -ge 1 -and [int]$choice -le $users.Count) {
                                $user = $users[[int]$choice - 1]
                                Write-Host "Selected user: $($user.Name)" -ForegroundColor Green
                                break
                            }
                            else {
                                Write-Host "Invalid choice. Enter 0-$($users.Count)." -ForegroundColor Red
                                $choice = $null
                            }
                        } while ($choice -eq $null)
                        break
                    }
                }
                catch { }
            }
        }
        else {
            Write-Host "Please enter a userID, first name, or full name (First Last)." -ForegroundColor Yellow
        }
    } # end catch name-lookup

    # Display expiration info if we found $user
    if ($user) {
        try {
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

            if ($user.PasswordNeverExpires) {
                $passwordExpiresFormatted = "Never"
            }
            elseif ($user.PasswordLastSet) {
                $passwordExpires = $user.PasswordLastSet.AddDays($maxPasswordAge.Days)
                $passwordExpiresFormatted = $passwordExpires.ToString("yyyy-MM-dd HH:mm:ss")
                $daysUntilExpiration     = ($passwordExpires - (Get-Date)).Days
                if ($daysUntilExpiration -lt 0) {
                    $expirationStatus = "EXPIRED $([Math]::Abs($daysUntilExpiration)) days ago"
                    $statusColor      = "Red"
                }
                elseif ($daysUntilExpiration -le 7) {
                    $expirationStatus = "Expires in $daysUntilExpiration days (WARNING)"
                    $statusColor      = "Yellow"
                }
                else {
                    $expirationStatus = "Expires in $daysUntilExpiration days"
                    $statusColor      = "Green"
                }
            }
            else {
                $passwordExpiresFormatted = "Unknown (Password never set)"
                $expirationStatus         = "Unknown"
                $statusColor              = "Yellow"
            }

            Write-Host "`n=== Password Expiration Information ===" -ForegroundColor Cyan
            Write-Host "User:                   $($user.Name)"                     -ForegroundColor White
            Write-Host "SAM Account Name:       $($user.SamAccountName)"           -ForegroundColor White
            Write-Host "Password Last Set:      $($user.PasswordLastSet)"          -ForegroundColor White
            Write-Host "Password Never Expires: $($user.PasswordNeverExpires)"     -ForegroundColor White
            Write-Host "Password Expires:       $passwordExpiresFormatted"        -ForegroundColor White
            if (-not $user.PasswordNeverExpires -and $user.PasswordLastSet) {
                Write-Host "Status:                 $expirationStatus"               -ForegroundColor $statusColor
            }
            Write-Host "======================================="                 -ForegroundColor Cyan
        }
        catch {
            Write-Host "Error calculating password expiration: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "`nUser not found. Please check:" -ForegroundColor Red
        Write-Host "- User ID (SAMAccountName)"        -ForegroundColor Red
        Write-Host "- First name or full name format" -ForegroundColor Red
        Write-Host "- User exists in the domain"       -ForegroundColor Red
    }

    # Ask to continue or exit
    Write-Host "`nWhat next?" -ForegroundColor Cyan
    Write-Host "[1] Check another user" -ForegroundColor White
    Write-Host "[2] Exit"                -ForegroundColor White
    do {
        $next = Read-Host "Enter choice (1 or 2)"
    } while ($next -notin '1','2')

} while ($next -eq '1')

Write-Host "`nGoodbye!" -ForegroundColor Cyan
