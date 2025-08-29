# PS_Scripts

Collection of PowerShell tools for Active Directory and Exchange-related tasks used for searching users, inspecting attributes, and checking password expiration.

## Overview

This repo contains small utilities that automate common AD lookups and checks used by administrators in hybrid environments (on-prem AD + Exchange/Exchange Online). Each script is standalone and intended to be run from a workstation with the ActiveDirectory PowerShell module available.

## Prerequisites

- Windows PowerShell or PowerShell 7+ (some AD cmdlets require Windows PowerShell)
- ActiveDirectory module (RSAT)
- Appropriate AD permissions for queries and updates
- Network access to your domain controllers / Global Catalog
- Credentials for Exchange-related lookups when needed

## Scripts

- Search_MultipleUsers_in_AD.ps1  
  - Purpose: Search multiple users (comma-separated) by display name. Prompts for search scope (entire directory via Global Catalog or a specific DC). Returns selected AD attributes (DisplayName, sAMAccountName, DistinguishedName, employeeType, employeeNumber, Manager, extensionAttribute13, mail). Intended for quick bulk lookups.

- Get_PasswordExpiration_date.ps1  
  - Purpose: Query a user's password expiration information. Accepts SAMAccountName, first name, or full name and reports PasswordLastSet, PasswordNeverExpires and calculates expiration where applicable. Includes retry flows for different name formats.

- AD_User_UPN_Management_Script_TEST-v1.ps1  
  - Purpose: Utility scaffolding for UPN / mailbox attribute management (test/dev version). Contains input helpers, header/colored output helpers and functions to find and manage users. Prompts for credentials to update mail attributes.

## Usage examples

Search multiple users (interactive prompt):
```powershell
.\Search_MultipleUsers_in_AD.ps1
# Follow prompts: choose scope, enter comma-separated names (e.g. "Jane Doe, J Smith")
```

Check password expiry:
```powershell
.\Get_PasswordExpiration_date.ps1
# Enter SAMAccountName, first name or "First Last" at the prompt
```

Run UPN/attribute management script (test):
```powershell
.\AD_User_UPN_Management_Script_TEST-v1.ps1
# Follow interactive prompts; supplies credential when requested
```

## Notes and explanations (useful AD attributes)

- AD Web Services (ADWS)  
  Used by many AD cmdlets to communicate with domain controllers. Scripts may discover DCs using `Get-ADDomainController -Discover -Service ADWS` or `-Service GlobalCatalog,ADWS`.

- AD LDS (Active Directory Lightweight Directory Services)  
  Use when your target directory data lives in an AD LDS instance (application-specific directories or testing) instead of AD DS.

- msExchRemoteRecipientType  
  Exchange attribute used in hybrid scenarios to indicate remote mailbox type. Example: `8` typically indicates an Equipment mailbox in Exchange Online. If the attribute is absent, the user likely has no remote mailbox.

- msRTCSIP-PrimaryUserAddress  
  Stores the SIP address for Skype for Business / Teams sign-in (format: `sip:user@domain`).

- proxyAddresses with an X.500 entry  
  An X.500 proxy (e.g. `x500:/o=ExchangeLabs/...`) commonly appears for migrated mailboxes or to preserve legacy addressing; it's normal in hybrid/migration scenarios.

## Caveats & best practices

- Run scripts with least privilege required for intended operations.
- Test on non-production accounts before making bulk modifications.
- When using `-Server` with AD cmdlets, ensure the correct port for Global Catalog (3268) if searching the GC.
- Input parsing: scripts expect comma-separated names; whitespace is trimmed.

## Contributing

- Open an issue or submit a PR with fixes/improvements.
- Keep changes minimal and document behavior changes in the script header.

## License

Add your preferred license file to the repo (e.g. MIT) or adjust internal distribution rules
