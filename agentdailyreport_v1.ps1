################################################################################################
# script: Report_AVI-DC_Users
# Function: Create csv report of all AVI-DC accounts with the following EmployeeType:
# "Employee","Secondary","Contractor","Test","Agent","AdminAccount"
# created by: Gert Geeraerts
################################################################################################


################################################################################################
# Load modules 
################################################################################################

Import-Module ActiveDirectory

################################################################################################
# AD and ADLDS servers
################################################################################################

$AdldsLocalServer = "bezavwseds1001.prg-dc.dhl.com"
$AdWsEnabledGc = (Get-ADDomainController -Discover -Service GlobalCatalog,adws -SiteName "bru-hub"       ).hostname[0] + ":3268"
$AdWsEnabledDc = (Get-ADDomainController -Discover -Service adws  -SiteName "bru-hub" ).hostname[0]


################################################################################################
# Code
################################################################################################

$Global:Log=@()

# Get the current Date
$date=get-date -format "yyyyMMdd"

# Create output file variable based on date
$filename= "C:\temp\allagentacc_v1"+$date+".csv" 

# Account employeetypes on scope
$AccountTypes=@("Agent")

# Searchscope of script
$Searchbase = "DC=avi-dc,DC=dhl,DC=com" 
#Searchbase = "OU=NonDHL,DC=avi-dc,DC=dhl,DC=com"
#$Searchbase = "OU=Users,OU=BRU,OU=BE,DC=avi-dc,DC=dhl,DC=com"

# Retrieve all users in scope
$AllUsers=get-aduser -filter * -Properties employeetype -searchbase $Searchbase | Where-object {($_.employeetype -in $AccountTypes) -and ($_.DistinguishedName -notmatch "OU=EDSdeprovisioned,DC=avi-dc,DC=dhl,DC=com") }

Foreach ($Account in $AllUsers)
    {
    $UserAccount=$Account.SamAccountName
    # Get AD info of all accounts in scope
    $ADInfo = Get-ADUser -Filter {samaccountname -eq $UserAccount} -Properties dpwncorpdivision,userPrincipalName,extensionAttribute4,extensionAttribute12,extensionAttribute13,Samaccountname,sn,GivenName,employeeType,Manager,mail,Displayname,Enabled,CanonicalName,dpwncrestcode,co,country,company -Server $AdWsEnabledDc
    # Get ADLDS (EDS) info of all accounts in scope
    $ADLDSInfo = Get-ADObject -Filter {uid -eq $UserAccount} -Server $AdldsLocalServer -SearchBase "o=dhl.com" -properties CN,dhlAccountARP,dhlGID

   # Get manager's display name by querying AD
    $ManagerDisplayName = ""
    if ($ADInfo.Manager) {
        try {
            $ManagerInfo = Get-ADUser -Identity $ADInfo.Manager -Properties DisplayName -Server $AdWsEnabledDc
            $ManagerDisplayName = $ManagerInfo.DisplayName
        }
        catch {
            # If manager lookup fails, fall back to extracting username from DN
            if ($ADInfo.Manager -match "CN=([^,]+)") {
                $ManagerDisplayName = $matches[1]
            }
        }
    }

    $Global:Log += New-Object -Type PSObject -Property (
            @{
            "GID" = $ADLDSInfo.dhlGID
            "UID" = $ADInfo.Samaccountname
            "CanonicalName" = $ADInfo.CanonicalName
            "Name"  = $ADInfo.sn
            "FirstName" = $ADInfo.GivenName
            "ARP" = $ADLDSInfo.dhlAccountARP
            "Manager" = $ManagerDisplayName            
            "Displayname" = $ADInfo.Displayname
            "Email" = $ADInfo.Mail
            "EmployeeType" = $ADInfo.employeeType
            "ADEnabled" =  $ADInfo.Enabled
            "CrestCode" = $ADInfo.dpwncrestcode
            "DPWNCorpDivision" = $ADInfo.dpwncorpdivision
            "Country"= $ADInfo.Country
            "CO" = $ADInfo.co
            "Company"= $ADInfo.Company
            "ExtensionAttribute12" = $ADInfo.extensionAttribute12
            "ExtensionAttribute13" = $ADInfo.extensionAttribute13 
            "ExtensionAttribute4" = $ADInfo.extensionAttribute4 
            "UPN" = $ADInfo.userPrincipalName

            }
            )
    }


# Export to .csv file in c:\temp folder
$Global:Log | export-csv $filename -NoTypeInformation -Delimiter ";" -force

# Export to gridview on screen
$Global:Log | out-gridview