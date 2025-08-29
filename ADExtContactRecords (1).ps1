################################################################################################
# script: Report_AVI-DC_Users
# Function: Create csv report of all ExternalContacts
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

# Get the current Date
$date=get-date -format "yyyyMMdd"

# Account employeetypes on scope
$contacts=@("ExternalContacts")

# Searchscope of script
$Searchbase = "OU=ExternalContacts,OU=DE,DC=prg-dc,DC=dhl,DC=com"
#Searchbase = "DC=avi-dc,DC=dhl,DC=com" 
#Searchbase = "OU=NonDHL,DC=avi-dc,DC=dhl,DC=com"
#$Searchbase = "OU=Users,OU=BRU,OU=BE,DC=avi-dc,DC=dhl,DC=com"

Get-ADObject -Filter 'objectClass -eq "contact"' -properties *| select displayname, name, mail | Export-CSV -path "C:\temp\ADExtContactRecords_$date.csv"
