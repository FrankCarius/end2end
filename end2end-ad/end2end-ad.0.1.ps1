#
# END2END AD
#
# Dieses Skript schreibt alle 5 Sekunden ein vorbestimmtes AD-Objekt auf einem DC und liest von allen anderen GCs den Wert aus.
# Damit kann zum einen die Latenzzeit und eine ausbleibende Replikation ermittelt werden.
#
# Not using any additional modules. Pure Powershell 2
#
# Using Windows PowerShell to list domain controllers in each of the domains within an ADS forest
# http://www.techietips.net/Using-Windows-PowerShell-list-domain-controllers-each-domains-within-ADS-forest
#
# PowerShell : How do I find all DCs in my forest ?
# http://www.shariqsheikh.com/blog/index.php/201102/powershell-how-do-i-find-all-dcs-in-my-forest/
#
# PowerShell Commands to list domain controllers in Domain
# http://techibee.com/active-directory/powershell-commands-to-list-domain-controllers-in-domain/160
#
# List of All Domain Controllers in Your Domain
# http://bsonposh.com/archives/1043
#[DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().DomainControllers

param (
	[string]$testobject = "" ,  # try cn=object,ou=orgunit,dc=domain,dc=tld
	[string]$testdc = "",       # add the DNSname, short name or IP-Address of the DC to write to
	[int]$maxdelay = 500,     # maximum latency to acccept
	[int]$idletime = 5,     # maximum latency to acccept
	[string]$smtpserver = "localhost",  # Mailserver to send alerts to
	[string]$smtpfrom = "end2endad",   # SMTP from Address
	[string]$smtpto = "heldesk@company.tld"  # SMTP Recipient
)

Set-PSDebug -strict
$ErrorActionPreference = "Continue";

write-host "end2end-ad: Start"

$myForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

# global catalogs
$gclist = $myForest.GlobalCatalogs   # every GC has property name (fqdn), sitename, IP-Address, CurrentTime, HighestCommittedUsn
# all DCs by parsing though all sites
# $myforest.Sites | % { $_.Servers } | Select Name, Domain  | out-file mydcs.txt

do {
	[object[]]$result=@()
	foreach ($gc in $gclist){
		$adsipath = "GC://"+$dc.name+"/"+$testobject
		write-host" Querying $adsipath"
		$adsiobject = [ADSI]$adsipath
		write-host "ADSI:Description" $adsiobject.description
	}
	
	write-host "Sleeping $idletim Seconds"
	start-sleep -seconds $idletime
} until ($false)

Vergleiche DC-Zeit mit aktueller Get-Date zeit da abfrage "länger "dauern kann


#Ausgabe als Out-GRIDView oder csv ode rtext mit
#DCName, CurrentTime, delta
#
#
#
#