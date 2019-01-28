# end2end-ews
#
# Dieses Script versucht den Posteingang und Kalender eines Postfach 
# Über EWS Impersonation zu Öffnen 
#
# Voraussettungen 
# - Exchange 2007 SP1+ 
# - NET 3.5 
# - "Microsoft.Exchange.WebServices.dll
#
# Getting started with the EWS Managed API # http://msdn.microsoft.com/en-us/library/dd633626(v=exchg.80).aspx
# 20180306 inital Version based on test-ews 1.2 
# 20180404 Fix zu BEServer-Ausgaben
# 20180504 PRTG Push Ausgabe addiert, FE Ausgabe addiert
# 20180505 PRTGPush mit Try/Catch abgefangen
#
param(
	[string]$MailboxSMTP = "frank.carius@netatwork.de",	# must be primary SMTP for impersonation
	[string]$Username    = "frank.carius@netatwork.de",				# use default credentials, if empty
	[string]$Domain      = "",				# Domain of the authentication user
	[double]$interval = 10,   # seconds to measure one interval
	[string]$Password    = ((get-credential -username $Username -message "Password for $Domain\$Username").getnetworkcredential().password),  # password, if $username is set
	[string]$Serviceurl  = "",						# Use Autodiscover, if empty https://exchange.msxfaq.local/EWS/Exchange.asmx",
	[string]$archivcsv = ".\end2end-ews.csv",
	[string]$prtgpushurl = "http://192.168.178.11:5050/end2end-ews",
	[switch]$useImpersonation = $false,						# forces impersonation 
	[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\nProgram Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll", # EWS DLL
	#[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll", # 
	[switch]$EWSTrace,				 				# enable tracing of EWS to STDOUT
	[switch]$Verbose								# enable verbose output
)

if ($Verbose) {
	$VerbosePreference = "continue" 
}
Write-verbose "Test-EWS: Start"
Write-Verbose "Loading EWS DLL"
Write-Verbose "Username: $username"
Write-Verbose "Domain  : $domain"
Write-Verbose "password: <not visible>"

[void][Reflection.Assembly]::LoadFile($dllpath)
Write-Verbose "Creating EWS Service Class"
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService('Exchange2007_SP1')
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
if ($ewsTrace) {
	Write-Verbose " Tracing: enabled"
	$service.TraceEnabled = $true
	#Tracing EWS requests http://msdn.microsoft.com/en-us/library/dd633676(v=exchg.80).aspx
}
# --------------------------- Credentials and Impersonation --------------------------- 
if ($username -eq "") {  Write-Verbose "Credentials: UseDefaultCredentials"
	$service.UseDefaultCredentials = $true
}
else  {
	Write-Verbose "Credentials: use alternate Credentials"
	$service.UseDefaultCredentials = $false  
	$service.Credentials = New-Object System.Net.NetworkCredential($username, $password, $domain) 
} 
if ($useImpersonation) {
	Write-Verbose "Credentials: use impersonation"
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxSMTP);

	#Configuring Exchange Impersonation (Exchange Web Services)  #http://msdn.microsoft.com/en-us/library/bb204095(v=exchg.80).aspx
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SID,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.PrincipalName,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SmtpAddress,wert)
}
# --------------------------- ServiceURI --------------------------- 

write-verbose "Checking Autodiscover for Mailbox"
try{
	if ($serviceurl -eq "") {  
		Write-Verbose "ServiceURL: Using Autodiscover for $mailboxSMTP"
		$service.AutodiscoverUrl($mailboxSMTP,{$true})
	}
	else {
		Write-Verbose "ServiceURL: using specified $serviceurl"
		$service.Url = New-Object System.Uri($serviceurl) 
	} 
	write-verbose -Message ("ServiceURI="+$service.Url.AbsoluteUri)
}
Catch {
	Write-Error "Unable to determinte AutoDisover URL with Error `r`n $_.Exception.Message"
	exit 1
}
	# --------------------------- Connect --------------------------- 

# Optional. Query Autodiscover for testing and retrieving Archive mailboxes etc
# write-verbose "Checking Autodiscover for Archive Mailbox"
#$autod = new-object Microsoft.Exchange.WebServices.Autodiscover.autodiscoverservice($mailboxSMTP.split("@")[1])
$autod = new-object Microsoft.Exchange.WebServices.Autodiscover.autodiscoverservice
$autod.EnableScpLookup = $false
#$autod.Credentials =
$autod.RedirectionUrlValidationCallback = {$true}   # ignore Redirect requests

Write-Verbose "Connecting to Inbox for SMTP $($mailboxSMTP)"
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
write-host "Number or Unread Messages : " $inbox.UnreadCount 

#ä Initializing Result Hashtable for storing values and generate output later
$statistics = New-Object PSObject -Property @{
	timestamp   = (get-date)
	durationavg = [long]9999
	durationmax = [long]0
	durationmin = [long]9999
	beserver    = [string]""
}
# Initialize Counters for Summary
[double]$durationsum=0
[double]$count=0
$result = $inbox.finditems(1)  # first search not measured to miss long latency due initial call

# Start Endless Loop
while ($true){
	$count++
	$start = get-date
	$result = $inbox.finditems(1)
	$duration = ((get-date) - $start).totalmilliseconds
	write-host "FE1:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[0])" -nonewline
	write-host "   FE2:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[1])" -nonewline
	write-host "   BE:$($result.service.HttpResponseHeaders["X-BEServer"])" -nonewline
	write-host "   RTT(ms)$($duration)"
	if ($statistics.beserver -ne $result.service.HttpResponseHeaders["X-BEServer"]){
		$statistics.beserver = $result.service.HttpResponseHeaders["X-BEServer"]
		write-host "New X-BEServer $($statistics.beserver) detected"
	}
	$durationsum+=$duration
	if ($statistics.durationmax -lt $duration){
		$statistics.durationmax = $duration
	}
	if ($statistics.durationmin -gt $duration){
		$statistics.durationmin = $duration
	}

	if ($count -ge $interval){
		$statistics.durationavg = ($durationsum/$count)
		$durationsum=0
		$statistics.timestamp = $start
		Write-Host "Summary Message"
		$statistics
		Write-Host "Export Results to archivcsv $($archivcsv)"
		$statistics | export-csv $archivcsv -Append -NoTypeInformation
		if ($prtgpushurl -ne ""){
			[string]$prtgresult = "<?xml version=""1.0"" encoding=""UTF-8"" ?>`r`n"
			$prtgresult+= "<prtg>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationmin</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationmin))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationmax</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationmax))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationavg</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationavg))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "  <text>OK</text>`r`n"
			$prtgresult+= "</prtg>"
			
			#_$prtgresult

			$prtgurl = ($prtgpushurl+ "?content=" + $prtgresult)
			Write-Host "Sending Data to $prtgpushurl" -nonewline
			try {
				$Answer=Invoke-Webrequest `
					-method "GET" `
					-URI $prtgurl
				if ($answer.Statuscode -ne 200) {
					write-warning "Request to PRTG $($prtgpushurl) failed"
				}
				elseif (($answer.content | convertfrom-json)."Matching Sensors" -eq "0") {
					write-warning " No matching sensor found for $($prtgpushurl)"
				}
				else {
					write-host "OK"
					# $result=$answer.content | convertfrom-json | select url,status,"Matching Sensors" 
					# $result.url = $prtgurl
					# $result
				}
			}
			catch {
				write-Error "Request to PRTG failed with error `r`n $($_.Exception.Message)"
			}
		}
		$count = 0
		$statistics.durationmin=9999
		$statistics.durationmax=0
	} 
	start-sleep -Milliseconds (1000-(get-date).Millisecond)
}

Write-Verbose "test-ewsimpersonation: End"