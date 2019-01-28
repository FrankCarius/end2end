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
# 20180508 sendto-prtg eingebaut 
# 20180712 FE Reihenfolge gefixt, UseDefaultCredentials addiert, Autodiscover Fehler abgefangen
#
param(
	[string]$MailboxSMTP = "frank.carius@netatwork.de",	# must be primary SMTP for impersonation
	[string]$Username    = "",				# use MailboxSMTP if empty and not userDefaultCredentials
	[string]$Password    = "", # password, if $username is set
	[switch]$useDefaultCredentials = $false,
	[string]$Domain      = "",				# Domain of the authentication user
	[double]$interval = 60,   # seconds to measure one interval
	[string]$Serviceurl  = "",						# Use Autodiscover, if empty https://exchange.msxfaq.local/EWS/Exchange.asmx",
	[string]$archivcsv = ".\end2end-ews_$($env:COMPUTERNAME).csv",
	[string]$prtgpushurl = "", #"http://prtg:5050/end2end-ews_$($env:COMPUTERNAME)",
	[switch]$useImpersonation = $false,						# forces impersonation 
	[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll", # EWS DLL
	#[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll", # 
	[switch]$EWSTrace,				 				# enable tracing of EWS to STDOUT
	[switch]$Verbose								# enable verbose output
)



# -----------------------------------------------------
# sendto-prtg   helper to send data to PRTG
# -----------------------------------------------------
function sendto-prtg (
	[string]$prtgresult,   # XML Structure
	[string]$prtgpushurl)  # HTTP-PUSH-url
{
	if ($prtgpushurl -ne "" ) {
		write-host "SendTo-PRTG: Post Result to PRTGProbe $prtgpushurl"
	
		try {
		$Answer=Invoke-RestMethod `
			-method "GET" `
			-URI ("$($prtgpushurl)?content=$($prtgresult)")
		}
		catch {
			write-Warning "SendTo-PRTG:Catch Request"
		}
		if ($answer."Matching Sensors" -eq "1") {
			write-host "SendTo-PRTG:Found 1 Sensors  OK"
		}
		elseif ($answer."Matching Sensors" -eq "0") {
			write-warning "SendTo-PRTG:Found 0 matching sensors. Retry next run"
		}
		else {
			write-Warning "SendTo-PRTG:Invalid reply"
			$answer
		}
	}
}


if ($Verbose) {
	$VerbosePreference = "continue" 
}

Write-Host "End2End-EWS: Start"
Write-Host " Loading EWS DLL"
Write-Host " Username: $username"
Write-Host " Domain  : $domain"
Write-Host " Password: <not visible>"
Write-Host " UseDefaultCredentials: $UseDefaultCredentials"

[void][Reflection.Assembly]::LoadFile($dllpath)
Write-Verbose "End2End-EWS:Creating EWS Service Class"
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService('Exchange2007_SP1')
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
if ($ewsTrace) {
	Write-Verbose "End2End-EWS:Tracing: enabled"
	$service.TraceEnabled = $true
	#Tracing EWS requests http://msdn.microsoft.com/en-us/library/dd633676(v=exchg.80).aspx
}
# --------------------------- Credentials and Impersonation --------------------------- 
if ($UseDefaultCredentials) {  
	Write-Verbose "Credentials: UseDefaultCredentials"
	$service.UseDefaultCredentials = $true
}
else  {
	Write-Verbose "End2End-EWS:Credentials: use alternate Credentials"
	if ($username -eq "") {
		Write-verbose " use MailboxSMTP $($MailboxSMTP) as Username "
		$username=$mailboxsmtp
	}
	$service.UseDefaultCredentials = $false  
	if ($password -eq "") {
		$password = ((get-credential -username $Username -message "Password for $Username").getnetworkcredential().password)
	}
	$service.Credentials = New-Object System.Net.NetworkCredential($username, $password, $domain) 
} 
if ($useImpersonation) {
	Write-Verbose "End2End-EWS:Credentials: use impersonation"
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxSMTP);

	#Configuring Exchange Impersonation (Exchange Web Services)  #http://msdn.microsoft.com/en-us/library/bb204095(v=exchg.80).aspx
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SID,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.PrincipalName,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SmtpAddress,wert)
}
# --------------------------- ServiceURI --------------------------- 

write-Host "End2End-EWS:Checking Autodiscover for Mailbox"
try{
	if ($serviceurl -eq "") {  
		Write-Verbose "End2End-EWS:ServiceURL: Using Autodiscover for $mailboxSMTP"
		$service.AutodiscoverUrl($mailboxSMTP,{$true})
	}
	else {
		Write-Verbose "End2End-EWS:ServiceURL: using specified $serviceurl"
		$service.Url = New-Object System.Uri($serviceurl) 
	} 
	write-Verbose "End2End-EWS:ServiceURI=$($service.Url.AbsoluteUri)"
}
Catch {
	Write-Error "End2End-EWS:Unable to determinte AutoDisover URL with Error `r`n $_.Exception.Message"
	$prtgresult = '
		<prtg>
		 <error>1</error>
		 <text>ERROR: Unable to determinte AutoDisover URL with Error `r`n '+$_.Exception.Message +' not found</text>
	  </prtg>'	
	sendto-prtg $prtgresult $prtgpushurl
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

Write-Host "End2End-EWS:Connecting to Inbox for SMTP $($mailboxSMTP)"
try {
	$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
	write-host " Number of Unread Messages : " $inbox.UnreadCount 
}
catch {
	write-host "Unable to connect to Inbox."
}

# Initializing Result Hashtable for storing values and generate output later
$statistics =  [pscustomobject][ordered] @{
	timestamp   = ([datetime]::UtcNow).tostring("yyyy-MM-dd HH:mm:ss")
	durationavgms = [long]9999
	durationmaxms = [long]0
	durationminms = [long]9999
	totalchecks = 0
	beserver    = [string]""
	Message     = "OK"
}
# Initialize Counters for Summary
[double]$durationsum=0
[double]$count=0
$result = $inbox.finditems(1)  # first search not measured to miss long latency due initial call
$statistics.beserver = $result.service.HttpResponseHeaders["X-BEServer"]  # preload Backend Server

# Start Loop until key pressed
Write-host "End2End-EWS: Start Monitoring Loop"
$continue=$True
while ($continue){
	$count++
	$start = get-date
	$result = $inbox.finditems(1)
	$duration =  [math]::Round(((get-date) - $start).totalmilliseconds)
	write-host "End2End-EWS ($($count)):FE1:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[1])" -nonewline
	write-host "   FE2:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[0])" -nonewline
	write-host "   BE:$($result.service.HttpResponseHeaders["X-BEServer"])" -nonewline
	write-host "   RTT(ms)$($duration)"
	if ($statistics.beserver -ne $result.service.HttpResponseHeaders["X-BEServer"]){
		$statistics.beserver = $result.service.HttpResponseHeaders["X-BEServer"]
		write-host "End2End-EWS:New X-BEServer $($statistics.beserver) detected"
		$statistics.Message="End2End-EWS:New X-BEServer $($statistics.beserver) detected"
	}
	$durationsum+=$duration
	if ($statistics.durationmaxms -lt $duration){
		$statistics.durationmaxms = $duration
	}
	if ($statistics.durationminms -gt $duration){
		$statistics.durationminms = $duration
	}

	if ($count -ge $interval){
		$statistics.durationavgms =  [math]::Round($durationsum/$count)
		$statistics.totalchecks= $count
		$durationsum=0
		$statistics.timestamp = ([datetime]::UtcNow).tostring("yyyy-MM-dd HH:mm:ss")
		Write-Host "End2End-EWS:Summary Message"
		$statistics
		Write-Host "End2End-EWS:Export Results to archivcsv $($archivcsv)"
		$statistics | export-csv $archivcsv -Append -NoTypeInformation
		if ($prtgpushurl -ne ""){
			[string]$prtgresult = "<?xml version=""1.0"" encoding=""UTF-8"" ?>`r`n"
			$prtgresult+= "<prtg>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationminms</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationminms))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationmaxms</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationmaxms))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "<result>`r`n"
			$prtgresult+= "  <channel>durationavgms</channel>`r`n"
			$prtgresult+= "  <value>$([int]($statistics.durationavgms))</value>`r`n"
			$prtgresult+= "  <unit>Custom</unit>`r`n"
			$prtgresult+= "  <customunit>ms</customunit>`r`n"
			$prtgresult+= "</result>`r`n"
			$prtgresult+= "  <text>OK</text>`r`n"
			$prtgresult+= "</prtg>"
			sendto-prtg $prtgresult $prtgpushurl

			write-host "End2End-EWS:Sleeping $($sleeptime) Seconds. Press ""X""-key to stop after next try"
		}
		$count = 0
		$statistics.durationminms=9999
		$statistics.durationmaxms=0
		$statistics.Message="OK"
	} 
	start-sleep -Milliseconds (1000-(get-date).Millisecond)

	if ([console]::KeyAvailable) {
		write-host " Key detected"
		$keycode = [System.Console]::ReadKey() 
		if ($keycode.key -eq "X") {
			write-host " Terminating Script"
			$continue = $false
		}
	} 
}

Write-Verbose "End2End-EWS:impersonation: End"