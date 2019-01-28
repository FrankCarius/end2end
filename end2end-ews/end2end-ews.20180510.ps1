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
# 20180508 sendto-prtg eingebaut und 
#
param(
	[string]$MailboxSMTP = "frank.carius@netatwork.de",	# must be primary SMTP for impersonation
	[string]$Username    = "frank.carius@netatwork.de",				# use default credentials, if empty
	[string]$Domain      = "",				# Domain of the authentication user
	[double]$interval = 10,   # seconds to measure one interval
	[string]$Password    = ((get-credential -username $Username -message "Password for $Domain\$Username").getnetworkcredential().password),  # password, if $username is set
	[string]$Serviceurl  = "",						# Use Autodiscover, if empty https://exchange.msxfaq.local/EWS/Exchange.asmx",
	[string]$archivcsv = ".\end2end-ews_$($env:COMPUTERNAME).csv",
	[string]$prtgpushurl = "http://prtg:5050/end2end-ews_$($env:COMPUTERNAME)",
	[switch]$useImpersonation = $false,						# forces impersonation 
	[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\nProgram Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll", # EWS DLL
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
	if ($prtgpushurl -eq "" ) {
		$Scriptname = (split-path $MyInvocation.ScriptName -Leaf).replace(".ps1","")
		$prtgpushurl=  "http://prtg:5050/$($scriptname)_$($env:COMPUTERNAME)"
	}
	write-host "SendTo-PRTG: Post Result to PRTGProbe $prtgpushurl"
	
	$Answer=Invoke-RestMethod `
		-method "GET" `
		-URI ("$($prtgpushurl)?content=$($prtgresult)")
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


if ($Verbose) {
	$VerbosePreference = "continue" 
}
Write-verbose "End2End-EWS: Start"
Write-Verbose " Loading EWS DLL"
Write-Verbose " Username: $username"
Write-Verbose " Domain  : $domain"
Write-Verbose " Password: <not visible>"

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
if ($username -eq "") {  Write-Verbose "Credentials: UseDefaultCredentials"
	$service.UseDefaultCredentials = $true
}
else  {
	Write-Verbose "End2End-EWS:Credentials: use alternate Credentials"
	$service.UseDefaultCredentials = $false  
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
		Write-Host "End2End-EWS:ServiceURL: Using Autodiscover for $mailboxSMTP"
		$service.AutodiscoverUrl($mailboxSMTP,{$true})
	}
	else {
		Write-Host "End2End-EWS:ServiceURL: using specified $serviceurl"
		$service.Url = New-Object System.Uri($serviceurl) 
	} 
	write-Host "End2End-EWS:ServiceURI=$($service.Url.AbsoluteUri)"
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
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
write-host " Number of Unread Messages : " $inbox.UnreadCount 

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

# Start Loop until key pressed
Write-host "End2End-EWS: Start Monitoring Loop"
$continue=$True
while ($continue){
	$count++
	$start = get-date
	$result = $inbox.finditems(1)
	$duration = ((get-date) - $start).totalmilliseconds
	write-host "End2End-EWS:FE1:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[0])" -nonewline
	write-host "   FE2:$($result.service.HttpResponseHeaders["X-FEServer"].split(",")[1])" -nonewline
	write-host "   BE:$($result.service.HttpResponseHeaders["X-BEServer"])" -nonewline
	write-host "   RTT(ms)$($duration)"
	if ($statistics.beserver -ne $result.service.HttpResponseHeaders["X-BEServer"]){
		$statistics.beserver = $result.service.HttpResponseHeaders["X-BEServer"]
		write-host "End2End-EWS:New X-BEServer $($statistics.beserver) detected"
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
		Write-Host "End2End-EWS:Summary Message"
		$statistics
		Write-Host "End2End-EWS:Export Results to archivcsv $($archivcsv)"
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
			sendto-prtg $prtgresult $prtgpushurl

			write-host "End2End-EWS:Sleeping $($sleeptime) Seconds. Press ""X""-key to stop after next try"
		}
		$count = 0
		$statistics.durationmin=9999
		$statistics.durationmax=0
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