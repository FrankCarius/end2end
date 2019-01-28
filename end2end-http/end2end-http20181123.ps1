# end2end HTTP
#
# Einfaches Skript, welches immer wieder die gleiche HTML-Seite ausliest und den 
#
#
# 20121201  Version 1.0 FC Initale Version
# 20121212  Version 1.1 FC Ausgsbe der Message auf erste 10 Zeichen beschränkt
# 20180108  Version 1.2 FC Sehr viele Umbauten und erweiterungen
# 20181212  Version 1.3 FC report2screeen addiert
# 20180307  Version 1.4 FC httperrorcode numeric, small output changes
# 20180328  Version 1.5 FC Parameterurls und report2screen=True
# 20180426  Version 20180426 FC Versionnummer to yyymmdd
# 2018058   Ausgabe optimiert, maxiteration entfernt und endlos laufen lassen, reportcsv auf export-csv, SendTo-PRTG
# 20181123  Kosmetische Korrekturen

[cmdletbinding()]
param(
	[string]$URL			= "https://outlook.office365.com/owa/favicon.ico",  			# url to connect and load data from
	[ValidateSet("favicon","smime","health","",$null)][String]$urltemplate = "favicon",						# use predefined set
	[long]$slowabsms  		= 1000,															# Absolute maximum time in msec to rate a request als Slow
	[long]$slowavgfactor    = 5,   															# Factor for average slow calculation
	[string]$reportcsv		= "$PWD\end2end-http_$($env:COMPUTERNAME).csv",  # Path to CSV File. Clear to disable CSV
	[switch]$report2pipe    = $false,														# enable to send output to stdout pipeline 
	[switch]$report2screen 	= $true,														# enable to send output to stdout pipeline 
	[string]$smtpserver		= "",															# Enter Smarthost to send Alert Mail
	[string]$smtpto			= "user1@msxfaq.de",									# Recipient for Alert Mails
	[string]$smtpfrom		= "end2end-http@msxfaq.de",									# SMTP Sender Adress
	[long]$sleeptime		= 1,															# time between the requests
	[long]$maxiteration		= 86400,														# maximum number of checks, about 24h
	[string]$prtgpushurl	= "http://prtg:5050/end2end-outlook.office365.com_$($env:COMPUTERNAME)", 	# url to PRTG server with sensor prefix
	[int]$reportinterval 	= 60   															# time in seconds to send a summary
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


$error.clear()
write-host "End2End-HTTP: Start"

switch ($urltemplate){
	"favicon" 	{$url = "https://outlook.office365.com/owa/favicon.ico"}
	"smime"		{$url = "https://outlook.office365.com/owa/smime/owasmime.msi"}
	"health"	{$url = "https://outlook.office365.com/owa/healthcheck.htm"}
}
write-host "End2End-HTTP: URL = $($url)"

$lastreporttime = get-date
[long]$reportsuccesscount=0;  	# count all successful request
[long]$reportslowcount=0;         # count slow requests
[long]$reportfailcount=0;   		# count all with hard failuer
[long]$reportrequestsum=0;   		# sum of all received bytes
[long]$prtgoverlimitcount=0;  	# number of request over average 
[long]$prtgaveragetime10=0;       # maximum response time of last 10 checks
[long]$prtgmintime=99999;       # minimum respsonse time of last checks
[long]$prtgmaxtime=0;       # average resonse time of last checks
[long]$prtgtotalbytes=0; 	# sum of all received bytes

$continue=$true
while ($continue) {
	$Error.Clear();

	Write-verbose "End2End-HTTP: Initializing Result Object"
	$result = New-Object PSObject -Property @{
				Timestamp	= (Get-Date -Format "dd.MM.yyyy HH:mm:ss")
				httpstatuscode  = ""
				httpbytes      = 0
				Status		   = ""
				Message		   = ""
				Duration	   = $null
				avg10          = 0
				url            = $url
	}
	
	try {
		write-verbose "End2End-HTTP: Downloading $url"
		[datetime]$TimeStart = [datetime]::Now;
		$webresult = invoke-webrequest `
						-uri $url `
						-Timeoutsec 2 `
						-usebasicparsing
		write-host "." -nonewline
		$result.Duration = [long]([datetime]::Now - $TimeStart).TotalMilliseconds
		write-verbose "End2End-HTTP: Duration: $($result.Duration)"
		write-verbose "End2End-HTTP: Statucode $($webresult.statuscode)"
		$result.httpstatuscode = $webresult.statuscode
		$result.httpbytes = $webresult.RawContentLength
		$prtgtotalbytes+=$webresult.RawContentLength
		$result.Message = ""
		if ($result.Duration -ge $slowabsms){
			$result.status = "SLOWABS"
			$reportslowcount++
		} 
		else {
			$result.status = "OK"
			$reportsuccesscount++
		}

		if ($result.Duration -gt $prtgmaxtime) {$prtgmaxtime = $result.Duration}
		if ($result.Duration -lt $prtgmintime) {$prtgmintime = $result.Duration}

		$reportrequestsum += $result.Duration 
		if ($prtgaveragetime10 -eq 0) {
			write-verbose "End2End-HTTP: Set initial Average to first result $($result.Duration)"
			$prtgaveragetime10 = $result.Duration
		}
		else {
			$prtgaveragetime10 = ((($prtgaveragetime10*9) + $result.Duration)/10)
		}

		$result.avg10 = $prtgaveragetime10
		if ($result.Duration -gt ($prtgaveragetime10*$slowavgfactor)) {
			$prtgoverlimitcount++
			$result.status = "SLOWAVG"
		}
	}
	catch {
		write-verbose "End2End-HTTP: Failed"
		write-host "E" -nonewline
		$result.Message = $Error[0].tostring()
		$result.Duration = -1
		$result.status = "FAIL"
		$result.httpstatuscode = "Error"
		$result.httpbytes = 0
		$reportfailcount++
		$error.clear()
	}

	if (($result.status -ne "OK") -and ($smtpserver -ne "")){
		send-mailmessage `
			-from $smtpfrom `
			-to $smtpto `
			-subject "End2End-htp: $url slower than $slowabsms msec" `
			-body "End2End-http: $url slower than $slowabsms msec" `
			-smtpServer $smtpserver           
	}

	if ($result.httpstatuscode -eq "Error"){
		$result.httpstatuscode = 600
	}
	# Send results to file
	if ($reportcsv) {
		$result | export-csv -path $reportcsv -append -notypeinformation
	}

	if ($report2pipe){
		# sending every single result to pipeline
		$result
	}

	if (((get-date) - $lastreporttime).totalseconds -ge $reportinterval) {
		write-host ""   # linebreak after status-dots
		if ($prtgpushurl) {
			write-verbose "End2End-HTTP: PRTG Push URL Found"
			if ($result.Duration -lt $prtgmintime) {$prtgmintime = $result.Duration}
			if ($prtgmintime -eq 99999) {$prtgmintime=""};
			Write-host "P" -nonewline
			[string]$prtgresult = '<?xml version="1.0" encoding="UTF-8" ?>
									<prtg>
										<result>
											<channel>Requests OK</channel>
											<value>'+$reportsuccesscount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Requests to slow</channel>
											<value>'+$reportslowcount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Requests failed</channel>
											<value>'+$reportfailcount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Request much slower than average</channel>
											<value>'+$prtgoverlimitcount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Total Received Bytes</channel>
											<value>'+$prtgtotalbytes+'</value>
											<customunit>Bytes</customunit>
											<mode>Absolute</mode>
										</result>
										<result>
											<channel>Average Request Time</channel>
											<value>'+[long]($reportrequestsum/($reportsuccesscount+$reportslowcount+$reportfailcount+$prtgoverlimitcount)) +'</value>
											<customunit>ms</customunit>
											<mode>Absolute</mode>
										</result>
										<result>
											<channel>Longest Request</channel>
											<value>'+$prtgmaxtime +'</value>
											<customunit>ms</customunit>
											<mode>Absolute</mode>
										</result>
										<result>
											<channel>Fastest Request</channel>
											<value>'+$prtgmintime +'</value>
											<customunit>ms</customunit>
											<mode>Absolute</mode>
										</result>
										<text>End2End HTTP Monitoring</text>
									</prtg>'

			Sendto-prtg $prtgresult $prtgpushurl
		} 
		else {
			write-verbose "End2End-HTTP:No PRTG Push URL Found - SKIP"
		}
		
		if ($report2screen){
			write-host ""
			#Write-Host "Summaryreport to Screen"
			write-host " SuccessCount        : $($reportsuccesscount)"
			Write-Host " Requests to slow    : $($reportslowcount)"
			Write-Host " Requests failed     : $($reportfailcount)"
			Write-Host " much slower than avg: $($prtgoverlimitcount)"
			Write-Host " Total Received Bytes: $($prtgtotalbytes)"
			Write-Host " Average Request Time: $([long]($reportrequestsum/($reportsuccesscount+$reportslowcount+$reportfailcount+$prtgoverlimitcount))) ms"
			Write-Host " Longest Request     : $($prtgmaxtime ) ms"
			Write-Host " Fastest Request     : $($prtgmintime) ms"
		}

		$lastreporttime = Get-Date
		[long]$reportsuccesscount=0;  	# count all successful request
		[long]$reportslowcount=0;         # count slow requests
		[long]$reportfailcount=0;   		# count all with hard failuer
		[long]$reportrequestsum=0;   		# number of reuests to calculate the average
		[long]$prtgoverlimitcount=0;  	# number of request over average 
		[long]$prtgaveragetime=0;       # maximum response time of last 10 checks
		[long]$prtgmintime=99999;       # minimum respsonse time of last checks
		[long]$prtgmaxtime=0;       # average resonse time of last checks
		[long]$prtgtotalbytes=0; 	# sum of all received bytes
		write-host " Press ""X""-key to stop after next try"
	}
	else {
		write-verbose "End2End-HTTP:Report generation Timeout not expired - SKIP"
	}

	start-sleep -seconds $sleeptime
	if ([console]::KeyAvailable) {
		write-host " Key detected"
		$keycode = [System.Console]::ReadKey() 
		if ($keycode.key -eq "X") {
			write-host " Terminating Script"
			$continue = $false
		}
	} 
} 
Write-host "End2End-HTTP: End"