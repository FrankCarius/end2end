﻿# end2end HTTP
#
# Einfaches Skript, welches immer wieder die gleiche HTML-Seite ausliest und den 
#
#
# 20121201  Version 1.0 FC Initale Version
# 20121212  Version 1.1 FC Ausgsbe der Message auf erste 10 Zeichen beschränkt
# 20180108  Version 1.2 FC Sehr viele Umbauten und erweiterungen

[cmdletbinding()]
param(
	[string]$URL			= "https://outlook.office365.com/owa/favicon.ico",  			# url to connect and load data from
	[long]$slowabsms  		= 1000,															# Absolute maximum time in msec to rate a request als Slow
	[long]$slowavgfactor    = 5,   															# Factor for average slow calculation
	[string]$reportcsv		= "$PWD\end2endhttp"+(get-date -Format yyyMMdd-HHmmss)+".csv",  # Path to CSV File. Clear to disable CSV
	[switch]$report2pipe    = $false,														# enable to send output to stdout pipeline 
	[string]$smtpserver		= "",															# Enter Smarthost to send Alert Mail
	[string]$smtpto			= "frank.carius@netatwork.de",									# Recipient for Alert Mails
	[string]$smtpfrom		= "end2endhttp@netatwork.de",									# SMTP Sender Adress
	[long]$sleeptime		= 1000,															# time between the requests
	[long]$maxiteration		= 86400,														# maximum number of checks, about 24h
	[string]$prtgpushurl	= "http://192.168.178.11:5050/end2end-outlook.office365.com", 	# url to PRTG server with sensor prefix
	[int]$prtginterval 		= 60   															# time in seconds to send a PRTGPush
)

$error.clear()
write-host "Start End2End HTTP"

if ($reportcsv) {
	write-Host " initializing CSV-File $($reportcsv)"
	$csvfile = New-Object System.IO.StreamWriter $reportcsv, $true  # append
	$csvfile.WriteLine("timestamp,status,duration,avg10,url,httpstatuscode,httpbytes")
}

$prtglastreporttime = get-date
[long]$prtgsuccesscount=0;  	# count all successful request
[long]$prtgslowcount=0;         # count slow requests
[long]$prtgfailcount=0;   		# count all with hard failuer
[long]$prtgrequestsum=0;   		# sum of all received bytes
[long]$prtgoverlimitcount=0;  	# number of request over average 
[long]$prtgaveragetime10=0;       # maximum response time of last 10 checks
[long]$prtgmintime=99999;       # minimum respsonse time of last checks
[long]$prtgmaxtime=0;       # average resonse time of last checks
[long]$prtgtotalbytes=0; 	# sum of all received bytes
foreach ($iteration in (1..$maxiteration)) {
	Write-Verbose "Iteration $iteration";
	$Error.Clear();

	Write-verbose "Initializing Result Object"
	$result = New-Object PSObject -Property @{
				Timestamp	= (Get-Date -Format "dd.MM.yyyy HH:mm:ss")
				httpstatuscode  = ""
				httpbytes      = 0
				Status		   = ""
				Message		   = ""
				Duration	   = $null
				avg10          = 0
	}
	
	try {
		write-verbose " Downloading $url"
		[datetime]$TimeStart = [datetime]::Now;
		$webresult = invoke-webrequest `
						-uri $URL `
						-Timeoutsec 2 `
						-usebasicparsing
		write-host "." -nonewline
		$result.Duration = [long]([datetime]::Now - $TimeStart).TotalMilliseconds
		write-verbose " Duration: $($result.Duration)"
		write-verbose " Statucode $($webresult.statuscode)"
		$result.httpstatuscode = $webresult.statuscode
		$result.httpbytes = $webresult.RawContentLength
		$prtgtotalbytes+=$webresult.RawContentLength
		$result.Message = ""
		if ($result.Duration -ge $slowabsms){
			$result.status = "SLOWABS"
			$prtgslowcount++
		} 
		else {
			$result.status = "OK"
			$prtgsuccesscount++
		}

		if ($result.Duration -gt $prtgmaxtime) {$prtgmaxtime = $result.Duration}
		if ($result.Duration -lt $prtgmintime) {$prtgmintime = $result.Duration}

		$prtgrequestsum += $result.Duration 
		if ($prtgaveragetime10 -eq 0) {
			write-verbose "Set initial Average to first result $($result.Duration)"
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
		write-verbose " Failed"
		write-host "E" -nonewline
		$result.Message = $Error[0].tostring()
		$result.Duration = -1
		$result.status = "FAIL"
		$result.httpstatuscode = "Error"
		$result.httpbytes = 0
		$prtgfailcount++
		$error.clear()
	}

	if (($result.status -ne "OK") -and ($smtpserver -ne "")){
		send-mailmessage `
			-from $smtpfrom `
			-to $smtpto `
			-subject "End2EndHTML: $url slower than $slowabsms msec" `
			-body "End2EndHTML: $url slower than $slowabsms msec" `
			-smtpServer $smtpserver           
	}

	# Send results to file
	if ($reportcsv) {
		$csvfile.WriteLine($result.timestamp + "," +$result.status + "," + $result.Duration +"," + $result.avg10 +"," + $url + "," +  $result.httpstatuscode  + "," +$result.httpbytes)
	}

	if ($report2pipe){
		$result
	}

	if ($prtgpushurl) {
		write-verbose "PRTG Push URL Found"
		if (((get-date) - $prtglastreporttime).totalseconds -ge $prtginterval) {
			write-verbose "SendTo-PRTG: Process"

			if ($result.Duration -lt $prtgmintime) {$prtgmintime = $result.Duration}
			if ($prtgmintime -eq 99999) {$prtgmintime=""};
			Write-host "P" -nonewline
			[string]$prtgresult = '<?xml version="1.0" encoding="UTF-8" ?>
									<prtg>
										<result>
											<channel>Requests OK</channel>
											<value>'+$prtgsuccesscount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Requests to slow</channel>
											<value>'+$prtgslowcount+'</value>
											<mode>Absolute</mode>
											<unit>Count</unit>
										</result>
										<result>
											<channel>Requests failed</channel>
											<value>'+$prtgfailcount+'</value>
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
											<value>'+[long]($prtgrequestsum/($prtgsuccesscount+$prtgslowcount+$prtgfailcount+$prtgoverlimitcount)) +'</value>
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
			$prtgurl = ($prtgpushurl+$sensorname+ "?content=" + $prtgresult)
			write-verbose "Sending Data to $prtgurl"
			try {
				$prtgAnswer = Invoke-Webrequest `
					-method "GET" `
					-timeoutsec 5 `
					-usebasicparsing `
					-URI ($prtgpushurl+$sensorname+ "?content=" + $prtgresult)
				if ($prtganswer.Statuscode -ne 200) {
					write-warning "Request to PRTG failed"
				}	
				elseif (($prtganswer.content | convertfrom-json)."Matching Sensors" -eq "0") {
					write-warning " No matching sensor found for $($prtgpushurl+$sensorname)"
				}
				else {
					write-verbose "Sending Data OK"
					#$prtganswer.content | convertfrom-json | select-object url,status,"Matching Sensors" 
				}
			}
			catch {
				write-verbose "Failure sending PRTG-Data to $prtgurl"
				write-verbose $Error[0].exception
				write-host "!" -nonewline
			}
			# Prepare Counters for next measure interval
			$prtglastreporttime = Get-Date
			[long]$prtgsuccesscount=0;  	# count all successful request
			[long]$prtgslowcount=0;         # count slow requests
			[long]$prtgfailcount=0;   		# count all with hard failuer
			[long]$prtgrequestsum=0;   		# number of reuests to calculate the average
			[long]$prtgoverlimitcount=0;  	# number of request over average 
			[long]$prtgaveragetime=0;       # maximum response time of last 10 checks
			[long]$prtgmintime=99999;       # minimum respsonse time of last checks
			[long]$prtgmaxtime=0;       # average resonse time of last checks
			[long]$prtgtotalbytes=0; 	# sum of all received bytes
		} 
		else {
			write-verbose "PRTGPush Timeout not expired - SKIP"
		}
	}
	else {
		write-verbose "No PRTG Push URL Found - SKIP"
	}
	write-verbose "Wait $sleeptime"
	start-sleep -Milliseconds $sleeptime
	if ([console]::KeyAvailable) {
		if (([console]::readkey()).key.tostring().tolower() -eq "x") {
			write-host "Keypress detected - exiting"
			break;  # leave for loop
		}
		else {
			write-host "Press X to exit monitor"
		}
	}
} 
if ($reportcsv) {
	Write-host "Closing CSV-File $($reportcsv)"
	$csvfile.Close();
}
Write-host "End2End HTTP finished"