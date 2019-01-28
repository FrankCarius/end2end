# end2end-ping
#
# Ping a given Endpoint every second
# measure roudntrip time and generate summaries
# avg, min, max, lost  (<1Sek)
#
# 20181114 initial Version based on end2end-pathping and end2end-udp3478
# 20190117  Fix bei IP-Adresse
# 20190119  Ausgabe schöner
#
# Pending
#   Currently IPv$ DNS only
#  Parallelisierung f�r mehrere Hosts
#
# System.Net.NetworkInformation  PingSend-Method
# https://docs.microsoft.com/de-de/dotnet/api/system.net.networkinformation.ping.send?view=netframework-4.7.2
# System.Net.NetworkInformation  PingReply  Class
# https://docs.microsoft.com/de-de/dotnet/api/system.net.networkinformation.pingreply?view=netframework-4.7.2
#
#  IPStatus Enum
# https://docs.microsoft.com/de-de/dotnet/api/system.net.networkinformation.ipstatus
#

param (
	[string]$hostname = "internetbeacon.msedge.net",  # samples  google = 8.8.8.8   SfBOnlineEdge = 13.107.8.2, outlook.office365.com, internetbeacon.msedge.net
	[byte]$pingcount=60,       # total number of Ping to collect in one dataset
	[long]$pingtimeout = 200,  # ms timeout to wait for a reply. anything longer will be treated as loss
	[long]$buffersize = 160,   # PING Buffersize
	[string]$csvfilename		= "$PWD\end2end-ping_$($env:COMPUTERNAME).csv"	# Path to CSV File. Clear to disable CSV
)


set-psdebug -strict

Write-Host "end2end-ping: Start"
Write-Host "  Param Hostname   : $($hostname)"
Write-Host "  Param Pingcount  : $($pingcount)"
Write-Host "  Param PingTimeout: $($pingtimeout)"
Write-Host "  Param Buffersize : $($buffersize)"
Write-Host "  Param CSVFile    : $($csvfilename)"

Write-Host "Check IP-Address"
$remoteip = $null
if ($hostname -as [ipaddress]) {
	write-host "IP-Address ok"
	$remoteip = $hostname
}
else {
	Write-Host "Try to resolve $($hostname) to IP-Address"
	$error.clear()
	$remoteip = ((resolve-dnsname $hostname -type A -ErrorAction SilentlyContinue)|?{$_.ipaddress})[0].ipaddress
}
if ($remoteip -ne $null) {
	Write-Host "End2End-Ping:Initialize ICMP Object"
	$ping = new-object System.Net.NetworkInformation.Ping
	 
	Write-Host "End2End-Ping:Initialize ICMP Buffer with Size:"$buffersize
	$asciiencoder = [system.Text.Encoding]::ASCII
	$buffer = $asciiencoder.GetBytes([string]("www.msxfaq.de end2end-ping").padright(($buffersize),"x")) 

	if ($csvfilename -ne "") {
		Write-Host "End2End-Ping:Initialize CSV-File"$csvfilename
		if (test-path $csvfilename -pathtype leaf){
			Write-Host "End2End-Ping:Adding CSV-Header"
			"Timestamp,RemoteIP,Min,Avg,Max,Total,Lost,statuspacketloss,statusavg2min,statusmax2avg" | out-file $csvfilename -append
		}
	}

	Write-Host "Initialize Result Property"
	$result = [pscustomobject][ordered]@{
			date = get-date -Format yyyy-MM-dd
			time = get-date -Format hh:mm:ss
			remoteip = [string]$($remoteip)
			max = [long]0
			min = [long]9999999
			avg = [long]0
			total = [long]0
			loss = [long]0
			statuspacketloss=0
			statusavg2min=0
			statusmax2avg=0
	}

	Write-Host "Start Endless ICMP Ping to $($remoteip) - press CTRL-C to Stop"
	Write-host  "Info:  .=Succesful   T=Timeout  E=ExpiredTTL  U=Unreachable  S=Size ?=Unknown  "
	Write-host  "Colorcode:" -nonewline
	Write-Host "<=5ms" -nonewline -backgroundcolor blue
	Write-Host "<=10ms" -nonewline -backgroundcolor green -foregroundcolor black
	Write-Host "<=50ms" -nonewline -backgroundcolor yellow -foregroundcolor black
	Write-Host "<=100ms." -nonewline -backgroundcolor magenta
	Write-Host ">100ms" -backgroundcolor red
	[long]$Avgsum = 0
	while ($true) {
		$result.total++
		# ping remote host
		$pingresult=$ping.send($remoteip,$pingtimeout,$buffer)
		
		switch ($pingresult.Status.tostring())	{
			"Success" {
				switch ($pingresult.roundtriptime) {
					{$_ -le 5} 		{Write-Host "." -nonewline -backgroundcolor blue;break;}
					{$_ -le 10} 	{Write-Host "." -nonewline -backgroundcolor green -foregroundcolor black;break;}
					{$_ -le 50} 	{Write-Host "." -nonewline -backgroundcolor yellow -foregroundcolor black;break;}
					{$_ -le 100} 	{Write-Host "." -nonewline -backgroundcolor magenta ;break;}
					default 		{Write-Host "." -nonewline -backgroundcolor red}
				}
				if ($pingresult.roundtriptime -le $result.min) {$result.min = $pingresult.roundtriptime} 
				if ($pingresult.roundtriptime -ge $result.max) {$result.max = $pingresult.roundtriptime}
				$Avgsum+=$pingresult.roundtriptime
			}
			"TimedOut" {
				Write-Host "T" -nonewline -backgroundcolor red
				$result.loss++
			}
			"TtlExpired" {
				Write-Host "E" -nonewline -backgroundcolor red
				$result.loss++
			}
			"DestinationPortUnreachable" {
				Write-Host "U" -nonewline -backgroundcolor red
				$result.loss++
			}
			"DestinationHostUnreachable" {
				Write-Host "U" -nonewline -backgroundcolor red
				$result.loss++
			}
			"PacketTooBig" {
				Write-Host "S" -nonewline -backgroundcolor red
				$result.loss++
			}
			default {
				Write-Host "?" -nonewline
				Write-Host $pingresult.Status 
				$result.loss++
			}
		}

		if ($result.total -ge $pingcount){
			# calculating Average
			$result.avg = [long]($Avgsum/$result.total)
			[long]$Avgsum = 0
			
			if (($result.loss*10) -gt $result.total) {
				write-warning "PacketLoss > 10%"
				$result.statuspacketloss="1"
			}

			if (($result.avg) -gt (($result.min+1)*10)) {  #ignore 0 as min
				# add 0,1
				write-warning "Avg 10 Times higher than Min"
				$result.statusavg2min="1"
			}
			if (($result.max) -gt ($result.avg*10)) {
				write-warning "Max 10 Times higher than Avg"
				$result.statusmax2avg="1"
			}
			Write-Host "$($result.timestamp) Min=$($result.min) Avg=$($result.avg) Max=$($result.max) Total=$($result.total) Loss=$($result.loss) Status=$($result.statuspacketloss),$($result.statusavg2min),$($result.statusmax2avg)"
			#$csvfile.Writeline("$($result.timestamp),$($result.RemoteIP),$($result.max),$($result.min),$($result.avg/$result.total),$($result.total),$($result.loss),$($result.statuspacketloss),$($result.statusavg2min),$($result.statusmax2avg)")
			if ($csvfilename -ne "") {
				"$($result.timestamp),$($result.RemoteIP),$($result.min),$($result.avg),$($result.max),$($result.total),$($result.loss),$($result.statuspacketloss),$($result.statusavg2min),$($result.statusmax2avg)" | out-file $csvfilename -append
			}
			$result
			$result.date = get-date -Format yyyy-MM-dd
			$result.time = get-date -Format hh:mm:ss
			$result.remoteip = $($remoteip)
			$result.max=0
			$result.min=9999999
			$result.avg = 0
			$result.total = 0
			$result.loss = 0
			$result.statuspacketloss=0
			$result.statusavg2min=0
			$result.statusmax2avg=0
		}
		
		# Wait for next second
		$waittime = 1000 - (get-date).millisecond
		if ($waittime -ge 0) {
			start-sleep -milliseconds $waittime 
		}
	}
}
else {
	Write-warning "end2end-ping: Stop - no IP-Address found"
}
Write-Host "end2end-ping: End"