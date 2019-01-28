#end2endpathping
#
# Simple Script to to a regular ICMP-paket and measure the roundtrip.
# Check intermediate Hops, if roudntrip fails
# Output to CSV-File
#
# Version 1 war eine VBScript-Variante
# Ver 2.1  mit CSV-Ausgabe
# ver 2.2 Rename to "PathPing"
 
param (
#	$target="avedge.netatwork.de" ,
#	$hoplist=("62.154.33.34","213.155.135.88","195.12.254.230") ,
	$target="internetbeacon.msedge.net",  # samples  google = 8.8.8.8   SfBOnlineEdge = 13.107.8.2, outlook.office365.com, internetbeacon.msedge.net
	$hoplist=("192.168.178.1","62.155.242.220"," 217.5.116.98","87.128.239.210","209.85.249.182") ,
	$buffersize = 160,
	$timeout=1000,
	$sleeptime = 500,
	$csvfilename = ".\end2end-pathping.csv"
)
 
write-host "end2end-pathping:START"
write-host "end2end-pathping:csvfilename  $($csvfilename)"
 
write-host "end2end-pathping:Initialize ICMP Object"
$ping = new-object System.Net.NetworkInformation.Ping
 
write-host "end2end-pathping:Initialize Buffer Size:"$buffersize
$enc = [system.Text.Encoding]::ASCII
$buffer = $enc.GetBytes([string]("").padright([int]($buffersize),"x")) 

write-host "end2end-pathping:Pinging target:"$target
$min=99999; $max=0 ; $avg=0 ;$total=0
$start=get-date
$result = "" | select datetime,target,status 
while ($true) {
	if ((get-date).second -ne $start.second) {
		# start only, if second is new
		$start = get-date
		write-host "min $min  max: $max   avg: "($avg/10)"  TotalByes/Sec: "$total"`n"(get-date -Format yyyyMMddhhmmss) -nonewline
		$result | export-csv $csvfilename -append -notypeinformation
		$min=99999; $max=0 ; $avg=0 ;$total=0
	}
	$pingresult= $ping.send($target,$timeout,$buffer)
	$total+=$buffersize
	if ($pingresult.Status.tostring() -eq "Success") {
		write-host "." -nonewline
		#write-progress -Activity "ICMP" -Percentcomplete ($pingresult.roundtriptime/10) -Status $pingresult.roundtriptime
		if ($pingresult.roundtriptime -le $min) {$min = $pingresult.roundtriptime}
		if ($pingresult.roundtriptime -ge $max) {$max = $pingresult.roundtriptime}
		$avg+=$pingresult.roundtriptime
	}
	elseif ($pingresult.Status.tostring() -eq "TimedOut") {
		write-host "TIMEOUT, checking Hops"
		$now=(get-date -Format yyyyMMddhhmmss)
		$result.datetime = $now
		$result.target = $target
		$result.status = "Timeout"
		$result | export-csv $csvfilename -append -notypeinformation
		write-host "$now $target timeout"
		foreach ($hop in $hoplist) {
			$pingresult= $ping.send($hop,5000,$buffer)
			$result.datetime = $now
			$result.hop = $hop
			$result.status = $pingresult.roundtriptime
			$result | export-csv $csvfilename -append -notypeinformation
			write-host "$now $hop " $pingresult.roundtriptime
		}
	}
	else {
		write-host $pingresult.Status -nonewline
	}
	$waittime = $sleeptime - ((get-date) - $start).totalmilliseconds
	if ($waittime -ge 0) {start-sleep -milliseconds $waittime }
}
write-host "end2end-pathping:Closing CSV-File"

write-host "end2end-pathping:END"