#end2endICMP
#
# Simple Script to to a regular ICMP-paket and measure the roundtrip.
# Check intermediate Hops, if roudntrip fails
# Output to CSV-File
#
# Version 1 war eine VBScript-Variante
 
param (
#	$target="avedge.netatwork.de" ,
#	$hoplist=("62.154.33.34","213.155.135.88","195.12.254.230") ,
	$target="www.google.de" ,
	$hoplist=("62.154.47.66","209.85.249.182") ,
	$buffersize = 160,
	$timeout=1000,
	$sleeptime = 100,
	$csvfilename = ".\End2EndPing.csv"
)
 
write-host "End2EndPing:START"
 
write-host "End2EndPing:Initialize ICMP Object"
$ping = new-object System.Net.NetworkInformation.Ping
 
write-host "End2EndPing:Initialize Buffer Size:"$buffersize
$enc = [system.Text.Encoding]::ASCII
$buffer = $enc.GetBytes([string]("").padright([int]($buffersize),"x")) 

write-host "End2EndPing:Initialize CSV-File"$csvfilename
[Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath
$csvfile = New-Object System.IO.StreamWriter($csvfilename)
if ((get-item $csvfilename).length -eq 0){
	write-host "End2EndPing: Adding CSV-Header"
	$csvfile.Writeline("datetime,ipaddress,roundtrip")
}

write-host "End2EndPing:Pinging target:"$target
$min=99999; $max=0 ; $avg=0 ;$total=0
$start=get-date
while (!$host.UI.RawUI.KeyAvailable) {
	if ((get-date).second -ne $start.second){
		$start=get-date
		write-host "min $min  max: $max   avg: "($avg/10)"  TotalByes/Sec: "$total"`n"(get-date -Format yyyyMMddhhmmss) -nonewline
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
		$csvfile.Writeline($now+","+$target+","+"Timeout")
		write-host "$now $target timeout"
		foreach ($hop in $hoplist) {
			$pingresult= $ping.send($hop,5000,$buffer)
			$csvfile.Writeline($now+","+$hop+","+$pingresult.roundtriptime)
			write-host "$now $hop " $pingresult.roundtriptime
		}
	}
	else {
		write-host $pingresult.Status -nonewline
	}
	#
	$waittime = $sleeptime - ((get-date) - $start).totalmilliseconds
	if ($waittime -ge 0) {start-sleep -milliseconds $waittime }
}
write-host "End2EndPing:Closing CSV-File"
$csvfile.close()
write-host "End2EndPing:END"