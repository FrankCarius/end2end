# End2End-udp3478
# Simple Skype for Business Online Edge Turn test
#
# 20180414 Added PRTG output
# 20180416 added RTT min/max/avg
# 20180425 fixes for PowerShell2 and PRTG detection and using 50019 as source port. 50000 is used to often
# 20180504 Erweiterung ArchivCSV

param (
	[int]$sourceudpport=50019,     # Source port to send UDP from. SfB normally uses 50000-50019
	[int]$remoteudpport=3478,      # Default STUN/TURN Port
	[string]$remoteip = "13.107.8.2",   # Anycast IP of Office 365 TURN Servers
	[byte]$maxttl=32,               # start with TTL32 for Countdown
	[int]$maxretries = 5,           # repeatcount if packet was not received
	[int]$averagecount=5,           # measure response with max TTL and calcualte average and max
	[string]$archivcsv = ".\End2End-udp3478.csv",   # store results
	[switch]$prtg=$false
)

Write-host "End2End-UDP3478:Start UDP-Client on $($sourceudpport)"
$udpClient = new-Object System.Net.Sockets.Udpclient($sourceudpport) 
$udpClient.Client.ReceiveTimeout = 500   
#$udpClient.Client.Blocking = $true

# STUN Packet from SfB Network Assessment Tool
$byteBuffer = @(0x00,0x03,0x00,0x54,0x21,0x12,0xa4,0x42,0xd2,0x79,0xaa,0x56,0x87,0x86,0x48,
               0x73,0x8f,0x92,0xef,0x58,0x00,0x0f,0x00,0x04,0x72,0xc6,0x4b,0xc6,0x80,0x08,
               0x00,0x04,0x00,0x00,0x00,0x04,0x00,0x06,0x00,0x30,0x04,0x00,0x00,0x1c,0x00,
               0x09,0xbe,0x58,0x24,0xe4,0xc5,0x1c,0x33,0x4c,0xd2,0x3f,0x50,0xf1,0x5d,0xce,
               0x81,0xff,0xa9,0xbe,0x00,0x00,0x00,0x01,0xeb,0x15,0x53,0xbd,0x75,0xe2,0xca,
               0x14,0x1e,0x36,0x31,0xbb,0xe3,0xf5,0x4a,0xa1,0x32,0x45,0xcb,0xf9,0x00,0x10,
               0x00,0x04,0x00,0x00,0x01,0x5e,0x80,0x06,0x00,0x04,0x00,0x00,0x00,0x01)
$RemoteIpEndPoint = New-Object System.Net.IPEndPoint([system.net.IPAddress]::Parse("0.0.0.0"),0);

Write-host " Connect to $($remoteip):$($remoteudpport)"
$udpclient.Connect($remoteip, $remoteudpport)

write-host " Init summary Table"
$summary = New-Object PSObject -Property @{
		remoteip=$remoteip
		remoteudpport =$remoteudpport
		maxroundtrip=[int]0
		minroundtrip=[int]9999
		averagertt=[int]0
		successrttcount=[int]0
		answerreceived=[boolean]$false
}

[int]$summary.maxroundtrip=0
[int]$summary.minroundtrip=9999
[int]$summary.averagertt=0
[int]$summary.successrttcount=0
[boolean]$summary.answerreceived=$false
write-host " Calculate average Roundtrip "
for ($loopcount=0; $loopcount -lt $averagecount; $loopcount++) {
	$starttime = get-date
	$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
	try {
		$null = $udpClient.Receive([ref]$remoteIpendpoint)
		write-host "Answer received" -nonewline
		$summary.successrttcount++
		$roundtriptimems = [int](((get-date) - $starttime).totalmilliseconds)
		write-host " $($roundtriptimems ) ms"
		$summary.averagertt = $summary.averagertt + $roundtriptimems
		if ($summary.maxroundtrip -lt $roundtriptimems) {
			$summary.maxroundtrip = $roundtriptimems
		}
		if ($summary.minroundtrip -gt $roundtriptimems) {
			$summary.minroundtrip = $roundtriptimems
		}
	}
	catch {
		#$_
		write-host "NO Answer received."
	}
}

if ($summary.successrttcount -eq 0) {
	$summary.averagertt = 0
}
else {
	$summary.averagertt = [int]($summary.averagertt / $summary.successrttcount)
	write-host "  RTT Avg: $($summary.averagertt)"
	write-host "  RTT Max: $($summary.maxroundtrip)"
	write-host "  RTT Min: $($summary.minroundtrip)"

	write-host " Start TTL Distance Check"
	[int]$retrycount=1
	$result=@{}
	for ($ttl=$maxttl; $ttl -gt 0 ; $ttl--) {
		Write-host "  Send STUN-Request TTL $($ttl) " -nonewline
		$udpclient.ttl = $ttl
		$starttime = get-date
		$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
		try {
			$result[$ttl]=$udpClient.Receive([ref]$remoteIpendpoint)
			write-host "  Answer received"
			$summary.answerreceived=$true
		}
		catch {
			#$_
			write-host "  NO Answer received. Try $($retrycount)"
			if ($retrycount -ge $maxretries){
				break
			}
			$retrycount++
			$ttl++
		}
	}
	if ($summary.answerreceived) {
		write-host " Summary: Distance $($ttl+1)"
	}
	else {
		write-host " Summary: NoConnection"
	}
}

Write-host " Closing UDP"
$udpClient.close()
# $udpClient.dispose()  # method not avaiable in .net 2.0


if ($archivcsv -ne ""){
	write-host " Export Data to ArchivCSV $($archivcsv)"
$summary
	$summary | export-csv -path $archivcsv -append -notypeinformation
}


Write-host "End2End-UDP3478:Done"
if ($myinvocation.MyCommand.Definition -match "\\EXEXML\\" -or $prtg) {
	if ($summary.answerreceived) {
	Write-Host "<prtg>"
	Write-Host "  <result>"
	Write-Host "    <channel>Hopcount</channel>"
	Write-Host "      <value>$($ttl+1)</value>"
	Write-Host "      <unit>Custom</unit>" 
	Write-Host "      <customunit>Hops</customunit>"
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Avg</channel>"
	Write-Host "      <value>$($summary.averagertt)</value>"
	Write-Host "      <unit>Custom</unit>"
	Write-Host "      <customunit>ms</customunit>"
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Max</channel>"
	Write-Host "      <value>$($summary.maxroundtrip)</value>"
	Write-Host "      <unit>Custom</unit>"
	Write-Host "      <customunit>ms</customunit>"
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Min</channel>"
	Write-Host "      <value>$($summary.minroundtrip)</value>"
	Write-Host "      <unit>Custom</unit>"
	Write-Host "      <customunit>ms</customunit>"
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <error>0</error>"
	Write-Host "  <text>Reply got from $($remoteip):$($remoteudpport) in $($ttl) hops</text>"
	Write-Host "</prtg>"
	}
	else {
		Write-Host "<prtg>"
		Write-Host "  <error>1</error>"
		Write-Host "  <text>Unabled to Connect to $($remoteip):$($remoteudpport) in $($maxttl) hops</text>"
		Write-Host "</prtg>"
	}
}
