# Simple Skype for Business Online Edge Turn test
#
# 20180414 Added PRTG output
# 20180416 added RTT min/max/avg

param (
	[int]$sourceudpport=50000,
	[int]$remoteudpport=3478,
	[string]$remoteip = "13.107.8.2",
	[byte]$maxttl=32,
	[int]$maxretries = 5,
	[int]$averagecount=5,  # measure response with max TTL and calcualte average and max
	[switch]$prtg=$false
)

Write-host "Start UDP-Client on $($sourceudpport)"
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
$result=@{}
Write-host "Connect to $($remoteip):$($remoteudpport)"
$udpclient.Connect($remoteip, $remoteudpport)


write-host "Calculate average Rounttrip "
[int]$maxroundtrip=0
[int]$minroundtrip=9999
[int]$averagertt=0
[int]$successrttcount=0
#$udpclient.ttl = 32

for ($loopcount=0; $loopcount -lt $averagecount; $loopcount++) {
	$starttime = get-date
	$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
	try {
		$null = $udpClient.Receive([ref]$remoteIpendpoint)
		write-host "Answer received" -nonewline
		$successrttcount++
		$roundtriptimems = [int](((get-date) - $starttime).totalmilliseconds)
		write-host " $($roundtriptimems ) ms"
		$averagertt = $averagertt + $roundtriptimems
		if ($maxroundtrip -lt $roundtriptimems) {
			$maxroundtrip = $roundtriptimems
		}
		if ($minroundtrip -gt $roundtriptimems) {
			$minroundtrip = $roundtriptimems
		}
	}
	catch {
		$_
		write-host "NO Answer received."
	}
}
if ($successrttcount -gt 0) {
	$averagertt = [int]($averagertt / $successrttcount)
		write-host "RTT Avg: $($averagertt)"
		write-host "RTT Max: $($maxroundtrip)"
		write-host "RTT Min: $($minroundtrip)"
}
else {
	$averagertt = 0
}
write-host "Start TTL Distance Check"
[boolean]$answerreceived=$false
[int]$retrycount=0
for ($ttl=$maxttl; $ttl -gt 0 ; $ttl--) {
	Write-host "Send STUN-Request with TTL $($ttl) " -nonewline
	$udpclient.ttl = $ttl
	$starttime = get-date
	$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
	try {
		$result[$ttl]=$udpClient.Receive([ref]$remoteIpendpoint)
		write-host "Answer received"
		$answerreceived =$true
	}
	catch {
		write-host "NO Answer received. Retry $($retrycount)"
		if ($retrycount -ge $maxretries){
			break
		}
		$retrycount++
		$ttl++
	}
}
Write-host "Closing UDP"
$udpClient.close()
$udpClient.dispose()
if ($answerreceived) {
	write-host "Result: Distance $($ttl+1)"
}
else {
	write-host "Result: NoConnection"
}

Write-host "Done"
if ($PSScriptRoot.EndsWith("EXEXML") -or $prtg) {
	if ($answerreceived) {
	Write-Host "<prtg>"
	Write-Host "  <result>"
	Write-Host "    <channel>Hopcount</channel>"
	Write-Host "      <value>$($ttl+1)</value>"
	Write-Host "      <unit>Hops</unit>" 
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Avg</channel>"
	Write-Host "      <value>$($averagertt)</value>"
	Write-Host "      <unit>ms</unit>" 
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Max</channel>"
	Write-Host "      <value>$($maxroundtrip)</value>"
	Write-Host "      <unit>ms</unit>" 
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <result>"
	Write-Host "    <channel>RTT Min</channel>"
	Write-Host "      <value>$($minroundtrip)</value>"
	Write-Host "      <unit>ms</unit>" 
	Write-Host "    <float>0</float>" 
	Write-Host "  </result>"
	Write-Host "  <error>0</error>"
	Write-Host "  <text>Reply got from $($remoteip):$($remoteport) in $($ttl) hops</text>"
	Write-Host "</prtg>"
	}
	else {
		Write-Host "<prtg>"
		Write-Host "  <error>1</error>"
		Write-Host "  <text>Unabled to Connect to $($remoteip):$($remoteport) in $($maxttl) hops</text>"
		Write-Host "</prtg>"
	}
}
