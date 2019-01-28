# End2End-udp3478
# Simple Skype for Business Online Edge Turn test
#
# 20180414 Added PRTG output
# 20180416 added RTT min/max/avg
# 20180425 fixes for PowerShell2 and PRTG detection and using 50019 as source port. 50000 is used to often
# 20180504 Erweiterung ArchivCSV
# 20180507 Umstellen auf PRTG Push und Dauerbetrieb
# 20181123 Fix bei Aufruf von sendto-prtg
# 20181205 Erweiterung um Analyse der R�ckantwort
# 20181211 erweiterung um Teams IP
# 20190122 Pretty Output Colorcoding

param( 
	[cmdletbinding()]
	[int]$sourceudpport=50019,      # Source port to send UDP from. SfB normally uses 50000-50019
	[int]$remoteudpport=3478,       # Default STUN/TURN Port
	[string]$remoteip = "o365",     # use static IP-Address or O365 or Teams
	[byte]$maxttl=24,               # start with TTL32 for Countdown
	[int]$maxretries = 5,           # repeatcount if packet was not received
	[int]$averagecount=5,           # measure response with max TTL and calcualte average and max
	[long]$sleeptime = 10,          # time in seconds to sleep between two checks
	[switch]$sendtopipeline=$false, # 
	[string]$archivcsv = ".\End2End-udp3478_$($env:COMPUTERNAME).csv",   # store results
	[string]$prtgpushurl = "http://prtg:5050/end2end-udp3478_$($ENV:COMPUTERNAME)",
	#[string]$prtgpushurl = "",
	[switch]$prtg=$false
)

# -----------------------------------------------------
# sendto-prtg   helper to send data to PRTG
# -----------------------------------------------------
function sendto-prtg (
	[string]$prtgresult,   # XML Structure
	[string]$prtgpushurl)  # HTTP-PUSH-url
{
		if ($prtg){
		if ($prtgpushurl -eq "") {
			$Scriptname = (split-path $MyInvocation.ScriptName -Leaf).replace(".ps1","")
			$prtgpushurl=  "http://prtg:5050/$($scriptname)_$($env:COMPUTERNAME)"
		}
		write-host "SendTo-PRTG: Post Result to PRTGProbe $prtgpushurl"
		
		$Answer=Invoke-RestMethod `
			-method "GET" `
			-timeout 5 `
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
}

write-host "End2End-UDP3467:Start"

if ($remoteip -eq "o365") {
	write-host "End2End-UDP3467:Using Skype for Business Online Servers" -NoNewline
	[string]$remoteip = "13.107.8.2"   # Anycast IP of Office 365 TURN Servers
}
elseif ($remoteip -eq "teams") {
	write-host "End2End-UDP3467:Using Office 365 Microsoft Teams Server" -NoNewline
	[string]$remoteip = "52.113.193.5"   # Anycast IP of a teams Turn Server
}
else {
	write-host "End2End-UDP3467:Using given IP-Addresse" -NoNewline
}
Write-host " TURN-Server:$($remoteip)"


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

try{
	Write-host "End2End-UDP3478:Connect UDPClient to $($remoteip):$($remoteudpport)"
	$udpclient.Connect($remoteip, $remoteudpport)
}
catch {
	Write-host "Unable to Connect UDP-Socket SourcePort: $($sourceudpport) Remote: $($remoteip):$($remoteudpport)"
	exit 1
}

$continue=$true
write-host "End2End-UDP3478: Press X-key to stop graceful" -BackgroundColor Magenta
while ($continue) {

	$udpclient.ttl = $maxttl
	
	write-verbose " Init summary Table"
	$summary = New-Object PSObject -Property @{
			remoteip=$remoteip
			remoteudpport =$remoteudpport
			maxroundtrip=[int]0
			minroundtrip=[int]9999
			averagertt=[int]0
			successrttcount=[int]0
			failedrttcount=[int]0
			answerreceived=[boolean]$false
	}

	[int]$summary.maxroundtrip=0
	[int]$summary.minroundtrip=9999
	[int]$summary.averagertt=0
	[int]$summary.successrttcount=0
	[int]$summary.failedrttcount=0
	[boolean]$summary.answerreceived=$false

	write-verbose " Calculate average Roundtrip "
	for ($loopcount=0; $loopcount -lt $averagecount; $loopcount++) {
		$starttime = get-date
		$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
		try {
			$receiveddata = $udpClient.Receive([ref]$remoteIpendpoint)
			
			if ($receiveddata){
				write-Verbose "Answer received"
				$ASCIIEncoder = new-object System.Text.ASCIIEncoding
				if ($ASCIIEncoder.GetString($receiveddata).Contains("The request did not contain a Message-Integrity attribute")) {
					write-host "M" -ForegroundColor Green -nonewline
					$summary.successrttcount++
					$roundtriptimems = [int](((get-date) - $starttime).totalmilliseconds)
					write-verbose " $($roundtriptimems ) ms"
					$summary.averagertt = $summary.averagertt + $roundtriptimems
					if ($summary.maxroundtrip -lt $roundtriptimems) {
						$summary.maxroundtrip = $roundtriptimems
					}
					if ($summary.minroundtrip -gt $roundtriptimems) {
						$summary.minroundtrip = $roundtriptimems
					}
				} else {
					write-host "F" -ForegroundColor red -NoNewline
					$summary.failedrttcount++
				}
			}
			else {
				Write-host "?" -foregroundcolor yellow -NoNewline
				$summary.failedrttcount++
			}
		}
		catch {
			#$_
			write-host "NO Answer Received 2" -foregroundcolor yellow
			$summary.failedrttcount++
		}
	}

	if ($summary.successrttcount -eq 0) {
		$summary.averagertt = 0
	}
	else {
		$summary.averagertt = [int]($summary.averagertt / $summary.successrttcount)
		write-host "Min/Avg/Max:$($summary.minroundtrip)/$($summary.averagertt)/$($summary.maxroundtrip)" -nonewline
		write-verbose " Start TTL Distance Check"
		[int]$retrycount=1
		$result=@{}
		for ($ttl=$maxttl; $ttl -gt 0 ; $ttl--) {
			Write-Verbose "  Send STUN-Request TTL $($ttl) "
			$udpclient.ttl = $ttl
			$starttime = get-date
			$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length)
			try {
				$result[$ttl]=$udpClient.Receive([ref]$remoteIpendpoint)
				write-host " $($ttl)" -BackgroundColor green -ForegroundColor black -NoNewline
				$summary.answerreceived=$true
			}
			catch {
				#$_
				write-host " $($ttl)" -BackgroundColor red -ForegroundColor black -NoNewline
				write-verbose "  NO Answer received. Try $($retrycount)"
				if ($retrycount -ge $maxretries){
					break
				}
				$retrycount++
				$ttl++
			}
		}
		if ($summary.answerreceived) {
			write-host " Distance $($ttl+1)"
		}
		else {
			write-host " NoConnection"
		}
	}

	Write-Verbose " Closing UDP"

	if ($archivcsv -ne ""){
		write-verbose " Export Data to ArchivCSV $($archivcsv)"
		if ($sendtopipeline){
			$summary
		}
		$summary | export-csv -path $archivcsv -append -notypeinformation
	}

	Write-Verbose "End2End-UDP3478:Done"
	#if ($myinvocation.MyCommand.Definition -match "\\EXEXML\\" -or $prtg) {
	if ($prtgpushurl -ne "") {
		write-verbose "end2end-udp3478: Build PRTG XML"   
		$prtgresult = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		if ($summary.answerreceived) {
			$prtgresult+= "<prtg>"
			$prtgresult+=  "  <result>"
			$prtgresult+=  "    <channel>Hopcount</channel>"
			$prtgresult+=  "      <value>$($ttl+1)</value>"
			$prtgresult+=  "      <unit>Custom</unit>" 
			$prtgresult+=  "      <customunit>Hops</customunit>"
			$prtgresult+=  "    <float>0</float>" 
			$prtgresult+=  "  </result>"
			$prtgresult+=  "  <result>"
			$prtgresult+=  "    <channel>RTT Avg</channel>"
			$prtgresult+=  "      <value>$($summary.averagertt)</value>"
			$prtgresult+=  "      <unit>Custom</unit>"
			$prtgresult+=  "      <customunit>ms</customunit>"
			$prtgresult+=  "    <float>0</float>" 
			$prtgresult+=  "  </result>"
			$prtgresult+=  "  <result>"
			$prtgresult+=  "    <channel>RTT Max</channel>"
			$prtgresult+=  "      <value>$($summary.maxroundtrip)</value>"
			$prtgresult+=  "      <unit>Custom</unit>"
			$prtgresult+=  "      <customunit>ms</customunit>"
			$prtgresult+=  "    <float>0</float>" 
			$prtgresult+=  "  </result>"
			$prtgresult+=  "  <result>"
			$prtgresult+=  "    <channel>RTT Min</channel>"
			$prtgresult+=  "      <value>$($summary.minroundtrip)</value>"
			$prtgresult+=  "      <unit>Custom</unit>"
			$prtgresult+=  "      <customunit>ms</customunit>"
			$prtgresult+=  "    <float>0</float>" 
			$prtgresult+=  "  </result>"
			$prtgresult+=  "  <error>0</error>"
			$prtgresult+=  "  <text>Reply got from $($remoteip):$($remoteudpport) in $($ttl) hops</text>"
			$prtgresult+=  "</prtg>"
		}
		else {
			$prtgresult+=  "<prtg>"
			$prtgresult+=  "  <error>1</error>"
			$prtgresult+=  "  <text>Unabled to Connect to $($remoteip):$($remoteudpport) in $($maxttl) hops</text>"
			$prtgresult+=  "</prtg>"
		}
		
		if (Get-Command "sendto-prtg" -errorAction SilentlyContinue) {
			SendTo-PRTG $summary
			Sendto-PRTG $prtgresult $prtgpushurl
		}
	}

	write-Verbose "Sleeping $($sleeptime) Seconds. Press ""X""-key to stop after next try"
	
	for($count = $sleeptime; $count -gt 0; $count--) {
		write-progress -activity "Sleeping  X=eXit script,  C=Continue" -SecondsRemaining $count
		#write-host "`b`b$($count) " -NoNewline
		start-sleep -seconds 1
		if ([console]::KeyAvailable) {
			$keycode = ([System.Console]::ReadKey()).key
			if ($keycode -eq "X") {
				write-host " Stopping Script"
				$count = 0
				$continue = $false
			}
			elseif ($keycode -eq "C") {
				$count = 0
			}
		}
	}
	write-progress -activity "Sleeping  X=eXit script,  C=Continue" -Completed
}
write-host "End2End-UDP3478:Close UDP-Socket"
# $udpClient.dispose()  # method not avaiable in .net 2.0
$udpClient.close()
write-host "End2End-UDP3478:End"
