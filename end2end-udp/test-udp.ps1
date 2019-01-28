# send-udp
#
# Simple Tool to send UDP-Ports to a range with a various source port range.
# you can use Netmon/Wireshark on the target to see the incomping traffic.
# i have NO listener to listen on all ports

param (
	[Validatepattern("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")]
	[string]$remoteip = "13.107.8.2",	# IP to send to   

	[ValidateRange(1,65535)]
	[int]$remoteudpport=3478,	# port to send to

	[ValidateRange(0,65535)]
	[int]$sourceudpport = 0,					# SourcePort, 0 uses an available port

	[string]$buffer = "SendUDP Message by msxfaq",

	[int]$packetcount = 100,					# packetcount

	[int]$delayms = 10,				# throttling in ms
	[byte}$maxttl=10
)

set-psdebug -strict
$error.clear()

write-host "send-udp:Sending Packet"
write-host "send-udp:RemoteIPAddr :"$remoteip 
write-host "send-udp:RemoteUDPort :"$remoteudpport
write-host "send-udp:SourceUDPPort:"$sourceudpport
write-host "send-udp:Buffer       :"$buffer
write-host "send-udp:Packet2Send  :"$packetcount
write-host "send-udp:Delayms      :"$delayms
write-host "send-udp:MaxTTL       :"$maxttl

try {
	$udpClient = new-Object system.Net.Sockets.Udpclient($sourceudpport)
	$byteBuffer  = [System.Text.Encoding]::ASCII.GetBytes($Buffer)
	$udpClient.ttl = $maxttl
	for ($i=0; $i -lt $packetcount; $i++) {
		$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length, $remoteip, $remoteudpport)
		if ($sentbytes -ne  $byteBuffer.length) {
			write-host "send-udp:Send Bytes Mismatch"
		}
		start-sleep -milliseconds $delayms
	}
}
catch {
	write-host "send-udp:Error found "$error
}
finally {
	write-host "send-udp:Closing UDPSocket"
	$udpclient.close()
	write-host "send-udp:End"
}
