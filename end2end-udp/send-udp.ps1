# send-udp
#
# Simple Tool to send UDP-Ports to a range with a various source port range.
# you can use Netmon/Wireshark on the target to see the incomping traffic.
# i have NO listener to listen on all ports
#
# 20181210. Enhanced to send Timestamp and Sequencenumber and HostFQDN

param (
	[Validatepattern("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")]
	[string]$remoteip = "192.168.100.1",	# IP to send to 

	[ValidateRange(1,65535)]
	[int]$remoteudpport=50002,		# port to send to
	[ValidateRange(0,65535)]
	[int]$sourceudpport = 0,		# SourcePort, 0 uses an available port
	[int]$packetcount = 1000,		# packetcount
	[int]$delayms = 20,				# throttling in ms
	[string]$payload="SendUDP Message by www.msxfaq.de 123456789012345678901234567890123456789012345678901234567890123456789012345678901234"   # additional Payload
)

set-psdebug -strict
$error.clear()

write-host "send-udp:Sending Packet"
write-host "send-udp:RemoteIPAddr :"$remoteip 
write-host "send-udp:RemoteUDPort :"$remoteudpport
write-host "send-udp:SourceUDPPort:"$sourceudpport
write-host "send-udp:Packet2Send  :"$packetcount
write-host "send-udp:Delayms      :"$delayms

try {
	$udpClient = new-Object system.Net.Sockets.Udpclient($sourceudpport)
	for ($i=1; $i -le $packetcount; $i++) {
		if (($i%100) -eq 0) {
			Write-host "Total Packets Sent $($i)"
		}
		$timestamp = get-date
		[string]$buffer = "$($timestamp.tostring("o")),$($i.ToString("00000000")),$($payload)"
		$byteBuffer  = [System.Text.Encoding]::ASCII.GetBytes(($Buffer))
		$sentbytes = $udpClient.Send($byteBuffer, $byteBuffer.length, $remoteip, $remoteudpport)
		if ($sentbytes -ne  $byteBuffer.length) {
			write-host "send-udp:Send Bytes Mismatch"
		}
		start-sleep -milliseconds ([math]::Max(0,[int]($delayms - ((get-date) - $timestamp).TotalMilliseconds)))
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
