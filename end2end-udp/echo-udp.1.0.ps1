# Echo-udp
#
# Start a UDP-Listener ona given Ip:Port and wait for packets to echo bach
#
# Links: http://pshscripts.blogspot.de/2008/12/send-udpdatagramps1.html

param (
	[int]$listenudpport = 50000,
	[boolean]$debug = $false
)

set-psdebug -strict
$error.clear()

write-host "echo-udp: Start"
try {
	$udpClient = New-Object system.Net.Sockets.Udpclient($listenudpport)
	$RemoteIpEndPoint = New-Object System.Net.IPEndPoint([system.net.IPAddress]::Parse("0.0.0.0")  , 0);
	Write-host "echo-udp:Wait for Data on Port: $listenudpport"
	while ($true) {
		$data= $udpclient.receive([ref]$RemoteIpEndPoint)  # wait for data arriving from any ip. Alternativ kann man eine Liste der erlaubten IPs angeben
		write-host "." -nonewline
		if ($debug) {
			write-host "echo-udp:Received packet from IP " $RemoteIpEndPoint.address ":" $RemoteIpEndPoint.Port
			write-host "echo-udp:Content" ([string]::join("",([System.Text.Encoding]::ASCII.GetChars($buffer))))
		}

		$sentbytes = $udpClient.Send($data, $data.length, $RemoteIpEndPoint.address, $RemoteIpEndPoint.port)
		if ($sentbytes -ne  $data.length) {
			write-host "echo-udp:Send Bytes Mismatch"
		}
	}
}
catch {
	write-host "echo-udp:Error occured $error"
}
finally {
	write-host "echo-udp:Closing"
	write-host "echo-udp: End"
	$udpclient.close()
}

