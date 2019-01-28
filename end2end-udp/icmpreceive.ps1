#ICMP-Receiver
$socket = New-Object System.Net.Sockets.Socket(
						[System.Net.Sockets.AddressFamily]::InterNetwork,
						[System.Net.Sockets.SocketType]::Raw,
						[System.Net.Sockets.ProtocolType]::Unspecified)
$RemoteIpEndPoint = New-Object System.Net.IPEndPoint([system.net.IPAddress]::Parse("0.0.0.0"),0);
$socket.Bind($RemoteIpEndPoint);
$socket.iocontrol([ Net.Sockets.IOControlCode]::ReceiveAll, 
					[BitConverter]::GetBytes(1),
					$null) 
$ReceiveBuffer = New-Object Byte[] 256
while ($true) {
	write-host "Wait for ICMP Packet"
	try {
		$inBytes = $socket.ReceiveFrom($ReceiveBuffer, $ReceiveBuffer.length, 0, [ref]$RemoteIpEndPoint);
		if ($ReceiveBuffer[20] -eq 11) {   # ICMP type = Delivery failed
			Write-host "TTL Exceeded"
			Write-host "  Sender: $($RemoteIpEndPoint.ToString())"
			Write-host "Destination: $($ReceiveBuffer[44]).$($ReceiveBuffer[45]).$($ReceiveBuffer[46]).$($ReceiveBuffer[47])"
		}
		else {
			Write-host "ICMP-Paket $($ReceiveBuffer[20]) received"
		}
	}
	catch {
		Write-host "ICMP-Paket Exception $_ received"
	}
}
$socket.close()
$socket.dispose()