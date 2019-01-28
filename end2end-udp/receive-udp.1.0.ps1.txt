# receive-udp
#
# Start a UDP-Listener on a given Ip:Port and wait for packets

param (
	[int]$udplistenport = 50002,
	[boolean]$debug = $false
)

set-psdebug -strict
$error.clear()

try {
	Write-host "Receive-UDP:Start"
	$udpClient = New-Object system.Net.Sockets.Udpclient($udplistenport)
	$RemoteIpEndPoint = New-Object System.Net.IPEndPoint([system.net.IPAddress]::Parse("0.0.0.0")  , 0);
	Write-host "Receive-UDP:Wait for Data on Port: $udplistenport"
	while ($true) {
		$data= $udpclient.receive([ref]$RemoteIpEndPoint)  # wait for data arriving from any ip. Alternativ kann man eine Liste der erlaubten IPs angeben
		# Achtung: UDPClient puffert, d.h. wenn weiter Daten  kommen, werden diese gequeue
		if ($debug) {
			write-host "Received packet from IP " $RemoteIpEndPoint.address ":" $RemoteIpEndPoint.Port
			write-host "Content" ([string]::join("",([System.Text.Encoding]::ASCII.GetChars($buffer))))
		}
		
		# send structured Data to pipeline
		New-Object PSObject -Property @{
			RemoteIP = $RemoteIpEndPoint.address
			RemotePort = $RemoteIpEndPoint.Port
			bytes = $buffer
			text = ([string]::join("",([System.Text.Encoding]::ASCII.GetChars($buffer))))
			timestamp = (get-date -Format "yyyy/MM/dd HH:mm:ss.fff")
		}
	}
}
catch {
	write-host "Receive-UDP:Error occured $error"
}
finally {
	write-host "Receive-UDP:Closing"
	$udpclient.close()
	Write-host "Receive-UDP:End"
}