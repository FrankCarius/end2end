# receive-udp
#
# Start a UDP-Listener on a given Ip:Port and wait for packets

param (
	[int]$listenport = 50002,
	[long]$maxlatency = 100,  # max allowed latency
	[boolean]$debug = $false
)

set-psdebug -strict
$error.clear()

write-host "send-udp:Waiting for Packet"
write-host "send-udp:ListenUDPPort: $($listenport)"
write-host "send-udp:maxlatency   : $($maxlatency)"

$lastpacket=@{}

try {
	Write-host "Receive-UDP:Start"
	$udpClient = New-Object system.Net.Sockets.Udpclient($listenport)
	$RemoteIpEndPoint = New-Object System.Net.IPEndPoint([system.net.IPAddress]::Parse("0.0.0.0") , 0);
	Write-host "Receive-UDP:Wait for Data on Port: $listenport"
	$count = 0
	while ($true) {
		$count++
		[string]$reportinfo = ""
		# wait for data arriving from any ip. Alternativ kann man eine Liste der erlaubten IPs angeben
		$data= $udpclient.receive([ref]$RemoteIpEndPoint)
		
		#Collect Timestamp
		$receivetimestamp = get-date -format o
		$receivetimestampdt = get-date $receivetimestamp
		
		# split payload to get values
		[string]$payload= ([string]::join("",([System.Text.Encoding]::ASCII.GetChars($data))))
		$Sender = "$($RemoteIpEndPoint.Address.IPAddressToString):$($RemoteIpEndPoint.Port)" 
		$sendtimestamp = $payload.split(",")[0]
		$sendtimestampdt = get-date $sendtimestamp
		[long]$sendersequence = $payload.split(",")[1]

		# Add Entry to Hash table, if not exisiting
		if (!$lastpacket.item($sender)) {
			write-host "  New Sender found from $($sender)"
			$lastpacket[$sender]= [pscustomobject][ordered]@{
					sender = [string]$sender
					sequence = $sendersequence - 1
					received = [long]0
					Missed = [long]0
					last = [long]0
					OutOfOrder = [long]0
					avglatencyms = [long]0
					sendtimestamp = [string]$sendtimestamp
					sendtimestampdt = $sendtimestampdt
					receivetimestamp = $receivetimestamp
					receivetimestampdt = $receivetimestampdt
			}
		}
		$lastpacket[$sender].received++

		if ($sendersequence -gt ($lastpacket[$sender].sequence + 1)) {
			write-host "Missed Packet from $($sender)  sendersequence:$($sendersequence)   LastpaketUSN $($lastpacket[$sender].sequence)" -foregroundcolor red
			[string]$reportinfo = "Missed Packet from $($sender)  sendersequence:$($sendersequence)   LastpaketUSN $($lastpacket[$sender].sequence)"
			$lastpacket[$sender].sequence = $sendersequence
			$lastpacket[$sender].Missed++
		}
		elseif ($sendersequence -lt ($lastpacket[$sender].sequence + 1)) {
			write-host "OutOfOrder Packet from $($sender)  sendersequence:$($sendersequence)   LastpaketUSN $($lastpacket[$sender].sequence)" -foregroundcolor mangeta
			[string]$reportinfo = "OutOfOrder Packet from $($sender)  sendersequence:$($sendersequence)   LastpaketUSN $($lastpacket[$sender].sequence)"
			$lastpacket[$sender].OutOfOrder++
		}
		else {
			# packet in order.
			$lastpacket[$sender].sequence = $sendersequence

			# oneway time. depends on accurate time on sender and recipient
			$onewaytime =(get-date $sendtimestamp) - $receivetimestampdt

			$latency = ($receivetimestampdt - $sendtimestampdt).totalmilliseconds
			
			if ($lastpacket[$sender].avglatencyms -eq 0 ) {
				# initial latency
				$lastpacket[$sender].avglatencyms = $latency
			}
			else {
				$lastpacket[$sender].avglatencyms = [math]::round((($lastpacket[$sender].avglatencyms) *9 + $latency)/10)
			}
			
			if ($latency -gt $maxlatency ) {
				write-host "Later than maxlatency Packet from $($sender) $($latency) MaxLatency=$($maxlatency)"
				[string]$reportinfo =  "Later than maxlatency Packet from $($sender) $($latency) MaxLatency=$($maxlatency)"
			}
			else {
				# latency OK, Update Timestamp
				$lastpacket[$sender].sendtimestampdt = $sendtimestampdt
				$lastpacket[$sender].sendtimestamp = $sendtimestamp
			}

			#Timegap between current and former packet
			$jitter = ($receivetimestampdt - $lastpacket[$sender].receivetimestampdt).totalmilliseconds
		}

		# Update TimeStamps
		$lastpacket[$sender].receivetimestamp = $receivetimestamp
		$lastpacket[$sender].receivetimestampdt = $receivetimestampdt

		if ($reportinfo -ne "") {
			write-host "Warnung $($reportinfo)" -foregroundcolor yellow
			$lastpacket.values
		}
		if (($count%100) -eq 0) {
			Write-host "Total Pakets received $($count)"
			$lastpacket.values
		}
		# send structured Data to pipeline
	}
}
catch {
	write-host "Receive-UDP:Error occured $error"
	$error
}

finally {
	write-host "Receive-UDP:Closing"
	$udpclient.close()
	Write-host "Receive-UDP:End"
}