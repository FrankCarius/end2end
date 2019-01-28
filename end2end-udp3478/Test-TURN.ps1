# Based on: http://blogs.technet.com/b/nettracer/archive/2013/10/13/testing-stun-turn-server-connectivity-from-external-networks-through-udp-3478-by-using-powershell.aspx

function ResolveLyncNames {
	param($LyncFqdn)
	try {
		$ipaddr = [System.Net.Dns]::GetHostAddresses($LyncFqdn)
		Write-Host "Successfully resolved $LyncFqdn to $ipaddr" -ForegroundColor Green
		return $ipaddr
	} catch {
		$exception = New-Object system.net.sockets.socketexception
		$errorcode = $exception.ErrorCode
		write-host "Requested name '$LyncFqdn' could not be resolved, error code:$errorcode" -ForegroundColor Red
		write-host "Error details: $exception" -ForegroundColor Red
		return 0
	}
}

function GetLogTime {
	return (Get-Date).ToString("HHmmss.ff")
}

if($args[0] -eq $NULL) {
	Write-host "Please specify a STUN server name or IP address to test"
	Write-Host "Example: Test-TURN.ps1 av.contoso.com"
	return
}

$stunserver1 = $args[0]
$job_logfile = "~\Documents\Test-TURN.log.txt"

if (ResolveLyncNames($stunserver1)) {
	try {
		$j = Start-Job -ScriptBlock {
			$stunserver2 = $args[0]
			$logfile = "~\Documents\Test-TURN.log.txt"
			Write-Output $null > $logfile
			
			$udpclient = new-Object System.Net.Sockets.UdpClient
			$udpclient.Connect($stunserver2,3478)
			# Added ReceiveTimeout to facilitate timely cleanup of the job.
            $udpclient.client.ReceiveTimeout = 15000 
			Write-Output "$((Get-Date).ToString("HHmmss.ff")) Connected to $stunserver2 port 3478" >> $logfile
			
			$listenport = $udpclient.client.localendpoint.port
			$endpoint = new-object System.Net.IPEndPoint([IPAddress]::Any, $listenport)
			Write-Output "$((Get-Date).ToString("HHmmss.ff")) Listening on port $listenport" >> $logfile
				
			# Minimal STUN/TURN datagram with 0-byte Username attribute.
			# Generates a response from a MS Lync 2010 Edge server.
			# For more info:
			# [MS-TURN]: Traversal Using Relay NAT (TURN) Extensions
			# https://msdn.microsoft.com/en-us/library/vs/alm/cc431507(v=office.12).aspx
			#
			# Decoded datagram: Total bytes: 32 (20-byte header + 12-byte message)
			# BEGIN TURN MESSAGE HEADER
			# 0x00,0x03				Message Type: Allocate Request
			# 0x00,0x0C				Message Length: 12 bytes (+ 20 byte header = 32 bytes total)
			# 0x21,0x12,0xa4,0x42	Magic Cookie: Always 0x2112A442
			# 0xf1,0x3c,0x08,0x4b	Transaction ID (Random number < 2^96, see section 9.3 of https://www.ietf.org/rfc/rfc3489.txt)
			# 0x80,0x18,0x17,0x72	Transaction ID (continued)
			# 0x47,0x49,0x30,0x65	Transaction ID (continued)
			# END TURN MESSAGE HEADER
			# 0x00,0x0f				Attribute Type: Magic Cookie
			# 0x00,0x04				Attribute Length: 4 bytes
			# 0x72,0xc6,0x4b,0xc6	Attribute Value: Always 0x72c64bc6
			# 0x00,0x06				Attribute Type: Username
			# 0x00,0x00				Attribute Length: 0 bytes

			[Byte[]] $payload =
			0x00,0x03,0x00,0x0c,0x21,0x12,0xa4,0x42,0xf1,0x3c,0x08,0x4b,
			0x80,0x18,0x17,0x72,0x47,0x49,0x30,0x65,0x00,0x0f,0x00,0x04,
			0x72,0xc6,0x4b,0xc6,0x00,0x06,0x00,0x00

			$bytesSent = $udpclient.Send($payload,$payload.length)
			Write-Output "$((Get-Date).ToString("HHmmss.ff")) Sent datagram: $bytesSent bytes" >> $logfile
			
			Write-Output "$((Get-Date).ToString("HHmmss.ff")) Listening for response..." >> $logfile
			$content = $udpclient.Receive([ref]$endpoint)

			if ($content) { Write-Output "$((Get-Date).ToString("HHmmss.ff")) Received response: $($content.length) bytes" >> $logfile}
			else { Write-Output "$((Get-Date).ToString("HHmmss.ff")) Null response." }
			
			write-Output "(Begin Raw Response)" >> $logfile
			if ($content) { Write-Output "$([BitConverter]::ToString($content))" >> $logfile }
			write-Output "(End Raw Response)" >> $logfile

			$Encoding = "ASCII"

			switch ( $Encoding.ToUpper() ) {
				"ASCII" { $EncodingType = "System.Text.ASCIIEncoding" }
				"UNICODE" { $EncodingType = "System.Text.UnicodeEncoding" }
				"UTF7" { $EncodingType = "System.Text.UTF7Encoding" }
				"UTF8" { $EncodingType = "System.Text.UTF8Encoding" }
				"UTF32" { $EncodingType = "System.Text.UTF32Encoding" }
				Default { $EncodingType = "System.Text.ASCIIEncoding" }
			}

			$Encode = new-object $EncodingType
			
			write-Output "(Begin Decoded Response)" >> $logfile
			if ($content) { Write-Output "$($Encode.GetString($content))" >> $logfile }
			write-Output "(End Decoded Response)" >> $logfile

			if ($content) {
				if ($Encode.GetString($content).Contains("The request did not contain a Message-Integrity attribute")) {
					write-host "Received response from STUN Server!" -ForegroundColor Green
				} else {
					write-host "Received unfamiliar response from STUN Server!" -ForegroundColor Green
				}
			} else {
				write-host "STUN Server either is not reachable or doesn't respond." -ForegroundColor Red
			}
			$udpclient.Close()
		} -ArgumentList $stunserver1
		
		write-host "Sending TURN server port allocation request at UDP port 3478, it will be checked after 10 seconds to see if a response is received or not ..." -ForegroundColor Green
		Start-Sleep -Seconds 10

		Write-Host "Job Log:"
		Get-Content $job_logfile | Write-Host
		
		if( $j.JobStateInfo.State -ne "completed" ) {
			Write-Host "$(GetLogTime) Job State: $($j.JobStateInfo.State)"
			Write-Host "The request timed out, STUN Server '$stunserver1' is not reachable or doesn't respond to the request." -ForegroundColor Red
			Write-host "Please check if UDP 3478 is allowed between the client and the STUN server" -ForegroundColor Red
			Write-Host "Also please check if STUN server (MediaRelaySvc.exe) is running on the Edge server and listening on UDP 3478" -ForegroundColor Red
		} else {
			$results = Receive-Job -Job $j
			$results
		}
		
		# Cleanup
		# Note: If UDP Client ReceiveTimeout is not set, then Stop-Job can take a couple of minutes to complete.
		Stop-Job -Job $j
		Write-Host "$(GetLogTime) Stopped Job."
		
		Remove-Job -Job $j -Force
		Write-Host "$(GetLogTime) Removed Job."
	} catch {
		$_.exception.message
	}
}

