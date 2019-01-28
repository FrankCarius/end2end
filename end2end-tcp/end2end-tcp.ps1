# END2END-TCP
#
# Meaure the time to connect to a TCP-Port without TLS
# https://www.msxfaq.de/signcrypt/zertifikate/get-tcpcert.htm
# https://www.powershellgallery.com/packages/TestTCPConnection/1.0.0.0/Content/TestTCPConnection.psm1  implementiert Callback falsch
# https://blogs.msdn.microsoft.com/joncole/2007/06/13/sample-asynchronous-sslstream-clientserver-implementation/
#
# Pending: Output as Object for "sendto-prtg", export-csv etc
#
param (
	[string[]] $hostlist = @("outlook.office365.com:443","sip.pstnhub.microsoft.com:5061","sip2.pstnhub.microsoft.com:5061","sip3.pstnhub.microsoft.com:5061"),
	[long]$tcptimeout = 1000
)

#$IP = [System.Net.Dns]::GetHostAddresses($hostname)
#$Address = [System.Net.IPAddress]::Parse($IP[0])
#$Address = $IP[0].IPAddressToString
write-host "Start Checks:  tcptimeout = $($tcptimeout)"
foreach ($hostname in $hostlist) {
	write-host "Host: $($hostname) " -nonewline
	$Socket = New-Object System.Net.Sockets.TCPClient
	$start = get-date
	$null = $Socket.BeginConnect(($hostname.split(":")[0]),($hostname.split(":")[1]),$null,$null) # no callback, no state
	do {
		$duration = ((get-date) - $start).totalmilliseconds
	} while (!$socket.Connected -and ($duration -lt $tcptimeout))

	if ($Socket.connected ) {
		$duration = ((get-date) - $start).totalmilliseconds
		if ($duration -lt 50) { 
			write-host "LatencyConnect: $($duration) " -nonewline -BackgroundColor Green -ForegroundColor Black
		}
		elseif ($duration -lt 200) { 
			write-host "LatencyConnect: $($duration) " -nonewline -BackgroundColor yellow -ForegroundColor Black
		}
		elseif ($duration -lt 500) { 
			write-host "LatencyConnect: $($duration) " -nonewline -BackgroundColor Magenta -ForegroundColor Black
		}
		else { 
			write-host "LatencyConnect: $($duration) " -nonewline -BackgroundColor red -ForegroundColor white
		}

		if ($hostname.split(":")[2] -eq "tls") { 
			$tcpstream = $Socket.GetStream()
			write-verbose " Get SSL-Stream"
			$sslStream = New-Object System.Net.Security.SslStream($tcpstream,$false,{$true})  # verification callback always true
			#$sslStream
			try {
				#$sslStream.AuthenticateAsClient($null)
				$sslStream.BeginAuthenticateAsClient($null)
				#write-host " TLSOK" -ForegroundColor green -nonewline
				$durationtls = ((get-date) - $start).totalmilliseconds
				write-host "LatencyTLS:" -nonewline
				if ($durationtls -lt 50) { 
						write-host "$($durationtls) " -nonewline -BackgroundColor Green -ForegroundColor Black
				}
				elseif ($durationtls -lt 200) { 
					write-host "$($durationtls) " -nonewline -BackgroundColor yellow -ForegroundColor Black
				}
				elseif ($durationtls -lt 500) { 
					write-host "$($durationtls) " -nonewline -BackgroundColor Magenta -ForegroundColor Black
				}
				else { 
					write-host "$($durationtls) " -nonewline -BackgroundColor red -ForegroundColor white
				}
				$duration = $durationtls
			}
			catch {
				$_
				write-host " TLSFail" -ForegroundColor Yellow -nonewline
				$duration=-3
			}
		}
	}
	else {
		write-host "Timeout on connection - port not open" -BackgroundColor red
		$duration=-2
	}
	$Socket.Close() 
	write-host " Done"
}
write-host "End Checks Duration: $($duration)"