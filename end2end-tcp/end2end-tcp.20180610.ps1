# END2END-TCP
#
# Meaure the time to connect to a TCP-Port without TLS
# https://www.msxfaq.de/signcrypt/zertifikate/get-tcpcert.htm
#
# Pending: OUtput as Object for "sendto-prtg", export-csv etc
#
param (
	[string[]] $hostlist = @("outlook.office365.com:443","sip.pstnhub.microsoft.com:5061","sip2.pstnhub.microsoft.com:5061","sip3.pstnhub.microsoft.com:5061"),
	[string] $Port = "443",
	[long]$timeout = 1000,
	[switch]$tlshandshake=$false
)

#$IP = [System.Net.Dns]::GetHostAddresses($hostname)
#$Address = [System.Net.IPAddress]::Parse($IP[0])
#$Address = $IP[0].IPAddressToString
write-host "Start Checks"
foreach ($hostname in $hostlist) {
	$Socket = New-Object System.Net.Sockets.TCPClient
	$Socket.Sendtimeout=100
	$Socket.Receivetimeout=100
	write-host "Host: $($hostname) " -nonewline
	$start = get-date
	try{
		# try TCP Connection
		$Socket.Connect($hostname.split(":")[0],$hostname.split(":")[1])
	}
	catch {
		$_
		write-host "Error waiting on connection"
		$duration=-1
	}
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
				$sslStream.AuthenticateAsClient($hostname.split(":")[0])
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
		write-host "Timeout on connection - port not open"
		$duration=-2
	}
	$Socket.Close() 
	write-host " Done"
}
write-host "End Checks Duration: $($duration)"