# END2END-TCP
#
# Meaure the time to connect to a TCP-Port without TLS

param (
	[string[]] $hostlist = @("outlook.office365.com:443","sip.pstnhub.microsoft.com:5061","sip2.pstnhub.microsoft.com:5061","sip3.pstnhub.microsoft.com:5061"),
	[string] $Port = "443",
	[long]$timeout = 1000
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
		$Connection = $Socket.Connect($hostname.split(":")[0],$hostname.split(":")[1])
		if ($Socket.connected ) {
			$duration = ((get-date) - $start).totalmilliseconds
			write-host "Latency: $duration"
		}
		else {
			write-host "Timeout on connection - port not open"
			$duration=-1
		}
	}
	catch {
		write-host "Error waiting on connection"
		$duration=-1
	}
	$Socket.Close() 
}
write-host "End Checks"