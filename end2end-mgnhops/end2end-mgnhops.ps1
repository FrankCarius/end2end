# end2end MGN-Hops
#
# Does a ping to outlook.office365.com and count the hops to the MGN
# Test-Netconnnection -traceroute is pretty slow
#
# Ping Class 
# https://msdn.microsoft.com/en-us/library/system.net.networkinformation.ping(v=vs.110).aspx 

param(
	[String]$pingtarget = "outlook.office365.com",
	[int]$maxdelay = 1000
)

$ping = new-object System.Net.NetworkInformation.Ping
$result1 = $ping.send($pingtarget,$maxdelay)

if ($result1.status -ne "Success") {
	write-error "Unable to Ping, check Network connection or target $(pingtarget)"
}
else {
	write-host "  Sucess to $($result1.Address)"
	write-host "  Roundtrip $($result1.RoundtripTime)"
	
	$pingoptions = new-object System.Net.NetworkInformation.PingOptions
	$buffer = [system.Text.Encoding]::ASCII.getbytes("end2end-mgnhops by www.msxfaq.de")
	foreach ($ttl in (1..($result1.roundtriptime))) {
		write-host " Ping with TTL $($ttl)" -nonewline
		$pingoptions.ttl = $ttl
		$result2 = $ping.send($result1.Address,$maxdelay,$buffer,$pingoptions)
	$result2 | select *
		if ($result2.status -eq "Success") {
			write-host "  Hop $($result2.address)"
		}
	}
}