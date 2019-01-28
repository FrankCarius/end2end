#
#      .SYNOPSIS
#            Test script to generate and analyze VoIP like load
#
#      .DESCRIPTION
#            generates and receives UDP-packets to check VoIP Problems
#
#      .PARAMETER  port
#            Specify the UDP-Port to send to and listen on 
#
#      .PARAMETER  Listen
#            Set to true, to listen for connections
#
#	   .PARAMETER  target
#            specify the name or ip-address of the receiving system
#
#      .EXAMPLE
#            PS C:\> adpt -cookiefilename C:\adpt\dirsync1.txt -ldappath LDAP://server/dc=domain.dc=tld
#            'This is the output'
#            This example shows how to call the Name function with positional parameters.
#
#      .INPUTS
#            System.String,System.Int32
#
#      .OUTPUTS
#            Objects with modifications
#
#      .NOTES
#            For more information about advanced functions, call Get-Help with any
#            of the topics in the links listed below.
#
#       .LINK
#           See http://www.msxfaq.net/adpt/index.htm
#
# Version 0.1-0.9 many iterations and tests


param (
	[string]$target = $null,	# Target IP-
	[int]$port = 11223,   		# udp-port to send to and listen from
	[switch]$listen  			# set to true to start receiving part	
)

Set-PSDebug -Strict

if (($target -eq $null) -and !$listen) {
	write-error "Specify target or use listenmode"
	break
}


if (!$listen) {
	write-host "Sending UDP pakets to $target on port $port"
	$udpobject = new-Object system.Net.Sockets.Udpclient
	$a = new-object system.text.asciiencoding
	[string]$sessionguid = [guid]::NewGuid().tostring()
	[string]$timestamp = (get-date).touniversaltime().tostring()
	[int]$usn=0
	do {
		$usn = $usn + 1
		[string]$timestamp = (get-date).touniversaltime().tostring()
		[string]$message = "end2endvoip!$timestamp!$usn!$sessionguid"
		if ($usn % 50 -eq 1) {write-host "`nSending $message" -nonewline}
		else {write-host "." -nonewline}
		$byte = $a.GetBytes($message)
		[void]$udpobject.Send($byte,$byte.length,$target,$port)
		start-sleep -milliseconds 100
	}
	until ($Host.UI.RawUI.KeyAvailable) 
	$Host.UI.RawUI.FlushInputBuffer()
	Write-Host "Closing UDP-Client"
	$udpobject.Close()
}
else {
	write-host "Listen for UDP pakets to $target on port $port"
	$udpobject = new-Object system.Net.Sockets.Udpclient($port)
	$ipendpoint = New-Object system.net.ipendpoint([system.net.ipaddress]::Any,$port)		
	do {
		write-host "Wait for incoming packets"
		#Blocks until a message returns on this socket from a remote host.
		[byte[]]$receivebytes = $udpobject.Receive([ref]$ipendpoint)
		[string]$receivestring = [Text.Encoding]::ASCII.GetString($receivebytes)
		Write-Host "SenderIP:port : $($ipendpoint.address.ToString()):$($ipendpoint.Port.ToString())"
		#Write-Host "Message   : $($receivestring)"
		if ($receivestring.startswith("end2endvoip!")) {
			Write-Host "OK end2endvoip Message found"
			$message = $receivestring.split("!")
			if ($message.count -eq 4 ){
				[string]$timestamp = $message[1]
				[int]$usn= $message[2]
				[string]$sessionguid = $message[3]
				Write-Host "Timestamp: $timestamp USN: $usn  GUID:$sessionguid"
			}
			else {
				Write-Host "SKIP Message not complete"
			}
		}
		else {
			Write-Host "SKIP Invalid message found $receivestring"
		}
	}
	until ($Host.UI.RawUI.KeyAvailable) 
	$Host.UI.RawUI.FlushInputBuffer()
	Write-Host "Closing UDP-Client"
	$udpobject.Close()
}

	 