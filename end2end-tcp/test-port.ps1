# http://powershell.codeplex.com
# http://poshcode.org/2455

#- NOTES
#Author : Christophe CREMON (uxone) � http://powershell.codeplex.com
#Requires : PowerShell V2
#
#Test if a TCP Port is open or not.
#
#- EndPoint can be a hostname or an IP address
#
#- EXAMPLE
#
#Test-TCPPort -EndPoint server1 -Port 80 
# Return true if port is open, false otherwise

Function Test-TCPPort
{
	param ( [ValidateNotNullOrEmpty()]
	[string] $EndPoint = $(throw "Please specify an EndPoint (Host or IP Address)"),
	[string] $Port = $(throw "Please specify a Port") )
	
	$TimeOut = 1000
	$IP = [System.Net.Dns]::GetHostAddresses($EndPoint)
	$Address = [System.Net.IPAddress]::Parse($IP)
	$Socket = New-Object System.Net.Sockets.TCPClient
	$Connect = $Socket.BeginConnect($Address,$Port,$null,$null)
	if ( $Connect.IsCompleted )
	{
		$Wait = $Connect.AsyncWaitHandle.WaitOne($TimeOut,$false)
		if(!$Wait) 
		{
			$Socket.Close() 
			return $false 
		} 
		else
		{
			$Socket.EndConnect($Connect)
			$Socket.Close()
			return $true
		}
	}
	else
	{
		return $false
	}
}