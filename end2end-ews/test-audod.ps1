param (
	$mailboxsmtp = "frank.carius@netatwork.de"
)

Write-host "Lade EWS.DLL"
[void][Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll")
Write-host "End2End-EWS:Creating EWS Service Class"
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
Write-host "User Default Credentials"
$service.UseDefaultCredentials = $true
#$service.Credentials = New-Object System.Net.NetworkCredential($username, $password, $domain)
write-host "Disable SCP Lookup"
$service.EnableScpLookup = $false
Write-host " Enable Tracing"
#$service.TraceEnabled = $true
#$service.TraceListener = TraceListenerInstance 

Write-Host "Start Autodiscover"
$service.AutodiscoverUrl($mailboxSMTP,{$true})
