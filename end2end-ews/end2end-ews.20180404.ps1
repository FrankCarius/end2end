# end2end-ews
#
# Dieses Script versucht den Posteingang und Kalender eines Postfach # �ber EWS Impersonation zu �ffnen # # Voraussettungen # - Exchange 2007 SP1+ # - NET 3.5 # - "Microsoft.Exchange.WebServices.dll
# Getting started with the EWS Managed API # http://msdn.microsoft.com/en-us/library/dd633626(v=exchg.80).aspx
# 20180306 inital Version based on test-ews 1.2 
# 20180404 Fix zu BEServer-Ausgaben
#
param(
	[string]$MailboxSMTP = "frank.carius@netatwork.de",	# must be primary SMTP for impersonation
	[string]$Username    = "frank.carius@netatwork.de",				# use default credentials, if empty
	[string]$Domain      = "",				# Domain of the authentication user
    [double]$interval = 10,   # seconds to measure one interval
	[string]$Password    = ((get-credential -username $Username -message "Password for $Domain\$Username").getnetworkcredential().password),  # password, if $username is set
	[string]$Serviceurl  = "",						# Use Autodiscover, if empty https://exchange.msxfaq.local/EWS/Exchange.asmx",
    [string]$csvfile = ".\end2end-ews.csv",
	[switch]$useImpersonation = $false,						# forces impersonation 
	#[string]$dllpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll",
	#[string]$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll", # 
	#[string]$dllpath = "C:\nProgram Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll", # EWS DLL
	[switch]$EWSTrace,				 				# enable tracing of EWS to STDOUT
	[switch]$Verbose								# enable verbose output
)

if ($Verbose) {
	$VerbosePreference = "continue" 
}
Write-verbose "Test-EWS: Start"
Write-Verbose "Loading EWS DLL"
Write-Verbose "Username: $username"
Write-Verbose "Domain  : $domain"
Write-Verbose "password: <not visible>"

[void][Reflection.Assembly]::LoadFile($dllpath)
Write-Verbose "Creating EWS Service Class"
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService('Exchange2007_SP1')
#$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
if ($ewsTrace) {
	Write-Verbose " Tracing: enabled"
	$service.TraceEnabled = $true
	#Tracing EWS requests http://msdn.microsoft.com/en-us/library/dd633676(v=exchg.80).aspx
}
# --------------------------- Credentials and Impersonation --------------------------- 
if ($username -eq "") {  Write-Verbose "Credentials: UseDefaultCredentials"
	$service.UseDefaultCredentials = $true
}
else  {
	Write-Verbose "Credentials: use alternate Credentials"
	$service.UseDefaultCredentials = $false  
	$service.Credentials = New-Object System.Net.NetworkCredential($username, $password, $domain) 
} 
if ($useImpersonation) {
	Write-Verbose "Credentials: use impersonation"
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxSMTP);

	#Configuring Exchange Impersonation (Exchange Web Services)  #http://msdn.microsoft.com/en-us/library/bb204095(v=exchg.80).aspx
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SID,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.PrincipalName,wert)
	#$service.impersonatedUserID = new impersonatedUserID(ConnectingIDType.SmtpAddress,wert)
}
# --------------------------- ServiceURI --------------------------- 
if ($serviceurl -eq "") {  
	Write-Verbose "ServiceURL: Using Autodiscover for $mailboxSMTP"
	$service.AutodiscoverUrl($mailboxSMTP,{$true})
}
else {
	Write-Verbose "ServiceURL: using specified $serviceurl"
	$service.Url = New-Object System.Uri($serviceurl) 
} 
write-verbose -Message ("ServiceURI="+$service.Url.AbsoluteUri)
# --------------------------- Connect --------------------------- 

# Optional. Query Autodiscover for testing and retrieving Archive mailboxes etc
write-verbose "Checking Autodiscover"
#$autod = new-object Microsoft.Exchange.WebServices.Autodiscover.autodiscoverservice($mailboxSMTP.split("@")[1])
$autod = new-object Microsoft.Exchange.WebServices.Autodiscover.autodiscoverservice
$autod.EnableScpLookup = $false
#$autod.Credentials =
$autod.RedirectionUrlValidationCallback = {$true}   # ignore Redirect requests
$userSettings = ( $autod.GetUserSettings($mailboxSMTP,"AlternateMailboxes")).settings
$AlternateMailboxlist = $UserSettings["AlternateMailboxes"].entries
$archivmailboxSMTP = ($AlternateMailboxlist | ?{$_.Type -eq "Archive"}).smtpaddress
Write-Verbose "ArchivSMTP = $($archivmailboxSMTP)"

$mbMailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($archivmailboxSMTP)

#$mbMailbox = new-object Microso\\dsft.Exchange.WebServices.Data.Mailbox($mailboxSMTP)
Write-Verbose "Binding Inbox"
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
write-host "Number or Unread Messages : " $inbox.UnreadCount 

 $statistics = New-Object PSObject -Property @{
    timestamp = (get-date)
    durationavg=[double]9999
    durationmax=[double]0
    durationmin=[double]9999
    beserver   =[string]""
  }

[double]$durationsum=0
[double]$count=0
$result = $inbox.finditems(1)  # first search not measured to miss long latency due initial call

while ($true){
    $count++
    $start = get-date
    $result = $inbox.finditems(1)
    $duration = ((get-date) - $start).totalmilliseconds
    if ($statistics.beserver -ne $result.service.HttpResponseHeaders["X-BEServer"]){
        $statistics.beserver = $result.service.HttpResponseHeaders["X-BEServer"]
        write-host "New X-BEServer $($statistics.beserver) detected"
    }
    $durationsum+=$duration
    if ($statistics.durationmax -lt $duration){
        $statistics.durationmax =$duration
    }
    if ($statistics.durationmin -gt $duration){
        $statistics.durationmin = $duration
    }
    if ($count -ge $interval){
        $statistics.durationavg = ($durationsum/$count)
        $durationsum=0
        $statistics.timestamp = $start
        $statistics
        $statistics | export-csv $csvfile -Append -NoTypeInformation
        $count = 0
        $statistics.durationmin=9999
        $statistics.durationmax=0
    } 
    write-host $duration
    start-sleep -Milliseconds (1000-(get-date).Millisecond)
}

Write-Verbose "test-ewsimpersonation: End"