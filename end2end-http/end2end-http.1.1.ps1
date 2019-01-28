# end2end HTTP
#
# Einfaches Skript, welches immer wieder die gleiche HTML-Seite ausliest und den 
#
#
# 20121201  Version 1.0 FC Initale Version
# 20121212  Version 1.1 FC Ausgsbe der Message auf erste 10 Zeichen beschränkt
param(
	[string]$URL			= "https://www.netatwork.de/loadbalancertest.html",  # url to test
	[long]$maxduration		= 100,										# maximum time in msec
	[string]$reportcsv		= "$PWD\end2endhttp.csv",					# CSV File
	[string]$smtpserver		= "",
	[string]$smtpto			= "frank.carius@netatwork.de",
	[string]$smtpfrom		= "end2endhttp@netatwork.de",
	[long]$sleeptime		= 1000
)

write-host "Start End2End HTTP"

write-host " initializing CSV-File"
$csvfile = New-Object System.IO.StreamWriter $reportcsv, $true  # append
$csvfile.WriteLine("timestamp,status,duration,url,Message")

write-host " initializing WebClient"
$webclient = New-Object Net.WebClient
$webclient.CachePolicy = new-object System.Net.Cache.RequestCachePolicy ("bypasscache")  # Skip Cache
$webclient.UseDefaultCredentials=$true

Write-host "Initializing Result Object"
$result = New-Object PSObject -Property @{
                Timestamp	= $null
                Status		= ""
                Message		= ""
                Duration	= $null
}

Write-host "Starting Tests.  end with any key"
while (!$host.UI.RawUI.KeyAvailable) {
                [datetime] $TimeStart = [datetime]::Now
                
                $Error.Clear()
                $result.timestamp                         = (Get-Date -Format "dd.MM.yyyy HH:mm:ss")
                
                try {
                               $result.Message             = $webclient.DownloadString($URL)
                }
                catch {
                               $result.Message             = $Error[0]
                               $error.clear()
                }
                $result.Duration             = ([datetime]::Now - $TimeStart).Milliseconds
                
                if ($result.Duration -ge $maxduration){
                               $result.status = "Fail"
                               if ($smtpserver -ne "") {
                                               send-mailmessage `
                                                               -from $smtpfrom `
                                                               -to $smtpto `
                                                               -subject "End2EndHTML: $url slower than $axduration msec" `
                                                               -body "End2EndHTML: $url slower than $maxduration msec" `
                                                               -smtpServer $smtpserver           
                               }
                }
                else {
                               $result.status = "OK"
                }
                $result.timestamp + "," +$result.status+ "," + $result.Duration +"," + $url + "," +$result.message.substring(0,10)
                $csvfile.WriteLine($result.timestamp + "," +$result.status+ "," + $result.Duration +"," + $url + "," + $result.message)
                start-sleep -Milliseconds $sleeptime
} 

Write-host "Closing CSV-File"
$csvfile.Close();
Write-host "End2End HTTP finished"

