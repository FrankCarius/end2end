Write-host "Generate-EXOLoad:Start"
Add-Type -AssemblyName System.Net.Http;
$winhttpclient = new-object System.Net.Http.HttpClient;
$winhttpclient.DefaultRequestHeaders.add("User-Agent", "Mozilla/5.0");
1..100|% {
	write-progress "Loop $($_) of 100"
 	$start=get-date
	$result = $winhttpclient.GetStringAsync("https://outlook.office365.com/owa/smime/owasmime.msi");
	$totalbytes = $result.Result.Length
	$duration = ((get-date)-$start).totalmilliseconds
	write-host "Go $($totalbytes) Bytes with kbit/Sek: $([int](($result.Result.Length)/$duration*10))"
}
Write-host "Generate-EXOLoad:End"
