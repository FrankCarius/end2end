Add-Type -AssemblyName System.Net.Http;
$winhttpclient = new-object System.Net.Http.HttpClient;
$winhttpclient.DefaultRequestHeaders.add("User-Agent", "Mozilla/5.0");
1..10000|% {
    write-progress "Loop $($_) of 50000"
    $null = $winhttpclient.GetStringAsync("https://outlook.office365.com/owa/smime/owasmime.msi");
}

