Add-Type -AssemblyName System.Net.Http;
write-host " Start TCP Connections $((get-nettcpconnection).count)"
1..1000|% {
   write-progress "Loop $($_) of 1000"
   $winhttpclient = new-object System.Net.Http.HttpClient;
   $winhttpclient.DefaultRequestHeaders.add("User-Agent", "Mozilla/5.0");
   $null = $winhttpclient.GetStringAsync("https://outlook.office365.com/favicon.ico");
}
write-host " End TCP Connections $((get-nettcpconnection).count)"
