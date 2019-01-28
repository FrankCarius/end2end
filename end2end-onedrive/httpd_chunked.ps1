# HTTPD-Chunked
#
# Simple HTTPListener to send back a timestamp with configurable intervals
# Helps to check chunked transfers and TCP-Timeouts
# Not MultiTheaded !

param (
    $interval = (1,5,10,65,95,125,245,305,605)
)
write-host "HTTPD-Chunked: Start"
$HttpListener = New-Object System.Net.HttpListener
$HttpListener.Prefixes.Add("http://+:8888/")

write-host "HTTPD-Chunked: Start Listener"
$HttpListener.Start()

While ($HttpListener.IsListening) {
    write-host "HTTPD-Chunked: Wait for Request"
    $HttpContext = $HttpListener.GetContext()
    $HttpRequest = $HttpContext.Request
    write-host "Request received: HttpRequest"
    Write-host "URL            : $($HttpRequest.URL)"
    Write-host "URL            : $($HttpRequest.URL)"
    Write-host "RemoteEndpoint : $($HttpRequest.RemoteEndPoint)"
    
    if($HttpRequest.HasEntityBody) {
    $RequestBody = New-Object System.IO.StreamReader($HttpRequest.InputStream)
        write-host " Request Body "
        Write-Output $RequestBody.ReadToEnd()
    }
    write-host "Building Reply"
    $HttpResponse = $HttpContext.Response
    $HttpResponse.Headers.Add("Content-Type","text/plain")
    $HttpResponse.SendChunked = $true
    
    write-host "Sending 5000 Spaces to force IE rendering chunked content"
    $ResponseBuffer = [System.Text.Encoding]::UTF8.GetBytes(((" ")*5000 + "Start: $(get-date)`r`n"))
    $HttpResponse.OutputStream.Write($ResponseBuffer,0,$ResponseBuffer.Length)
    $HttpResponse.OutputStream.flush()

    $interval | %{
        write-host " Sending next Chunk in $($_) Seconds"
        start-sleep -seconds $_
        $ResponseBuffer = [System.Text.Encoding]::UTF8.GetBytes("Nach $_ Sec: $(get-date)`r`n")
        $HttpResponse.OutputStream.Write($ResponseBuffer,0,$ResponseBuffer.Length)
        $HttpResponse.OutputStream.flush()
    }
    write-host " Done"
    #$HttpResponse.ContentLength64 = $ResponseBuffer.Length
    $HttpResponse.StatusCode = 200
    $HttpResponse.Close()
    Write-Output "" # Newline
    $HttpListener.Stop()
}
$HttpListener.Stop()
write-host "HTTPD-Chunked: End"
