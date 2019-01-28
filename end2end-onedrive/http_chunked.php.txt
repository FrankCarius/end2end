<?php
// Simple PHP Paket to return the timestamp with increasing delays to see TCP-Timeouts

// set Header to Plain Text
header("Content-Type: text/plain");

//Minutes to pause between Data
$interval = array(1,5,10,65,95,125,245,305,605);  

echo "MSXFAQ Session Timeout Test with HTTP Chunked Data \r\n";
flush();
ob_flush();
echo "Your IP Address and Port" .getenv('REMOTE_ADDR').":".getenv('REMOTE_PORT') ."\r\n";
echo str_pad("",5000," ");
// send 5000 character to make sure, that browsers start rendering
ob_end_flush();  // Disable output puffering
flush();
ob_flush();
echo "Start    " .date("d.m.Y - H:i:s",time()) ."\r\n";
foreach ($interval as $delay) {
    echo("Waiting $delay Seconds \r\n");
    flush();
    ob_flush();
    sleep($delay);
    echo "Timestamp" .date("d.m.Y - H:i:s",time()) ."\r\n";
}
echo "End     " .date("d.m.Y - H:i:s",time()) ."\r\n";
flush();
ob_flush();
?>
