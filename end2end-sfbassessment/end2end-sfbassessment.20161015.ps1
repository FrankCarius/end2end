# PRTG-SFBAssessment.ps1
#
# Uses the Microsoft Network Assessment Tool to check 
#
# 20161015 frank@carius.de
#	Initial Version

param (
   $prtgserveruri = "http://192.168.100.1:5050/sfbassessment"
)


while ($true) {
   Start-Process -FilePath .\NetworkAssessmentTool.exe -NoNewWindow -Wait

   $result = Import-Csv .\results.tsv -Delimiter "`t"
   

   $prtgresult = '<?xml version="1.0" encoding="UTF-8" ?>
      <prtg>
         <result>
            <channel>Packet loss rate</channel>
            <value>' + (($result | select packetlossrate -ExpandProperty packetlossrate | sort)[-2]).replace(",",".") +'</value>
            <float>1</float>
         </result>
         <result>
            <channel>RTT latency</channel>
            <value>' + (($result | select RoundTripLatencyInMs -ExpandProperty RoundTripLatencyInMs | sort)[-2]).replace(",",".") +'</value>
            <float>1</float>
         </result>
         <result>
            <channel>Jitter</channel>
            <value>' + (($result | select AverageJitterInMs -ExpandProperty AverageJitterInMs | sort)[-2]).replace(",",".") +'</value>
            <float>1</float>
         </result>
         <result>
            <channel>Packet reorder ratio</channel>
            <value>' + (($result | select PacketReorderRatio -ExpandProperty PacketReorderRatio | sort)[-2]).replace(",",".") +'</value>
            <float>1</float>
         </result>
      </prtg>'

#$prtgresult | out-file result.xml

   $Answer=Invoke-Webrequest `
      -method "GET" `
      -URI ($prtgserveruri + "?content=$prtgresult")
   if ($answer.Statuscode -ne 200) {
      write-warning "Request to PRTG failed"
      exit 1
   }
}
