# PRTG-SFBAssessment.ps1
# Uses the Microsoft Network Assessment Tool to check 
#
# 20161015 frank@carius.de
#	Initial Version
# 20180203 frank@carius.de
#	Updates for updated SFB Assessment tool Nov 2017
# 

param (
	$prtgserveruri = "http://192.168.100.59:5050/sfbassessment1",
	$resultcsvname = "$($env:LOCALAPPDATA)\Microsoft Skype for Business Network Assessment Tool\performance_results.tsv" 
)

write-host "PRTG-SFBAssessment:Start"

write-host " Parameter resultcsvname: $($resultcsvname)"
write-host " Parameter prtgserveruri: $($prtgserveruri)"

while ($true) {
   write-host "PRTG-SFBAssessment: Execute Assessment tool"
   Start-Process -FilePath .\NetworkAssessmentTool.exe -NoNewWindow -Wait

   write-host "PRTG-SFBAssessment: Parsing Result TSV"
   $result = Import-Csv $resultcsvname -Delimiter "`t"

   write-host " PacketLossRate      : "+(($result | select packetlossrate -ExpandProperty packetlossrate | sort)[-2]).replace(",",".")
   write-host " RoundTripLatencyInMs:"+(($result | select RoundTripLatencyInMs -ExpandProperty RoundTripLatencyInMs | sort)[-2]).replace(",",".")
   write-host " AverageJitterInMs   :"+(($result | select AverageJitterInMs -ExpandProperty AverageJitterInMs | sort)[-2]).replace(",",".")
   write-host " PacketReorderRatio  :"+(($result | select PacketReorderRatio -ExpandProperty PacketReorderRatio | sort)[-2]).replace(",",".")


   write-host "PRTG-SFBAssessment: Build PRTG XML"   
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

   write-host "PRTG-SFBAssessment: Post Result to PRTG-Server"
   $Answer=Invoke-Webrequest `
      -method "GET" `
      -URI ($prtgserveruri + "?content=$prtgresult")
   if ($answer.Statuscode -ne 200) {
      write-warning "Request to PRTG failed"
      exit 1
   }
}

write-host "PRTG-SFBAssessment:End"
