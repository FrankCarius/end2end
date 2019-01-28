# PRTG-SFBAssessment.ps1
# Uses the Microsoft Network Assessment Tool to check 
#
# 20161015 frank@carius.de
#	Initial Version
# 20180203 frank@carius.de
#	Updates for updated SFB Assessment tool Nov 2017
# 20180423 frank@carius.de
#	Added parsing of result codes
#	Bugfix if only one measurement done and avg/max calculation
#	parameter sfbpath added
param (
	$prtgserveruri = "http://prtg:5050/sfbassessment1",
	$resultcsvname = "$($env:LOCALAPPDATA)\Microsoft Skype for Business Network Assessment Tool\performance_results.tsv",
	$sfbpath = "c:\program Files\Microsoft Skype for Business Network Assessment Tool\"
)

write-host "PRTG-SFBAssessment:Start"

write-host " Parameter resultcsvname: $($resultcsvname)"
write-host " Parameter prtgserveruri: $($prtgserveruri)"

while ($true) {
   write-host "PRTG-SFBAssessment: Execute Assessment tool"
   Start-Process -FilePath ($sfbpath + "NetworkAssessmentTool.exe") -NoNewWindow -Wait -workingDirectory $sfbpath

   write-host "PRTG-SFBAssessment: Parsing Result TSV"
   $result = Import-Csv $resultcsvname -Delimiter "`t"

   write-host "PRTG-SFBAssessment: Calculating results"
   
   $Packetssent         = ($result | %{$_.PacketsSent.replace(",",".")} | Measure-Object -sum)
   $PacketLossRate      = ($result | %{$_.packetlossrate.replace(",",".")} | Measure-Object -Average -Maximum)
   $RoundTripLatencyInMs= ($result | %{$_.RoundTripLatencyInMs.replace(",",".")} | Measure-Object -Average -Maximum)
   $AverageJitterInMs   = ($result | %{$_.AverageJitterInMs.replace(",",".")} | Measure-Object -Average -Maximum)
   $PacketReorderRatio  = ($result | %{$_.PacketReorderRatio.replace(",",".")} | Measure-Object -Average -Maximum)
   
   write-host " Runs                     : $($Packetssent.count)"
   write-host " Packetssent              : $($Packetssent.sum)"
   write-host " PacketLossRate      (avg): $($PacketLossRate.Average)"
   write-host " RoundTripLatencyInMs(avg): $($RoundTripLatencyInMs.Average)"
   write-host " AverageJitterInMs   (avg): $($AverageJitterInMs.Average)"
   write-host " PacketReorderRatio  (avg): $($PacketReorderRatio.Average)"

   write-host "PRTG-SFBAssessment: Build PRTG XML"   
   $prtgresult = '<?xml version="1.0" encoding="UTF-8" ?>
      <prtg>
         <result>
            <channel>Runs</channel>
            <value>' + $totalPacketsSent.count +'</value>
            <float>0</float>
            <unit>Count</unit>
         </result>         <result>
            <channel>Packet loss rate avg</channel>
            <value>' + $totalPacketsSent.sum +'</value>
            <float>0</float>
            <unit>Count</unit>
         </result>
         <result>
            <channel>Packet loss rate avg</channel>
            <value>' + $PacketLossRate.Average +'</value>
            <float>1</float>
            <unit>count</unit>
         </result>
         <result>
            <channel>Packet loss rate max</channel>
            <value>' + $PacketLossRate.Maximum +'</value>
            <float>1</float>
            <unit>count</unit>
         </result>
         <result>
            <channel>RTT latency avg</channel>
            <value>' + $RoundTripLatencyInMs.Average +'</value>
            <float>1</float>
            <unit>Custom</unit>
            <customunit>Milliseconds</customunit>
         </result>
         <result>
            <channel>RTT latency Max</channel>
            <value>' + $RoundTripLatencyInMs.Maximum +'</value>
            <float>1</float>
            <unit>Custom</unit>
            <customunit>Milliseconds</customunit>
         </result>
         <result>
            <channel>Jitter</channel>
            <value>' + $AverageJitterInMs.Average +'</value>
            <float>1</float>
            <unit>Custom</unit>
            <customunit>Milliseconds</customunit>
         </result>
         <result>
            <channel>Jitter Max</channel>
            <value>' + $AverageJitterInMs.Maximum +'</value>
            <float>1</float>
            <unit>Custom</unit>
            <customunit>Milliseconds</customunit>
         </result>
         <result>
            <channel>Packet reorder ratio</channel>
            <value>' + $PacketReorderRatio.Average +'</value>
            <float>1</float>
            <unit>count</unit>
         </result>
         <result>
            <channel>Packet reorder ratio max</channel>
            <value>' + $PacketReorderRatio.Maximum +'</value>
            <float>1</float>
            <unit>count</unit>
         </result>
      </prtg>'

   #$prtgresult | out-file result.xml

   write-host "PRTG-SFBAssessment: Post Result to PRTG-Server"
   try {
   $Answer=Invoke-RestMethod `
      -method "GET" `
      -URI ($prtgserveruri + "?content=$prtgresult")
      if ($answer."Matching Sensors" -eq "1") {
         write-host "Found 1 Sensors  OK"
      }
      else {
         write-Warning "Invalid reply"
         $answer
#         exit 1
      }
   }
   catch {
      write-Warning "Unable to invoke-Restmethod  $($_.Exception.Message)"
   }
}

write-host "PRTG-SFBAssessment:End"
