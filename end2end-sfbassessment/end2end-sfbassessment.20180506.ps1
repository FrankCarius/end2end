# End2End-SFBAssessment.ps1
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
# 20180506 frank@carius.de
#	Addieren ArchivCSV, Umbau auf PRTG Dauerläufer

param (
	[string]$sfbpath = "c:\program Files\Microsoft Skype for Business Network Assessment Tool\",
	[string]$resultcsvname = "$($env:LOCALAPPDATA)\Microsoft Skype for Business Network Assessment Tool\performance_results.tsv",
	[string]$prtgserveruri = "http://prtg:5050/sfbassessment1",
	[string]$archivcsv = ".\end2end-SFBAssessment-archiv.csv",
	[string]$summarycsv = ".\end2end-SFBAssessment-summary.csv"
)

write-host "end2end-SFBAssessment:Start"

write-host " Parameter resultcsvname: $($resultcsvname)"
write-host " Parameter prtgserveruri: $($prtgserveruri)"

while ($true) {
	write-host "end2end-SfBAssessment: Execute Assessment tool"
	Start-Process -FilePath ($sfbpath + "NetworkAssessmentTool.exe") -NoNewWindow -Wait -workingDirectory $sfbpath

	write-host "end2end-SfBAssessment: Parsing Result TSV"
	[array]$result = Import-Csv $resultcsvname -Delimiter "`t"

	if ($archivcsv -ne "") {
		Write-host "end2end-SfBAssessment: Adding to Archivcsv $($archivcsv)"
		$result | export-csv -path $archivcsv -append -notypeinformation
	}

	write-host "end2end-SfBAssessment: Calculating results"
	$statistics = New-Object PSObject -Property @{
		Packetssentcount          = ($result.count)
		Packetssentsum          = ($result | %{$_.PacketsSent.replace(",",".")} | Measure-Object -sum).Sum
		PacketLossRateAvg       = ($result | %{$_.packetlossrate.replace(",",".")} | Measure-Object -Average).Average
		PacketLossRateMax       = ($result | %{$_.packetlossrate.replace(",",".")} | Measure-Object -Maximum).Maximum
		RoundTripLatencyAvgInMs = ($result | %{$_.RoundTripLatencyInMs.replace(",",".")} | Measure-Object -Average).Average
		RoundTripLatencyMaxInMs = ($result | %{$_.RoundTripLatencyInMs.replace(",",".")} | Measure-Object -Maximum).Maximum
		AverageJitterInMsAvg    = ($result | %{$_.AverageJitterInMs.replace(",",".")} | Measure-Object -Average).Average
		AverageJitterInMsMax    = ($result | %{$_.AverageJitterInMs.replace(",",".")} | Measure-Object  -Maximum).Maximum
		PacketReorderRatioAvg   = ($result | %{$_.PacketReorderRatio.replace(",",".")} | Measure-Object -Average).Average
		PacketReorderRatioMax   = ($result | %{$_.PacketReorderRatio.replace(",",".")} | Measure-Object -Maximum).Maximum
	}
   
	write-host " PacketssentCount         : $($statistics.Packetssentcount)"
	write-host " PacketssentSum           : $($statistics.Packetssentsum)"
	write-host " PacketLossRate      (avg): $($statistics.PacketLossRateAvg)"
	write-host " RoundTripLatencyInMs(avg): $($statistics.RoundTripLatencyAvgInMs)"
	write-host " AverageJitterInMs   (avg): $($statistics.AverageJitterInMsAvg)"
	write-host " PacketReorderRatio  (avg): $($statistics.PacketReorderRatioAvg)"

	if ($summarycsv -ne "") {
		Write-host "end2end-SfBAssessment: Adding to Summary $($Summarycsv)"
		$statistics | export-csv -path $summarycsv -append -notypeinformation
	}
 
	if ($prtgserveruri -ne "") {
		write-host "end2end-SfBAssessment: Build PRTG XML"   
		$prtgresult = '<?xml version="1.0" encoding="UTF-8" ?>
			<prtg>
			 <result>
				<channel>Runs</channel>
				<value>' + $statistics.totalPacketsSentcount +'</value>
				<float>0</float>
				<unit>Count</unit>
			 </result>         <result>
				<channel>Packet loss rate avg</channel>
				<value>' + $statistics.totalPacketsSentsum +'</value>
				<float>0</float>
				<unit>Count</unit>
			 </result>
			 <result>
				<channel>Packet loss rate avg</channel>
				<value>' + $statistics.PacketLossRateAvg +'</value>
				<float>1</float>
				<unit>count</unit>
			 </result>
			 <result>
				<channel>Packet loss rate max</channel>
				<value>' + $statistics.PacketLossRateMax +'</value>
				<float>1</float>
				<unit>count</unit>
			 </result>
			 <result>
				<channel>RTT latency avg</channel>
				<value>' + $statistics.RoundTripLatencyAvgInMs +'</value>
				<float>1</float>
				<unit>Custom</unit>
				<customunit>Milliseconds</customunit>
			 </result>
			 <result>
				<channel>RTT latency Max</channel>
				<value>' + $statistics.RoundTripLatencyMaxInMs +'</value>
				<float>1</float>
				<unit>Custom</unit>
				<customunit>Milliseconds</customunit>
			 </result>
			 <result>
				<channel>Jitter</channel>
				<value>' + $statistics.AverageJitterInMsAvg +'</value>
				<float>1</float>
				<unit>Custom</unit>
				<customunit>Milliseconds</customunit>
			 </result>
			 <result>
				<channel>Jitter Max</channel>
				<value>' + $statistics.AverageJitterInMsMax +'</value>
				<float>1</float>
				<unit>Custom</unit>
				<customunit>Milliseconds</customunit>
			 </result>
			 <result>
				<channel>Packet reorder ratio</channel>
				<value>' + $statistics.PacketReorderRatioAvg +'</value>
				<float>1</float>
				<unit>count</unit>
			 </result>
			 <result>
				<channel>Packet reorder ratio max</channel>
				<value>' + $statistics.PacketReorderRatioMax +'</value>
				<float>1</float>
				<unit>count</unit>
			 </result>
		  </prtg>'

	   #$prtgresult | out-file result.xml

	   write-host "end2end-SfBAssessment: Post Result to PRTG-Server"
	   $Answer=Invoke-RestMethod `
		  -method "GET" `
		  -URI ($prtgserveruri + "?content=$prtgresult")
	   if ($answer.Statuscode -ne 200) {
		  write-warning "Request to PRTG failed - Statuscode not 200"
		  exit 1
	   }
	   elseif ($answer."Matching Sensors" -eq "1") {
		  write-host "Found 1 Sensors  OK"
		  exit 0
	   }
	   else {
		  write-Warning "Invalid reply"
		  $answer
		  exit 1
		}
	}
}

write-host "end2end-SfBAssessment:End"
