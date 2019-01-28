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
# 20180507 frank@carius.de
#	Modify to work with "numiterations=1
#	PRTG PushURL mit Computername versehen
#	PRTGPush ErrorHandlng
#	Milliseconds rounded to integer
#	CSv-Files with Computername

param (
	[string]$sfbpath = "c:\program Files\Microsoft Skype for Business Network Assessment Tool\",
	[string]$resultcsvname = "$($env:LOCALAPPDATA)\Microsoft Skype for Business Network Assessment Tool\performance_results.tsv",
	[string]$prtgpushuri = "http://192.168.102.69:5050/sfbassessment_$($env:COMPUTERNAME)",
	[string]$archivcsv = ".\end2end-SFBAssessment_$($env:COMPUTERNAME)-archiv.csv",
	[string]$summarycsv = ".\end2end-SFBAssessment_$($env:COMPUTERNAME)-summary.csv"
)

write-host "end2end-SFBAssessment:Start"

write-host " Parameter resultcsvname: $($resultcsvname)"
write-host " Parameter prtgpushuri: $($prtgpushuri)"

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
		Packetssentcount        = ($result.count)
		Packetssentsum          = [long]($result | %{$_.PacketsSent.replace(",",".")} | Measure-Object -sum).Sum
		PacketLossRateAvg       = [long]($result | %{$_.packetlossrate.replace(",",".")} | Measure-Object -Average).Average
		PacketLossRateMax       = [long]($result | %{$_.packetlossrate.replace(",",".")} | Measure-Object -Maximum).Maximum
		RoundTripLatencyAvgInMs = [long]($result | %{$_.RoundTripLatencyInMs.replace(",",".")} | Measure-Object -Average).Average
		RoundTripLatencyMaxInMs = [long]($result | %{$_.RoundTripLatencyInMs.replace(",",".")} | Measure-Object -Maximum).Maximum
		AverageJitterInMsAvg    = [long]($result | %{$_.AverageJitterInMs.replace(",",".")} | Measure-Object -Average).Average
		AverageJitterInMsMax    = [long]($result | %{$_.AverageJitterInMs.replace(",",".")} | Measure-Object  -Maximum).Maximum
		PacketReorderRatioAvg   = [long]($result | %{$_.PacketReorderRatio.replace(",",".")} | Measure-Object -Average).Average
		PacketReorderRatioMax   = [long]($result | %{$_.PacketReorderRatio.replace(",",".")} | Measure-Object -Maximum).Maximum
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
 
	if ($prtgpushuri -ne "") {
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

	   write-host "end2end-SfBAssessment: Post Result to PRTGProbe $prtgpushuri"
	   $Answer=Invoke-RestMethod `
		  -method "GET" `
		  -URI ($prtgpushuri + "?content=$prtgresult")
		if ($answer."Matching Sensors" -eq "1") {
		  write-host " PRTG-Reply:Found 1 Sensors  OK"
		}
		elseif ($answer."Matching Sensors" -eq "0") {
		  write-warning " PRTG-Reply:Found 0 matching sensors. Retry next run"
		}
		else {
			write-Warning " PRTG-Reply:Invalid reply"
			$answer
		}
	}
}

write-host "end2end-SfBAssessment:End"
