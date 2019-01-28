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
# 20180508 frank@carius.de
#	Kontrolle ob SFB Tool EXE vorhanden ist, Call Failed handling
# 20181123 EXE-Pfad auf 64bit Client angepasst.

param (
	[string]$sfbpath = "c:\program Files(x86)\Microsoft Skype for Business Network Assessment Tool\",
	[string]$resultcsvname = "$($env:LOCALAPPDATA)\Microsoft Skype for Business Network Assessment Tool\performance_results.tsv",
	[string]$prtgpushurl = "http://prtg:5050/end2end-sfbassessment_$($env:COMPUTERNAME)",
	[string]$archivcsv = ".\end2end-SFBAssessment_$($env:COMPUTERNAME)-archiv.csv",
	[string]$summarycsv = ".\end2end-SFBAssessment_$($env:COMPUTERNAME)-summary.csv"
)

# -----------------------------------------------------
# sendto-prtg   helper to send data to PRTG
# -----------------------------------------------------
function sendto-prtg (
	[string]$prtgresult,   # XML Structure
	[string]$prtgpushurl)  # HTTP-PUSH-url
{
	if ($prtgpushurl -eq "" ) {
		$Scriptname = (split-path $MyInvocation.ScriptName -Leaf).replace(".ps1","")
		$prtgpushurl= "http://prtg:5050/$($scriptname)_$($env:COMPUTERNAME)"
	}
	write-host "SendTo-PRTG: Post Result to PRTGProbe $prtgpushurl"
	
	$Answer=Invoke-RestMethod `
		-method "GET" `
		-URI ("$($prtgpushurl)?content=$($prtgresult)")
	if ($answer."Matching Sensors" -eq "1") {
		write-host "SendTo-PRTG:Found 1 Sensors  OK"
	}
	elseif ($answer."Matching Sensors" -eq "0") {
		write-warning "SendTo-PRTG:Found 0 matching sensors. Retry next run"
	}
	else {
		write-Warning "SendTo-PRTG:Invalid reply"
		$answer
	}
}

write-host "End2End-SFBAssessment:Start"
write-host "End2End-SFBAssessment:Parameter resultcsvname: $($resultcsvname)"
write-host "End2End-SFBAssessment:Parameter prtgpushurl: $($prtgpushurl)"

while ($true) {
	$sfbtoolexe = ($sfbpath + "NetworkAssessmentTool.exe")
	$sfbtoolconfig = ($sfbpath + "NetworkAssessmentTool.exe.config")
	if (!(test-path -Path $sfbtoolexe -PathType leaf)) {
		write-host "End2End-SFBAssessment:ERROR: EXE at $($sfbtoolexe) not found" -foregroundcolor red
		$prtgresult = '
			<prtg>
			 <error>1</error>
			 <text>ERROR: EXE at '+$($sfbtoolexe)+' not found</text>
		  </prtg>'	
		sendto-prtg $prtgresult $prtgpushurl
		write-host "End2End-SFBAssessment: Sleeping 60 Seconds to retry"
		start-sleep -seconds 60
	}
	elseif (!(test-path -Path $sfbtoolconfig -PathType leaf)) {
			write-host "End2End-SFBAssessment:ERROR: Config at $($sfbtoolconfig) not found" -foregroundcolor red
			$prtgresult = '
				<prtg>
				 <error>1</error>
				 <text>ERROR: $sfbtoolconfig at '+$($sfbtoolexe)+' not found</text>
			  </prtg>'	
			sendto-prtg $prtgresult	$prtgpushurl
			write-host "End2End-SFBAssessment: Sleeping 60 Seconds to retry"
			start-sleep -seconds 60
	}
	else {
		write-host "End2End-SFBAssessment: Execute Assessment tool"
		Start-Process -FilePath ($sfbpath + "NetworkAssessmentTool.exe") -NoNewWindow -Wait -workingDirectory $sfbpath
		if (!(test-path -path $resultcsvname -pathtype leaf)) {
			write-host "  Unable to find CSV at $($resultcsvname)" -foregroundcolor red
			$prtgresult = '
				<prtg>
				 <error>1</error>
				 <text>ERROR: Unable to find Resultfile at '+$($resultcsvname)+' not found</text>
			  </prtg>'	
			sendto-prtg $prtgresult	$prtgpushurl
			write-host "End2End-SFBAssessment: Sleeping 60 Seconds to retry"
			start-sleep -seconds 60
		}
		else {
			write-host "End2End-SFBAssessment:Parsing Result TSV"
			[array]$result = Import-Csv $resultcsvname -Delimiter "`t"

			if (!$result) {
				write-host "  CSV Empty. Assume Call failed!" -foregroundcolor red	
				$prtgresult = '
					<prtg>
					 <error>1</error>
					 <text>ERROR: Resultfile empty - Call failed ?</text>
				  </prtg>'	
				sendto-prtg $prtgresult	$prtgpushurl
				write-host "End2End-SFBAssessment: Sleeping 60 Seconds to retry"
				start-sleep -seconds 60				
			}
			else {
				if ($archivcsv -ne "") {
					Write-host "End2End-SFBAssessment: Adding to Archivcsv $($archivcsv)"
					$result | export-csv -path $archivcsv -append -notypeinformation
				}

				write-host "End2End-SFBAssessment: Calculating results"
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
					Write-host "End2End-SFBAssessment: Adding to Summary $($Summarycsv)"
					$statistics | export-csv -path $summarycsv -append -notypeinformation
				}
			 
				if ($prtgpushurl -ne "") {
					write-host "End2End-SFBAssessment: Build PRTG XML"   
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
						 <text>OK</text>
					  </prtg>'
					sendto-prtg $prtgresult $prtgpushurl
				}
			}
		}
	}
}
write-host "End2End-SFBAssessment:End"
