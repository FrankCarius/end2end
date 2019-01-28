# end2endo365
#
#Simple Skript to measure Office 365 Response time by loading the favicon

while ($true) {
	write-host "." -nonewline
	$starttime = get-date	
	$result = Invoke-WebRequest "https://outlook.office365.com/favicon.ico"
	$durationms =  ((get-date) -$starttime).TotalMilliseconds
	$bytes = $result.RawContentLength
	"" | Select-Object `
		@{Name="Date";Expression={$starttime.ToUniversalTime().tostring()}},`
		@{Name="durationms";Expression={[math]::round($durationms)}}, `
		@{Name="KBit/sec";Expression={[math]::round($bytes/$durationms*10)}}
	start-sleep -milliseconds ([math]::Max(1000-$durationms,0))
}