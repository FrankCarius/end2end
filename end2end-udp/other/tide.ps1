# Tide simulator für RRDTool
#
# Generieren von synthetischen Daten für RRDTool
#
# Erde dreht sich 1x am tag um die eigene Achse = 2x Ebbe und 2xFlut
# Ein Tag hat 1440 = 288 x 5 Minuten = 720 Grad oder 4x 360
#-> 2x Ebbe und 2xFlut
#
# minute = Minutes from 0
# singrad = 720%minutes
# dayTide = Sin(singrad)
#
# Einfluss des Mondes Umlaufzeit 27,32 Tage  nicht relevant
#Mondumlauf  http://de.wikipedia.org/wiki/Mondumlauf
# "glitch" für Mond fehlt noch. 1/27tel pro Tag abweichung  

# Spring/Nippflut 2x pro Monat, 
 
function grad2rad ($grad) {
	[float]$grad * [Math]::PI / 180
}

$time19700101 = [datetime]"01.01.1970"  # Unix timecounter in seconds from 01.01.1970
$time20110101 = [datetime]"1.1.2011"
$basetime = [int](($time20110101-$time19700101).totalseconds)


write-host "Create RRD"
& .\rrdtool.exe create "C:\group\Technik\Skripte\end2end\end2endudp\tide.db" `
   -start $basetime `
    DS:tide:GAUGE:5:0:2 RRA:AVERAGE:0:360:576 RRA:AVERAGE:0:30:576 RRA:AVERAGE:0:7:576 RRA:AVERAGE:0:1:576

# schleife 720 =  tag


for ($day=1; $day -le 365;$day+=1){
	write-host "Day $day"
	for ($minute=0; $minute -le (1440*$day); $minute+=5) {

		#Tide des Tages ohne Mond
		$grad = ($minute/2)
		$rad =(grad2rad $grad)
		$sin = 1+([math]::sin($rad)) + (0.2*([math]::sin($rad/14)))
	#	Write-Host "Minute:$minute  Grad: "$grad"  Rad:" $rad "   SIN:" $sin
		#write-host "." -nonewline
		$timestamp = $basetime + ($minute *60)
		
		#"$timestamp;$sin"
		.\rrdtool.exe updatev "C:\group\Technik\Skripte\end2end\end2endudp\tide.db" "$timestamp@$sin"
	}
}

.\rrdtool.exe graphv "C:\group\Technik\Skripte\end2end\end2endudp\tide.png"