#'-------------------------------------------------------------------------
#' end2dend-file 1.2
#'
#' Beschreibung
#' Schreibt kontinuierlich in eine Datei und misst die Dauer dafür
#'
#' Vorgehensweise
#' - Initialisiert einen String mit vorgegebener Größe und schreibt diesen mit einstellbarer Pause auf die Festplatte
#'
#' Voraussetzung
#' - Schreibrechte im Zielverzeichnis
#'
#' Achtung: Das Skript kann nur mit CTRL-C abgebrochen werden
#' Timer() gibt die milliseconds zwischen 00:00 und der aufrufzeit in Sekunden zurück (single)
#'
#' Version 1.0 (09. Nov 2007 )
#'	Erste Version abgeleitet von DiskTest 1.1
#' Version 1.1 (14. Dez 2007 )
#'	Mehr Detailausgabe
#' Version 1.2 (24. Jul 2010)
#'	parametrisierung
#'-------------------------------------------------------------------------#

param(
	[string]$Testfilename = 'end2end-file.tmp',
	[string]$Reportfilename = 'end2end-file.csv',
	[int]$IdleTime = 100,
	[int]$Buffersize = 1024,
	[int]$Alarmdelta = 1000
)

Set-PSDebug -strict
$ErrorActionPreference = "Continue";
$verbosepreference = "SilentlyContinue";
#$verbosepreference = "Continue";
$DebugPreference = "SilentlyContinue";
#$DebugPreference = "Continue";
$WarningPreference="Continue";
$Error.Clear();

#Write-EventLog -Message "End2End-File started with $Testfilename" -EntryType Information -EventId 0 -Source end2end-File -LogName Application

write-host "end2end-file: gestartet"
write-host "end2end-file: testfile  :" $Testfilename
write-host "end2end-file: reportfile:" $Reportfilename
write-host "end2end-file: idletime  :" $IdleTime
write-host "end2end-file: buffersize:" $Buffersize
write-host "end2end-file: alarmdelta:" $Alarmdelta

write-host  "end2end-file: Initialize Buffer START"
$buffer = [System.String]$buffer.PadLeft($Buffersize,"a")

#write-host  ":Initialize Logfile", 0
#dim fs, file, reportfile
#Set fs = CreateObject("Scripting.FileSystemObject")
#Set reportfile = fs.OpenTextFile(strReportfilename, ForWriting, ForOverwriting)
"timestamp;performance" | Out-File -FilePath $Reportfilename

[int]$mittelwert = 0 
[int]$count  = 0 
[int]$max = 0 
[int]$message = "" 
[int]$alive=0
write-host  "end2end-file: Start Writing"

while ($true) {
	if ($alive > 6000) { # dump regular "alive" messages to eventlog nearly every 10+WriteTime Minutes 
		write-host  "end2end-file: alive",
		#Write-EventLog -Message "End2End-File $Testfilename Alive" -EntryType Error -EventId 1 -Source End2End-File -LogName Application
		$alive = 0
	else
		$alive = $alive + 1
	}

	$performance = (measure-command {
		$buffer.ToString() | Out-File -FilePath $Testfilename
	}).Milliseconds
	
	if ($mittelwert -eq 0 ) {
		$mittelwert = $performance  # First run
	}
	elseif ($performance -lt 0) {
		$performance = $mittelwert  # simply skip the midnight rollover
	}
	else {
		if ($performance - $mittelwert -gt  $Alarmdelta) {
			$message = "Diskprobe ALARM: Fileaccess exceeded limit `r`n"
			$message = $message + "`t Mittelwert:  `t" + ($mittelwert/1000) + "s `r`n"
			$message = $message + "`t Aktueller Wert:  `t" + $performance/1000 + "s `r`n"
			$message = $message + "`t Buffersize:  `t" + $Buffersize + " Bytes `r`n"
			$message = $message + "`t Alarmdelta: `t" + $alarmdelta/1000 + " s `r`n"
			$message = $message + "`t Idletime: `t" + $IdleTime + " s `r`n"
			$message = $message + "`t Testfilename: `t" + $Testfilename + " s" 
			write-host  $message
			#Write-EventLog -Message $message -EntryType Warning -EventId 2 -Source end2end-File -LogName Application
		}
		else {
			$mittelwert = $mittelwert + ($performance - $mittelwert)/10  # verschiebe mittelwert median by 10% 
		}
	}

	if ($performance -gt $max){
		$max = $performance
	}

	if ($count -gt 9) {
		" Mittel: $mittelwert ms  Max: $Max ms  Speed:{0:N2} MB/Sek" -f (10/($mittelwert +0.1))
		$count = 0 
		$max = 0 
	}
	else{
		$count = $count  + 1
		write-host "." -NoNewline
	}
	(Get-Date).tostring() + "; $performance" | Out-File -FilePath $Reportfilename -Append -NoClobber
	Start-Sleep -Milliseconds $IdleTime
}
