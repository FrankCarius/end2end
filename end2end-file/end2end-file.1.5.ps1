#'-------------------------------------------------------------------------
#' end2dend-file 1.4
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
#'      Erste Version abgeleitet von DiskTest 1.1
#' Version 1.1 (14. Dez 2007 )
#'      Mehr Detailausgabe
#' Version 1.2 (24. Jul 2010)
#'      parametrisierung
#' Version 1.3 (01. Jul 2012)
#       bugfixing. 
#' Version 1.4 (15. Jan 2013)
#       undefined variables, missing brackes result in wrong counting of max and median, SMTPMail
#' Version 1.4 (05. Jul 2015)
#       Umstellen auf ByteBuffer und schreiben mit [system.io.file]
#'-------------------------------------------------------------------------#
 
param(
	[string]$Testfilename = 'end2end-file.tmp',		# testfile to write, can be any UNC or local path
	[string]$Reportfilename = 'end2end-file.csv',	# CSVFile to write results
	[int]$IdleTime = 100,							# time in MS to sleep between two writes
	[int]$Buffersize = 1024,						# size in bytes of buffer to write
	[int]$Alarmdelta = 1000,						# limit in ms to reach for triggering an alarm
	[string]$smtpserver="",							# Specify SMTP-Servername, sender and recipient to get a message
	[string]$smtpfrom="",
	[string]$smtpto=""
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
[string]$buffer=""
[string]$buffer = ([System.String]$buffer).PadLeft($Buffersize,"a")
$buffer2 = [system.text.encoding]::ASCII.GetBytes($buffer);
 
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
		$alive+= 1
	}

	$performance = (measure-command {
		$buffer | Out-File -FilePath $Testfilename
		[system.io.file]::WriteAllBytes($Testfilename,$buffer2)  # ca 1ms lokal
	}).Milliseconds

	if ($mittelwert -eq 0 ) {
		$mittelwert = $performance  # First run
	}
#        elseif ($performance -lt 0) {
#                $performance = $mittelwert  # simply skip the midnight rollover
#        }
	else {
		if (($performance - $mittelwert) -gt  $Alarmdelta) {
			$message = "Diskprobe ALARM: Fileaccess exceeded limit `r`n"
			$message = $message + "`t Mittelwert:  `t" + ($mittelwert/1000) + "s `r`n"
			$message = $message + "`t Aktueller Wert:  `t" + $performance/1000 + "s `r`n"
			$message = $message + "`t Buffersize:  `t" + $Buffersize + " Bytes `r`n"
			$message = $message + "`t Alarmdelta: `t" + $alarmdelta/1000 + " s `r`n"
			$message = $message + "`t Idletime: `t" + $IdleTime + " s `r`n"
			$message = $message + "`t Testfilename: `t" + $Testfilename + " s" 
			write-host  $message
			#Write-EventLog -Message $message -EntryType Warning -EventId 2 -Source end2end-File -LogName Application
			if ($smtpserver -ne "") {
				write-host  "Sending SMTP-Alertmessage"
				Send-MailMessage `
					-SmtpServer $smtpserver `
					-From $smtpfrom `
					-To $smtptp `
					-Subject "End2EndFile Alarm" `
					-Body $message
			}
		}
		else {
			$mittelwert = $mittelwert + ($performance - $mittelwert)/10  # verschiebe mittelwert median by 10% 
		}
	}

	if ($performance -gt $max){
		$max = $performance
	}

	if ($count -gt 9) {
		" Mittel(last 10 samples): $mittelwert ms    Maximum Duration: $Max ms  Speed:{0:N2} MB/Sek" -f ($buffersize/($mittelwert+0.1))
		$count = 0 
		$max = 0 
		$mittelwert=0
	}
	else{
			$count = $count  + 1
			write-host "." -NoNewline
	}
	(Get-Date).tostring() + "; $performance" | Out-File -FilePath $Reportfilename -Append -NoClobber
	Start-Sleep -Milliseconds $IdleTime
}
