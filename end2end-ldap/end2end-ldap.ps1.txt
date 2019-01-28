#'-------------------------------------------------------------------------
#' end2end-ldap 1.0
#'
#' Beschreibung
#' Liest immer wieder HIghestUSN und misst Zeit
#'
#' Vorgehensweise
#' - Initialisiert einen String mit vorgegebener Größe und schreibt diesen mit einstellbarer Pause auf die Festplatte
#'
#' Voraussetzung
#' - Lesen im AD
#'
#' Achtung: Das Skript kann nur mit CTRL-C abgebrochen werden
#' Timer() gibt die milliseconds zwischen 00:00 und der aufrufzeit in Sekunden zurück (single)
#'
#' Version 1.0 (11. Mrz 2014 )
#'	Erste Version 
#
# Polling for Changes Using USNChanged (Windows)
# http://msdn.microsoft.com/en-us/library/windows/desktop/ms677627(v=vs.85).aspx
#'-------------------------------------------------------------------------#

param(
	[string]$dc="de.msxfaq.de" ,		# FQDN of a DC to ask
	[string]$reportfilename = 'end2end-ldap.csv',
	[int]$IdleTime = 1000,
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

#Write-EventLog -Message "end2end-ldap started with $Testfilename" -EntryType Information -EventId 0 -Source end2end-ldap -LogName Application

write-host "end2end-ldap: gestartet"
write-host "end2end-ldap: DC        :" $DC
write-host "end2end-ldap: reportfile:" $reportfilename
write-host "end2end-ldap: idletime  :" $IdleTime
write-host "end2end-ldap: alarmdelta:" $Alarmdelta

write-host  "end2end-ldap: Initialize ADSI"

$adsipath="LDAP://"+$dc+"/RootDSE"
write-host "end2end-ldap: LDAPPath" $adsipath 

#write-host  ":Initialize Logfile", 0
#dim fs, file, reportfile
#Set fs = CreateObject("Scripting.FileSystemObject")
#Set reportfile = fs.OpenTextFile(strReportfilename, ForWriting, ForOverwriting)
if (test-path -path $reportfilename -pathtype leaf) {
	write-host "end2end-ldap: Append to CSV File $reportfilename"
}
else {
	write-host "end2end-ldap: Create new CSV File $reportfilename"
	"timestamp;performance" | Out-File -FilePath $reportfilename
}

[int]$mittelwert = 0 
[int]$count  = 0 
[int]$max = 0 
[int]$message = "" 
[int]$alive=0
write-host  "end2end-ldap: Start Reading"

while ($true) {
	if ($alive > 6000) { # dump regular "alive" messages to eventlog nearly every 10+WriteTime Minutes 
		write-host  "end2end-ldap: alive",
		#Write-EventLog -Message "end2end-ldap $Testfilename Alive" -EntryType Error -EventId 1 -Source end2end-ldap -LogName Application
		$alive = 0
	else
		$alive = $alive + 1
	}

	$performance = (measure-command {
		$rootdse = [ADSI]$adsipath
		$rootdse.highestCommittedUSN | out-null
		$rootdse.psbase.dispose()
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
			$message = $message + "`t Alarmdelta: `t" + $alarmdelta/1000 + " s `r`n"
			$message = $message + "`t Idletime: `t" + $IdleTime + " s `r`n"
			$message = $message + "`t Testfilename: `t" + $Testfilename + " s" 
			write-host  $message
			#Write-EventLog -Message $message -EntryType Warning -EventId 2 -Source end2end-ldap -LogName Application
		}
		else {
			$mittelwert = $mittelwert + ($performance - $mittelwert)/10  # verschiebe mittelwert median by 10% 
		}
	}

	if ($performance -gt $max){
		$max = $performance
	}

	if ($count -gt 9) {
		" Mittel: $mittelwert ms  Max: $Max ms "
		$count = 0 
		$max = 0 
	}
	else{
		$count = $count  + 1
		write-host "." -NoNewline
	}
	(Get-Date).tostring() + "; $performance" | Out-File -FilePath $reportfilename -Append -NoClobber
	Start-Sleep -Milliseconds $IdleTime
}



