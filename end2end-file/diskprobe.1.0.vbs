Option Explicit
'-------------------------------------------------------------------------
' Diskprobe 1.0
'
' Beschreibung
' Schreibt kontinuierlich in eine Datei und misst die Dauaer dafür
'
' Vorgehensweise
' - Initialisiert einen String mit 1k Größe und schreibt diesen im Abstand von 0,1Sek auf die Festplatte
' Erzeugt also eine Dauerlast von ca 10kbyte/Sek, was einen Server nicht wirklich stören sollte
'
' Voraussetzung
' - Schreibrechte im Zielverzeichnis
'
' Achtung: Das Skript kann nur mit CTRL-C abgebrochen werden
' Timer() gibt die milliseconds zwischen 00:00 und der aufrufzeit in Sekunden zurück (single)
'
' Version 1.0 (09. Nov 2007 )
'	Erste Version abgeleitet von DiskTest 1.1
'-------------------------------------------------------------------------

' Konfigurationskonstanzen
const conTestfilename = "diskprobe.tmp"	'Testdatei, die beschrieben wird
'~ const conCSVfilename = "diskprobe.csv"	'Ausgabe der aktuellen  Messwerte als LOG
const conIdleTime = 100	' Ruhezeit zwischen zwei Schreibvorgängen
const Alarmdelta = 1000	'Maximale abweichung des Messwerts vom mittelwert in ms
const Buffersize = 10240	'Größe des zu schreibenden Buffers

' interne Konstanten
Const ForWriting = 2
Const ForOverwriting= true
Const ForAppending = false

'~ testfilename = wscript.arguments.named("file")
'if testfilename  = "" then testfilename = "disktest.tmp"

'~ dim maxsize
'~ maxsize = wscript.arguments.named("maxsize")
'~ if maxsize = "" then maxsize = 0 

wscript.echo "To stop that script, press CTRL-C"
wscript.echo "YOU HAVE TO REMOVE THE TEMPORARY TEST FILE MANUALLY !"
'wscript.stdin.readline

dim StrWriteBuffer
wscript.echo  now & ":Initialize Buffer START"
StrWriteBuffer = space (10240)

wscript.echo  now & ":Initialize Eventlog Writer"
Dim objShell 
Set objShell = CreateObject ("WScript.Shell") 
objShell.LogEvent 0, "Diskprobe started"

'~ wscript.echo  now & ":Initialize Logfile"
dim fs, file, logfile
Set fs = CreateObject("Scripting.FileSystemObject")
'~ Set logfile = fs.OpenTextFile(conCSVfilename, ForWriting, ForOverwriting)
'~ logfile.writeline "timestam;performance"

dim performance, mittelwert, count, max, dtstart, dtstop, message 
mittelwert = 0 : count  = 0 : max = 0 : message = ""
wscript.echo  now() & ":Start Writing"
do
	dtstart = clng(timer()*1000)
	Set file = fs.OpenTextFile(conTestfilename, ForWriting, ForOverwriting)
	file.Write(StrWriteBuffer)
	file.close
	dtstop = clng(timer()*1000)
	performance = (dtstop - dtstart)  ' in milliceonds
	
	if mittelwert = 0 then 
		mittelwert = performance  ' First run
	elseif performance < 0  then
		performance = mittelwert  ' simply skip the midnight rollover
	else
		if performance - mittelwert > Alarmdelta then
			message = "Diskprobe ALARM: Fileaccess exceeded limit" & vbcrlf &_
					vbtab & "Mittelwert: " & vbtab & formatnumber(mittelwert /1000,3) & "s" & vbcrlf &_
					vbtab & "Aktueller Wert: " & vbtab & performance /1000 &  "s" & vbcrlf &_
					vbtab & "Buffersize: " & vbtab & Buffersize & " Bytes"& vbcrlf &_
					vbtab & "Alarmdelta: " & vbtab & alarmdelta/1000 &  " s"& vbcrlf &_
					vbtab & "Idletime: " & vbtab & conIdleTime &  " s"  & vbcrlf &_
					vbtab & "Testfilename: " & vbtab & conTestfilename &  " s" 
			wscript.echo now & message 
			objShell.LogEvent 1, message 
		else
			mittelwert = mittelwert + (performance - mittelwert)/10  ' verschiebe mittelwert median by 10% 
		end if
	end if

	if performance > max then max = performance

	if count > 9 then 
		wscript.echo now & " Mittel:" & cint(mittelwert) & "ms  Max:" & Max & "ms " '& " Speed:" & clng(10/(mittelwert +0.1)) & "MB/Sek"
		count = 0 : max = 0 
	else
		count = count  + 1
		wscript.stdout.write "."
	end if 
		
		
	'~ wscript.echo now & " " & formatnumber(count/10,1,-1) & "Mb/Sek " & replace(space(count/50)," ","#")
	'~ logfile.writeline now & ";" & count/10
	wscript.sleep(conIdleTime)
loop