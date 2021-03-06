Option Explicit
'-------------------------------------------------------------------------
' Diskprobe 1.0
'
' Beschreibung
' Schreibt kontinuierlich in eine Datei und misst die Dauaer daf�r
'
' Vorgehensweise
' - Initialisiert einen String mit 1k Gr��e und schreibt diesen im Abstand von 0,1Sek auf die Festplatte
' Erzeugt also eine Dauerlast von ca 10kbyte/Sek, was einen Server nicht wirklich st�ren sollte
'
' Voraussetzung
' - Schreibrechte im Zielverzeichnis
'
' Achtung: Das Skript kann nur mit CTRL-C abgebrochen werden
' Timer() gibt die milliseconds zwischen 00:00 und der aufrufzeit in Sekunden zur�ck (single)
'
' Version 1.0 (09. Nov 2007 )
'	Erste Version abgeleitet von DiskTest 1.1
'-------------------------------------------------------------------------

' Konfigurationskonstanzen
const conTestfilename = "diskprobe.tmp"	'Testdatei, die beschrieben wird
'~ const conCSVfilename = "diskprobe.csv"	'Ausgabe der aktuellen  Messwerte als LOG
const conIdleTime = 100	' Ruhezeit zwischen zwei Schreibvorg�ngen
const Alarmdelta = 1000	'Maximale abweichung des Messwerts vom mittelwert in ms
const Buffersize = 10240	'Gr��e des zu schreibenden Buffers

' interne Konstanten
Const ForWriting = 2
Const ForOverwriting= true
Const ForAppending = false


dim objDebug
set objdebug = new DebugWriter 
objDebug.target = "file:6 console:3" ' errorlogging  0=only output, 1=Error 2=Warning 3=information  5++ =debug
objDebug.outFile = "end2end-file-" & Date() & "-" & Time() &".log"
objDebug.start
objDebug.writeln "end2end-file: gestartet", 0

dim StrWriteBuffer
objDebug.writeln "end2end-file: gestartet", 0
objDebug.writeln ":Initialize Buffer START",0
StrWriteBuffer = space (10240)

objDebug.writeln ":Initialize Eventlog Writer",0
Dim objShell 
Set objShell = CreateObject ("WScript.Shell") 
objShell.LogEvent 0, "end2end-file: gestartet"

'~ objDebug.writeln ":Initialize Logfile"
dim fs, file, logfile
Set fs = CreateObject("Scripting.FileSystemObject")
'~ Set logfile = fs.OpenTextFile(conCSVfilename, ForWriting, ForOverwriting)
'~ logfile.writeline "timestam;performance"

dim performance, mittelwert, count, max, dtstart, dtstop, message , alive
mittelwert = 0 : count  = 0 : max = 0 : message = "" : alive=0
objDebug.writeln ":Start Writing",0
do
	if alive > 6000 then	' dump regular "alive" messages to eventlog nearly every 10+WriteTime Minutes 
		objDebug.writeln "end2end-file: alive",0
		objShell.LogEvent 0, "end2end-file: alive"
		alive = 0
	else
		alive = alive+1
	end if

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
			objDebug.writeln message,1
			objShell.LogEvent 1, message 
		else
			mittelwert = mittelwert + (performance - mittelwert)/10  ' verschiebe mittelwert median by 10% 
		end if
	end if

	if performance > max then max = performance

	if count > 9 then 
		objDebug.writeln " Mittel:" & cint(mittelwert) & "ms  Max:" & Max & "ms ",3 '& " Speed:" & clng(10/(mittelwert +0.1)) & "MB/Sek"
		count = 0 : max = 0 
	else
		count = count  + 1
		wscript.stdout.write "."
	end if 
		
	'~ wscript.echo now & " " & formatnumber(count/10,1,-1) & "Mb/Sek " & replace(space(count/50)," ","#")
	'~ logfile.writeline now & ";" & count/10
	wscript.sleep(conIdleTime)
loop




class debugwriter
	' Generic Class for writing debugging information

	private objIE, file, fs, debugfilename, status, strline
	private debuglevelIE , debuglevelfile, debugleveleventlog, debuglevelConsole

	private Sub Class_Initialize
		status = "active" : strline = "" : debugfilename = ""
		debuglevelIE = -1
		debuglevelfile = -1 
		debugleveleventlog = -1
		debuglevelConsole = -1
	End Sub

	private Sub Class_Terminate()
		if isobject(OBJIE) then
			objie.document.write "</table></body></html>"
		end if
		if debugfilename <> "" then
			file.Close
		end if
	End Sub

	public sub start : status = "active": end sub
	public sub pause : status = "pause" : end sub

	public property let outfile(wert) 	
		if debugfilename <> "" then    'Close existing debug file
			file.close : file = nothing : fs = nothing 
		end if

		debugfilename = wert    ' open debug file
		Set fs = CreateObject("Scripting.FileSystemObject")
		Set file = fs.OpenTextFile(makefilename(debugfilename), 8, True)
	end property

	public property let setie (wert)  : set objIE = wert  : objie.visible = true  end property

	public property let target (wert)
		dim arrTemp, intcount
		arrTemp = split(wert," ")  ' spit by space
		for intcount = 0 to ubound(arrTemp)
			select case lcase(split(arrtemp(intcount),":")(0))
				case "ie" 		debuglevelIE = cint(right(arrtemp(intcount),1))
				case "file" 	debuglevelfile = cint(right(arrtemp(intcount),1))
				case "eventlog" debugleveleventlog = cint(right(arrtemp(intcount),1))
				case "console" 	debuglevelConsole = cint(right(arrtemp(intcount),1))
			end select
		next
	end property

	sub write(strMessage)  
		strline = strline & strMessage
	end sub

	Sub writeln(strMessage, intseverity)
	'F�gt einen Eintrag in die Log-Datei ein
		strMessage = strline & strMessage
		if (status = "active") Then
           if (debuglevelfile >= intseverity) and (debugfilename <> "") then
                file.Write(Now & ",")
                Select Case intseverity
                    Case 0  file.Write("Out0")
                    Case 1  file.Write("Err1")
                    Case 2  file.Write("Wrn2")
                    Case 3  file.Write("Inf3")
                    Case Else file.Write("Dbg"&intseverity)
                End Select
                file.WriteLine("," & Convert2Text(strMessage))
            end if

           if debugleveleventlog >=intSeverity then
                dim objWSHShell
				Set objWSHShell = Wscript.CreateObject("Wscript.Shell")
                Select Case intseverity
                    Case 0  objWSHShell.LogEvent 0, strMessage '           		Const EVENT_SUCCESS = 0
                    Case 1  objWSHShell.LogEvent 1, strMessage '           		const EVENT_ERROR = 1
                    Case 2  objWSHShell.LogEvent 2, strMessage '           		Const EVENT_WARNING = 2
                    Case else  objWSHShell.LogEvent 4, strMessage '           		Const EVENT_INFO = 4
                End Select
           end if

           if debuglevelconsole >=intSeverity then
                Select Case intseverity
                    Case 0  wscript.echo now() & ",OUT0:" & strMessage
                    Case 1  wscript.echo now() & ",ERR1:" & strMessage
                    Case 2  wscript.echo now() & ",WRN2:" & strMessage
                    Case 3  wscript.echo now() & ",INF3:" & strMessage
                    Case Else wscript.echo now() & ",DBG" & intseverity & ":" & strMessage
                End Select

           end if

           if debuglevelie >=intSeverity then
           		dim strieline
      			if  not isobject(objIE) then
      				Set objIE = CreateObject("InternetExplorer.Application")
           		    objIE.navigate("about:blank")
					objIE.visible = true
					Do While objIE.Busy
				    	WScript.Sleep 50
					Loop
					objIE.document.write "<html><head><title>DebugWriter Output</title></head><body>"
					objIE.document.write "<table  border=""1"" width=""100%""><tr><th>Time</th><th>intseverity</th><th>Description</th></tr>"
				end if
           		strieline = "<tr><td>" & now () & "</td>"
                Select Case intseverity
                    Case 0  strieline = strieLine & "<td bgcolor=""#00FF00"">Out0</td>"
                    Case 1  strieline = strieLine & "<td bgcolor=""#FF0000"">Err1</td>"
                    Case 2  strieline = strieLine & "<td bgcolor=""#FFFF00"">Wrn2</td>"
                    Case 3  strieline = strieLine & "<td>Inf3</td>"
                    Case Else strieline = strieLine & "<td>Dbg"&intseverity&"</td>"
                End Select
                strieline = strieline & "<td>" & strmessage & "</td></tr>"
				objIE.document.write cstr(strieline)
           end if

           '~ if (instr(DebugTarget,"mom") <>0) then
				'~ scriptContext.echo now() &","& intseverity &":"& strline & strMessage
           '~ end if

		end if  ' if status = active
		strline = ""
   	End Sub
	
	
	private function makefilename(wert)
		' Converts all invalid characters to valid file names
		wert = replace(wert,"\","-")
		wert = replace(wert,"/","-")
		wert = replace(wert,":","-")
		wert = replace(wert,"*","-")
		wert = replace(wert,"?","-")
		wert = replace(wert,"<","-")
		wert = replace(wert,"|","-")
		wert = replace(wert,"""","-")
		makefilename = wert
	end function
	
	private function Convert2Text(wert) 	' Converts non printable characters to "X" , so that Textfile is working
		dim loopcount, tempwert, inttest
		tempwert=""
		for loopcount = 1 to len(wert)   ' replace all unprintable characters  maybe easier and faster with RegEx
			tempwert = tempwert & chr(ascb(mid(wert,loopcount,1)))	
		next
		Convert2Text = tempwert
	end function
	
end class
