option explicit

'-------------------------------------------------------------------------
' end2end-msxlatency
'
' Beschreibung
' Sammelt per WMI die Exchange Store Latency und alarmiert bei Überschreitung
'
' Laufzeitfehler werde nicht abgefangen und beenden das Skript. !!!
'
' Das Skript wird mit den Berechtigungen des angemeldeten Benutzers
' ausgeführt. 
'
' (c)2007 Net at Work Netzwerksysteme GmbH
'
' Version 1.0 (20. Nov 2007) Frank Carius
'
' Version 1.1 (02. Dez 2008) Holger Daßler
'		zusätzliche Ausgabe der Ereignisse in CSV-Datei
'-------------------------------------------------------------------------

const conMinLatency = 2000  	' minimum latency required to get notification
const conAdjustment = 10  	' weight of new latency to be added to average
const conMaxDelta = 10  	' maximum delta factor between average and current latency 

dim objDebug
set objdebug = new DebugWriter 
objDebug.target = "file:6 console:3" ' errorlogging  0=only output, 1=Error 2=Warning 3=information  5++ =debug
objDebug.outFile = "end2end-msxlatency-" & Date() & "-" & Time() &".log"
objDebug.start

Dim boolNoCSVHeader
dim csvfile
Dim strCSVFileName

set csvfile = new clsCSVWriter
boolNoCSVHeader = False
strCSVFileName = "end2end-msxlatency.csv"

boolNoCSVHeader = csvfile.Exist(strCSVFileName)
csvfile.OpenFile strCSVFileName,csvfile.Append
csvfile.Delimiter = ";"
csvfile.AddHeader "ClientIP"
csvfile.AddHeader "ClientName"
csvfile.AddHeader "ClientMode"
csvfile.AddHeader "ClientVersion"
csvfile.AddHeader "LastOpTime"
csvfile.AddHeader "Latency"
csvfile.AddHeader "Average"
csvfile.AddHeader "MailboxDispName"
csvfile.AddHeader "MailboxLegacyDN"
csvfile.AddHeader "ServerName"
csvfile.AddHeader "StorageGroup"
csvfile.AddHeader "StoreName"
If Not boolNoCSVHeader Then csvfile.WriteHeader ""

objDebug.writeln "end2end-msxlatency: START", 0

objDebug.writeln "end2end-msxlatency: WMI to localhost INIT", 5
dim objSWbemServices
Set objSWbemServices = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\MicrosoftExchangeV2")
objDebug.writeln "end2end-msxlatency: WMI to localhost DONE", 0

objDebug.writeln "end2end-msxlatency: WMI ExecNotify START", 5
dim objEventSource
Set objEventSource = objSWbemServices.ExecNotificationQuery("SELECT * FROM __InstanceOperationEvent  WITHIN 10 WHERE TargetInstance ISA 'Exchange_Logon' ")
objDebug.writeln "end2end-msxlatency: WMI ExecNotify DONE", 0

objDebug.writeln "end2end-msxlatency: Create Dictionary START", 5
dim  dictlatency
set dictlatency = CreateObject("Scripting.dictionary")
objDebug.writeln "end2end-msxlatency: Create Dictionary DONE", 5

dim objEventObject, strMessage, intlatency, dictkey 

Dim objShell 
Set objShell = CreateObject ("WScript.Shell") 

dim count
count = 0
'while count < 500
While true 
	count = count +1
		objDebug.writeln "end2end-msxlatency: ---- Wait for WMI Event -----", 0
	Set objEventObject = objEventSource.NextEvent()
	
	intlatency = objEventObject.targetinstance.Latency
	objDebug.writeln "end2end-msxlatency: Latency = " & intlatency, 5
	dictkey = objEventObject.targetinstance.ClientIP & objEventObject.targetinstance.LoggedOnUserAccount & objEventObject.targetinstance.MailboxLegacyDN
	objDebug.writeln "end2end-msxlatency: dictkey = " & dictkey , 5

	Select Case objEventObject.Path_.Class
		Case "__InstanceDeletionEvent"
			objDebug.writeln "end2end-msxlatency: WMI __InstanceDeletionEvent fired", 5
			if dictlatency.exists (dictkey) then
				objDebug.writeln "end2end-msxlatency: DEL Latency to Dictionary", 0
				dictlatency.remove (dictkey) 
			end if
		Case "__InstanceCreationEvent"
			objDebug.writeln "end2end-msxlatency: WMI __InstanceCreationEvent fired", 5
		Case "__InstanceModificationEvent"
			objDebug.writeln "end2end-msxlatency: WMI __InstanceModificationEvent fired", 5
	End Select
	
	if isnumeric(intlatency) then
		if intlatency> 1 then  ' Skip entries with 0
			if not dictlatency.exists (dictkey) then
				objDebug.writeln "end2end-msxlatency: NEW Latency to Dictionary" & intlatency, 0
				dictlatency.add dictkey, intlatency
			else 
				if  (intlatency > (dictlatency(dictkey)* conMaxDelta)) AND (intlatency > conMinLatency) then  ' current latency 100 times higher than median
					strMessage = 	"end2end-msxlatency: High Latency detectes at" & vbcrlf & _
							"ClientIP	:" & objEventObject.targetinstance.ClientIP & vbcrlf & _
							"ClientName	:" & objEventObject.targetinstance.ClientName & vbcrlf & _
							"ClientMode	:" & objEventObject.targetinstance.ClientMode & vbcrlf & _
							"ClientVersion	:" & objEventObject.targetinstance.ClientVersion & vbcrlf & _
							"LastOpTime	:" & objEventObject.targetinstance.LastOperationTime & vbcrlf & _
							"Latency	:" & objEventObject.targetinstance.Latency & vbcrlf & _
							"Average	:" & dictlatency(dictkey) & vbcrlf & _
							"MailboxDispName:" & objEventObject.targetinstance.MailboxDisplayName & vbcrlf & _
							"MailboxLegacyDN:" & objEventObject.targetinstance.MailboxLegacyDN & vbcrlf & _
							"ServerName	:" & objEventObject.targetinstance.ServerName & vbcrlf & _
							"StorageGroup	:" & objEventObject.targetinstance.StorageGroupName & vbcrlf & _
							"StoreName	:" & objEventObject.targetinstance.StoreName
					objDebug.writeln strMessage, 2
					objShell.LogEvent 1, strMessage
					csvfile.StartLine
					csvfile.AddField "ClientIP", objEventObject.targetinstance.ClientIP
					csvfile.AddField "ClientName", objEventObject.targetinstance.ClientName
					csvfile.AddField "ClientMode", objEventObject.targetinstance.ClientMode
					csvfile.AddField "ClientVersion", objEventObject.targetinstance.ClientVersion
					csvfile.AddField "LastOpTime", objEventObject.targetinstance.LastOperationTime
					csvfile.AddField "Latency", objEventObject.targetinstance.Latency
					csvfile.AddField "Average", dictlatency(dictkey)
					csvfile.AddField "MailboxDispName", objEventObject.targetinstance.MailboxDisplayName
					csvfile.AddField "MailboxLegacyDN", objEventObject.targetinstance.MailboxLegacyDN
					csvfile.AddField "ServerName", objEventObject.targetinstance.ServerName
					csvfile.AddField "StorageGroup", objEventObject.targetinstance.StorageGroupName
					csvfile.AddField "StoreName", objEventObject.targetinstance.StoreName
					csvfile.Writeline
				else
					objDebug.writeln "end2end-msxlatency: UPDATE latency to Dictionary", 0
					dictlatency(dictkey) = (dictlatency(dictkey) + (intlatency-dictlatency(dictkey))/conAdjustment)   ' add current latency
				end if 
			end if
		else
			objDebug.writeln "end2end-msxlatency: SKIP Latency < 1  :" & intlatency, 5
		end if
	end if
	objDebug.writeln "end2end-msxlatency: End of WMI Event", 5
wend




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
	'Fügt einen Eintrag in die Log-Datei ein
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

class clsCSVWriter
	' Class to write generic CSVFiles
	' Version 1.0 Initial Version
	' Version 1.1  WriteHeader + Prefix, WriteRawLine
	' Version 1.2  Add "Exists"-Methode  und Fehler in "Append"-Methode
	' Last Modified: 30. Jan 2008
	' Pending: Quotataion of '"'-Character in Data !!
	
	private csvfilename, csvfs, csvfile, chrDelimiter, chrQuote, strline
	private dictLine

	private Sub Class_Initialize
		chrDelimiter =";" : chrQuote = """" : csvfilename = ""
		set dictLine = createobject("scripting.dictionary")
	End Sub	
	
	public property let Delimiter (wert)
	' Konfigure the delimiter. Default is ";"
		chrDelimiter =wert
	end property

	public property let Quote(wert)
	' Konfigure the sting enquoting. Default is "
		chrQuote = wert
	end property

	public property get Overwrite
	' Contant for Filemode
		Overwrite = 2
	end property

	public property get Append
	' Contant for Filemode
		Append = 8
	end property

	public property get Exist(strFile)
		Set csvfs = CreateObject("Scripting.FileSystemObject")
		if csvfs.FileExists(strFile) then
			Exist = true
		else
			Exist = False
		end if
	end property

	sub OpenFile(wert,intFileMode)
	' Open and start a new CSV-File
		if csvfilename <> "" then    'Close existing debug file
			csvfile.close : csvfile = nothing : csvfs = nothing
		end if

		csvfilename = wert    ' open debug file
		Set csvfs = CreateObject("Scripting.FileSystemObject")
		Set csvfile = csvfs.OpenTextFile(csvfilename, intFileMode, True)
	end sub

	sub AddHeader(strvalue)
	' Add a new column to the csv dataset
		if dictLine.exists(strvalue) then 
			objDebug.writeln "CSVWriter: duplicate Header definition:" & strValue, 1
		else
			dictLine.add strvalue, empty
		end if
	end sub	

	Sub WriteHeader(strPrefix)
	' Write the current Header Definition to the file. optional with a prefix 
	' Prefix can be used to fake IISLogs with "# Field: "
		dim key, strline 
		strline = ""
		for each key in dictLine.keys
			if strline <> "" then 
				strline = strline & chrDelimiter 
			end if 
			strline = strline & chrQuote & cstr(key) & chrQuote
		next
		csvfile.WriteLine(strPrefix & strline)
	End Sub

	sub AddField(strfieldname,strvalue)
	' add a valuue together with the field name
		if dictLine.exists(strfieldname) then 
			dictLine.item(strfieldname) = strvalue
		else
			objDebug.writeln "CSVWriter: Field not declared:" & strFieldname, 1
		end if
	end sub
	
	Sub WriteLine
	' Write the current filled fields to the disk and starte a new line
		dim key, strline 
		strline = ""
		for each key in dictLine.keys
			if strline <> "" then 
				strline = strline & chrDelimiter 
			end if 
			strline = strline & chrQuote & dictLine(key) & chrQuote
			dictLine.item(key) = empty
		next
		csvfile.WriteLine(strline)
   	End Sub

	Sub WriteRawLine(strLine)
	' Write line without any formatting etc. Ideal for comments and other custom output
		csvfile.WriteLine(strline)
   	End Sub
	
	sub StartLine
	' Start a new line. Remove all existing data of the current line
		dim key, strline 
		strline = ""
		for each key in dictLine.keys
			dictLine.item(key) = empty
		next
   	End Sub

end class
