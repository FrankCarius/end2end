'-------------------------------------------------------------------------
' end2end Mailbox
'
' Beschreibung
' Einfaches Sample um per CDO eine Mailbox dynamisch zu öffnen, die Inhalte zu lesen und die Zeit zu messen
'
' Voraussetzung
' - Ausführender Benutzer muss das Postfach öffnen können
' - Verbindung zum DC/GC und Mailserver

' Eingabe
'
' Ausgabe:
'
' Version 1.0 Initial
' Version 1.1  FC 18. 08.2008 
'	objDebug.writeln Fehler gefixt
'	Array Sample Werte gefixt
#
#  CODE IST NOCH NICHT READY
#
'-------------------------------------------------------------------------

Option Explicit
'on error resume next

dim Mailboxpath(1)  ' Anzahl der Pfade -1
Mailboxpath(0) = "srv01" & vbtab & "user1@msxfaq.local" & vbtab & "\\Top of Information Store\Kalender"
Mailboxpath(1) = "srv01" & vbtab & "Admin@msxfaq.local" & vbtab & "\\Top of Information Store\Kalender"

Const CdoDefaultFolderCalendar = 0
Const CdoDefaultFolderContacts = 5 
Const CdoDefaultFolderDeletedItems = 4 
Const CdoDefaultFolderInbox = 1 
Const CdoDefaultFolderJournal = 6 
Const CdoDefaultFolderNotes = 7
Const CdoDefaultFolderOutbox = 2 
Const CdoDefaultFolderSentItems = 3
const CdoDefaultFolderTasks = 8 

if InStr(1,WScript.FullName,"cscript",vbTextCompare) = 0 Then   ' Prüfung ob mit CSCRIPT gestartet wurde
	wscript.echo "Bitte mit CSCRIPT aufrufen"
	wscript.quit (255)
end if

wscript.echo "CDOSample: started"

' ----- Initialisierung der Debugging und Loggingoptionen
dim strOutFilePrefix
strOutFilePrefix = "end2end-mailbox" & makefilename(Date() & "-" & Time()) ' Pfad und Dateiname der Log-Datei
dim objDebug
set objdebug = new DebugWriter 
objDebug.target = "file:5 console:6" ' ie eventlog
objDebug.outFile = strOutFilePrefix &".log"
objDebug.start

objDebug.writeln "end2end-mailbox: gestartet", 0

dim csvfile
set csvfile = new clsCSVWriter
csvfile.OpenFile "end2end-mailbox.csv",csvfile.Overwrite	' Overwrite
csvfile.Delimiter = ";"
'~ csvfile.Quote = "'"
csvfile.WriteRawLine("# Dies ist ein Kommentar")
csvfile.WriteRawLine("# Created with MSXFAQ.CSVWriter")
csvfile.AddHeader "datetime" : csvfile.AddHeader "Homeserver" : csvfile.AddHeader "Mailbox" : csvfile.AddHeader "Folder"
csvfile.AddHeader "timeopen": csvfile.AddHeader "timefolder" : csvfile.AddHeader "timeread"

'~ ' Ab hier dann 
'~ csvfile.StartLine
'~ csvfile.AddField "field1","Feld1content"
'~ csvfile.AddField "field2","Feld2content"
'~ csvfile.AddField "field3","Feld3content"
'~ csvfile.Writeline	'Zeile rausschreiben

dim mapihelper
set mapihelper = new clsmapihelper

dim result, intcount, arrMBEntry
dim dtStart, dtOpen, dtFolder, dtItem
dim objMailbox, objFolder, objItem

dim WshShell
Set WshShell = CreateObject("WScript.Shell")

do
	for intcount = lbound (Mailboxpath) to ubound (Mailboxpath)
		arrMBEntry = Split(Mailboxpath(intcount),vbtab)
		csvfile.StartLine
		dtStart = now
		objDebug.writeln "datetime:" & dtStart ,5
		csvfile.AddField "datetime" , dtStart
		objDebug.writeln "Homeserver:" & arrMBEntry(0),5
		csvfile.AddField "Homeserver" , arrMBEntry(0)
		objDebug.writeln "Mailbox:" & arrMBEntry(1),5
		csvfile.AddField "Mailbox" , arrMBEntry(1)
		objDebug.writeln "Folder:" & arrMBEntry(2),5
		csvfile.AddField "Folder" , arrMBEntry(2)
		objMailbox = mapihelper.openmailbox(arrMBEntry(0) ,arrMBEntry(1))
		if mapihelper.status <> "OK" then
			objDebug.writeln "Error opening mailbox:" & mapihelper.status,5
			csvfile.AddField "ERROR:" & mapihelper.status
		else
			dtopen  = now - dtstart
			objDebug.writeln "timeopen:" & dtopen ,5
			csvfile.AddField "timeopen", dtopen 
						
			objFolder = mapihelper.OpenFolder(objMailbox, arrMBEntry(2)) 
			if mapihelper.status <> "OK" then
				objDebug.writeln "Error opening folder:" & mapihelper.status,5
				'csvfile.AddField "ERROR:" & mapihelper.status
			else
				dtfolder  = now - dtstart
				objDebug.writeln "timefolder:" & dtfolder,5
				csvfile.AddField "timefolder" , dtfolder
				
				' Optional - Parse every item
				for each objItem in objFolder.messages
					objDebug.writeln "  Item.subect:" & objItem.subject , 6
				next
				dtitem  = now - dtstart
				objDebug.writeln "timeread:" & dtitem,5
				csvfile.AddField "timeread" , dtitem
			end if
		end if
		csvfile.Writeline	'Zeile rausschreiben
	next

	result = WshShell.Popup("Waiting 5 Seconds", 5, "Stop Script", 33) 'OKCancel(1) + Question (32)
	If result = 2 Then
		exit do
	End If
loop
objDebug.writeln "end2end-mailbox: beendet", 0



' ==================================================  Auxilary procedures and classes ==========================


class clsmapihelper
	' Generic Class for handling MAPI/CDO

	dim strStatus
	dim infostores, mapifolder, mailboxfolder
	dim oMapiMessages, oMapiMessage, oMapiSession

	private Sub Class_Initialize
		strStatus = "OK" 
	End Sub

	private Sub Class_Terminate()
	End Sub

	public property get status
	' Contant for Filemode
		status = strStatus
	end property

	'~ public property let outfile(wert) 	
		'~ if debugfilename <> "" then    'Close existing debug file
			'~ file.close : file = nothing : fs = nothing 
		'~ end if

		'~ debugfilename = wert    ' open debug file
		'~ Set fs = CreateObject("Scripting.FileSystemObject")
		'~ Set file = fs.OpenTextFile(makefilename(debugfilename), 8, True)
	'~ end property

	'~ public property let setie (wert)  : set objIE = wert  : objie.visible = true  end property


	function openmailbox(strHomeServer,strMail)
		'~ on error resume next

		objDebug.writeln "Create MAPI Session", 3
		Set oMapiSession = CreateObject("MAPI.Session")
		if Err.Number <> 0 Then
			objDebug.writeln "  Error creating MAPI Session Skip Mailbox",1
			strStatus = "ERROR CDOObject"
			Exit function
		End If	
		
		objDebug.writeln "  Logon to mailbox " & strMail, 0
		oMapiSession.Logon "", "", False, True, 0, true, strHomeServer & vbLF & strMail
		if Err.Number <> 0 Then
			objDebug.writeln "  Error MAPI Logon",1
			strStatus = "ERROR MapiLogon"
			Exit function
		end if

		objDebug.writeln  "  Get Top Level folder", 3
		set infostores = oMapiSession.InfoStores
		dim infostore 
		for each infostore in infostores
			objDebug.writeln  "  Infostore found:" & infostore.name, 3
		next
		set mailboxfolder = infostores(2).RootFolder
		if Err.Number <> 0 Then
			objDebug.writeln "  Error MAPI Get Rootfolder",1
			strStatus = "ERROR MapiGetRootFolder"
			Exit function
		end if 	
		
		objDebug.writeln "Mailbox successful opened",0
		objDebug.writeln  "  Mailbox successful opened", 3
	end function
	

	function OpenFolder(objMbox, strPath)
	
		on error resume next
		dim child
		dim intItemcount, intItemSize, strItemClass, intItemUnread, dtLastUnreadDate 
		dim dtdauer
			
		objDebug.writeln "OpenFolder " & strPath, 5
		
		'~ xxxx hier muss noch Code rein
		
		if lcase(strpath & folder.name) = lcase(arrMBEntry(1)) then
			call xmlWriter.WriteStartElement("folder")
			call xmlWriter.WriteElementString("timestamp5", now)
			dtdauer = now
			call xmlWriter.WriteElementString("name", folder.name)  'OK
			call xmlWriter.WriteElementString("path", path)  'OK
			call xmlWriter.WriteElementString("itemcount", folder.messages.count)  'OK
			call xmlWriter.WriteElementString("itemclass", GetMapiProperty(folder, CdoPR_CONTAINER_CLASS))
			call xmlWriter.WriteElementString("itemsize", GetMapiProperty(folder, CdoPR_MESSAGE_SIZE))
			'call xmlWriter.WriteElementString("itemsize", folder.messages.size)

			intItemcount = 0 
			intItemSize = 0 
			intItemUnread = 0
			dtLastUnreadDate  = 0
			strItemClass = ""
			for each child in folder.messages
				objDebug.writeln "  Item.subect:" & child.subject , 5
				objDebug.writeln "  Item.Messageclass:" & GetMapiProperty(child, CdoPR_MESSAGE_CLASS), 5
				if instr(strItemClass,GetMapiProperty(child, CdoPR_MESSAGE_CLASS)) = 0 then
					strItemClass = strItemClass & GetMapiProperty(child, CdoPR_MESSAGE_CLASS) & "," 
				end if
				objDebug.writeln "  Item.Unread:" & child.unread, 5
				if child.unread = true then 
					intItemUnread = intItemUnread +1
						objDebug.writeln "  Item.Timereceived:" & child.Timereceived , 5  ' Vartype = 7  Date
						objDebug.writeln "  dtLastUnreadDate :" & dtLastUnreadDate , 5  ' Vartype = 7  Date
						err.clear
						if (child.Timereceived < dtLastUnreadDate)or (dtLastUnreadDate = 0) then 
							objDebug.writeln "  Item.Timereceived: Updated" , 5  ' Vartype = 7  Date
							dtLastUnreadDate = child.Timereceived
						end if 
				end if 
				intItemcount = intItemcount + 1
				intItemSize  = intItemSize  + child.size
				objDebug.writeln "  Done Child:", 6
			next
			objDebug.writeln "  Done Folder", 6
			wscript.echo ":"&strItemClass&":"
			if len(strItemClass) > 0 then 
				strItemClass = left(strItemClass,len(strItemClass)-1)  ' remove last ","
			end if
			call xmlWriter.WriteElementString("itemsize2", intitemsize)
			call xmlWriter.WriteElementString("intItemUnread", intItemUnread)
			call xmlWriter.WriteElementString("dtLastUnreadDate", dtLastUnreadDate)
			call xmlWriter.WriteElementString("itemcount2", intitemcount)
			call xmlWriter.WriteElementString("itemclass2", strItemClass) 
			call xmlWriter.WriteElementString("timestamp6", now)
			call xmlWriter.WriteElementString("dauerfolder", (now-dtdauer)*86400000)
			call xmlWriter.WriteEndElement ' of "folder"
			'check the child folders
		else
			objDebug.writeln "     SKIP", 6
		end if
		for each child in folder.folders
			call checkfolder(child, path & folder.name & "\")
		next
	end function

	function GetMapiProperty(oobj, property)
		on error resume next
		dim value
		value = oobj.Fields(property) 
		objDebug.writeln "  Getting MAPI Property" & property,6
		'~ if Err.Number <> 0 Then
			'~ objDebug.writeln "  Error Getting MAPI Property",1
			'~ err.clear
			'~ value = "MISSING"
		'~ End If	
		GetMapiProperty = value
		
	end function


	sub closemailbox
		set mailboxfolder = nothing
		set infostores = nothing
		objDebug.writeln "Logoff from Mailbox ", 0
		objDebug.writeln  "  Logoff from Mailbox ", 3
		oMapiSession.Logoff
		objDebug.writeln "Logoff complete, Releasing MAPI Session", 0
		objDebug.writeln  "  Logoff complete, Releasing MAPI Session", 3
		Set oMapiSession = Nothing
		objDebug.writeln "MAPI-Session released", 0
		objDebug.writeln  "  MAPI-Session released", 3
		status = "OK"
	end sub


end class

class debugwriter
	' Generic Class for writing debugging information and handling runtime errors
	' By default al Level 1 Messaegs are logged to the Console

	private objIE, file, fs, debugfilename, status, strline
	private debuglevelIE , debuglevelfile, debugleveleventlog, debuglevelConsole

	private Sub Class_Initialize
		status = "active" : strline = "" : debugfilename = ""
		debuglevelIE = 0
		debuglevelfile = 0
		debugleveleventlog = 0
		debuglevelConsole = 1
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
		dim blnerror
		
		if err <> 0 then  ' Sonderbehandlung als "ErrorHandler"
			blnerror = true
			strmessage= "RUNTIME ERROR  :" & strMessage & vbcrlf & _
						"ERR.Number     :" & err.number & vbcrlf & _
						"ERR.Description:" & err.Description & vbcrlf & _
						"ERR.Source     :" & err.source & vbcrlf & _
						"ERR.HelpFile   :" & err.HelpFile & vbcrlf & _
						"ERR.HelpContext:" & err.HelpContext & vbcrlf
			err.clear
		else
			blnerror = false
		end if
	
		if ((status = "active")  or blnerror) then 

			if (debuglevelfile > 0) and ((debuglevelfile >= intseverity) or blnerror) and (debugfilename <> "") then
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

			if (debugleveleventlog > 0) and ((debugleveleventlog >=intSeverity) or blnerror) then
                dim objWSHShell
				Set objWSHShell = Wscript.CreateObject("Wscript.Shell")
                Select Case intseverity
                    Case 0  objWSHShell.LogEvent 0, strMessage '           		Const EVENT_SUCCESS = 0
                    Case 1  objWSHShell.LogEvent 1, strMessage '           		const EVENT_ERROR = 1
                    Case 2  objWSHShell.LogEvent 2, strMessage '           		Const EVENT_WARNING = 2
                    Case else  objWSHShell.LogEvent 4, strMessage '           	Const EVENT_INFO = 4
                End Select
			end if

			if (debuglevelconsole > 0) and ((debuglevelconsole >=intSeverity) or blnerror) then
				Select Case intseverity
					Case 0  wscript.echo now() & ",OUT0:" & strMessage
					Case 1  wscript.echo now() & ",ERR1:" & strMessage
					Case 2  wscript.echo now() & ",WRN2:" & strMessage
					Case 3  wscript.echo now() & ",INF3:" & strMessage
					Case Else wscript.echo now() & ",DBG" & intseverity & ":" & strMessage
				End Select
			end if

           if (debuglevelie > 0) and ((debuglevelie >= intSeverity) or blnerror) then
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

           '~ if (instr(DebugTarget,"mom") <>0) or blnerror then
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


sub ForceCScript
    If InStr(1,WScript.FullName,"cscript",vbTextCompare) = 0 Then   ' Prüfung ob mit CSCRIPT gestartet wurde
        wscript.echo "Bitte mit CSCRIPT aufrufen"
        wscript.quit (255)
	end if
end sub


Sub abbruch(info,waittime)
	' Usage: call abbruch ("Script abbrechen" ,5)
	dim WshShell, result
	Set WshShell = CreateObject("WScript.Shell")
	result = WshShell.Popup("Continue script at position "& vbcrlf & info & vbcrlf & "Waiting "&waittime&" Seconds", waittime, "Stop Script", 33) 'OKCancel(1) + Question (32)
	If result = 2 Then
        WScript.echo "Abbruch durch Anwender (Exitcode = 255)"
        WScript.Quit(255)
    End If
End Sub


function makefilename(byVal wert)
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

function Convert2Text(wert)
	' Converts non printable characters to "X" , so that Textfile is working
	dim loopcount, tempwert, inttest
	tempwert=""
	for loopcount = 1 to len(wert)   ' replace all unprintable characters  maybe easier and faster with RegEx
		tempwert = tempwert & chr(ascb(mid(wert,loopcount,1)))	
	next
	Convert2Text = tempwert
end function



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
