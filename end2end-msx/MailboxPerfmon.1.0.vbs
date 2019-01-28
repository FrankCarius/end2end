'<SCRIPT LANGUAGE="VBScript">
'-------------------------------------------------------------------------
' MailboxPerfmon
'
' Beschreibung
' Lesen von Kalender mit Performancedaten und Unterbrechnungen
'
' Vorgehensweise
' - Liest den Kalender der angegebenen Postfächer und misst die Laufzeiten
'
' Voraussetzung
' - Ausführender Benutzer muss auf die Kalender LESEN haben
' - Verbindung zum DC/GC und Mailserver

' Eingabe
' 	Konfiguration der zu lesenden Kalender
'
' Ausgabe:
'	Ergebnisbericht als XML-Datei
'	Protokoll als Logdatei
'
' Version 1.0 (08 Mai 2006)
'	Erste Version basierend auf MB-Report
'
'-------------------------------------------------------------------------
'**Start Encode**

Option Explicit
'on error resume next


' http://www.cdolive.com/cdo10.htm
Const CdoPR_CONTAINER_CLASS = &H3613001E
Const CdoPR_CONTENT_COUNT = &H33020003
Const CdoPR_MESSAGE_SIZE = &HE080003
const CdoPR_MESSAGE_CLASS = &H001A001E

dim totalerr, total, totalread
dim strOutFilePrefix
dim objCommand, objConnection, objRecordSet
dim timestart, intReady 
dim strHomeServer, strMailbox, intcount
dim strResult
dim dtStarttime
dim arrMBEntry
dim dtdauer
dim WshShell, result
Set WshShell = CreateObject("WScript.Shell")

dim Mailboxpath(1)  ' Anzahl der Pfade -1
Mailboxpath(0) = "cn=user1,ou=test,dc=msxfaq,dc=local" & vbtab & "\\Top of Information Store\Kalender"
Mailboxpath(1) = "cn=user2ö\, t,ou=test,dc=msxfaq,dc=local" & vbtab & "\\Top of Information Store\Kalender"

call ForceCScript  ' must be rund with CSCRIPT
call abbruch ("MailboxPerfmon Script fortsetzen ?" ,5)  ' Last question to stop

' ----- Initialisierung der Debugging und Loggingoptionen
strOutFilePrefix = "MailboxPerfmon-" & makefilename(Date() & "-" & Time()) ' Pfad und Dateiname der Log-Datei
dim objDebug
set objdebug = new DebugWriter 
objDebug.target = "file:5 console:3" ' ie eventlog
objDebug.outFile = strOutFilePrefix &".log"
objDebug.start

objDebug.writeln "MailboxPerfmon: gestartet", 0

' Start XML-File
dim xmlWriter
set xmlWriter = new XmlTextWriter
xmlWriter.filename = strOutFilePrefix & ".xml"
xmlWriter.Indentation = 4
call xmlWriter.WriteStylesheet("MailboxPerfmon.xsl")
call writexslt("MailboxPerfmon.xsl")
call xmlWriter.WriteStartElement("MailboxPerfmon")
call xmlWriter.WriteElementString("starttime", now())
objDebug.writeln "MailboxPerfmon: XML-Writer initialisiert", 5

do
	call xmlWriter.WriteElementString("timestamp1", now)
	dtdauer = now
	for intcount = lbound (Mailboxpath) to ubound (Mailboxpath)
		arrMBEntry = Split(Mailboxpath(intcount),vbtab)
		call xmlWriter.WriteStartElement("object")
		call xmlWriter.WriteElementString("timestamp2", now)
		dtdauer = now 
		objDebug.writeln "User:" & arrMBEntry(0),0
		call xmlWriter.WriteElementString("user", arrMBEntry(0))
		strMailbox = getUserMailbox(arrMBEntry(0))
		objDebug.writeln "Mailbox :" & strMailbox ,0
		strHomeServer = getUserHomeServer(arrMBEntry(0))
		objDebug.writeln "Homeserver:" & strHomeServer,0
		strResult = inspectmailbox(strHomeServer ,strMailbox)  ' ---Ok lets inspect that mailbox ---
		'~ call xmlWriter.WriteElementString("result", strResult) ' ergebnis protokollieren
		call xmlWriter.WriteElementString("timestamp3", now)
		call xmlWriter.WriteElementString("dauerobject", (now-dtdauer)*86400000)
		call xmlWriter.WriteEndElement  ' of Object
	next
	call xmlWriter.WriteElementString("timestamp4", now)
	call xmlWriter.WriteElementString("dauergesamt", (now-dtdauer)*86400000)

	objDebug.writeln "Idle 5 Sekunden",0

	result = WshShell.Popup("Waiting 5 Seconds", 5, "Stop Script", 33) 'OKCancel(1) + Question (32)
	If result = 2 Then
		exit do
	End If
loop

call xmlWriter.WriteElementString("total", total)
objDebug.writeln "Total Mailboxes checked:" & total, 0
call xmlWriter.WriteElementString("totalread", totalread)
objDebug.writeln "# inspected:" & totalread, 0
call xmlWriter.WriteElementString("totalerr", totalerr)
objDebug.writeln "# Error:" & totalerr, 0
call xmlWriter.WriteElementString("endtime", now())
call xmlWriter.WriteEndElement()
call xmlWriter.Close
wscript.quit(0)


function inspectmailbox (strHomeServer,strMail)
	'~ on error resume next

	dim infostores, mapifolder, mailboxfolder
	dim oMapiMessages, oMapiMessage, oMapiSession

	objDebug.writeln "Create MAPI Session", 3
	Set oMapiSession = CreateObject("MAPI.Session")
	if Err.Number <> 0 Then
		objDebug.writeln "  Error creating MAPI Session Skip Mailbox",1
		inspectmailbox = "ERROR CDOObject"
		totalerr = totalerr + 1
		Exit function
	End If	
	
	objDebug.writeln "  Logon to mailbox " & strMail, 0
	objDebug.writeln  "Logon to Mailbox", 3
	oMapiSession.Logon "", "", False, True, 0, true, strHomeServer & vbLF & strMail
	if Err.Number <> 0 Then
		objDebug.writeln "  Error MAPI Logon",1
		inspectmailbox = "ERROR MapiLogon"
		totalerr = totalerr + 1
		Exit function
	end if

	objDebug.writeln  "  Get Top Level folder", 3
	set infostores = oMapiSession.InfoStores
	set mailboxfolder = infostores(2).RootFolder
	if Err.Number <> 0 Then
		objDebug.writeln "  Error MAPI Get Rootfolder",1
		inspectmailbox = "ERROR MapiGetRootFolder"
		totalerr = totalerr + 1
		Exit function
	end if 	
	
	objDebug.writeln "Mailbox successful opened - Checking Folders",0
	objDebug.writeln  "  Mailbox successful opened - Checking Folders", 3

	call checkfolder(mailboxfolder, "\\")   ' lets parse all folders recursively

	totalread = totalread + 1
	objDebug.writeln " Done - Cleanup",0
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
	inspectmailbox = "OK"
end function

sub checkfolder(folder, path)
	on error resume next
	dim child
	dim intItemcount, intItemSize, strItemClass, intItemUnread, dtLastUnreadDate 
	dim dtdauer
		
	objDebug.writeln "Checking folder " & folder.name , 5
	objDebug.writeln  "- " & path & folder.name & "...", 4
	if lcase(path & folder.name) = lcase(arrMBEntry(1)) then
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

end sub

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

function GetUserHomeServer(userdn)
	'~ on error resume next
	' ermittelt den WINS-Namen des HomeServers aus dem HomeMDB-Feldformat
	'z.B. CN=Postfachspeicher (SRV01),CN=Erste Speichergruppe,CN=InformationStore,CN=SRV01,CN=Servers,CN=AG1,CN=Administrative Groups,CN=MSXFAQ,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=msxfaq,DC=local
	dim oHomeUser, oHomeMDB, oHomeServer
	set oHomeUser = getobject("LDAP://" & replace(userdn,"/","\/"))
	set oHomeMDB = getobject("LDAP://" & (oHomeUser.get("homeMDB")))
	set oHomeServer = getobject("LDAP://" & replace(oHomeMDB.get("msExchOwningServer"),"/","\/"))
	getuserhomeserver = oHomeServer.get("name")
end function

function GetUserMailbox(userdn)
	'~ on error resume next
	' ermittelt den WINS-Namen des HomeServers aus dem HomeMDB-Feldformat
	'z.B. CN=Postfachspeicher (SRV01),CN=Erste Speichergruppe,CN=InformationStore,CN=SRV01,CN=Servers,CN=AG1,CN=Administrative Groups,CN=MSXFAQ,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=msxfaq,DC=local
	dim oHomeUser, oHomeMDB, oHomeServer
	set oHomeUser = getobject("LDAP://" & replace(userdn,"/","\/"))
	GetUserMailbox = oHomeUser.get("mail")
end function



sub writexslt(strfilename)
	on error resume next
	dim txtxsl
	txtxsl = _
		"<?xml version=""1.0"" encoding=""UTF-8"" ?>" & vbcrlf & _
		"<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & vbcrlf & _
		"<xsl:template match=""/"">" & vbcrlf & _
		"<html>" & vbcrlf & _
		"<head><title>MailboxPerfmon Status</title></head>" & vbcrlf & _
		"<body>" & vbcrlf & _
		"<h1>MailboxPerfmon Status</h1> " & vbcrlf & _
		"<h2>Parameters</h2> " & vbcrlf & _
		"<table border=""1"" width=""100%"">" & vbcrlf & _
		"	<tr><th>Starttime</th><td><xsl:value-of select=""MailboxPerfmon/starttime"" /></td></tr>" & vbcrlf & _
		"	<tr><th>EndTime</th> <td><xsl:value-of select=""MailboxPerfmon/endtime"" /></td></tr>" & vbcrlf & _
		"	<tr><th>Objekte insgesamt</th> <td><xsl:value-of select=""MailboxPerfmon/total"" /></td></tr>" & vbcrlf & _
		"	<tr><th>gelesene Objekte</th> <td><xsl:value-of select=""MailboxPerfmon/totalread"" /></td></tr>" & vbcrlf & _
		"	<tr><th>Fehlerhafte Objekte</th> <td bgcolor=""#FF0000""><xsl:value-of select=""MailboxPerfmon/totalerr"" /></td></tr>" & vbcrlf & _
		"</table>" & vbcrlf & _
		"<h2>Details</h2> " & vbcrlf & _
		"<table border=""1"" width=""100%"">" & vbcrlf & _
		"<tr bgcolor=""#808080"">" & vbcrlf & _
		"	<th>Mailbox:</th> " & vbcrlf & _
		"	<th>Pfad:</th> " & vbcrlf & _
		"	<th>Ordnername:</th> " & vbcrlf & _
		"	<th>Klasse</th> " & vbcrlf & _
		"	<th>Klasse2</th> " & vbcrlf & _
		"	<th>Elemente</th> " & vbcrlf & _
		"	<th>Elemente2</th> " & vbcrlf & _
		"	<th>Size</th> " & vbcrlf & _
		"	<th>Size2</th> " & vbcrlf & _
		"	<th>Ungelesen</th> " & vbcrlf & _
		"	<th>Oldest Unread</th> " & vbcrlf & _
		"	<th>Ergebnis</th> " & vbcrlf & _
		"  </tr>" & vbcrlf & _
		"<xsl:for-each select=""MailboxPerfmon/object/folder""> " & vbcrlf & _
		"	<tr>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""../mailbox"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""path"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""name"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemclass"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemclass2"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemcount"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemcount2"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemsize"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""itemsize2"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""intItemUnread"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"	<td>" & vbcrlf & _
		"		<xsl:value-of select=""dtLastUnreadDate"" /> " & vbcrlf & _
		"	</td>" & vbcrlf & _
		"		<xsl:choose>" & vbcrlf & _
		"		<xsl:when test=""../result='OK'""><td><xsl:value-of select=""../result"" /></td></xsl:when>" & vbcrlf & _
		"		<xsl:otherwise><td bgcolor=""#FF0000""><xsl:value-of select=""../result"" /></td></xsl:otherwise>" & vbcrlf & _
		"		</xsl:choose>" & vbcrlf & _
		"	</tr>" & vbcrlf & _
		"</xsl:for-each>" & vbcrlf & _
		"</table>" & vbcrlf & _
		"</body>" & vbcrlf & _
		"</html>" & vbcrlf & _
		"</xsl:template>" & vbcrlf & _
		"</xsl:stylesheet> "

	Const ForWriting = 2
	dim fs, file
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set file = fs.OpenTextFile(strfilename, ForWriting, True)
	file.writeline Convert2Text(txtxsl)
	file.Close
end sub

' ==================================================  Auxilary procedures and classes ==========================

function findGCPath
	on error resume next
	objDebug.write "Looking for GC"
	dim oCont, oGC, strGCPath
	Set oCont = GetObject("GC:")
	For Each oGC In oCont
		findGCPath = oGC.ADsPath
	Next
	objDebug.writeln "strGCPath=" & strGCPath, 5
end function


class XMLTextWriter
'~ Usage in VBScript. Please define StyleSheet and filename first and than start writing the data
'~ set xmlWriter = new XmlTextWriter
'~ xmlWriter.filename = "filename.xml"
'~ xmlWriter.Indentation = 4
'~ call xmlWriter.WriteStylesheet("stylesheet.xsl")
'~ call xmlWriter.WriteStartElement("Root")
'~ call xmlWriter.WriteElementString("starttime", now())
'~ call xmlWriter.WriteEndElement
'~ call xmlWriter.close

	dim xmlfilename	'stores the filename
	dim xmldom  	'DOM Object
	dim xmlparent	'currentNode
	dim xmlroot		'RootNode
	dim xmlobject	'PArent of currentNode
	dim intIndentation

	'~ xmlfile.write "<?xml version='1.0' encoding='ISO-8859-1'?>" & vbcrlf

	private Sub Class_Initialize
		dim xmldeclaration
		Set xmlDom = CreateObject("Microsoft.XMLDOM")
		xmlDom.loadxml "<?xml version='1.0'?>"
		set xmlobject = xmlDom
	End Sub

	private Sub Class_Terminate
		'xmldom.LoadXML getFormattedXML
		xmldom.save(xmlfilename)
	End Sub
	
	public Property let filename(wert)    
		xmlfilename = wert
	End Property
	
	public Property let Indentation(wert)
		' only for Backwards compatibility
	End Property

	public Property let Formatting(wert)
		' writer.Formatting = Formatting.Indented   ' Funktioniert nur mit .nEtg ?
	End Property

	sub Writestylesheet (item)
		dim stylePI 
		Set stylePI = xmlDom.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href="""&item & """")
		xmlDom.appendChild(stylePI)
	end sub

	sub WriteStartElement(item)
		dim xmlobject2
		set xmlobject2 = xmlDom.createElement(item)
		xmlobject.appendchild xmlobject2
		set xmlobject = xmlobject2
	end sub


	sub WriteAttributeString(name,value)  ' ergänzt eine ID zum aktuellen Element
		if isnull(value) then value = ""
		xmlobject.setAttribute name, value
	end sub
	

	sub WriteElementString(item,value)  ' add XML tag and Data
		dim xmldata
		set xmldata = xmlDom.createElement(item)
		if isnull(value) then value = ""
		xmldata.text = value
		xmlobject.appendchild(xmldata)
	end sub
	
	sub WriteEndElement
	' Schliesse den aktuellen Client und gehe ein objekt höher
		set xmlobject = xmlobject.parentnode
	end sub

	sub DeleteEndElement
	' Entferne den letzten Client komplett
		dim xmlobject2
		set xmlobject2 = xmlobject
		set xmlobject = xmlobject.parentnode
		xmlobject.removechild(xmlobject2)
	end sub


	function getXML
	' gebe die aktuelle XML-Information unformtiert aus
		getxml = xmldom.xml
	end function

	function LoadXML(strxml)
	' ersetze die Information durch eine neue XML-Information
		xmldom.loadXML(strxml)
	end function

	sub Flush()
	' Schreibe die aktuelle XML-Struktur als Datei heraus
		xmldom.LoadXML getFormattedXML
		xmldom.save(xmlfilename)
	end sub


	function getFormattedXML
	' Gebe die XML-Struktur formatiert und besser lesbar aus
		dim oStylesheet
		set oStylesheet = CreateObject("Microsoft.XMLDOM")
		oStylesheet.async = False
		oStylesheet.loadXML ("<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf & _
						"<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & vbcrlf & _
						"<xsl:output method=""xml"" indent=""yes"" encoding=""UTF-8"" />" & vbcrlf & _
						"<xsl:template match=""@* | node()"">" & vbcrlf & _
						"   <xsl:copy>" & vbcrlf & _
						"        <xsl:apply-templates select=""@* | node()"" />" & vbcrlf & _
						"   </xsl:copy>" & vbcrlf & _
						"</xsl:template>" & vbcrlf & _
						"</xsl:stylesheet>")
'		getFormattedXML = xmlDOM.transformNode(oStylesheet)
		xmlDOM.transformNodetoObject oStylesheet,getFormattedXML
	end function


	sub close()
		'~ xmldom.LoadXML getFormattedXML
		xmldom.save(xmlfilename)
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


'</SCRIPT>

