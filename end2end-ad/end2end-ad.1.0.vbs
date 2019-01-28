' End2End-dcupdate.vbs
'
' Vbscript to start on a DC using Taskmanager as local system
' Modifies the "description"-Property of the DC Computerobject on th DC itself
' put a timestamp in there
'
' You can then check that timestamp on any other DC an see the last update timestamp
' Helps you to find latencies.
'
' Test interactive as local system with Sysinternals:  "psexec -i -s cmd.exe"
'
' Get the computerName of PC
' there are many ways. see http://www.robvanderwoude.com/vbstech_network_names_computer.php
' Retrieving Basic Logon and Computer Information 
' http://technet.microsoft.com/en-us/library/ee198776.aspx
'
' 20130118  Version 1.0  Initial Version

option explicit

' Property to modify
' Using a GC Property is a good idea to check cross domain replication
' http://msdn.microsoft.com/en-us/library/ms675094.aspx

' const adproperty = "Description"  ' Self has no permission to update
' const adproperty = "street"  ' Self can update and is in GC
const adproperty = "telephoneNumber"  ' in GC and personal but may disturb CTI/Lync and others
const dc = "localhost"

wscript.echo "INF:Loading ADSystemInfo"
dim objSysInfo,strComputerDN
Set objSysInfo = CreateObject("ADSystemInfo")
strComputerDN = objSysInfo.ComputerName
WScript.Echo "INF:Computer Name: " & strComputerDN

dim oObject

On Error Resume Next
Err.Clear
wscript.echo "INF:Binding LDAP-Object using DC:" & dc
set oObject = GetObject("LDAP://" & dc & "/"+ strComputerDN)
'set oObject = GetObject("LDAP://"& strComputerDN)
If (Err.number <> 0) Then
	WScript.Echo "ERR:Unable to bind ComputerAccount"
	WScript.Echo " Error: " & Err.Number
	WScript.Echo Err.Descritption 
	Wscript.Quit 2
Else
	wscript.echo "INF:Property:" &  adproperty &  " = " &  oObject.get(adproperty)
	If (Err.number <> 0) Then
		WScript.Echo "WRN:Property not found:" & Err.Number
		WScript.Echo Err.Description 
		Err.Clear
	End If

	Dim strTs
	strTs = FormatDateTime(Date(), 0) & " " & FormatDateTime(Time(), 3)
	wscript.echo "OK :Local timestamp:" & strTs

	const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20
	dim objWMIService,colItems,objItem,ShowUTCTimeNow
	Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UTCTime", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objItem In colItems
		ShowUTCTimeNow = objItem.Year & "/" & Right("0" & objItem.Month,2) & "/" & Right("0" & objItem.Day,2) & " " & Right("0" & objItem.Hour,2) & ":"& Right("0" & objItem.Minute,2) & ":" & Right("0" & objItem.Second,2)
	Next
	If (Err.number <> 0) Then
		WScript.Echo "ERR:Unable to generate UTC Timestamp" & Err.Number
		WScript.Echo Err.Description 
		Err.Clear
		Wscript.Quit 3
	else
		wscript.echo "INF:UTC   timestamp:" & ShowUTCTimeNow
		wscript.echo "INF:Update Property:" & adproperty & " = " & ShowUTCTimeNow
		oObject.put adproperty,"end2end-dcupdate:"&ShowUTCTimeNow
		oObject.setinfo
		If (Err.number <> 0) Then
			WScript.Echo "ERR:Unable to update Property:" & Err.Number
			WScript.Echo Err.Description 
			Err.Clear
			Wscript.Quit 4
		else
			WScrpt.Echo "OK :Timezone updated"
		End If
	End If
End If

WScript.Echo "INF:End of Script"

Wscript.Quit 0





