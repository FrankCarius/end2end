'===============================================================
' Purpose:   Display each Exchange_Logon found for Exchange server,
'            and show all properties on the Exchange_Logon
'            objects
' Change:    cComputerName [string] the computer to access
' Output:    Displays the name of each Exchange_Logon and properties
' Based on http://msdn2.microsoft.com/en-us/library/aa143724.aspx
'===============================================================

On Error Resume Next

Dim strWinMgmts		' Connection string for WMI
Dim objWMIExchange	' Exchange Namespace WMI object
Dim listExchange_Logons	' ExchangeLogons collection
Dim objExchange_Logon		' A single ExchangeLogon WMI object

' Create the object string, indicating WMI (winmgmts), using the
' current user credentials (impersonationLevel=impersonate),
' on the computer specified in the constant cComputerName, and
' using the CIM namespace for the Exchange provider.
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//./root/MicrosoftExchangeV2"
Set objWMIExchange =  GetObject(strWinMgmts)
' Verify we were able to correctly set the object.
If Err.Number <> 0 Then
	WScript.Echo "ERROR: Unable to connect to the WMI namespace."
Else
	'
	' The Resources that currently exist appear as a list of
	' Exchange_Logon instances in the Exchange namespace.
	Set listExchange_Logons = objWMIExchange.InstancesOf("Exchange_Logon")
	'
	' Were any Exchange_Logon Instances returned?
	If (listExchange_Logons.count > 0) Then
		' If yes, do the following:
		' Iterate through the list of Exchange_Logon objects.
		For Each objExchange_Logon in listExchange_Logons
			Wscript.Echo"--------------------------------------"
			WScript.echo "AdapterSpeed        		["&TypeName(objExchange_Logon.AdapterSpeed)&"] "& objExchange_Logon.AdapterSpeed
			WScript.echo "ClientIP            		["&TypeName(objExchange_Logon.ClientIP)&"] "& objExchange_Logon.ClientIP
			WScript.echo "ClientName          		["&TypeName(objExchange_Logon.ClientName)&"] "& objExchange_Logon.ClientName
			WScript.echo "ClientMode          		["&TypeName(objExchange_Logon.ClientMode)&"] "& objExchange_Logon.ClientMode
			WScript.echo "ClientVersion       		["&TypeName(objExchange_Logon.ClientVersion)&"] "& objExchange_Logon.ClientVersion
			WScript.echo "CodePageID          		["&TypeName(objExchange_Logon.CodePageID)&"] "& objExchange_Logon.CodePageID
			WScript.echo "FolderOperationRate 		["&TypeName(objExchange_Logon.FolderOperationRate)&"] "& objExchange_Logon.FolderOperationRate
			WScript.echo "HostAddress         		["&TypeName(objExchange_Logon.HostAddress)&"] " & objExchange_Logon.HostAddress
			WScript.echo "LastOperationTime   		["&TypeName(objExchange_Logon.LastOperationTime)&"] "& objExchange_Logon.LastOperationTime
			WScript.echo "Latency        	  		["&TypeName(objExchange_Logon.Latency)&"] "& objExchange_Logon.Latency
			WScript.echo "LocaleID            		["&TypeName(objExchange_Logon.LocaleID)&"] " & objExchange_Logon.LocaleID
			WScript.echo "LoggedOnUserAccount 		["&TypeName(objExchange_Logon.LoggedOnUserAccount)&"] "& objExchange_Logon.LoggedOnUserAccount
			WScript.echo "LoggedOnUsersMailboxLegacyDN= 	["&TypeName(objExchange_Logon.LoggedOnUsersMailboxLegacyDN)&"] "& objExchange_Logon.LoggedOnUsersMailboxLegacyDN
			WScript.echo "LogonTime       			["&TypeName(objExchange_Logon.LogonTime)&"] "& objExchange_Logon.LogonTime
			WScript.echo "MacAddress     			["&TypeName(objExchange_Logon.MacAddress)&"] "& objExchange_Logon.MacAddress
			WScript.echo "MailboxDisplayName     		["&TypeName(objExchange_Logon.MailboxDisplayName)&"] "& objExchange_Logon.MailboxDisplayName
			WScript.echo "MailboxLegacyDN 			["&TypeName(objExchange_Logon.MailboxLegacyDN)&"] "& objExchange_Logon.MailboxLegacyDN
			WScript.echo "MessagingOperationRate 		["&TypeName(objExchange_Logon.MessagingOperationRate)&"] "& objExchange_Logon.MessagingOperationRate
			WScript.echo "OpenAttachmentCount    		["&TypeName(objExchange_Logon.OpenAttachmentCount)&"] "& objExchange_Logon.OpenAttachmentCount
			WScript.echo "OpenFolderCount 			["&TypeName(objExchange_Logon.OpenFolderCount)&"] "& objExchange_Logon.OpenFolderCount
			WScript.echo "OpenMessageCount 			["&TypeName(objExchange_Logon.OpenMessageCount)&"] "& objExchange_Logon.OpenMessageCount
			WScript.echo "OtherOperationRate     		["&TypeName(objExchange_Logon.OtherOperationRate)&"] "& objExchange_Logon.OtherOperationRate
			WScript.echo "ProgressOperationRate  		["&TypeName(objExchange_Logon.ProgressOperationRate)&"] "& objExchange_Logon.ProgressOperationRate
			WScript.echo "RowID           			["&TypeName(objExchange_Logon.RowID)&"] "& objExchange_Logon.RowID
			WScript.echo "RPCSucceeded    			["&TypeName(objExchange_Logon.RPCSucceeded)&"] "& objExchange_Logon.RPCSucceeded
			WScript.echo "ServerName      			["&TypeName(objExchange_Logon.ServerName)&"] "& objExchange_Logon.ServerName
			WScript.echo "StorageGroupName			["&TypeName(objExchange_Logon.StorageGroupName)&"] "& objExchange_Logon.StorageGroupName
			WScript.echo "StoreName       			["&TypeName(objExchange_Logon.StoreName)&"] "& objExchange_Logon.StoreName
			WScript.echo "StoreType       			["&TypeName(objExchange_Logon.StoreType)&"] "& objExchange_Logon.StoreType
			WScript.echo "StreamOperationRate      		["&TypeName(objExchange_Logon.StreamOperationRate)&"] "& objExchange_Logon.StreamOperationRate
			WScript.echo "TableOperationRate       		["&TypeName(objExchange_Logon.TableOperationRate)&"] "& objExchange_Logon.TableOperationRate
			WScript.echo "TotalOperationRate       		["&TypeName(objExchange_Logon.TotalOperationRate)&"] "& objExchange_Logon.TotalOperationRate
			WScript.echo "TransferOperationRate    		["&TypeName(objExchange_Logon.TransferOperationRate)&"] "& objExchange_Logon.TransferOperationRate
		Next
	Else
		' If no Exchange_Logon instances were returned,
		' display that.
		WScript.Echo "WARNING: No Exchange_Logon instances were returned."
	End If
End If

