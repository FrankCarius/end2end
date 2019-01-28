' http://msdn2.microsoft.com/en-us/library/aa143724.aspx

'===============================================================
' Purpose:   Display each Exchange_Logon found for Exchange server,
'            and show all properties on the Exchange_Logon
'            objects
' Change:    cComputerName [string] the computer to access
' Output:    Displays the name of each Exchange_Logon and properties
'===============================================================

On Error Resume Next
Dim cComputerName
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_Logon"
cComputerName = "MyComputerNETBIOSName"

Dim strWinMgmts		' Connection string for WMI
Dim objWMIExchange	' Exchange Namespace WMI object
Dim listExchange_Logons	' ExchangeLogons collection
Dim objExchange_Logon		' A single ExchangeLogon WMI object

' Create the object string, indicating WMI (winmgmts), using the
' current user credentials (impersonationLevel=impersonate),
' on the computer specified in the constant cComputerName, and
' using the CIM namespace for the Exchange provider.
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//"& _
cComputerName&"/"&cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
' Verify we were able to correctly set the object.
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Unable to connect to the WMI namespace."
Else
  '
  ' The Resources that currently exist appear as a list of
  ' Exchange_Logon instances in the Exchange namespace.
  Set listExchange_Logons = objWMIExchange.InstancesOf(cWMIInstance)
  '
  ' Were any Exchange_Logon Instances returned?
  If (listExchange_Logons.count > 0) Then
    ' If yes, do the following:
    ' Iterate through the list of Exchange_Logon objects.
    For Each objExchange_Logon in listExchange_Logons
   Wscript.Echo""
   Wscript.Echo""
       '
       '
       ' Display the value of the AdapterSpeed property.
       WScript.echo "AdapterSpeed            = "& _
        " ["&TypeName(objExchange_Logon.AdapterSpeed)&"] "& _
       objExchange_Logon.AdapterSpeed
       '
       '
       ' Display the value of the ClientIP property.
       WScript.echo "ClientIP            = "& _
        " ["&TypeName(objExchange_Logon.ClientIP)&"] "& _
       objExchange_Logon.ClientIP
       '
       '
       ' Display the value of the ClientName property.
       WScript.echo "ClientName            = "& _
        " ["&TypeName(objExchange_Logon.ClientName)&"] "& _
       objExchange_Logon.ClientName
       '
       '
       ' Display the value of the ClientMode property.
       WScript.echo "ClientMode               = "& _
        " ["&TypeName(objExchange_Logon.ClientMode)&"] "& _
       objExchange_Logon.ClientMode
       '
       '
       ' Display the value of the ClientVersion property.
       WScript.echo "ClientVersion            = "& _
        " ["&TypeName(objExchange_Logon.ClientVersion)&"] "& _
       objExchange_Logon.ClientVersion
       '
       '
       ' Display the value of the CodePageID property.
       WScript.echo "CodePageID               = "& _
        " ["&TypeName(objExchange_Logon.CodePageID)&"] "& _
       objExchange_Logon.CodePageID
       '
       '
       ' Display the value of the FolderOperationRate property.
       WScript.echo "FolderOperationRate      = "& _
        " ["&TypeName(objExchange_Logon.FolderOperationRate)&"] "& _
       objExchange_Logon.FolderOperationRate
       '
       '
       ' Display the value of the HostAddress property.
       WScript.echo "HostAddress              = "& _
        " ["&TypeName(objExchange_Logon.HostAddress)&"] "& _
       objExchange_Logon.HostAddress
       '
       '
       ' Display the value of the LastOperationTime property.
       WScript.echo "LastOperationTime        = "& _
        " ["&TypeName(objExchange_Logon.LastOperationTime)&"] "& _
       objExchange_Logon.LastOperationTime
       '
       '
       ' Display the value of the Latency property.
       WScript.echo "Latency        = "& _
        " ["&TypeName(objExchange_Logon.Latency)&"] "& _
       objExchange_Logon.Latency
       '
       '
       ' Display the value of the LocaleID property.
       WScript.echo "LocaleID                 = "& _
        " ["&TypeName(objExchange_Logon.LocaleID)&"] "& _
       objExchange_Logon.LocaleID
       '
       '
       ' Display the value of the LoggedOnUserAccount property.
       WScript.echo "LoggedOnUserAccount      = "& _
        " ["&TypeName(objExchange_Logon.LoggedOnUserAccount)&"] "& _
       objExchange_Logon.LoggedOnUserAccount
       '
       '
       ' Display the value of the LoggedOnUsersMailboxLegacyDN property.
       WScript.echo "LoggedOnUsersMailboxLegacyDN= "& _
        " ["&TypeName(objExchange_Logon.LoggedOnUsersMailboxLegacyDN)&"] "& _
       objExchange_Logon.LoggedOnUsersMailboxLegacyDN
       '
       '
       ' Display the value of the LogonTime property.
       WScript.echo "LogonTime                = "& _
        " ["&TypeName(objExchange_Logon.LogonTime)&"] "& _
       objExchange_Logon.LogonTime
       '
       '
       ' Display the value of the MacAddress property.
       WScript.echo "MacAddress       = "& _
        " ["&TypeName(objExchange_Logon.MacAddress)&"] "& _
       objExchange_Logon.MacAddress
       '
       '
       ' Display the value of the MailboxDisplayName property.
       WScript.echo "MailboxDisplayName       = "& _
        " ["&TypeName(objExchange_Logon.MailboxDisplayName)&"] "& _
       objExchange_Logon.MailboxDisplayName
       '
       '
       ' Display the value of the MailboxLegacyDN property.
       WScript.echo "MailboxLegacyDN          = "& _
        " ["&TypeName(objExchange_Logon.MailboxLegacyDN)&"] "& _
       objExchange_Logon.MailboxLegacyDN
       '
       '
       ' Display the value of the MessagingOperationRate property.
       WScript.echo "MessagingOperationRate   = "& _
        " ["&TypeName(objExchange_Logon.MessagingOperationRate)&"] "& _
       objExchange_Logon.MessagingOperationRate
       '
       '
       ' Display the value of the OpenAttachmentCount property.
       WScript.echo "OpenAttachmentCount      = "& _
        " ["&TypeName(objExchange_Logon.OpenAttachmentCount)&"] "& _
       objExchange_Logon.OpenAttachmentCount
       '
       '
       ' Display the value of the OpenFolderCount property.
       WScript.echo "OpenFolderCount          = "& _
        " ["&TypeName(objExchange_Logon.OpenFolderCount)&"] "& _
       objExchange_Logon.OpenFolderCount
       '
       '
       ' Display the value of the OpenMessageCount property.
       WScript.echo "OpenMessageCount         = "& _
        " ["&TypeName(objExchange_Logon.OpenMessageCount)&"] "& _
       objExchange_Logon.OpenMessageCount
       '
       '
       ' Display the value of the OtherOperationRate property.
       WScript.echo "OtherOperationRate       = "& _
        " ["&TypeName(objExchange_Logon.OtherOperationRate)&"] "& _
       objExchange_Logon.OtherOperationRate
       '
       '
       ' Display the value of the ProgressOperationRate property.
       WScript.echo "ProgressOperationRate    = "& _
        " ["&TypeName(objExchange_Logon.ProgressOperationRate)&"] "& _
       objExchange_Logon.ProgressOperationRate
       '
       '
       ' Display the value of the RowID property.
       WScript.echo "RowID                    = "& _
        " ["&TypeName(objExchange_Logon.RowID)&"] "& _
       objExchange_Logon.RowID
       '
       '
       ' Display the value of the RPCSucceeded property.
       WScript.echo "RPCSucceeded                    = "& _
        " ["&TypeName(objExchange_Logon.RPCSucceeded)&"] "& _
       objExchange_Logon.RPCSucceeded
       '
       '
       ' Display the value of the ServerName property.
       WScript.echo "ServerName               = "& _
        " ["&TypeName(objExchange_Logon.ServerName)&"] "& _
       objExchange_Logon.ServerName
       '
       '
       ' Display the value of the StorageGroupName property.
       WScript.echo "StorageGroupName         = "& _
        " ["&TypeName(objExchange_Logon.StorageGroupName)&"] "& _
       objExchange_Logon.StorageGroupName
       '
       '
       ' Display the value of the StoreName property.
       WScript.echo "StoreName                = "& _
        " ["&TypeName(objExchange_Logon.StoreName)&"] "& _
       objExchange_Logon.StoreName
       '
       '
       ' Display the value of the StoreType property.
       WScript.echo "StoreType                = "& _
        " ["&TypeName(objExchange_Logon.StoreType)&"] "& _
       objExchange_Logon.StoreType
       '
       '
       ' Display the value of the StreamOperationRate property.
       WScript.echo "StreamOperationRate      = "& _
        " ["&TypeName(objExchange_Logon.StreamOperationRate)&"] "& _
       objExchange_Logon.StreamOperationRate
       '
       '
       ' Display the value of the TableOperationRate property.
       WScript.echo "TableOperationRate       = "& _
        " ["&TypeName(objExchange_Logon.TableOperationRate)&"] "& _
       objExchange_Logon.TableOperationRate
       '
       '
       ' Display the value of the TotalOperationRate property.
       WScript.echo "TotalOperationRate       = "& _
        " ["&TypeName(objExchange_Logon.TotalOperationRate)&"] "& _
       objExchange_Logon.TotalOperationRate
       '
       '
       ' Display the value of the TransferOperationRate property.
       WScript.echo "TransferOperationRate    = "& _
        " ["&TypeName(objExchange_Logon.TransferOperationRate)&"] "& _
       objExchange_Logon.TransferOperationRate
       '
    Next
  Else
    ' If no Exchange_Logon instances were returned,
    ' display that.
    WScript.Echo "WARNING: No Exchange_Logon instances were returned."
  End If
End If

