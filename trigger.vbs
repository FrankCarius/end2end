strComputer = "."
Set objSWbemServices = GetObject("winmgmts:" &_
    "{impersonationLevel=impersonate}!" &_
    "\\" & strComputer & "\root\cimv2")

Set objEventSource = objSWbemServices.ExecNotificationQuery( _
    "SELECT * FROM __InstanceOperationEvent " &_
    "WITHIN 1 " &_
    "WHERE TargetInstance " &_
    "ISA 'Win32_Process' " &_
    "AND TargetInstance.Name = 'notepad.exe'")

Set objEventObject = objEventSource.NextEvent()
Select Case objEventObject.Path_.Class
    Case "__InstanceCreationEvent"
        Wscript.Echo "Instanz von Notepad.exe gestartet."
    Case "__InstanceDeletionEvent"
        Wscript.Echo "Instanz von Notepad.exe beendet."
    Case "__InstanceModificationEvent"
        Wscript.Echo "Instanz von Notepad.exe geändert."
End Select
