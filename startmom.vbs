Option Explicit
Dim intRefreshRt : intRefreshRt = 1000 '1 Second Refresh Rate in Milliseconds

GetScripts
'GetCDPkey
StartMoM

Sub GetScripts
    On Error Resume Next
    
    Dim wmiQuery, colItems, objItem, strCmdLine, intPID, strComputer, imgName
    strComputer = GetLocCompName 'if input box is empty, use local computer name by calling a function
    
    If Ping(strComputer) = False Then 'test connectivity with ping function
        'alert "Computer specified is unreachable!!"
        Exit Sub
    End If
    
    wmiQuery = "Select * From Win32_Process Where (Name='mshta.exe')"  'define WMI query
    
    Set colItems = objWMI(strComputer, wmiQuery) 'retrieve WMI collection by calling the objWMI function
    
    For Each objItem In colItems 'loop through the collection
        strCmdLine = Replace(Replace(objItem.CommandLine, Chr(34), ""), "mshta.exe", "")
        strCmdLine = Trim(strCmdLine) 'format string to pull out the script name being run
		'MsgBox( strCmdLine )
        'WriteFile(strCmdLine)
        intPID = objItem.ProcessID
        imgName = Right( strCmdLine, Len(strCmdLine) - InStrRev(strCmdLine, "\") )
        'MsgBox(imgName)
        If InStr(imgName,"oom.hta")>0 Then
            KillScript intPID, 0
        End If
    Next
    
    Set colItems = Nothing
End Sub

Sub GetCDPkey
    On Error Resume Next
    
    Dim wmiQuery, colItems, objItem, strCmdLine, intPID, strComputer, imgName
    strComputer = GetLocCompName 'if input box is empty, use local computer name by calling a function
    
    If Ping(strComputer) = False Then 'test connectivity with ping function
        'alert "Computer specified is unreachable!!"
        Exit Sub
    End If

    wmiQuery = "Select * From Win32_Process Where (Name='cdpkey.exe')"  'define WMI query
    
    Set colItems = objWMI(strComputer, wmiQuery) 'retrieve WMI collection by calling the objWMI function
    
    For Each objItem In colItems 'loop through the collection
        strCmdLine = Replace(Replace(objItem.CommandLine, Chr(34), ""), "silent", "")
        strCmdLine = Trim(strCmdLine) 'format string to pull out the script name being run
		'MsgBox( strCmdLine )
        'WriteFile(strCmdLine)
        intPID = objItem.ProcessID
        imgName = Right( strCmdLine, Len(strCmdLine) - InStrRev(strCmdLine, "\") )
        'MsgBox(imgName)
        If imgName = "cdpkey.exe" Then
            KillScript intPID, 1
        End If
    Next
	
    Set colItems = Nothing
End Sub

Sub KillScript(intPID, intSub)
    On Error Resume Next
    
    Dim wmiQuery, colItems, objItem, strResponse, strComputer
    
    'strResponse = MsgBox("Are you sure you want to terminate the script with Process ID of: " & intPID & "?", 36, "Confirm Script Termination")
    'If strResponse = vbNo Then Exit Sub
    
    strComputer = GetLocCompName
    
    wmiQuery = "Select * From Win32_Process Where ProcessID = " & intPID 'define WMI query
    
    Set colItems = objWMI(strComputer, wmiQuery)'retrieve WMI collection by calling the objWMI function
    
    For Each objItem In colItems
        objItem.Terminate 'terminate process with PID specified
    Next
    Set colItems = Nothing

	If intSub = 0 Then
		GetScripts 'refresh the process list
	Else
		GetCDPkey
	End If
End Sub

Function GetLocCompName
    'On Error Resume Next
    
    Dim objNetwork
    
    Set objNetwork = CreateObject("WScript.Network")
    GetLocCompName = UCase(objNetwork.ComputerName) 'get local computer name
    Set objNetwork = Nothing
End Function

Function Ping(strRmtPC)
    'On Error Resume Next
    
    Dim wmiQuery, objPing, objStatus, blnStatus
    
    wmiQuery = "Select * From Win32_PingStatus Where Address = '" & strRmtPC & "'"
    
    Set objPing = objWMI(".", wmiQuery)
    
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
            blnStatus = False 'Not Reachable
        Else
            blnStatus = True 'Reachable
        End If
    Next
    
    Ping = blnStatus
    Set objPing = Nothing
End Function

Function objWMI(strComputer, strWQL)
    'On Error Resume Next
    
     Dim wmiNS, objWMIService
    
     wmiNS = "\root\cimv2"
    
     Set objWMIService = GetObject("winmgmts:\\" & strComputer & wmiNS) 'connect to WMI
     Set objWMI = objWMIService.ExecQuery(strWQL) 'execute query
     Set objWMIService = Nothing
End Function

Function ConvertDT(strDT)
    'On Error Resume Next
    
     Dim objTime
    
     Set objTime = CreateObject("WbemScripting.SWbemDateTime")
     objTime.Value = strDT
     ConvertDT = objTime.GetVarDate 'convert UTC to Standard Time
     Set objTime = Nothing
End Function

Function StartMoM
   Dim strComputer, objWMIService, colProcesses
   strComputer = "."
   Set objWMIService = GetObject("winmgmts:" _
       & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   Set colProcesses = objWMIService.ExecQuery _
       ("Select * from Win32_Process Where Name = 'oommenu.html'")
   
   If colProcesses.Count = 0 Then
       	Dim objShell, oShell ' Objects
   		Dim Command, programs, oommenu, parameter ' Variables
   		Set oShell = CreateObject( "WScript.Shell" )
   
   		programs=oShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")
   		oommenu = Chr(34) + programs + "\cdp\oom\oom.hta" + Chr(34)
   		parameter = Chr(34) + "analyzedisk" + Chr(34)
   		'parameter = parameter + " " + Chr(34) + "2" + Chr(34)
   		Set objShell = CreateObject("WScript.Shell")
   
   		Command = "%WinDir%\system32\mshta.exe " + oommenu + " " + parameter
   		'MsgBox(Command)
   		objShell.Run Command
   Else
       MsgBox("MoM Menu is already loaded!")
   End If
End Function
