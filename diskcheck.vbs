'VBScript file diskcheck.vbs
Option Explicit

Dim objWMIService, oShell, alluser, objFso, objFolder, LogFile, DisplayFile, rx, RecNumParse, TimestampParse ' Objects
Dim strComputer, DisplayFileName, LogFileName, strFolder, DiplayFileFullName, LogFileFullName, strMessageBody, LineInput ' Strings
Dim colLoggedEvents, intEvent, objEvent, dtmEventDate, strTimeWritten, intRecordNum, colDisks, objDisk ' Variables
Dim result, result2, result3, result4, PrevHighShownRecNum, PrevHighShownTimestamp, NumNewRecs ' Variables

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")

'For Each objDisk in colDisks
'	drvType = objDisk.DeviceID
'	drvDesc = objDisk.Description
'Next


Set oShell = CreateObject( "WScript.Shell" )
alluser=oShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")
DisplayFileName = "disklog.html"
LogFileName = "disklog.txt"
strFolder = alluser & "\CDP\SnapBack\APPs\analyzedisk\"
DiplayFileFullName = strFolder & DisplayFileName
LogFileFullName = strFolder & LogFileName

' Section to create folder to hold file.
Set objFso = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(strFolder) Then
    Set objFolder = objFSO.GetFolder(strFolder)
Else
   Set objFolder = objFSO.CreateFolder(strFolder)
End If


If objFso.FileExists(DiplayFileFullName) Then
   objFso.DeleteFile DiplayFileFullName
End If 

PrevHighShownRecNum = 0
PrevHighShownTimestamp = ""

Set rx = new RegExp
rx.Pattern = "^Record Number: (\d+)$"
rx.Global = False
rx.IgnoreCase = False


' Before deleting the old log file (the .txt) we look for its first Record Num and corresponding Time Written
' to remember how high we reached previously. That way we can deduce which records are new. Then we will only 
' show those new recotds in the display file (the .html) but still repeat ALL errors found into the new
' log file (the .txt) so they will still be available for reference. 

' We now use the time written as the primary determination of new records, only resorting to the record number
' for more granularity when the times are exactly equal. This allows for the case of the user resetting the 
' System Error log which could cause record numbers to revert to 1 thus invalidating the record number as an
' indication of newer records.

' Also NOTE the following

' Assumption: We use the first Record Num we find as the highest previous. This seems safe because
' ==========  it appears the system log is always presented to usin time descending order.


' Here we capture the previous high
 
If objFso.FileExists(LogFileFullName) Then
   Set LogFile = objFso.OpenTextFile(LogFileFullName, ForReading)
   
   While ( LogFile.AtEndOfStream <> True and PrevHighShownRecNum = 0 ) 
      LineInput = LogFile.ReadLine()
      
      Set RecNumParse = rx.Execute(LineInput)
      If RecNumParse.Count > 0 then
         PrevHighShownRecNum = CLng(RecNumParse(0).SubMatches(0))

         While ( LogFile.AtEndOfStream <> True and PrevHighShownTimestamp = "" ) 
            LineInput = LogFile.ReadLine()
            
            rx.Pattern = "^Time Written : ([\d /:]+)$"
            Set TimestampParse = rx.Execute(LineInput)
            If TimestampParse.Count > 0 then
               PrevHighShownTimestamp = TimestampParse(0).SubMatches(0)
            End If
         Wend

      End If
   Wend
   
   LogFile.Close
   objFso.DeleteFile LogFileFullName
End If 

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent " _
        & "Where Logfile = 'System' " _
        & " AND (   (SourceName = 'disk' )" _
             & " OR (EventCode = 55 AND SourceName = 'ntfs' ) ) ")

'Start writing to disklog.html
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set DisplayFile = objFso.OpenTextFile(DiplayFileFullName, ForWriting, True)
DisplayFile.WriteLine("<!DOCTYPE html>")
DisplayFile.WriteLine("<!-- saved from url=(0014)about:internet -->")
DisplayFile.WriteLine("<html>")
DisplayFile.WriteLine("<body>")
DisplayFile.WriteLine("<big style='font-family: Comic Sans MS;'>Analyze Disk Results:</big>&nbsp;")

If colLoggedEvents.Count > 0 Then

'  Here we count the new records so that we can show how many there are at the head of the display

	NumNewRecs = 0
	For Each objEvent in colLoggedEvents
		If WMIDateStringToDate(objEvent.TimeWritten) < PrevHighShownTimestamp then Exit For
		If WMIDateStringToDate(objEvent.TimeWritten) = PrevHighShownTimestamp then 
		   If objEvent.RecordNumber <= PrevHighShownRecNum then Exit For
		End If
		NumNewRecs = NumNewRecs + 1
	Next

   DisplayFile.WriteLine("<span style='font-weight: bold; color: rgb(232, 47, 15);'>There are " & NumNewRecs & " new errors reported</span>")
   DisplayFile.WriteLine("<br><br>")

   If NumNewRecs <> 0 then
   
      DisplayFile.WriteLine("<table>")
      DisplayFile.WriteLine("   <tr>")
      DisplayFile.WriteLine("      <bold><td>Date</td><td>Event ID</td><td>Description</td><td>Source</td></bold>")
      DisplayFile.WriteLine("   </tr>")

   End If
   
	If objFso.FileExists(LogFileFullName) Then
       	objFso.DeleteFile LogFileFullName
	End If 

	Set LogFile = objFso.CreateTextFile(LogFileFullName, True) 
	strComputer = "."

	' Next section loops through ID properties
	intEvent = 1
	LogFile.WriteLine "Your System Event Log has recorded the following errors."
	LogFile.WriteLine "Please refer to the 'Event Code' and 'Message' for details"
	LogFile.WriteLine "of each error. For more detail click on the appropriate"
	LogFile.WriteLine "link on the Results page."
	LogFile.WriteLine "    Prev hi recnum = " & PrevHighShownRecNum
	LogFile.WriteLine "    Prev hi timestamp = " & PrevHighShownTimestamp
	LogFile.WriteLine ""

	For Each objEvent in colLoggedEvents
    	'For Each objDisk in colDisks
		dtmEventDate = objEvent.TimeWritten
		strTimeWritten = WMIDateStringToDate(dtmEventDate)
		LogFile.WriteLine ("Record Number: " & objEvent.RecordNumber)
		LogFile.WriteLine ("Event Type   : " & objEvent.SourceName)
		LogFile.WriteLine ("Event Code   : " & objEvent.EventCode)
		LogFile.WriteLine ("Time Written : " & WMIDateStringToDate(dtmEventDate))
		LogFile.WriteLine ("Message      : " & objEvent.Message)
		LogFile.WriteLine (" ")

'  Here we test for the new records to decide whether to display them

		If WMIDateStringToDate(objEvent.TimeWritten) > PrevHighShownTimestamp or _
		   ( WMIDateStringToDate(objEvent.TimeWritten) = PrevHighShownTimestamp and _ 
		     objEvent.RecordNumber > PrevHighShownRecNum ) then 

            DisplayFile.WriteLine("   <tr>")
            DisplayFile.WriteLine("      <td>" & CDate(strTimeWritten) & "</td><td style='text-align:center'>" & objEvent.EventCode & "</td><td>" & objEvent.Message & "</td><td style='text-align:center'>" & objEvent.SourceName & "</td>")
            DisplayFile.WriteLine("   </tr>")
      End If

		intRecordNum = intRecordNum + 1
		IntEvent = intEvent + 1
	Next

    DisplayFile.WriteLine("</table>")
Else
   DisplayFile.WriteLine("<span style='font-weight: bold; color: rgb(232, 47, 15);'>There are " & colLoggedEvents.Count & " errors reported</span>")
   DisplayFile.WriteLine("<br><br>")
	Set LogFile = objFso.CreateTextFile(LogFileFullName, True) 
	LogFile.WriteLine "Your System Event Log has recorded NO disk errors."
End If

DisplayFile.WriteLine("</html>")
DisplayFile.WriteLine("</body>")

DisplayFile.Close 
Set DisplayFile = nothing


'		' This function converts the date to a readable format (i.e. 8/15/2007 4:25:23 PM)
Function WMIDateStringToDate(dtmEventDate)
'	    WMIDateStringToDate = dtmEventDate
'	    WMIDateStringToDate = CDate(Mid(dtmEventDate, 5, 2) & "/" & _
	    WMIDateStringToDate = Left(dtmEventDate, 4) & "/" & _
	        Mid(dtmEventDate, 5, 2) & "/" & Mid(dtmEventDate, 7, 2) _
	            & " " & Mid (dtmEventDate, 9, 2) & ":" & _
	                Mid(dtmEventDate, 11, 2) & ":" & Mid(dtmEventDate, _
	                    13, 2)
'	                    13, 2))
End Function
	
WScript.Quit
