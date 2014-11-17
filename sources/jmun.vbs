'get location of desktop

dim WSHShell, desktop, pathstring, objFSO
set objFSO=CreateObject("Scripting.FileSystemObject")
Set WSHshell = CreateObject("WScript.Shell")
desktop = WSHShell.SpecialFolders("Desktop")
pathstring = objFSO.GetAbsolutePathName(desktop)


strURL="http://192.168.1.1/jmun_ajax.php?string="
Set objHTTP = CreateObject("MSXML2.XMLHTTP") 
dim NIC1, Nic, StrIP, CompName
Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
max=9000000
min=0

' Greetings

MsgBox "Please wait while the software connects to the MainFrame Server. You will be greeted as soon as you are connected. Stay tuned..."

For Each Nic in NIC1
if Nic.IPEnabled then
StrIP = Nic.IPAddress(i)
Set WshNetwork = WScript.CreateObject("WScript.Network")
CompName= WshNetwork.Computername
End if
next

'change here
Call objHTTP.Open("GET", strURL + "insert into register (name, ip) values ('" & CompName & "', '" &StrIP & "')&id=0", FALSE) 
objHTTP.Send

Function ProgressMsg( strMessage, strWindowTitle )
    Set wshShell = WScript.CreateObject( "WScript.Shell" )
    strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    If strMessage = "" Then
        On Error Resume Next
        objProgressMsg.Terminate( )
        On Error Goto 0
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"

    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "MsgBox""" & strMessage & """, 4096, """ & strWindowTitle & """" )
    objTempMessage.Close

    On Error Resume Next
    objProgressMsg.Terminate( )
    On Error Goto 0

    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )

    Set wshShell = Nothing
    Set objFSO   = Nothing
End Function

disarm = 1
last = ""

Sub HTTPDownload( myURL, myPath, title )
  ' Standard housekeeping
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg
    Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ' Create a File System Object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ' Check if the specified target file or folder exists,
    ' and build the fully qualified path of the target file
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERROR: Target folder not found."
        Exit Sub
    End If

    ' Create or open the target file
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )

    ' Create an HTTP object
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )

    ' Download the specified URL
    objHTTP.Open "GET", myURL, False
    objHTTP.Send
	MsgBox "Starting download of file titled '" & title & "'. Please be patient, your file will be downloaded in the background. You will be notified as soon as the download gets over!"
    ' Write the downloaded byte stream to the target file
    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next

    ' Close the target file

	MsgBox "Successfully downloaded the file! The file is placed at your Desktop for convinience ( " & pathstring & " )"
End Sub

while(disarm = 1)
Randomize
Call objHTTP.Open("GET", strURL + "select * from message where id = (SELECT MAX(id)  FROM message)&id=0&rnd=" & (Int((max-min+1)*Rnd+min)), FALSE) 
objHTTP.Send
field0 = objHTTP.ResponseText

Call objHTTP.Open("GET", strURL + "select * from message where id = (SELECT MAX(id)  FROM message)&id=1&rnd=" & (Int((max-min+1)*Rnd+min)), FALSE) 
objHTTP.Send
field1 = objHTTP.ResponseText

Call objHTTP.Open("GET", strURL + "select * from message where id = (SELECT MAX(id)  FROM message)&id=2&rnd=" & (Int((max-min+1)*Rnd+min)), FALSE) 
objHTTP.Send
field2 = objHTTP.ResponseText

Call objHTTP.Open("GET", strURL + "select * from message where id = (SELECT MAX(id)  FROM message)&id=4&rnd=" & (Int((max-min+1)*Rnd+min)), FALSE) 
objHTTP.Send
field4 = objHTTP.ResponseText

if field1 = "#@^&" Then
if last <> field0 Then
last = field0
HTTPDownload field2, pathstring, field4
end if
end if

if field1 = "everyone" OR field1 = CompName Then
if last <> field0 Then
last = field0
MsgBox field2
end if
end if

Call objHTTP.Open("GET", strURL + "select * from shutdown&id=0&rnd=" & (Int((max-min+1)*Rnd+min)), FALSE) 
objHTTP.Send
field0 = objHTTP.ResponseText
if field0 = "-1" Then
disarm = 0
MsgBox "Successfully Disarmed software. Thank you for using it!"

end if

if field0 = "1" Then
ProgressMsg "Crisis Detected! Shutting down the pc in 5 seconds", "Crisis"
Wscript.Sleep 5000
Dim objshell
set objShell = CreateObject("WScript.Shell") 
strShutdown = "shutdown -s -t 0 -f -m \\" & "."
objShell.Run strShutdown
Wscript.Quit
End if

Wscript.Sleep 1000
Wend