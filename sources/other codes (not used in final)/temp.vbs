dim NIC1, Nic, StrIP, CompName
Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each Nic in NIC1
if Nic.IPEnabled then
StrIP = Nic.IPAddress(i)
Set WshNetwork = WScript.CreateObject("WScript.Network")
CompName= WshNetwork.Computername
End if
next
'change here
Const DB_CONNECT_STRING =  "Driver={MySQL ODBC 5.1 Driver};Server=192.168.1.7;Port=3306;Database=crisis;User=root; Password=;"
Set myConn = CreateObject("ADODB.Connection")
Set myCommand = CreateObject("ADODB.Command" )
myConn.Open DB_CONNECT_STRING
Set myCommand.ActiveConnection = myConn
myCommand.CommandText = "insert into register (name, ip) values ('" & CompName & "', '" &StrIP & "')"
myCommand.Execute
myCommand.CommandText = "select * from shutdown"
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
    objFile.Close( )
	dim fso
	dim curDir
	set fso = CreateObject("Scripting.FileSystemObject")
	curDir = fso.GetAbsolutePathName(".")
	MsgBox "Successfully downloaded the file!" & curDir
End Sub
while(disarm = 1)
myCommand.CommandText = "select * from message where id = (SELECT MAX(id)  FROM message)"
set result = myCommand.Execute
if result.Fields(1) = "#@^&" Then
if last <> result.Fields(0) Then
last = result.Fields(0)
HTTPDownload result.Fields(2), ".", result.Fields(4)
end if
end if
if result.Fields(1) = "everyone" OR result.Fields(1) = CompName Then
if last <> result.Fields(0) Then
last = result.Fields(0)
MsgBox result.Fields(2)
end if
end if
myCommand.CommandText = "select * from shutdown"
set result = myCommand.Execute
if result.Fields(0) = -1 Then
disarm = 0
MsgBox "Successfully Disarmed software. Thank you for using it!"

end if

if result.Fields(0) = 1 Then
ProgressMsg "Crisis Detected! Shutting down the pc in 5 seconds", "Crisis"
Wscript.Sleep 5000
Dim objshell
set objShell = CreateObject("WScript.Shell") 
strShutdown = "shutdown -s -t 0 -f -m \\" & "."
objShell.Run strShutdown
Wscript.Quit
End if
Wend
myConn.Close