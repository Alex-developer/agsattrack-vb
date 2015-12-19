VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl engine 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "engine.ctx":0000
   PropertyPages   =   "engine.ctx":0442
   ScaleHeight     =   540
   ScaleWidth      =   480
   ToolboxBitmap   =   "engine.ctx":0458
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   330
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' November 27,2000
' Daniel W Ankrom
'FTPx  ActiveX control

Public Event Error()
'Default Property Values:
Const m_def_TimeOut = 60
Const m_def_UserName = "anonymous"
Const m_def_Password = "myemail@company.com"
Const m_def_LastErrorCode = 0
Const m_def_LastErrorDesc = "none"
Const m_def_RemoteHost = "ftp.microsoft.com"
Const m_def_RemotePort = 21
'Property Variables:
Dim m_TimeOut As Integer
Dim m_UserName As Variant
Dim m_Password As Variant
Dim m_LastErrorCode As Long
Dim m_LastErrorDesc As String
Dim m_RemoteHost As String
Dim m_RemotePort As Long

Dim Connected As Boolean           'Flag to indicate connect status
Dim TimedOut As Boolean            'Flag to indicate time out occured
Dim TimePeriod As Integer          ' the number of seconds to timeout
Dim Welcomed As Boolean            'Flag to indicate the FTP server responded with a welcome
Dim LastServerCode As Integer      'The last numeric code received from the FTP server
Dim LastResponse As String         'The last text response from the FTP server
Dim FileReceiveDone As Boolean     'Flag to indicate a file dload transfer is complete
Dim SendComplete As Boolean        'Flag to indicate a file upload transfer is complete
Dim Socks As Integer               'The number of data socks loaded
Dim BlockSize As Long              'For upload, the calculated max block size
Dim RemainSize As Long             'For upload the remainder for odd sized files
Dim TotalBlocks As Long            'For upload, the total blocks to send
Dim RecPosition As Long            'For upload, the current block number
Dim FileSendingSize As Long        'For upload, the size of the file to send
Dim BlockCount As Integer          'For upload, the number of blocks sent
Dim Cancel_Operation As Boolean    'Flag to indicate a Cancel request

'Event Declarations:
Public Event ReceiveProgress(ByVal BytesToGet As Long, BytesGot As Long)
Public Event SendProgress(BytesSent As Long, BytesTotal As Long)
Public Event ServerResponse(Response As String)
Public Event CommandSentToServer(CommandSent As String)






Public Function Connect() As Boolean
Attribute Connect.VB_Description = "Connect to the ftp server. Username, Password, Port must have been previously set. Returns true if successful, False otherwise."
'connects the control to the remote server
'assumes the server address, username, password, port
' and timeout period is already set
'returns false on failure
'writes to registry the failure code

If Connected Then Exit Function
ClearLastError
TimePeriod = TimeOut
TimedOut = False
Timer1.Enabled = True
Winsock(0).RemoteHost = RemoteHost
Winsock(0).RemotePort = RemotePort
Winsock(0).Connect
While Not Connected And Not TimedOut
 DoEvents
Wend
If TimedOut Then
 Connect = False
 SetError 1510, "Connection Request TimedOut"
 Winsock(0).Close
 Exit Function
End If
While Not Welcomed And Not TimedOut
 DoEvents
Wend
If TimedOut Then
 Connect = False
 SetError 1510, "Connection Request TimedOut"
 Winsock(0).Close
 Exit Function
End If
'we are now connected, send user and pass info
SendCommand "USER", UserName, 331
Winsock(0).SendData "PASS " & Password & vbCrLf

While (InStr(1, LastResponse, "230") = 0 And Not TimedOut) And LastServerCode <> 530
 DoEvents
Wend
Timer1.Enabled = False
If InStr(1, LastResponse, "230") <> 0 Then 'all connected well
 Connect = True
 Exit Function
End If
SetError LastServerCode, LastResponse & " " & UserName
Connect = False
Connected = False
Winsock(0).Close
End Function


Public Function Disconnect() As Boolean
Attribute Disconnect.VB_Description = "Disconnects from the remote file."
Dim x As Long
Disconnect = True
If Not Connected Then Exit Function
'Winsock(0).SendData "QUIT" & vbCrLf
'For x = 0 To 50000: Next
Do While Not (SendCommand("QUIT", "", 221))
Loop
Winsock(0).Close
Connected = False
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the name used to identify the remote computer"
Attribute RemoteHost.VB_ProcData.VB_Invoke_Property = "FTPx_Properties"
    RemoteHost = m_RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    m_RemoteHost = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
Attribute RemotePort.VB_ProcData.VB_Invoke_Property = "FTPx_Properties"
    RemotePort = m_RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    m_RemotePort = New_RemotePort
    PropertyChanged "RemotePort"
End Property

Private Sub Timer1_Timer()
Static Ticks As Integer
Ticks = Ticks + 1
If Ticks > TimePeriod Then
 TimedOut = True
 Timer1.Enabled = False
 Ticks = 0
End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_RemoteHost = m_def_RemoteHost
    m_RemotePort = m_def_RemotePort
    m_LastErrorCode = m_def_LastErrorCode
    m_LastErrorDesc = m_def_LastErrorDesc
    m_UserName = m_def_UserName
    m_Password = m_def_Password
    m_TimeOut = m_def_TimeOut
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_RemoteHost = PropBag.ReadProperty("RemoteHost", m_def_RemoteHost)
    m_RemotePort = PropBag.ReadProperty("RemotePort", m_def_RemotePort)
    m_LastErrorCode = PropBag.ReadProperty("LastErrorCode", m_def_LastErrorCode)
    m_LastErrorDesc = PropBag.ReadProperty("LastErrorDesc", m_def_LastErrorDesc)
    m_UserName = PropBag.ReadProperty("UserName", m_def_UserName)
    m_Password = PropBag.ReadProperty("Password", m_def_Password)
    m_TimeOut = PropBag.ReadProperty("TimeOut", m_def_TimeOut)
End Sub

Private Sub UserControl_Terminate()
'make sure we are disconnected
If Winsock(0).State > 1 Then
 Winsock(0).Close
End If
If Socks Then
 Winsock(Socks).Close
End If

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("RemoteHost", m_RemoteHost, m_def_RemoteHost)
    Call PropBag.WriteProperty("RemotePort", m_RemotePort, m_def_RemotePort)
    Call PropBag.WriteProperty("LastErrorCode", m_LastErrorCode, m_def_LastErrorCode)
    Call PropBag.WriteProperty("LastErrorDesc", m_LastErrorDesc, m_def_LastErrorDesc)
    Call PropBag.WriteProperty("UserName", m_UserName, m_def_UserName)
    Call PropBag.WriteProperty("Password", m_Password, m_def_Password)
    Call PropBag.WriteProperty("TimeOut", m_TimeOut, m_def_TimeOut)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function GetFile(RemoteDirectory As String, RemoteFileName As String, LocalfileName As String, AsciiMode As Boolean) As Boolean
Attribute GetFile.VB_Description = "Downloads the remote file. Returns True if sucessful, False otherwise."
'gets the remote file specified by RemoteFileName
'navigates to the remote directory specified by RemoteDirectory
'transfer the file in the mode specified by AsciiMode
'expects the control is already connected
'fires event of ReceiveProgress


Dim strData As String: Dim FileNum As Integer
Dim FileSize As Long: GetFile = False

If Not CheckConnected("GET") Then Exit Function
Cancel_Operation = False
FileSize = RemoteFileSize(RemoteFileName)
If FileSize = 0 Then
 SetError 500, "No Such File!"
 Exit Function
End If

'check for directory change
If RemoteDirectory <> "" Then
 RemoteDirectory = Replace(RemoteDirectory, "/", "\", 1, -1)
 SendCommand "CWD", RemoteDirectory, 250
End If

If AsciiMode Then
 strData = "TYPE A"
Else
 strData = "TYPE I"
End If
If Not SendCommand(strData, "", 200) Then Exit Function

Socks = Socks + 1
Load Winsock(Socks)
Winsock(Socks).Tag = Str(FreeFile()) & " File Receiver" & "[" & Str(FileSize) & "]"
strData = Winsock(0).LocalIP
strData = Replace(strData, ".", ",", 1, -1) & ","
Winsock(Socks).Listen
strData = strData & Str(Fix(Winsock(Socks).LocalPort / 256)) & ","
strData = strData & Str(Winsock(Socks).LocalPort - (Fix(Winsock(Socks).LocalPort / 256) * 256))
strData = Replace(strData, " ", "", 1, -1)

If Not SendCommand("PORT", strData, 200) Then
 Unload Winsock(Socks)
 Socks = Socks - 1
 Exit Function
End If

FileNum = FreeFile(): FileReceiveDone = False
Open LocalfileName For Output As #FileNum
Winsock(0).SendData "RETR " & RemoteFileName & vbCrLf

While Not FileReceiveDone And Not Cancel_Operation And LastServerCode <> 425
 DoEvents
Wend
If Cancel_Operation Or LastServerCode = 425 Then
 If LastServerCode = 425 Then 'connection failed
  Unload Winsock(Socks)
  Socks = Socks - 1
 End If
 Close #FileNum
Else
 GetFile = True
End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function PutFile(RemoteDirectory As String, RemoteFileName As String, LocalfileName As String, AsciiMode As Boolean) As Boolean
Attribute PutFile.VB_Description = "Sends a file. Returns true if sucessful, false otherwise."
Dim strData As String: Dim FileSize As Long
Dim Buffer As String * 8196
PutFile = False: Cancel_Operation = False

If Not CheckConnected("PUT") Then Exit Function
'check for directory change
If RemoteDirectory <> "" Then
 If Not SendCommand("CWD", RemoteDirectory, 250) Then Exit Function
End If

If AsciiMode Then
 strData = "TYPE A"
Else
 strData = "TYPE I"
End If

If Not SendCommand(strData, "", 200) Then Exit Function
Socks = Socks + 1
Load Winsock(Socks)
FileSize = FileLen(LocalfileName)
FileSendingSize = FileSize
Winsock(Socks).Tag = Str(FreeFile()) & " File Sender" & "[" & Str(FileSize) & "]"
strData = Winsock(0).LocalIP
strData = Replace(strData, ".", ",", 1, -1) & ","
Winsock(Socks).Listen
strData = strData & Str(Fix(Winsock(Socks).LocalPort / 256)) & ","
strData = strData & Str(Winsock(Socks).LocalPort - (Fix(Winsock(Socks).LocalPort / 256) * 256))
strData = Replace(strData, " ", "", 1, -1)
If Not SendCommand("PORT", strData, 200) Then Exit Function
FileNum = FreeFile(): SendComplete = False
Open LocalfileName For Binary As #FileNum
BlockSize = 8196
While BlockSize > FileSize
 BlockSize = BlockSize - 1000
Wend
If BlockSize < 1000 Then BlockSize = FileSize
TotalBlocks = Fix(FileSize / BlockSize)
RemainSize = FileSize - (TotalBlocks * BlockSize)
BlockCount = 1: RecPosition = 1
SendCommand "STOR", RemoteFileName, 150
While Winsock(Socks).State <> 7
 DoEvents
Wend
 Get #FileNum, BlockCount, Buffer
 Winsock(Socks).SendData Left(Buffer, BlockSize)
 RecPosition = RecPosition + BlockSize
 RaiseEvent SendProgress(BlockSize, FileSize)
 BlockCount = BlockCount + 1
 
While Not SendComplete And Not Cancel_Operation
 DoEvents
Wend
If Cancel_Operation Then
 Close #FileNum
Else
 Winsock(Socks).Close
 Unload Winsock(Socks)
 PutFile = True
 Socks = Socks - 1
End If
SendCommand "PWD", "", 257
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,0,0
Public Property Get LastErrorCode() As Long
Attribute LastErrorCode.VB_Description = "Returns the last error code that occured."
    LastErrorCode = m_LastErrorCode
End Property

Public Property Let LastErrorCode(ByVal New_LastErrorCode As Long)
    If Ambient.UserMode Then Err.Raise 382
    m_LastErrorCode = New_LastErrorCode
    PropertyChanged "LastErrorCode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,0
Public Property Get LastErrorDesc() As String
Attribute LastErrorDesc.VB_Description = "Returns the last string error description that occured."
    LastErrorDesc = m_LastErrorDesc
End Property

Public Property Let LastErrorDesc(ByVal New_LastErrorDesc As String)
    If Ambient.UserMode Then Err.Raise 382
    m_LastErrorDesc = New_LastErrorDesc
    PropertyChanged "LastErrorDesc"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get UserName() As Variant
Attribute UserName.VB_ProcData.VB_Invoke_Property = "FTPx_Properties"
    UserName = m_UserName
End Property

Public Property Let UserName(ByVal New_UserName As Variant)
    m_UserName = New_UserName
    PropertyChanged "UserName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Password() As Variant
Attribute Password.VB_ProcData.VB_Invoke_Property = "FTPx_Properties"
    Password = m_Password
End Property

Public Property Let Password(ByVal New_Password As Variant)
    m_Password = New_Password
    PropertyChanged "Password"
End Property

Private Sub ClearLastError()
SaveSetting "Engine", vbNullChar, "LastErrorCode", 0
SaveSetting "Engine", vbNullChar, "LastErrorStr", ""
End Sub

Private Sub SetError(ErrorCode As Integer, ErrorString As String)
SaveSetting "Engine", vbNullChar, "LastErrorCode", ErrorCode
SaveSetting "Engine", vbNullChar, "LastErrorStr", ErrorString
End Sub

Private Sub Winsock_Connect(Index As Integer)
If Index = 0 Then
 Connected = True
End If

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Winsock(Index).Close
Winsock(Index).Accept requestID
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal BytesTotal As Long)

Dim strData As String
Static FileNum As Integer
Static SizeToGet As Long
Static SizeGot As Long

Select Case Index
 Case 0
  Winsock(Index).GetData strData
  LastServerCode = Val(Left(strData, 3))
  LastResponse = strData
  RaiseEvent ServerResponse(strData)
  If LastServerCode = 220 Then
   Welcomed = True
  End If
 Case 1
  If InStr(1, Winsock(1).Tag, "File Receiver") <> 0 Then
   FileNum = Val(Winsock(1).Tag)
   SizeToGet = Val(Mid(Winsock(1).Tag, InStr(1, Winsock(1).Tag, "[") + 1))
   SizeGot = SizeGot + BytesTotal
   Winsock(1).GetData strData
   Print #FileNum, strData;
   RaiseEvent ReceiveProgress(SizeToGet, SizeGot)
   If SizeGot >= SizeToGet Then
    FileReceiveDone = True
    Winsock(1).Close
    Unload Winsock(Socks)
    SizeGot = 0
    Close #FileNum
    Socks = Socks - 1
   End If
   Exit Sub
  End If
  If InStr(1, Winsock(1).Tag, "Lister") <> 0 Then
   Winsock(1).GetData strData
   FileNum = Val(Winsock(1).Tag)
   Print #FileNum, strData;
   Exit Sub
  End If
    
 End Select
End Sub



Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'
End Sub

Public Function GetLastErrorCode() As String
GetLastErrorCode = GetSetting("Engine", vbNullChar, "LastErrorCode")
End Function

Public Function GetLastErrorString() As String
GetLastErrorString = GetSetting("Engine", vbNullChar, "LastErrorStr")
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function KillFile(RemoteDirectory As String, RemoteFileName As String) As Boolean
Attribute KillFile.VB_Description = "Deletes the File Specified by RemoteFileName"
KillFile = False
If Not CheckConnected("KILL") Then Exit Function
If RemoteDirectory <> "" Then
 SendCommand "CWD", RemoteDirectory, 250
End If
KillFile = SendCommand("DELE", RemoteFileName, 250)
End Function

Private Function CheckConnected(operation As String) As Boolean
If Not Connected Then
 SetError 500, "Could not " & operation & " file. Not Connected."
 CheckConnected = False
Else
 CheckConnected = True
End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,60
Public Property Get TimeOut() As Integer
Attribute TimeOut.VB_Description = "Time out period (in seconds) for the control to time out. Default is 60 seconds."
Attribute TimeOut.VB_ProcData.VB_Invoke_Property = "FTPx_Properties"
    TimeOut = m_TimeOut
End Property

Public Property Let TimeOut(ByVal New_TimeOut As Integer)
    m_TimeOut = New_TimeOut
    PropertyChanged "TimeOut"
End Property

Private Function SendCommand(strCommand As String, strParameters As String, intCodeToWaitFor As Integer) As Boolean
SendCommand = False
TimePeriod = TimeOut
TimedOut = False
Timer1.Enabled = True
LastServerCode = 0
If strParameters = "" Then
 Winsock(0).SendData strCommand & vbCrLf
Else
 Winsock(0).SendData strCommand & " " & strParameters & vbCrLf
End If

While LastServerCode = 0 And Not TimedOut
 DoEvents
Wend
If TimedOut Then
 SetError 500, "Timed out on " & strCommand & " command."
 Exit Function
End If
If intCodeToWaitFor <> -1 Then
 If LastServerCode <> intCodeToWaitFor Then
  SetError LastServerCode, strCommand & " command failed."
  Exit Function
 End If
End If
SendCommand = True
RaiseEvent CommandSentToServer(strCommand & " " & strParameters)
End Function

Private Sub Winsock_SendComplete(Index As Integer)
Static LastBlockSent As Boolean
Dim FileNum As Integer
Dim Buffer As String * 8196
If InStr(1, Winsock(Index).Tag, "File Sender") <> 0 Then
 If BlockCount <= TotalBlocks Then
   FileNum = Val(Winsock(Index).Tag)
   Get #FileNum, RecPosition, Buffer
   Winsock(Index).SendData Left(Buffer, BlockSize)
   RecPosition = RecPosition + BlockSize
   RaiseEvent SendProgress(RecPosition, FileSendingSize)
   BlockCount = BlockCount + 1
   LastBlockSent = False
   Exit Sub
 Else
   If BlockSize <> Val(Mid(Winsock(Index).Tag, InStr(1, Winsock(Index).Tag, "[") + 1)) And Not LastBlockSent Then
     FileNum = Val(Winsock(Index).Tag)
     Get #FileNum, RecPosition, Buffer
     Winsock(Index).SendData Left(Buffer, RemainSize)
     RecPosition = RecPosition + BlockSize
     Close #FileNum
     LastBlockSent = True
     RaiseEvent SendProgress(FileSendingSize, FileSendingSize)
     Exit Sub
   End If
 End If
 Winsock(Index).Close
 SendComplete = True
End If
End Sub

Private Function RemoteFileSize(RemoteFileName As String) As Long
'assumes already connected
'assumes working directory has already been set
'uses Winsock_DataArrival Event
'Tags the winsock with Filenum and Lister
'waits for the server to terminate the conn to indicate
' the listing transfer is complete
'deletes the tempory file used to retreive the listing
'returns 0 if file not found
'returns filesize of RemoteFileName

Dim FileNum As Integer: FileNum = FreeFile()
Dim strData As String
Dim Filname As String: Dim tmpRemoteFileSize As Long
Dim x As Integer: Dim Marker1 As Integer: Dim Marker2 As Integer
RemoteFileSize = 0

'1st try the SIZE command, if not implemented then do it the hard way
SendCommand "TYPE", "I", 200
LastResponse = ""
Winsock(0).SendData "SIZE " & RemoteFileName & vbCrLf
While LastResponse = ""
 DoEvents
Wend
If Mid(LastResponse, 1, 3) = "550" Then 'the command was successful but No Such File
 Exit Function
End If
If Mid(LastResponse, 1, 3) = "213" Then 'the command was successfule
 RemoteFileSize = Val(Mid(LastResponse, 4))
 Exit Function
End If

'do it the hard way by getting a listing and searching it
SendCommand "TYPE", "A", 200
Socks = Socks + 1
Load Winsock(Socks)
Winsock(Socks).Tag = Str(FileNum) & " Lister"
Winsock(Socks).Listen
Open App.Path & "\dirxca.dat" For Output As #FileNum
DirectoryReceived = False
strData = Winsock(0).LocalIP
strData = Replace(strData, ".", ",", 1, -1) & ","
strData = strData & Str(Fix(Winsock(Socks).LocalPort / 256)) & ","
strData = strData & Str(Winsock(Socks).LocalPort - (Fix(Winsock(Socks).LocalPort / 256) * 256))
strData = Replace(strData, " ", "", 1, -1)
If Not SendCommand("PORT", strData, 200) Then Exit Function
SendCommand "LIST", "", 150
While Winsock(Socks).State = 7
 DoEvents
Wend
Close #FileNum
Winsock(Socks).Close
Unload Winsock(Socks)
Socks = Socks - 1

Open App.Path & "\dirxca.dat" For Input As #FileNum
While Not EOF(FileNum)
 Input #FileNum, strData
 Marker1 = 1: Marker2 = 0
 While Marker2 <> 4
  While Mid(strData, Marker1, 1) <> " "
   Marker1 = Marker1 + 1
  Wend
  While Mid(strData, Marker1, 1) = " "
   Marker1 = Marker1 + 1
  Wend
  Marker2 = Marker2 + 1
 Wend
 tmpRemoteFileSize = Val(Mid(strData, Marker1))
 Marker1 = Len(strData): Marker2 = Marker1
 While InStr(Marker1, strData, " ") = 0
  Marker1 = Marker1 - 1
 Wend
 FileName = Mid(strData, Marker1 + 1)
 If FileName = RemoteFileName Then
  RemoteFileSize = tmpRemoteFileSize
 End If
 x = 0
Wend
Close #FileNum
Kill App.Path & "\dirxca.dat"
End Function

Public Function Execute(CommandToSend As String, ByRef Response As String) As Boolean
'sends the command specified by CommandToSend to the server
'returns True if executed, False on failure
'Returns Response information returned by server
'assumes control is already connected
Execute = False

If Winsock(0).State <> 7 Then
 Response = "500 Not Connected!"
 Exit Function
End If

LastServerCode = 0
LastResponse = ""
TimedOut = False
Timer1.Enabled = True

Winsock(0).SendData CommandToSend & vbCrLf
While LastResponse = "" And Not TimedOut
 DoEvents
Wend
Timer1.Enabled = False
If TimedOut Then
 Response = "500 Server Timed Out."
 Exit Function
End If
Execute = True
Response = LastResponse
End Function


Public Function GetRemoteFileSize(RemoteDirectory As String, RemoteFileName As String) As Long
'assumes already connected
'sets the working directory to RemoteDirectory and when
'complete returns it to the root
'uses Winsock_DataArrival Event
'Tags the winsock with Filenum and Lister
'waits for the server to terminate the conn to indicate
' the listing transfer is complete
'deletes the tempory file used to retreive the listing
'returns 0 if file not found
'returns filesize of RemoteFileName

Dim FileNum As Integer: FileNum = FreeFile()
Dim strData As String

GetRemoteFileSize = 0
If RemoteDirectory <> "" Then
 SendCommand "CWD", RemoteDirectory, 250
End If

GetRemoteFileSize = 0

'1st try the SIZE command, if not implemented then do it the hard way
SendCommand "TYPE", "I", 200
LastResponse = ""
Winsock(0).SendData "SIZE " & RemoteFileName & vbCrLf
While LastResponse = ""
 DoEvents
Wend
If Mid(LastResponse, 1, 3) = "550" Then 'the command was successful but No Such File
 Exit Function
End If
If Mid(LastResponse, 1, 3) = "213" Then 'the command was successfule
 GetRemoteFileSize = Val(Mid(LastResponse, 4))
 Exit Function
End If

'do it the hard way by getting a listing and searching it



SendCommand "TYPE", "A", 200
Socks = Socks + 1
Load Winsock(Socks)
Winsock(Socks).Tag = Str(FileNum) & " Lister"
Winsock(Socks).Listen
Open App.Path & "\dirxca.dat" For Output As #FileNum
DirectoryReceived = False
strData = Winsock(0).LocalIP
strData = Replace(strData, ".", ",", 1, -1) & ","
strData = strData & Str(Fix(Winsock(Socks).LocalPort / 256)) & ","
strData = strData & Str(Winsock(Socks).LocalPort - (Fix(Winsock(Socks).LocalPort / 256) * 256))
strData = Replace(strData, " ", "", 1, -1)
If Not SendCommand("PORT", strData, 200) Then GoTo PutBack
SendCommand "LIST", "", 150
While Winsock(Socks).State = 7
 DoEvents
Wend
Close #FileNum
Winsock(Socks).Close
Unload Winsock(Socks)
Socks = Socks - 1
Dim x As Integer: Dim Marker1 As Integer: Dim Marker2 As Integer
Open App.Path & "\dirxca.dat" For Input As #FileNum
While Not EOF(FileNum)
 Input #FileNum, strData
 Marker1 = InStr(1, strData, " ")
 While x < 7
  While Mid(strData, Marker1, 1) = " "
   Marker1 = Marker1 + 1
  Wend
  While Mid(strData, Marker1, 1) <> " "
   Marker1 = Marker1 + 1
  Wend
  x = x + 1
  If x = 3 Then Marker2 = Marker1 ' the size pointer
 Wend
 If Mid(strData, Marker1 + 1) = RemoteFileName Then
  GetRemoteFileSize = Val(Mid(strData, Marker2))
 End If
 x = 0
Wend
Close #FileNum
Kill App.Path & "\dirxca.dat"
PutBack:
If RemoteDirectory <> "" Then
 SendCommand "CDUP", RemoteDirectory, 250
End If
End Function


Public Sub Cancel()
If Socks > 0 Then
 Winsock(1).Close
 Unload Winsock(1)
 Socks = Socks - 1
 Cancel_Operation = True
End If
End Sub

Public Function DirListing(DirFileName As String) As Boolean
'assumes already connected
'uses Winsock_DataArrival Event
'Tags the winsock with Filenum and Lister
'waits for the server to terminate the conn to indicate
' the listing transfer is complete
'returns TRUE if sucessful


Dim FileNum As Integer: FileNum = FreeFile()
Dim strData As String
DirListing = False


SendCommand "TYPE", "A", 200
Socks = Socks + 1
Load Winsock(Socks)
Winsock(Socks).Tag = Str(FileNum) & " Lister"
Winsock(Socks).Listen
On Error GoTo CloseFile
TryAgain:
Open DirFileName For Output As #FileNum
DirectoryReceived = False
strData = Winsock(0).LocalIP
strData = Replace(strData, ".", ",", 1, -1) & ","
strData = strData & Str(Fix(Winsock(Socks).LocalPort / 256)) & ","
strData = strData & Str(Winsock(Socks).LocalPort - (Fix(Winsock(Socks).LocalPort / 256) * 256))
strData = Replace(strData, " ", "", 1, -1)
Debug.Print "Port info:" & strData
If Not SendCommand("PORT", strData, 200) Then
 Exit Function
End If

SendCommand "LIST", "", 150
While Winsock(Socks).State <> 8
 DoEvents
Wend
Close #FileNum
Winsock(Socks).Close
Unload Winsock(Socks)
Socks = Socks - 1
DirListing = True
Exit Function
CloseFile:
Close #FileNum
GoTo TryAgain
End Function

Public Function Rename(RemoteDirectory As String, OldFileName As String, NewFileName As String) As Boolean

Rename = False
If RemoteDirectory <> "" Then
 If Not SendCommand("CWD", RemoteDirectory, 250) Then Exit Function
End If

If Not SendCommand("RNFR", OldFileName, 350) Then Exit Function
Rename = SendCommand("RNTO", NewFileName, 250)
End Function
