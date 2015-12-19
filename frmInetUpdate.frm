VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInetUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Keps Update"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmInetUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbFiles 
      Height          =   315
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1650
      Width           =   1845
   End
   Begin VB.FileListBox lstFiles 
      Height          =   480
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4620
      TabIndex        =   5
      Top             =   1590
      Width           =   765
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Height          =   375
      Left            =   2333
      TabIndex        =   4
      Top             =   1590
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1590
      Width           =   765
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   146
      TabIndex        =   0
      Top             =   1200
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblFileDetails 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   139
      TabIndex        =   2
      Top             =   780
      Width           =   5145
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Waiting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   139
      TabIndex        =   1
      Top             =   60
      Width           =   5145
   End
End
Attribute VB_Name = "frmInetUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim bAbort As Boolean
'Dim strFile As String
'Dim strName As String
'Dim BeginTransfer As Single
'Dim strConnectType As String
'
'Private Type RASCONN
'  dwSize As Long
'  hRasConn As Long
'  szEntryName(256) As Byte
'  szDeviceType(16) As Byte
'  szDeviceName(128) As Byte
'End Type
'
'Private Declare Function RasEnumConnectionsA& Lib "RasApi32.DLL" (lprasconn As Any, lpcb&, lpcConnections&)
'
'Private WithEvents mFTP As cFTP
'Private cn As CNetConnect
'
'Private Sub cmdAbort_Click()
'  bAbort = True
'End Sub
'
'Private Sub cmdClose_Click()
'  Unload Me
'End Sub
'
'Private Sub cmdStart_Click()
'Dim nResult As Integer
'
''Set cn = New CNetConnect
''strConnectType = CheckInetConnected
''If strConnectType <> "" Then
''  nResult = MsgBox("You are not connected to the Internet. Please establish a connection and try again", vbInformation + vbAbortRetryIgnore, "Internet Error")
''  Me.lblStatus = "Waiting..."
''Else
''  nResult = vbIgnore
''End If
'nResult = MsgBox("Please ensure that your are connected to the Internet. If you are connected then select ok.", vbInformation + vbOKCancel + vbDefaultButton1, "Internet Connection")
'If nResult = vbOK Then
'  Me.cmdAbort.Visible = True
'  Me.cmdClose.Enabled = False
'  Me.cmdStart.Enabled = False
'  InetKepsUpdate
'End If
'
''Set cn = Nothing
'End Sub
'
'Public Function InetKepsUpdate()
'  On Error GoTo ERROR_InetKepsUpdate
'
'  Dim nFile As Integer
'  Dim nfile1 As Integer
'  Dim nfile2 As Integer
'  Dim vData() As Variant
'  Dim strLine As String
'  Dim strFTPServer As String
'  Dim strLastFTPServer As String
'  Dim strdir As String
'  Dim bOk As Boolean
'  Dim strTemp As String
'
'  Set mFTP = New cFTP
'  mFTP.SetModeActive
'  mFTP.SetTransferBinary
'
'  bAbort = False
'  nFile = FreeFile
'  Open App.Path & "\Internet Updates\" & Me.cmbFiles.Text For Input As #nFile
'
'  Do While Not (EOF(nFile))
'    Line Input #nFile, strLine
'    If Trim(strLine) <> "" Then
'      vData = StrParse(strLine, ",")
'      strName = vData(0)
'      strFile = vData(1)
'      strFTPServer = vData(2)
'      strdir = vData(3)
'      bOk = True
'
'      Me.lblStatus.Caption = "Connecting to FTP server (" & strConnectType & ")"
'      Me.lblStatus.Refresh
'      '    If strFTPServer <> strLastFTPServer Then
'      '      If strLastFTPServer <> "" Then
'      '        FTP.Disconnect
'      '      End If
'      If mFTP.OpenConnection(strFTPServer, "anonymous", "webmaster@hamsoftware.co.uk") Then
'        Me.lblStatus.Caption = "Connected to FTP server (" & strConnectType & ")"
'        strLastFTPServer = strFTPServer
'      Else
'        MsgBox mFTP.GetLastErrorMessage
'        bOk = False
'        bAbort = True
'      End If
'      '   End If
'
'      If bOk Then
'        Me.lblFileDetails = strName
'        Me.lblFileDetails.Refresh
'        BeginTransfer = Timer
'        GetFile strdir & "/" & strFile, App.Path & "\temp.dat"
'        nfile1 = FreeFile
'        Open App.Path & "\temp.dat" For Input As #nfile1
'        Line Input #nfile1, strTemp
'        strTemp = Replace(strTemp, Chr(10), vbCrLf)
'        Close #nfile1
'        Open App.Path & "\Elements\" & strFile For Output As #nfile1
'        Print #nfile1, strTemp
'        Close #nfile1
'        Me.PB.Value = 0
'      End If
'
'      If bAbort Then
'        Exit Do
'      End If
'      mFTP.CloseConnection
'      DoEvents
'    End If
'  Loop
'  On Error Resume Next
'  Kill App.Path & "\temp.dat"
'  On Error GoTo 0
'
'  If bAbort Then
'    Me.lblStatus = "Aborted"
'  Else
'    Me.lblStatus = "Complete Waiting..."
'  End If
'
'EXIT_InetKepsUpdate:
'  Me.lblFileDetails = ""
'  Me.cmdAbort.Visible = False
'  Me.cmdClose.Enabled = True
'  Me.cmdStart.Enabled = True
'
'  Close #nFile
'  mFTP.CloseConnection
'  Set mFTP = Nothing
'  Exit Function
'
'ERROR_InetKepsUpdate:
'  MsgBox "Error in ERROR_InetKepsUpdate : " & Error
'  Resume EXIT_InetKepsUpdate
'
'End Function
'
'Private Function GetFile(strSource As String, strDest As String) As Boolean
'
'  If Not mFTP.FTPDownloadFile(strDest, strSource) Then
'    MsgBox mFTP.GetLastErrorMessage
'    GetFile = False
'  Else
'    GetFile = True
'  End If
'End Function
'
'Private Sub Form_Load()
'  Dim i As Integer
'
'  CenterForm Me
'  lstFiles.Path = App.Path & "\Internet Updates"
'  lstFiles.Pattern = "*.dat"
'  Me.cmbFiles.Clear
'  For i = 0 To lstFiles.ListCount - 1
'    Me.cmbFiles.AddItem Me.lstFiles.List(i)
'  Next i
'  Me.cmbFiles.ListIndex = 0
'End Sub
'
'Private Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
'  On Error Resume Next
'  Dim strTemp As String
'  Dim TransferRate As Single
'
'  TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1000, "####.00")
'  PB.Max = lTotalBytes
'  PB.Min = 0
'  PB.Value = lCurrentBytes
'  PB.ToolTipText = PB.Value & " Bytes of " & PB.Max & " Bytes Transfered"
'  PB.Refresh
'
'  strTemp = strName & " "
'  strTemp = strTemp & PB.Value \ 1024 & " KB of " & PB.Max \ 1024 & " KB "
'  strTemp = strTemp & Format(TransferRate, "##.#0#") & " Kbps "
'  strTemp = strTemp & "Time Left: " & ConvertTime(Int(((PB.Max - PB.Value) / 1024) / TransferRate))
'  Me.lblFileDetails = strTemp
'
'End Sub
'
'Public Function ConvertTime(ByVal TheTime As Single) As String
'  Dim NewTime                         As String
'  Dim Sec                             As Single
'  Dim Min                             As Single
'  Dim H                               As Single
'  If TheTime > 60 Then
'    Sec = TheTime
'    Min = Sec / 60
'    Min = Int(Min)
'    Sec = Sec - Min * 60
'    H = Int(Min / 60)
'    Min = Min - H * 60
'    NewTime = H & ":" & Min & ":" & Sec
'    If H < 0 Then H = 0
'    If Min < 0 Then Min = 0
'    If Sec < 0 Then Sec = 0
'    NewTime = Format(NewTime, "HH:MM:SS")
'    ConvertTime = NewTime
'  End If
'  If TheTime < 60 Then
'    NewTime = "00:00:" & TheTime
'    NewTime = Format(NewTime, "HH:MM:SS")
'    ConvertTime = NewTime
'  End If
'End Function
'
'Private Function CheckInetConnected() As String
'
'  Me.lblStatus = "Checking Internet Connection"
'  Me.lblStatus.Refresh
'  If cn.Connected Then
'    CheckInetConnected = cn.ConnectModeDesc
'  Else
'    CheckInetConnected = ""
'  End If
'
'End Function
Private Sub Form_Load()

End Sub
