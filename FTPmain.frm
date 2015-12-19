VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFTPMain 
   HelpContextID   =  1400
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Keps Update"
   ClientHeight    =   2820
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "FTPmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Finishe the element update"
      Top             =   2220
      Width           =   915
   End
   Begin VB.FileListBox lstFiles 
      Height          =   480
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ComboBox cmbFiles 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "Select the element group to update"
      Top             =   2220
      Width           =   1605
   End
   Begin MSWinsockLib.Winsock sckDownload 
      Left            =   3060
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Timer tmrUpdateProgress 
      Interval        =   1
      Left            =   3180
      Top             =   1800
   End
   Begin VB.Timer tmrTimeLeft 
      Interval        =   1000
      Left            =   3120
      Top             =   1860
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Starts the element update"
      Top             =   2220
      Width           =   915
   End
   Begin VB.Frame fraDownloadProgress 
      Caption         =   " File Update Progress "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5775
      Begin VB.PictureBox picDownloadProgress 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5475
         TabIndex        =   20
         Top             =   900
         Width           =   5535
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtURL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Tag             =   "http://tucows.erols.com/files4/bzfinst.exe"
         Text            =   "http://www.celestrak.com/NORAD/elements/amateur.txt"
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   900
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lblSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblRecieve 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblElapsed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblRemaining 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   8
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Time Remaining:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   6
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recieved Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   0
      ToolTipText     =   "Stops the current element update"
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "Pauses the current element update"
      Top             =   2220
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmFTPMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sDATA                         As String
Private Percent                         As Integer
Private BeginTransfer                   As Single

Private Header                          As Variant
Private Status                          As String
Private TransferRate                    As Single

Private bFTPThroughProxy                As Boolean
Private WithEvents CFtpConnection       As CFtpConnection
Attribute CFtpConnection.VB_VarHelpID = -1
Private bFTPDownload                    As Boolean
Private bDownloadPaused                 As Boolean
Private bDownloadComplete               As Boolean



Public Function GETDATAHEAD(DATA As Variant, ToRetrieve As String)
    Dim EndBYTES                        As Integer
    Dim A                               As String
    Dim LENGTHEND                       As Integer
    Dim PART                            As Integer
    Dim Part2                           As Integer
    Dim RetrieveLength                  As Integer
    On Error Resume Next
    If DATA = "" Then Exit Function
    If InStr(DATA, ToRetrieve) > 0 Then
        LENGTHEND = Len(DATA)
        PART = InStr(DATA, ToRetrieve)
        RetrieveLength = Len(ToRetrieve)
        A = Right(DATA, LENGTHEND - PART - RetrieveLength)
        LENGTHEND = Len(A)
        If InStr(A, vbCrLf) > 0 Then
            Part2 = InStr(A, vbCrLf)
            A = Left(A, Part2 - 1)
        End If
        GETDATAHEAD = A
    End If
End Function

Public Function OutFileName(file$) As String
    Dim P                               As Integer
    P = InStr(file$, ".") 'Check for the period in the file
    If P = 0 Then
        OutFileName = file & "ext" & ".rsm" 'If no period then add a period and extension to it
        Exit Function
    End If
    If LCase(Right(file$, 3) = "rsm") Then 'Check to see if its extension is the resuming file extension used by downloader
        Dim Length                      As Integer
        Dim A                           As String
        Dim B                           As String
        P = InStr(file$, ".")
        A = Left(file$, P - 1) 'Trimming off the filename without added extension
        B = Right(A, 3) 'Getting extension of original filename
        Length = Len(A$)
        A = Left(A, Length - 3) 'get rid of the original extension
        OutFileName = A & "." & B 'add original extension back on with period
    Else 'if its not a resumable file then make it one!
        Dim Dot                         As Integer
        Dim One                         As String
        Dim Ext                         As String
        Dim SLength                     As Integer
        Dot = InStr(file$, ".") 'get position of period
        One = Left(file$, Dot - 1) 'Get the filename by itself
        Ext = Right(file$, 3) 'Get the extension by itself
        OutFileName = One & Ext & ".rsm" 'Put the rsm file extension onto the file!
    End If
End Function

Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

Public Function StartUpdate(ByVal strURL As String)
    Dim Pos                             As Integer
    Dim Length                          As Integer
    Dim NextPos                         As Integer
    Dim LENGTH2                         As Integer
    Dim POS2                            As Integer
    Dim POS3                            As Integer
    BytesAlreadySent = 1
    If strURL = "" Then
        Exit Function
    End If
    URL = strURL
    Pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    Length = Len(strURL) 'Length of the entire url
    If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, Length - LENGTH2 - Pos + 1) ' remove http:// or ftp://
    End If
    If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
        POS2 = InStr(strURL, "/") 'gets the position of the / mark
        '-----------------GET THE FILENAME-------------
        Dim strFile                     As String
        strFile = strURL 'load the variables into each other
        Do Until InStr(strFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(strFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(strFile, "/") 'find the / mark
            strFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
        Loop
        
            If InStr(strFile, ":") Then
                FileName = Left(strFile, InStr(strFile, ":") - 1)
            Else
                FileName = strFile
            End If
            
        '----------------END GET FILE NAME--------------
        If Not bProxy Then
            strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
        End If
    End If
    '-----------END TRIM THE URL FOR THE SERVER NAME-----------
End Function

Public Sub Reset()
    CloseSocket
    m_sDATA = ""
    Percent = 0
    BeginTransfer = 0
    BytesAlreadySent = 0
    BytesRemaining = 0
    Status = ""
    Header = ""
    RESUMEFILE = False
    UpdateProgress picDownloadProgress, 0
    cmdDownload.Enabled = True
    cmdPause.Enabled = False
    cmdStop.Enabled = False
End Sub

Public Sub CloseSocket()
    Do Until sckDownload.State = 0
        sckDownload.Close
        sckDownload.LocalPort = 0
        'Close #1
    Loop
End Sub

Private Sub CFtpConnection_DownloadProgress(lBytes As Long)

BytesAlreadySent = lBytes

 If RESUMEFILE = False Then
        'This is pretty straightforward if you ever taken math before you can tell what im doing!
        TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
    Else
        'If you dont subtract the difference you will get a really large and odd download speed hehe.
        TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
    End If

End Sub

Private Sub CFtpConnection_ReplyMessage(sMessage As String)
 '   frmHeader.txtHeader.SelText = sMessage
End Sub

Private Sub CFtpConnection_StateChanged(State As FTP_CONNECTION_STATES)
    Select Case State
        Case FTP_CONNECTION_RESOLVING_HOST
'            frmHeader.txtHeader.SelText = "FTP_CONNECTION_RESOLVING_HOST" & vbNewLine
        Case FTP_CONNECTION_HOST_RESOLVED
'            frmHeader.txtHeader.SelText = "FTP_CONNECTION_HOST_RESOLVED" & vbNewLine
        Case FTP_CONNECTION_CONNECTED
'            frmHeader.txtHeader.SelText = "FTP_CONNECTION_CONNECTED" & vbNewLine
        Case FTP_CONNECTION_AUTHENTICATION
'            frmHeader.txtHeader.SelText = "FTP_CONNECTION_AUTHENTICATION" & vbNewLine
        Case FTP_USER_LOGGED
'            frmHeader.txtHeader.SelText = "FTP_USER_LOGGED" & vbNewLine
        Case FTP_ESTABLISHING_DATA_CONNECTION
'            frmHeader.txtHeader.SelText = "FTP_ESTABLISHING_DATA_CONNECTION" & vbNewLine
        Case FTP_DATA_CONNECTION_ESTABLISHED
'            frmHeader.txtHeader.SelText = "FTP_DATA_CONNECTION_ESTABLISHED" & vbNewLine
        Case FTP_RETRIEVING_DIRECTORY_INFO
'            frmHeader.txtHeader.SelText = "FTP_RETRIEVING_DIRECTORY_INFO" & vbNewLine
        Case FTP_DIRECTORY_INFO_COMPLETED
'            frmHeader.txtHeader.SelText = "FTP_DIRECTORY_INFO_COMPLETED" & vbNewLine
        Case FTP_TRANSFER_STARTING
'            frmHeader.txtHeader.SelText = "FTP_TRANSFER_STARTING" & vbNewLine
        Case FTP_TRANSFER_COMLETED
'            frmHeader.txtHeader.SelText = "FTP_TRANSFER_COMLETED" & vbNewLine
            If Not bDownloadPaused Then
                bDownloadComplete = True
            End If
    End Select
End Sub

Private Sub CFtpConnection_UploadProgress(lBytes As Long)
    Stop
End Sub

Private Sub cmdDownload_Click()
Dim nResult As Integer

'nResult = MsgBox("Please ensure that your are connected to the Internet. If you are connected then select ok.", vbInformation + vbOKCancel + vbDefaultButton1, "Internet Connection")
'If nResult = vbOK Then
  Me.cmdStop.Enabled = True
  Me.cmdDownload.Enabled = False
  Me.cmdExit.Enabled = False
  bDownloadComplete = False
  InetKepsUpdate
  lblStatus.Caption = "Update Completed"

'End If

End Sub
Public Function InetKepsUpdate()
  'On Error GoTo ERROR_InetKepsUpdate

  Dim nFile As Integer
  Dim nfile1 As Integer
  Dim nfile2 As Integer
  Dim vData() As Variant
  Dim strLine As String
  Dim strTemp As String
  Dim strName As String
  Dim strFile As String
  Dim strFTPServer As String
  Dim strDIR As String
  Dim nCount As Integer
  
  nFile = FreeFile
  Open App.Path & "\Internet Updates\" & Me.cmbFiles.Text For Input As #nFile
  Do While Not (EOF(nFile))
    Line Input #nFile, strLine
    nCount = nCount + 1
  Loop
  Close #nFile
  
  Me.ProgressBar.Max = nCount
  Me.ProgressBar.Value = 0
  Open App.Path & "\Internet Updates\" & Me.cmbFiles.Text For Input As #nFile
  Do While Not (EOF(nFile))
    Line Input #nFile, strLine
    If Trim(strLine) <> "" Then
      vData = StrParse(strLine, ",")
      strName = vData(0)
      strFile = vData(1)
      strFTPServer = vData(2)
      strDIR = vData(3)
      GetFile strFTPServer & strDIR & "/" & strFile, App.Path & "\Elements\" & strFile
      Me.ProgressBar.Value = Me.ProgressBar.Value + 1
    End If
  Loop


EXIT_InetKepsUpdate:
  Me.cmdStop.Enabled = False
  Me.cmdDownload.Enabled = True
  Me.cmdExit.Enabled = True

  Close #nFile
  Exit Function

ERROR_InetKepsUpdate:
  MsgBox "Error in ERROR_InetKepsUpdate : " & Error
  Resume EXIT_InetKepsUpdate

End Function

Private Function GetFile(strSourceFilename As String, strDestFilename As String)
    'Are we useing a proxy
    bProxy = sProgramOptions.bFTPProxy
    
    txtURL.Text = strSourceFilename
    
    If bProxy Then
        'Yes
        strSvrURL = sProgramOptions.strFTPProxyURL
        strSvrPort = sProgramOptions.nFTPPPort
        bFTPThroughProxy = sProgramOptions.bFTPThruProxy
    Else
        'No
        strSvrURL = txtURL
        strSvrPort = 80
        bFTPThroughProxy = False
    End If
    
    StartUpdate txtURL
    
    FilePathName = strDestFilename
    bDownloadComplete = False
    StartDownload FilePathName
    While Not bDownloadComplete
      DoEvents
    Wend
    Reset
    lblStatus.Visible = True
    lblStatus.Caption = "Element Update Complete"
    picDownloadProgress.Visible = False
    CloseSocket
End Function

Private Sub cmdExit_Click()
  bDownloadComplete = True
  Unload Me
End Sub

Private Sub cmdPause_Click()
    cmdPause.Enabled = True
    cmdDownload.Enabled = False
    
    If BytesRemaining > BytesAlreadySent Then
        cmdStop.Enabled = False
        If cmdPause.Caption = "&Pause" Then
            cmdPause.Caption = "&Resume"
            bDownloadPaused = True
            tmrTimeLeft.Enabled = False
            
            If bFTPDownload Then
                picDownloadProgress.Visible = False
                lblStatus.Visible = True
                lblStatus.Caption = "Download Paused"
                CFtpConnection.CancelTransfer
            ElseIf sckDownload.State > 0 Then
                m_sDATA = ""
                BeginTransfer = 0
                Status = ""
                Header = ""
                CloseSocket
                picDownloadProgress.Visible = False
                lblStatus.Visible = True
                lblStatus.Caption = "Download Paused"
            End If
        Else
            cmdStop.Enabled = True
            cmdPause.Caption = "&Pause"
            bDownloadPaused = False
            tmrTimeLeft.Enabled = True
            If bFTPDownload Then
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                FileLength = FileLen(FilePathName)
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                RESUMEFILE = True
                StartFTPDownload
            ElseIf sckDownload.State < 0 Then
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                FileLength = FileLen(FilePathName)
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                RESUMEFILE = True
                sckDownload.Connect strSvrURL, strSvrPort
            End If
        End If
    End If
End Sub

Private Sub cmdStop_Click()
    If bFTPDownload Then
        bDownloadPaused = True
        If Not CFtpConnection Is Nothing Then
            picDownloadProgress.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Aborted"
            CFtpConnection.BreakeConnection
            Reset
        End If
    ElseIf sckDownload.State > 0 Then
        picDownloadProgress.Visible = False
        lblStatus.Visible = True
        lblStatus.Caption = "Download Aborted"
        CloseSocket
        Reset
    End If
    Me.cmdExit.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Integer

  RESUMEFILE = False


  UpdateProgress picDownloadProgress, 0

  CenterForm Me
  lstFiles.Path = App.Path & "\Internet Updates"
  lstFiles.Pattern = "*.dat"
  Me.cmbFiles.Clear
  For i = 0 To lstFiles.ListCount - 1
    Me.cmbFiles.AddItem Me.lstFiles.List(i)
  Next i
  Me.cmbFiles.ListIndex = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseSocket
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseSocket
End Sub

Private Sub tmrTimeLeft_Timer()
    'On Error Resume Next
    If BytesRemaining > 0 And BytesAlreadySent > 0 And TransferRate > 0 Then
        If BytesRemaining <= BytesAlreadySent Then
            lblSpeed = 0
            CloseSocket
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            Reset
            cmdDownload.Enabled = False
            picDownloadProgress.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Completed"
            bDownloadComplete = True
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
                Sec = 0
                Min = Min + 1
            ElseIf Min >= 60 Then
                Min = 0
                Hr = Hr + 1
            End If
            'cmdDownload.Enabled = True
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
            lblRemaining = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
            lblSpeed = Format(TransferRate, "##.#0#") & " Kbps"

        End If
    End If
End Sub

Private Sub tmrUpdateProgress_Timer()
'    On Error Resume Next
    If BytesAlreadySent > 0 Then 'And BytesRemaining > 0 Then

        lblRecieve = File_ByteConversion(BytesAlreadySent)
        If BytesRemaining = 0 Then
            lblSize = "Unknown"
        Else
            lblSize = File_ByteConversion(BytesRemaining)
        End If
            If lblSize <> "Unknown" Then
            Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
            UpdateProgress picDownloadProgress, Percent 'updates progress bar with new percentage rate
        End If
    End If
End Sub

Private Sub sckDownload_Close()
    FormsOnTop Me, False
    picDownloadProgress.Visible = False
    lblStatus.Visible = True
    lblStatus.Caption = "Download Completed"
    sckDownload.Close
End Sub

Private Sub sckDownload_Connect()
     On Error Resume Next
    Dim strCommand                      As String
    If Mid$(URL, 1, 6) = "ftp://" Then
        If InStr(7, URL, "@") <> 0 Then
            If InStr(InStr(7, URL, "@"), URL, ":") Then
                URL = Mid$(URL, 1, InStr(InStr(7, URL, "@"), URL, ":") - 1)
                Stop
            End If
        ElseIf InStr(7, URL, ":") <> 0 Then
            URL = Mid$(URL, 1, InStr(7, URL, ":") - 1)
        End If
    End If
    
    
    strCommand = "GET " + Right(URL, Len(URL) - Len(strSvrURL) - 7) + " HTTP/1.0" + vbCrLf
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    
    If RESUMEFILE = True Then
        strCommand = strCommand + "Range: bytes=" & FileLength & "-" & vbCrLf
    End If
    
    strCommand = strCommand + "User-Agent: Elucid Software Downloader" & vbCrLf
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
    
    strCommand = strCommand + vbCrLf
    sckDownload.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate
End Sub

Private Sub sckDownload_DataArrival(ByVal bytesTotal As Long)
    Dim Pos                             As Integer
    Dim Length                          As Integer
    Dim HEAD                            As String
    Dim nFile As Integer
    
    sckDownload.GetData m_sDATA, vbString
    
    If InStr(LCase(m_sDATA), "content-type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
        If RESUMEFILE = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
            If InStr(LCase(m_sDATA), "206 partial content") = 0 Then
                MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
                Reset
                CloseSocket
                Exit Sub
                End If
        End If
    
    If InStr(LCase(m_sDATA), "404 not found") > 0 Then
            MsgBox "The file requested was not found on the server!" & vbCrLf & vbCrLf & "Possible Reasons:" & vbCrLf & "- File Does Not Exist On Server" _
            & vbCrLf & "- URL Given Was Script And Data Returned Was Invalid" & vbCrLf & "- URL Entered Was Incorrect" & vbCrLf & "- Server Is Excessively Busy" _
            & vbCrLf & vbCrLf & "You may reattempt to download.  If its still failure then most likely invalid url.", , "File Not Found"
            Reset
            CloseSocket
            Exit Sub
   End If
   
        Pos = InStr(m_sDATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
        Length = Len(m_sDATA) 'get the length of the data chunk
        HEAD = Left(m_sDATA, Pos - 1) 'Get the header from the chunk of data and ignore the data content
        m_sDATA = Right(m_sDATA, Length - Pos - 3) 'Get the data from the first chunk that contains the header also
        Header = Header & HEAD 'Append the header to header text box
        
        If RESUMEFILE = True Then
            BytesAlreadySent = FileLength + 1
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
            BytesRemaining = BytesRemaining + FileLength
        Else
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
        End If
        
       ' frmHeader.txtHeader = Header
    End If
    '-----------BEGIN WRITE CHUNK TO FILE CODE--------
    nFile = FreeFile
    Open FilePathName For Binary Access Write As #nFile 'opens file for output
    Put #nFile, BytesAlreadySent, m_sDATA 'writes data to the end of file
    BytesAlreadySent = Seek(nFile)
    Close #nFile 'close file for now until next data chunk is available
    '--------------------------------------------------
    
    If RESUMEFILE = False Then
        'This is pretty straightforward if you ever taken math before you can tell what im doing!
        TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
    Else
        'If you dont subtract the difference you will get a really large and odd download speed hehe.
        TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
    End If
End Sub

Public Sub StartDownload(ByVal sTargetFile As String)
    Dim bRollback                       As Boolean
    Dim intRollback                     As Integer
    
    cmdPause.Enabled = True
    cmdStop.Enabled = True
    cmdDownload.Enabled = False

    bRollback = sProgramOptions.bFTPRollback
    intRollback = sProgramOptions.nFTPRollback

    FilePathName = sTargetFile
    bFTPDownload = False
    If Left$(LCase$(txtURL), 6) = "ftp://" Then
        If bFTPThroughProxy Then
            frmFTPMain.sckDownload.Connect strSvrURL, strSvrPort
        Else
            bFTPDownload = True
            StartFTPDownload
        End If
    ElseIf Left$(LCase$(txtURL), 7) = "http://" Then
        frmFTPMain.sckDownload.Connect strSvrURL, strSvrPort
    End If
End Sub
Private Sub StartFTPDownload()
    Dim sUsername                       As String
    Dim sPassword                       As String
    Dim sPort                           As String
    Dim sServer                         As String
    Dim sDirectory                      As String
    Dim sFile                           As String
    Dim sTemp                           As String
    Dim lStartAt                        As Long
    Dim lRet                            As Long
    Dim bSuccess                        As Boolean
    Dim intTimeout                      As Integer
    Dim bPasvMode                       As Boolean
    Set CFtpConnection = New CFtpConnection
    
    'URL = "ftp://10.1.1.10/Update/iqb00529.exe"
    'URL = "ftp://ftp:ftp@10.1.1.10/Update/iqb00529.exe"
    'URL = "ftp://ftp:ftp@10.1.1.10/Update/iqb00529.exe:21"
    'URL = "ftp://10.1.1.10/Update/iqb00529.exe:21"
    sTemp = URL
    sTemp = Mid(URL, 7)
    'Extract Server
    sServer = Mid$(sTemp, 1, InStr(1, sTemp, "/") - 1)
    If InStr(1, sServer, "@") <> 0 Then
        'Username / Password
        sUsername = Mid$(sServer, 1, InStr(1, sServer, ":") - 1)
        sServer = Mid$(sServer, Len(sUsername) + 2)
        sPassword = Mid$(sServer, 1, InStr(1, sServer, "@") - 1)
        sServer = Mid$(sServer, Len(sPassword) + 2)
    Else
        sUsername = "anonymous"
        sPassword = "downloader@thenet.com"
    End If
    
    If InStr(InStr(7, sTemp, "/"), sTemp, ":") <> 0 Then
        'FTP Port
        sPort = Mid$(sTemp, InStrRev(sTemp, ":") + 1)
    Else
        sPort = 21
    End If
    sDirectory = Mid(sTemp, InStr(7, sTemp, "/"))
    If InStr(InStr(7, sTemp, "/"), sTemp, ":") <> 0 Then
        sDirectory = Left$(sDirectory, Len(sDirectory) - (Len(sPort) + 1))
    End If
    sFile = Right(sDirectory, Len(sDirectory) - InStrRev(sDirectory, "/"))
    
    sDirectory = Left(sDirectory, Len(sDirectory) - (Len(sFile) + 1))
    If FileCheck(FilePathName) Then
        If RESUMEFILE Then
            lStartAt = FileLen(FilePathName)
'            FileLength = FileLen(FilePathName)
        Else
            Kill FilePathName
            lStartAt = 0
        End If
    End If
    
    intTimeout = sProgramOptions.nFTPTimeout
    bPasvMode = sProgramOptions.bFTPbPasvMode
    
    If intTimeout = 0 Then
      intTimeout = 30
    End If
    CFtpConnection.Timeout = intTimeout
    CFtpConnection.PassiveMode = bPasvMode
    
    CFtpConnection.UserName = sUsername
    CFtpConnection.Password = sPassword
  
    bSuccess = True
    Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
        If CFtpConnection.Connect(sServer, sPort) Then
            bSuccess = True
            Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
                If CFtpConnection.SetCurrentDirectory(sDirectory) Then
                    bSuccess = True
                    BeginTransfer = Timer
                    bDownloadComplete = False
                    Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
                        If CFtpConnection.DownloadFile(sFile, FilePathName, FTP_IMAGE_MODE, lStartAt) Then
                            bSuccess = True
                            bDownloadComplete = True
                        Else
                            If Mid$(CFtpConnection.GetLastServerResponse, 1, 3) = "504" Then
                                MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
                                Kill FilePathName
                                lStartAt = 0
                                bSuccess = True
                            ElseIf bDownloadPaused Then 'And _
                                (Mid$(CFTPConnection.GetLastServerResponse, 1, 3) = "426" Or _
                                Mid$(CFTPConnection.GetLastServerResponse, 1, 3) = "225") Then
                                '426 Transfger complete, 225 ABOR command received
                                'Ignore the error, the download should be canceld because we paused it
                            Else
                                lRet = MsgBox("Server returned the following error:" & vbNewLine & CFtpConnection.GetLastServerResponse & vbNewLine, vbRetryCancel)
                            End If
                        End If
                    Loop
                Else
                    lRet = MsgBox("Error occured while changing server directory to: " & vbNewLine & sDirectory, vbRetryCancel + vbCritical)
                End If
            Loop
        Else
            lRet = MsgBox("Error occured while conencting to server: " & _
                vbNewLine & sServer, vbRetryCancel + vbCritical)
        End If
    Loop
    If bDownloadComplete Then
        picDownloadProgress.Visible = False
        lblStatus.Visible = True
        lblStatus.Caption = "Download Completed"
    End If
    Set CFtpConnection = Nothing
End Sub

Private Sub txtURL_Change()
  txtURL = Trim(txtURL)
End Sub

