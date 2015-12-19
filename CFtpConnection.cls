VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class partially based on the FTPClient Sample Application from Oleg Gdalevich
'
Option Explicit
'Declarations of winsock objects.
'********************************************************************************
'Winsock Control for control connection
Private frmFTPContainer                 As frmFTPWinsockContainter
Private WithEvents wscControl           As MSWinsockLib.Winsock
Attribute wscControl.VB_VarHelpID = -1
'Winsock Control for data connection
Private WithEvents wscData              As MSWinsockLib.Winsock
Attribute wscData.VB_VarHelpID = -1
'********************************************************************************
'Local variables to hold values of the class properies
'********************************************************************************
Private m_strUserName                   As String
Private m_strPassword                   As String
Private m_varFtpServer                  As Variant
Private m_varFtpServerPort              As Variant
Private m_strCurrentDirectory           As String
Private m_bPassiveMode                  As Boolean
Private m_bBusy                         As Boolean
Private m_intTimeout                    As Integer
Private m_TransferMode                  As FtpTransferModes
'********************************************************************************
'Public Enums
'********************************************************************************
'various states of ftp connection
Public Enum FTP_CONNECTION_STATES
    FTP_CONNECTION_RESOLVING_HOST
    FTP_CONNECTION_HOST_RESOLVED
    FTP_CONNECTION_CONNECTED
    FTP_CONNECTION_AUTHENTICATION
    FTP_USER_LOGGED
    FTP_ESTABLISHING_DATA_CONNECTION
    FTP_DATA_CONNECTION_ESTABLISHED
    FTP_RETRIEVING_DIRECTORY_INFO
    FTP_DIRECTORY_INFO_COMPLETED
    FTP_TRANSFER_STARTING
    FTP_TRANSFER_COMLETED
End Enum
'all possible reply codes that can be sent by ftp server
Private Enum FTP_RESPONSE_CODES
    FTP_RESPONSE_RESTART_MARKER_REPLY = 110
    FTP_RESPONSE_SERVICE_READY_IN_MINUTES = 120
    FTP_RESPONSE_DATA_CONNECTION_ALREADY_OPEN = 125
    FTP_RESPONSE_FILE_STATUS_OK = 150
    FTP_RESPONSE_COMMAND_OK = 200
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_SUPERFLUOUS_AT_THIS_SITE = 202 'superfluous at this site
    FTP_RESPONSE_SYSTEM_STATUS_OR_SYSTEM_HELP_REPLY = 211
    FTP_RESPONSE_DIRECTORY_STATUS = 212
    FTP_RESPONSE_FILE_STATUS = 213
    FTP_RESPONSE_HELP_MESSAGE = 214
    FTP_RESPONSE_NAME_SYSTEM_TYPE = 215
    FTP_RESPONSE_SERVICE_READY_FOR_NEW_USER = 220
    FTP_RESPONSE_SERVICE_CLOSING_CONTROL_CONNECTION = 221
    FTP_RESPONSE_DATA_CONNECTION_OPEN = 225
    FTP_RESPONSE_CLOSING_DATA_CONNECTION = 226
    FTP_RESPONSE_ENTERING_PASSIVE_MODE = 227
    FTP_RESPONSE_USER_LOGGED_IN = 230
    FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED = 250
    FTP_RESPONSE_PATHNAME_CREATED = 257
    FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD = 331
    FTP_RESPONSE_NEED_ACCOUNT_FOR_LOGIN = 332
    FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO = 350
    FTP_RESPONSE_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION = 421
    FTP_RESPONSE_CANNOT_OPEN_DATA_CONNECTION = 425
    FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED = 426
    FTP_RESPONSE_REQUESTED_FILE_ACTION_NOT_TAKEN = 450
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED = 451
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN = 452
    FTP_RESPONSE_SYNTAX_ERROR_COMMAND_UNRECOGNIZED = 500
    FTP_RESPONSE_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS = 501
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED = 502
    FTP_RESPONSE_BAD_SEQUENCE_OF_COMMANDS = 503
    FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER = 504
    FTP_RESPONSE_NOT_LOGGED_IN = 530
    FTP_RESPONSE_NEED_ACCOUNT_FOR_STORING_FILES = 532
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE = 550
    FTP_RESPONSE_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN = 551
    FTP_RESPONSE_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION = 552
    FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED = 553
End Enum
'transfer modes
Public Enum FtpTransferModes
    FTP_ASCII_MODE
    FTP_IMAGE_MODE
End Enum
'********************************************************************************
'Class errors
'********************************************************************************
Public Enum FtpErrors
    ERROR_FTP_WINSOCK_AddressInUse
    ERROR_FTP_WINSOCK_AddressNotAvailable
    ERROR_FTP_WINSOCK_AlreadyComplete
    ERROR_FTP_WINSOCK_AlreadyConnected
    ERROR_FTP_WINSOCK_BadState
    ERROR_FTP_WINSOCK_ConnectAborted
    ERROR_FTP_WINSOCK_ConnectionRefused
    ERROR_FTP_WINSOCK_ConnectionReset
    ERROR_FTP_WINSOCK_GetNotSupported
    ERROR_FTP_WINSOCK_HostNotFound
    ERROR_FTP_WINSOCK_HostNotFoundTryAgain
    ERROR_FTP_WINSOCK_InProgress
    ERROR_FTP_WINSOCK_InvalidArg
    ERROR_FTP_WINSOCK_InvalidArgument
    ERROR_FTP_WINSOCK_InvalidOp
    ERROR_FTP_WINSOCK_InvalidPropertyValue
    ERROR_FTP_WINSOCK_MsgTooBig
    ERROR_FTP_WINSOCK_NetReset
    ERROR_FTP_WINSOCK_NetworkSubsystemFailed
    ERROR_FTP_WINSOCK_NetworkUnreachable
    ERROR_FTP_WINSOCK_NoBufferSpace
    ERROR_FTP_WINSOCK_NoData
    ERROR_FTP_WINSOCK_NonRecoverableError
    ERROR_FTP_WINSOCK_NotConnected
    ERROR_FTP_WINSOCK_NotInitialized
    ERROR_FTP_WINSOCK_NotSocket
    ERROR_FTP_WINSOCK_OpCanceled
    ERROR_FTP_WINSOCK_OutOfMemory
    ERROR_FTP_WINSOCK_OutOfRange
    ERROR_FTP_WINSOCK_PortNotSupported
    ERROR_FTP_WINSOCK_SetNotSupported
    ERROR_FTP_WINSOCK_SocketShutdown
    ERROR_FTP_WINSOCK_Success
    ERROR_FTP_WINSOCK_Timedout
    ERROR_FTP_WINSOCK_Unsupported
    ERROR_FTP_WINSOCK_WouldBlock
    ERROR_FTP_WINSOCK_WrongProtocol
    ERROR_FTP_PROTOCOL_SERVICE_READY_IN_MINUTES
    ERROR_FTP_PROTOCOL_USER_NAME_OK_NEED_PASSWORD
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_LOGIN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO
    ERROR_FTP_PROTOCOL_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
    ERROR_FTP_PROTOCOL_CANNOT_OPEN_DATA_CONNECTION
    ERROR_FTP_PROTOCOL_CONNECTION_CLOSED_TRANSFER_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_COMMAND_UNRECOGNIZED
    ERROR_FTP_PROTOCOL_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED
    ERROR_FTP_PROTOCOL_BAD_SEQUENCE_OF_COMMANDS
    ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER
    ERROR_FTP_PROTOCOL_NOT_LOGGED_IN
    ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_STORING_FILES
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN
    ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION
    ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED
    ERROR_FTP_USER_TIMEOUT
    ERROR_FTP_USER_TRANSFER_IN_PROGRESS
End Enum
'********************************************************************************
'Class events
'********************************************************************************
Public Event StateChanged(State As FTP_CONNECTION_STATES)
Public Event DownloadProgress(lBytes As Long)
Public Event UploadProgress(lBytes As Long)
Public Event ReplyMessage(ByRef sMessage As String)
'********************************************************************************
'Service constants and variables used inside the class
'********************************************************************************
Const RESPONSE_CODE_LENGHT = 3
Private m_LastError                     As FtpErrors
Private m_strLastErrorDesc              As String
Private m_strWinsockBuffer              As String
Private m_strDataBuffer                 As String
Private m_strLocalFilePath              As String
Private m_intLocalFileID                As Integer
Private m_bTransferInProgress           As Boolean
Private m_lDownloadedBytes              As Long
Private m_bUploadFile                   As Boolean
Private m_lUploadedBytes                As Long
Private m_strLastServerResponse         As String
Private m_objTimeOut                    As CTimeout
Private m_bFileIsOpened                 As Boolean
Private CHUNK_SIZE                      As Integer
'Trap for small files on fast connection
Private m_sToFast                       As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub DoEvents2()
    'Sleep 1 is to prevent 100% CPU Usage
    Sleep 1
    DoEvents
End Sub

Public Function FtpGetLastError() As FtpErrors
    FtpGetLastError = m_LastError
End Function

Public Function CurrentDirectory() As String
    CurrentDirectory = m_strCurrentDirectory
End Function

Public Function GetLastServerResponse() As String
    GetLastServerResponse = m_strLastServerResponse
End Function

Public Property Get TransferMode() As FtpTransferModes
    TransferMode = m_TransferMode
End Property

Public Property Let TransferMode(ByRef NewValue As FtpTransferModes)
    m_bBusy = True
    If Not (NewValue = m_TransferMode) Then
        If ProcessTYPECommand(NewValue) Then
            m_TransferMode = NewValue
        End If
    End If
    m_bBusy = False
End Property

Private Function ProcessLISTCommand() As Boolean
    On Error GoTo ProcessLISTCommand_Err_Handler
    Dim strResponse                     As String
    Dim strData                         As String
    wscControl.SendData "LIST" & vbCrLf
    RaiseEvent ReplyMessage("LIST")
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            If GetResponseCode(m_strWinsockBuffer) = 150 Or _
                GetResponseCode(m_strWinsockBuffer) = 125 Then
                'ignore 150 reply code
                m_strWinsockBuffer = Mid$(m_strWinsockBuffer, InStr(1, m_strWinsockBuffer, vbCrLf) + 2)
            Else
                strData = m_strWinsockBuffer
                m_strWinsockBuffer = ""
                Exit Do
            End If
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_CLOSING_DATA_CONNECTION Or _
        GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessLISTCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessLISTCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessLISTCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Property Get PassiveMode() As Boolean
    PassiveMode = m_bPassiveMode
End Property

Public Property Let PassiveMode(ByRef NewValue As Boolean)
    m_bPassiveMode = NewValue
End Property

Public Function EnumFiles(ByRef oFiles As CFTPFiles) As Boolean
    Dim bDataConnectionEstablished      As Boolean
    'On Error GoTo EnumFiles_Err_Handler
    m_bBusy = True
    If m_bPassiveMode Then
        'send PASV command
        bDataConnectionEstablished = ProcessPASVCommand
    Else
        'send PORT command
        bDataConnectionEstablished = ProcessPORTCommand
    End If 'm_bPassiveMode
    If bDataConnectionEstablished Then
        RaiseEvent StateChanged(FTP_RETRIEVING_DIRECTORY_INFO)
        If ProcessLISTCommand Then
            m_objTimeOut.StartTimer
            Do
                DoEvents2
                If m_objTimeOut.Timeout Then
                    m_LastError = ERROR_FTP_USER_TIMEOUT
                    If GetResponseCode(Left(m_strLastServerResponse, 3)) = FTP_RESPONSE_CLOSING_DATA_CONNECTION Then
                        Set oFiles = GetFileList(m_strDataBuffer)
                        EnumFiles = True
                        RaiseEvent StateChanged(FTP_DIRECTORY_INFO_COMPLETED)
                        m_strDataBuffer = ""
                    End If
                    Exit Do
                End If
                If wscData.State = sckClosing Or wscData.State = sckClosed Then
                    Set oFiles = Nothing
                    Set oFiles = GetFileList(m_strDataBuffer)
                    EnumFiles = True
                    RaiseEvent StateChanged(FTP_DIRECTORY_INFO_COMPLETED)
                    m_strDataBuffer = ""
                    Exit Do
                End If
            Loop
            m_objTimeOut.StopTimer
        Else
            'raise error - LIST command
        End If 'ProcessLISTCommand
    Else 'bDataConnectionEstablished
        'raise error - can't establish data connection
    End If 'bDataConnectionEstablished
Exit_Label:
    m_bBusy = False
    Exit Function
EnumFiles_Err_Handler:
    Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.EnumFiles", Err.Description
    GoTo Exit_Label
End Function

Public Function SetCurrentDirectory(ByRef sNewDirectory As String) As Boolean
    m_bBusy = True
    SetCurrentDirectory = ProcessCWDCommand(sNewDirectory)
    m_bBusy = False
End Function

Public Property Get FtpServer() As Variant
    FtpServer = m_varFtpServer
End Property

Public Property Let FtpServer(ByRef NewValue As Variant)
    m_varFtpServer = NewValue
End Property

Public Property Get FtpServerPort() As Variant
    FtpServerPort = m_varFtpServerPort
End Property

Public Property Let FtpServerPort(ByRef NewValue As Variant)
    m_varFtpServerPort = NewValue
End Property

Public Property Get Password() As String
    Password = m_strPassword
End Property

Public Property Let Password(ByRef NewValue As String)
    m_strPassword = NewValue
End Property

Public Property Get UserName() As String
    UserName = m_strUserName
End Property

Public Property Let UserName(ByRef NewValue As String)
    m_strUserName = NewValue
End Property
'********************************************************************************
'Purpose     :Establishes the connection to ftp server
'********************************************************************************

Public Function Connect(Optional ByVal sFTPServer As String, Optional ByVal sFTPServerPort As String) As Boolean
    On Error GoTo Connect_Err_Handler
    
    
    Dim strData                         As String
    m_strWinsockBuffer = ""
    m_bBusy = True
    If sFTPServer <> "" Then
        m_varFtpServer = sFTPServer
    End If
    If sFTPServerPort <> "" Then
        m_varFtpServerPort = sFTPServerPort
    End If
    If Len(m_varFtpServer) > 0 Then
        With wscControl
            .Close
            .LocalPort = 0
            .Connect m_varFtpServer, m_varFtpServerPort
            m_objTimeOut.StartTimer
            Do
                DoEvents2
                If m_objTimeOut.Timeout Then
                    m_LastError = ERROR_FTP_USER_TIMEOUT
                    Exit Do
                End If
                If .State = sckConnected Then
                    m_objTimeOut.StopTimer
                    RaiseEvent StateChanged(FTP_CONNECTION_CONNECTED)
                    m_objTimeOut.StartTimer
                    Do
                        DoEvents2
                        If m_objTimeOut.Timeout Then
                            m_LastError = ERROR_FTP_USER_TIMEOUT
                            Exit Do
                        End If
                        If Len(m_strWinsockBuffer) > (RESPONSE_CODE_LENGHT - 1) Then
                            strData = m_strWinsockBuffer
                            m_strWinsockBuffer = ""
                            Exit Do
                        End If
                    Loop
                    m_objTimeOut.StopTimer
                    Select Case GetResponseCode(strData)
                        Case FTP_RESPONSE_SERVICE_READY_FOR_NEW_USER
                            Select Case ProcessUSERCommand
                                Case FTP_RESPONSE_USER_LOGGED_IN
                                    Connect = True
                                Case FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD
                                    If ProcessPASSCommand = FTP_RESPONSE_USER_LOGGED_IN Then
                                        Connect = True
                                    End If
                            End Select
                            'Get working directory
                            If Connect Then
                                Call ProcessPWDCommand
                            End If
                        Case FTP_RESPONSE_SERVICE_READY_IN_MINUTES
                            '120 Service ready in nnn minutes.
                            m_LastError = ERROR_FTP_PROTOCOL_SERVICE_READY_IN_MINUTES
                        Case FTP_RESPONSE_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
                            '421 Service not available, closing control connection.
                            m_LastError = ERROR_FTP_PROTOCOL_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
                    End Select
                    Exit Do
                ElseIf .State = sckConnectAborted Then
                    m_LastError = ERROR_FTP_WINSOCK_ConnectAborted
                ElseIf .State = sckResolvingHost Then
                    RaiseEvent StateChanged(FTP_CONNECTION_RESOLVING_HOST)
                ElseIf .State = sckHostResolved Then
                    RaiseEvent StateChanged(FTP_CONNECTION_HOST_RESOLVED)
                End If
            Loop
            m_objTimeOut.StopTimer
        End With
    Else
        'raise error
        Connect = False
        Exit Function
    End If
Exit_Label:
    If Connect Then RaiseEvent StateChanged(FTP_USER_LOGGED)
    m_bBusy = False
    Exit Function
Connect_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.Connect", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Sub Class_Initialize()
    Set frmFTPContainer = New frmFTPWinsockContainter
    m_varFtpServerPort = 21
    Set wscControl = frmFTPContainer.sckControl
    Set wscData = frmFTPContainer.sckData
    Set m_objTimeOut = New CTimeout
    CHUNK_SIZE = 4096       'Default at 4k
End Sub

Private Function GetResponseCode(ByRef strResponse As String) As Integer
    If Len(strResponse) > (RESPONSE_CODE_LENGHT - 1) Then
        GetResponseCode = CInt(Left$(strResponse, 3))
    End If
End Function

Private Function ProcessUSERCommand() As FTP_RESPONSE_CODES
    Dim strData                         As String
    On Error GoTo ProcessUSERCommand_Err_Handler
    RaiseEvent StateChanged(FTP_CONNECTION_AUTHENTICATION)
    m_strUserName = IIf(Len(m_strUserName) > 0, m_strUserName, "anonymous")
    If Len(m_strPassword) = 0 Then
        If m_strUserName = "anonymous" Then
            m_strPassword = "guest@unknown.com"
        Else
            'raise error
            Exit Function
        End If
    End If
    wscControl.SendData "USER " & m_strUserName & vbCrLf
    RaiseEvent ReplyMessage("USER " & m_strUserName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If Len(m_strWinsockBuffer) > RESPONSE_CODE_LENGHT Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    Select Case GetResponseCode(strData)
        Case FTP_RESPONSE_USER_LOGGED_IN
            ProcessUSERCommand = FTP_RESPONSE_USER_LOGGED_IN
        Case FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD
            ProcessUSERCommand = FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD
        Case Else
            ProcessFtpResponse GetResponseCode(strData)
    End Select
Exit_Label:
    Exit Function
ProcessUSERCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessUSERCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessPASSCommand() As FTP_RESPONSE_CODES
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessPASSCommand_Err_Handler
    wscControl.SendData "PASS " & m_strPassword & vbCrLf
    RaiseEvent ReplyMessage("PASS " & "**********" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If Len(m_strWinsockBuffer) > RESPONSE_CODE_LENGHT Then
            strData = m_strWinsockBuffer
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_USER_LOGGED_IN Then
        Do
            DoEvents2
            If InStr(1, m_strWinsockBuffer, "230 ") > 0 Then
                ProcessPASSCommand = FTP_RESPONSE_USER_LOGGED_IN
                m_strWinsockBuffer = ""
                Exit Function
            End If
        Loop
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
    ProcessPASSCommand = GetResponseCode(strData)
Exit_Label:
    Exit Function
ProcessPASSCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessPASSCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessPWDCommand() As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessPWDCommand_Err_Handler
    wscControl.SendData "PWD" & vbCrLf
    RaiseEvent ReplyMessage("PWD" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_PATHNAME_CREATED Then
        Dim intPosA                     As Integer
        Dim intPosB                     As Integer
        intPosA = InStr(1, strData, Chr$(34)) + 1
        intPosB = InStr(intPosA, strData, Chr$(34))
        If intPosA > 1 And intPosB > 0 Then
            m_strCurrentDirectory = Mid$(strData, intPosA, intPosB - intPosA)
            ProcessPWDCommand = True
        Else
            'raise error - unknown response format
        End If
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessPWDCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessPWDCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Sub Class_Terminate()
    Call BreakeConnection
    Set wscData = Nothing
    Set wscControl = Nothing
    m_objTimeOut.StopTimer
    Set m_objTimeOut = Nothing
End Sub

Private Sub wscControl_DataArrival(ByVal bytesTotal As Long)
    Dim strData                         As String
    On Error GoTo ErrTrap
    Dim iRetryCounter                   As Integer
    wscControl.GetData strData
    m_strWinsockBuffer = m_strWinsockBuffer & strData
    m_strLastServerResponse = strData
    m_objTimeOut.Reset
    If GetResponseCode(strData) = 426 Then
        If m_bTransferInProgress Or m_bUploadFile Then
            wscData.Close
            Close m_intLocalFileID
            m_strDataBuffer = ""
            m_lDownloadedBytes = 0
            m_lUploadedBytes = 0
            m_bTransferInProgress = False
            m_bUploadFile = False
            m_bFileIsOpened = False
        End If
        wscControl.Close
        m_bBusy = False
    ElseIf InStr(strData, "425") Then
        'If m_bTransferInProgress Or m_bUploadFile Then
        '    wscData.Close
        '    Close m_intLocalFileID
        '    m_strDataBuffer = ""
        '    m_lDownloadedBytes = 0
        '    m_lUploadedBytes = 0
        '    m_bTransferInProgress = False
        '    m_bUploadFile = False
        '    m_bFileIsOpened = False
        'End If
        'wscControl.Close
        'm_bBusy = False
    End If
    'Debug.Print Left(strData, Len(strData) - 2)
    RaiseEvent ReplyMessage(Left(strData, Len(strData) - 2) & vbCrLf)
    Exit Sub
ErrTrap:
    If Err.Number = 40006 Then
        'Debug.Assert False
        iRetryCounter = iRetryCounter + 1
        If iRetryCounter <= 5 Then
            DoEvents2
            Resume
        End If
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Function ProcessPORTCommand() As Boolean
    Dim intPort                         As Integer
    Dim strIPAddress                    As String
    Dim colIPAddresses                  As New Collection
    Dim strSend                         As String
    Dim strData                         As String
    On Error Resume Next
    RaiseEvent StateChanged(FTP_ESTABLISHING_DATA_CONNECTION)
    Do
        intPort = GetFreePort
        'Debug.Print "intPort: " & intPort
        If wscData.State <> sckClosed Then wscData.Close
        wscData.LocalPort = intPort
        wscData.Listen
        If Not Err Then Exit Do
    Loop
    On Error GoTo ProcessPORTCommand_Err_Handler
    strIPAddress = CStr(wscControl.LocalIP)
    strSend = "PORT " & Replace(strIPAddress, ".", ",")
    strSend = strSend & "," & intPort \ 256 & "," & (intPort Mod 256)
    strSend = strSend & vbCrLf
    wscControl.SendData strSend
    If InStr(strData, "425") Then
        'Stop
    End If
    RaiseEvent ReplyMessage(Left(strSend, Len(strSend) - 2) & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_COMMAND_OK Then
        ProcessPORTCommand = True
        RaiseEvent StateChanged(FTP_DATA_CONNECTION_ESTABLISHED)
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessPORTCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessPORTCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function GetFreePort() As Integer
    'Zie mod ftp support
    'Static intPort As Integer
    If g_intPort = 0 Then
        g_intPort = 1100
    Else
        g_intPort = g_intPort + 1
    End If
    If g_intPort > 32768 Then
        g_intPort = 1100
    End If
    GetFreePort = g_intPort
End Function

Private Sub wscData_ConnectionRequest(ByVal requestID As Long)
    If wscData.State <> sckClosed Then
        wscData.Close
    End If
    wscData.Accept (requestID)
End Sub

Private Sub wscData_DataArrival(ByVal bytesTotal As Long)
    Dim strData                         As String
    If bytesTotal = 0 Then
        Exit Sub
    End If
    If Not wscData.State = sckConnected Then
        'Stop
    End If
    wscData.GetData strData
    Do Until strData = ""
        If m_bTransferInProgress Then
            If m_bFileIsOpened Then
                'write data to local file
                Put m_intLocalFileID, , strData
                'raise DownloadProgress event
                m_lDownloadedBytes = m_lDownloadedBytes + bytesTotal
                RaiseEvent DownloadProgress(m_lDownloadedBytes)
            Else
                m_sToFast = m_sToFast & strData
            End If
        Else
            m_strDataBuffer = m_strDataBuffer & strData
        End If
        strData = ""
        DoEvents2
        If wscData.State = sckConnected Then
            wscData.PeekData strData
            'Debug.Print "<" & strData
            If Len(strData) <> 0 Then
                wscData.GetData strData
            End If
        End If
    Loop
    m_objTimeOut.Reset
End Sub

Public Function RenameFile(ByRef sOldFileName As String, ByRef sNewFileName As String) As Boolean
    m_bBusy = True
    If ProcessRNFRCommand(sOldFileName) Then
        If ProcessRNTOCommand(sNewFileName) Then
            RenameFile = True
        End If
    End If
    m_bBusy = False
End Function

Public Function DeleteFile(ByRef sFileName As String) As Boolean
    m_bBusy = True
    DeleteFile = ProcessDELECommand(sFileName)
    m_bBusy = False
End Function

Public Function RemoveDirectory(ByRef sDirName As String) As Boolean
    m_bBusy = True
    RemoveDirectory = ProcessRMDCommand(sDirName)
    m_bBusy = False
End Function

Public Function CreateDirectory(ByRef sDirName As String) As Boolean
    m_bBusy = True
    CreateDirectory = ProcessMKDCommand(sDirName)
    m_bBusy = False
End Function

Private Function GetFileList(ByRef sListing As String) As CFTPFiles
    Dim vFiles                          As Variant
    Dim vFile                           As Variant
    Dim vComponents                     As Variant
    Dim oFtpFile                        As CFTPFile
    Dim oFtpFiles                       As New CFTPFiles
    Dim icounter                        As Integer
    Dim strFile                         As String
    On Error Resume Next
    Set GetFileList = Nothing
    vFiles = Split(sListing, vbCrLf)
    For Each vFile In vFiles
        Set oFtpFile = New CFTPFile
        'replace multiple whitespaces with single whitespace
        For icounter = 15 To 2 Step -1
            vFile = Replace(vFile, Space(icounter), " ")
        Next
        If Len(vFile) > 0 Then
            If Not LCase(Left(vFile, 5)) = "total" Then
                vComponents = Split(vFile, " ")
                If UBound(vComponents) > 7 Then
                    With oFtpFile
                        If Left(vComponents(0), 1) = "d" Then
                            oFtpFile.IsDirectory = True
                        ElseIf Left(vFile, 1) = "l" Then
                            .FilePath = vComponents(10)
                            If Not CBool(InStr(InStrRev(vComponents(10), "/") + 1, vComponents(10), ".")) Then
                                .IsDirectory = True
                            End If
                        End If
                        .Permissions = vComponents(0)
                        .Owner = vComponents(2)
                        .Group = vComponents(3)
                        .FileSize = vComponents(4)
                        
                        '.FileName = vComponents(8)
                        .FileName = ""
                        For icounter = 8 To UBound(vComponents)
                            .FileName = .FileName & vComponents(icounter) & Space(1)
                        Next
                        .FileName = Left(.FileName, Len(.FileName) - 1)
                        .LastWriteTime = GetDate(vComponents(6), vComponents(5), vComponents(7))
                        If Not (.FileName = "." Or .FileName = "..") Then
                            oFtpFiles.Add oFtpFile, oFtpFile.FileName
                        End If
                    End With
                Else
                    With oFtpFile
                        If vComponents(2) = "<DIR>" Then
                            .IsDirectory = True
                        Else
                            .FileSize = CLng(vComponents(2))
                        End If
                        If UBound(vComponents) > 3 Then
                            For icounter = 3 To UBound(vComponents)
                                strFile = strFile & " " & vComponents(icounter)
                            Next icounter
                            strFile = Mid$(strFile, 2)
                        Else
                            strFile = vComponents(3)
                        End If
                        .FileName = strFile
                        .LastWriteTime = CDate(vComponents(0) & " " & vComponents(1))
                        oFtpFiles.Add oFtpFile, oFtpFile.FileName
                    End With
                End If
                Set oFtpFile = Nothing
            End If
        End If
        strFile = ""
    Next
    Set GetFileList = oFtpFiles
    Set oFtpFiles = Nothing
End Function

Private Function GetDate(ByRef vDay, ByRef vMonth, ByRef vYear) As Date
    vYear = IIf(InStr(1, vYear, ":"), Year(Now), vYear)
    Select Case vMonth
        Case "Jan"
            vMonth = 1
        Case "Feb"
            vMonth = 2
        Case "Mar"
            vMonth = 3
        Case "Apr"
            vMonth = 4
        Case "May"
            vMonth = 5
        Case "Jun"
            vMonth = 6
        Case "Jul"
            vMonth = 7
        Case "Aug"
            vMonth = 8
        Case "Sep"
            vMonth = 9
        Case "Oct"
            vMonth = 10
        Case "Nov"
            vMonth = 11
        Case "Dec"
            vMonth = 12
    End Select
    GetDate = DateSerial(CInt(vYear), CInt(vMonth), CInt(vDay))
End Function

Private Function ProcessPASVCommand() As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessPASVCommand_Err_Handler
    RaiseEvent StateChanged(FTP_ESTABLISHING_DATA_CONNECTION)
    wscControl.SendData "PASV" & vbCrLf
    RaiseEvent ReplyMessage("PASV" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_ENTERING_PASSIVE_MODE Then
        ProcessPASVCommand = MakePassiveDataConnection(strData)
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessPASVCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessPASVCommand", Err.Description
    End If
    GoTo Exit_Label
End Function
'Example of the string passed with sData argument
'227 Entering Passive Mode (194,220,224,2,7,189)

Private Function MakePassiveDataConnection(ByRef sData As String) As Boolean
    Dim iPos                            As Integer
    Dim iPos2                           As Integer
    Dim sDataAddress                    As String
    Dim strIP                           As String
    Dim lPort                           As Long
    On Error GoTo MakePassiveDataConnection_Err_Handler
    iPos = InStr(1, sData, "(") + 1
    If Not CBool(iPos) Then
        Exit Function
    End If
    sDataAddress = Mid$(sData, iPos, InStr(1, sData, ")") - iPos)
    sDataAddress = Replace(sDataAddress, ",", ".", 1, 3)
    iPos = InStr(1, sDataAddress, ",")
    strIP = Left$(sDataAddress, iPos - 1)
    lPort = CLng(Mid$(sDataAddress, iPos + 1, InStr(iPos + 1, sDataAddress, ",") - iPos))
    lPort = lPort * 256
    lPort = lPort + CLng(Mid$(sDataAddress, InStrRev(sDataAddress, ",") + 1))
    wscData.Close
    wscData.LocalPort = 0
    wscData.Connect strIP, lPort
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If wscData.State = sckConnected Then
            MakePassiveDataConnection = True
            RaiseEvent StateChanged(FTP_DATA_CONNECTION_ESTABLISHED)
            RaiseEvent ReplyMessage("Connecting to: " & strIP & ":" & lPort & vbCrLf)
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
Exit_Label:
    Exit Function
MakePassiveDataConnection_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.MakePassiveDataConnection", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Function DownloadFile(ByRef sFileName As String, ByRef sLocalFileName As String, ByRef vTransferMode As FtpTransferModes, Optional ByRef lStartPoint As Long) As Boolean
    Dim bDataConnectionEstablished      As Boolean
    Dim CFTPFiles As CFTPFiles
    Dim CFTPFile As CFTPFile
    

        If EnumFiles(CFTPFiles) Then
        For Each CFTPFile In CFTPFiles
            If LCase$(CFTPFile.FileName) = LCase$(sFileName) Then
                BytesRemaining = CFTPFile.FileSize
'                MsgBox BytesRemaining
                Exit For
            End If
        Next
    End If


    
    m_sToFast = ""          'Clear small/fast file buffer
    m_bBusy = True
    If ProcessTYPECommand(vTransferMode) Then
        m_TransferMode = vTransferMode
    Else
        Exit Function
    End If
    If m_bPassiveMode Then
        bDataConnectionEstablished = ProcessPASVCommand
    Else
        bDataConnectionEstablished = ProcessPORTCommand
    End If
    If bDataConnectionEstablished Then
     
    
        If lStartPoint > 0 Then
            m_lDownloadedBytes = lStartPoint
            If Not ProcessRESTCommand(lStartPoint) Then
                'can't restart download
                DownloadFile = False
                Exit Function
            End If
        End If




        m_bTransferInProgress = True
        m_strLocalFilePath = sLocalFileName
        If ProcessRETRCommand(sFileName, lStartPoint) Then
            m_objTimeOut.StartTimer
            Do
                DoEvents2
                If m_objTimeOut.Timeout Then
                    m_LastError = ERROR_FTP_USER_TIMEOUT
                    Exit Do
                End If
                If wscData.State = sckClosed Or wscData.State = sckClosing Then
                    RaiseEvent StateChanged(FTP_TRANSFER_COMLETED)
                    'close file
                    Close #m_intLocalFileID
                    m_bFileIsOpened = False
                    m_bTransferInProgress = False
                    m_lDownloadedBytes = 0
                    If Left$(GetLastServerResponse, 3) = "426" Then
                        m_LastError = FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED
                        Call ProcessFtpResponse(FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED)
                        DownloadFile = False
                    Else
                        DownloadFile = True
                    End If
                    Exit Do
                End If
            Loop
            m_objTimeOut.StopTimer
        Else
            DownloadFile = False
            m_bTransferInProgress = False
            Close m_intLocalFileID
        End If
    End If
    m_bBusy = False
End Function

Private Function ProcessRETRCommand(ByRef sFileName As String, ByRef lStartPoint As Long) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessRETRCommand_Err_Handler
    m_strDataBuffer = ""
    wscControl.SendData "RETR " & sFileName & vbCrLf
    RaiseEvent ReplyMessage("RETR " & sFileName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If Not m_bTransferInProgress Then
            strData = m_strWinsockBuffer
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            If GetResponseCode(m_strWinsockBuffer) = 150 Or _
                GetResponseCode(m_strWinsockBuffer) = 125 Then
                If lStartPoint = 0 And FileExists(m_strLocalFilePath) Then
                    Kill m_strLocalFilePath
                End If
                m_intLocalFileID = FreeFile
                Open m_strLocalFilePath For Binary As m_intLocalFileID
                If lStartPoint > 0 Then
                    Seek m_intLocalFileID, lStartPoint + 1
                End If
                'turn on flag m_bFileIsOpened
                m_bFileIsOpened = True
                If m_sToFast <> "" Then     'Already fetched data before we had a chance to open the file
                    'write data to local file
                    Put m_intLocalFileID, , m_sToFast
                    'raise DownloadProgress event
                    'm_lDownloadedBytes = m_lDownloadedBytes + bytesTotal
                    'RaiseEvent DownloadProgress(m_lDownloadedBytes)
                End If
                'ignore 150 and 125 reply codes
                m_strWinsockBuffer = Mid$(m_strWinsockBuffer, InStr(1, m_strWinsockBuffer, vbCrLf) + 2)
                RaiseEvent StateChanged(FTP_TRANSFER_STARTING)
            Else
                strData = m_strWinsockBuffer
                m_strWinsockBuffer = ""
                Exit Do
            End If
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_CLOSING_DATA_CONNECTION Or _
        GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessRETRCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessRETRCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessRETRCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessRESTCommand(ByRef lStartPoint As Long) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessRESTCommand_Err_Handler
    wscControl.SendData "REST " & lStartPoint & vbCrLf
    RaiseEvent ReplyMessage("REST " & lStartPoint & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO Then
        ProcessRESTCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessRESTCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessRESTCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Sub BreakeConnection()
    On Error Resume Next
    If wscData <> sckClosed Then
        wscData.Close
    Else
        wscControl.Close
    End If
    If m_bTransferInProgress Or m_bUploadFile Then
        Close m_intLocalFileID
        m_strDataBuffer = ""
        m_lDownloadedBytes = 0
        m_lUploadedBytes = 0
        m_bTransferInProgress = False
        m_bUploadFile = False
    End If
    m_bFileIsOpened = False
    m_bBusy = False
    m_objTimeOut.StopTimer
End Sub

Private Function ProcessTYPECommand(ByRef vType As FtpTransferModes) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessTYPECommand_Err_Handler
    wscControl.SendData "TYPE " & IIf(vType = FTP_ASCII_MODE, "A", "I") & vbCrLf
    RaiseEvent ReplyMessage("TYPE " & IIf(vType = FTP_ASCII_MODE, "A", "I") & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_COMMAND_OK Then
        ProcessTYPECommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessTYPECommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessTYPECommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function FileExists(ByRef sFileName As String) As Boolean
    On Error GoTo ERROR_HANDLER
    FileExists = (Dir(sFileName) <> "") And (Dir(sFileName, vbDirectory) = "")
    ''FileExists = (GetAttr(sFileName) And vbDirectory) = 0
ERROR_HANDLER:
End Function

Private Function ProcessDELECommand(ByRef sFileName As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessDELECommand_Err_Handler
    wscControl.SendData "DELE " & sFileName & vbCrLf
    RaiseEvent ReplyMessage("DELE " & sFileName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessDELECommand = True
    Else
        ProcessFtpResponse (GetResponseCode(strData))
    End If
Exit_Label:
    Exit Function
ProcessDELECommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessDELECommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessMKDCommand(ByRef sDirName As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessMKDCommand_Err_Handler
    wscControl.SendData "MKD " & sDirName & vbCrLf
    RaiseEvent ReplyMessage("MKD " & sDirName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_PATHNAME_CREATED Then
        ProcessMKDCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessMKDCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessMKDCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessRMDCommand(ByRef sDirName As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessRMDCommand_Err_Handler
    wscControl.SendData "RMD " & sDirName & vbCrLf
    RaiseEvent ReplyMessage("RMD " & sDirName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessRMDCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessRMDCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessRMDCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessRNFRCommand(ByRef sFileName As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessRNFRCommand_Err_Handler
    wscControl.SendData "RNFR " & sFileName & vbCrLf
    RaiseEvent ReplyMessage("RNFR " & sFileName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO Then
        ProcessRNFRCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessRNFRCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessRNFRCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessRNTOCommand(ByRef sFileName As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessRNTOCommand_Err_Handler
    wscControl.SendData "RNTO " & sFileName & vbCrLf
    'Debug.Print "RNTO " & sFileName
    RaiseEvent ReplyMessage("RNTO " & sFileName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        '
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        '
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessRNTOCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessRNTOCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessRNTOCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Function UploadFile(ByRef sLocalFileName As String, ByRef sRemoteFileName As String, Optional ByRef vTransferMode As FtpTransferModes = FTP_IMAGE_MODE, Optional ByRef lStartPoint As Long = 0) As Boolean
    Dim bDataConnectionEstablished      As Boolean
    m_bBusy = True
    If Not (vTransferMode = m_TransferMode) Then
        If ProcessTYPECommand(vTransferMode) Then
            m_TransferMode = vTransferMode
        Else
            Exit Function
        End If
    End If
    If m_bPassiveMode Then
        bDataConnectionEstablished = ProcessPASVCommand
    Else
        bDataConnectionEstablished = ProcessPORTCommand
    End If
    If bDataConnectionEstablished Then
        '
        If Not IsMissing(lStartPoint) Then
            If Not ProcessRESTCommand(lStartPoint) Then
                UploadFile = False
                Exit Function
            End If
        End If
        '
        m_strLocalFilePath = sLocalFileName
        m_bUploadFile = True
        If ProcessSTORCommand(sLocalFileName, sRemoteFileName, lStartPoint) Then
            m_objTimeOut.StartTimer
            Do
                DoEvents2
                '
                If m_objTimeOut.Timeout Then
                    m_LastError = ERROR_FTP_USER_TIMEOUT
                    Exit Do
                End If
                '
                If wscData.State = sckClosing Or _
                    wscData.State = sckClosed Then
                    'clear winsock buffer
                    RaiseEvent StateChanged(FTP_TRANSFER_COMLETED)
                    Exit Do
                End If
            Loop
            m_objTimeOut.StopTimer
            UploadFile = True
        End If
    End If
    m_bBusy = False
End Function

Private Function ProcessSTORCommand(ByRef sLocalFileName As String, ByRef sRemoteFileName As String, ByRef lStartPoint As Long) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessSTORCommand_Err_Handler
    m_strDataBuffer = ""
    wscControl.SendData "STOR " & sRemoteFileName & vbCrLf
    'Debug.Print "STOR " & sRemoteFileName
    RaiseEvent ReplyMessage("STOR " & sRemoteFileName & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        '
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        '
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            If GetResponseCode(m_strWinsockBuffer) = 150 Or _
                GetResponseCode(m_strWinsockBuffer) = 125 Then
                m_strWinsockBuffer = ""
                RaiseEvent StateChanged(FTP_TRANSFER_STARTING)
                m_strLocalFilePath = sLocalFileName
                Call UploadData(lStartPoint)
            Else
                strData = m_strWinsockBuffer
                m_strWinsockBuffer = ""
                Exit Do
            End If
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_CLOSING_DATA_CONNECTION Or _
        GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessSTORCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessSTORCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessSTORCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Sub wscData_SendComplete()
    If m_bUploadFile Then
        Call UploadData(0)
    End If
    m_objTimeOut.Reset
End Sub

Private Sub wscData_SendProgress(ByVal bytesSent As Long, ByVal BytesRemaining As Long)
    m_lUploadedBytes = m_lUploadedBytes + bytesSent
    RaiseEvent UploadProgress(m_lUploadedBytes)
End Sub
'********************************************************************************
'Purpose     :Opens file, reads data from the file and
'             sends the data to remote computer by (CHUNK_SIZE) chunks(default at 4k(4096b)).
'Description :If file size is more than CHUNK_SIZE the procedure called one or
'             multiple times from wscFtpData_SendComplete event procedure.
'********************************************************************************

Private Sub UploadData(ByRef lStartPoint As Long)
    Static bFileIsOpen                  As Boolean      'flag variable
    Static lchunksCount                 As Long         'quantity of chunks to send
    Static lCounter                     As Long         'sent chunks counter
    Static intRemainder                 As Integer
    Dim strData                         As String       'data buffer to send
    On Error GoTo UploadData_Err_Handler
    'if bFileIsOpen = True, the procedure was called before
    If m_bFileIsOpened Then
        'if we have to send next chunk
        If lCounter < lchunksCount And lCounter > 0 Then
            'prepare the buffer
            strData = Space(CHUNK_SIZE)
            'increament counter
            lCounter = lCounter + 1
            'read data from file
            Get m_intLocalFileID, , strData
            'send data
            wscData.SendData strData
        Else
            'all the data is sent
            If lCounter = 0 Then
                '
                'close data connection to inform ftp server
                'that transfer is comlteted
                '
                wscData.Close
                '
                'close local file
                '
                Close #m_intLocalFileID
                '
                RaiseEvent StateChanged(FTP_TRANSFER_COMLETED)
                '
                'reset values of all static and module
                'level variables
                '
                m_lUploadedBytes = 0
                lchunksCount = 0
                intRemainder = 0
                m_bFileIsOpened = False
                m_bUploadFile = False
                '
            Else
                'all the chunks are sent
                'now we have to send the remainder
                '
                'prepare the buffer
                strData = Space(intRemainder)
                'reset the counter
                lCounter = 0
                'read data from file
                Get m_intLocalFileID, , strData
                'send data
                m_objTimeOut.StartTimer
                Do
                    DoEvents2
                    '
                    If m_objTimeOut.Timeout Then
                        m_LastError = ERROR_FTP_USER_TIMEOUT
                        Exit Do
                    End If
                    '
                    If wscData.State = sckConnected Then
                        wscData.SendData strData
                        Exit Do
                    End If
                Loop
                m_objTimeOut.StopTimer
            End If
        End If
    Else
        'if we are here, the procedure called at first time
        m_bFileIsOpened = True  'turn on flag variable
        m_intLocalFileID = FreeFile
        Open m_strLocalFilePath For Binary As m_intLocalFileID
        If lStartPoint > 0 Then
            Seek m_intLocalFileID, lStartPoint + 1
            m_lUploadedBytes = lStartPoint
            'get quantity of chancks to send
            lchunksCount = CLng((FileLen(m_strLocalFilePath) - lStartPoint) \ CHUNK_SIZE)
            'get remainder in bytes
            intRemainder = (FileLen(m_strLocalFilePath) - lStartPoint) Mod CHUNK_SIZE
        Else
            'get quantity of chancks to send
            lchunksCount = CLng(FileLen(m_strLocalFilePath) \ CHUNK_SIZE)
            'get remainder in bytes
            intRemainder = FileLen(m_strLocalFilePath) Mod CHUNK_SIZE
        End If
        If lchunksCount = 0 Then
            'if amount of data is less then 4Kb
            'prepare buffer to read data from a file
            strData = Space(intRemainder)
        Else
            'prepare buffer to read data from a file
            strData = Space(CHUNK_SIZE)
            'increament counter of sent chunks
            lCounter = 1
        End If
        'open file to read data
        'Open m_strLocalFilePath For Binary As #intFile
        'read data to buffer strData
        Get m_intLocalFileID, , strData
        'send data
        Do
            DoEvents2
            If wscData.State = sckConnected Then
                wscData.SendData strData
                Exit Do
                'ElseIf wscData.State = 0 Then
                '    UploadData lStartPoint
                '    Exit Sub
            End If
        Loop
        'If lCounter>0, file size if equal or less then chunk size
        'and we have to send more data. At the next time this sub will
        'be called from wscData_SendComplete event procedure to send
        'next chunk or remainder.
    End If
Exit_Label:
    Exit Sub
UploadData_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.UploadData", Err.Description
    End If
    'Close #intFile
    Close #m_intLocalFileID
    GoTo Exit_Label
End Sub

Private Function ShowTimeOut() As Boolean
    Dim intRetVal                       As Integer
    intRetVal = MsgBox("A time-out occurred while communicating with the server." & _
                "The server took too long to respond." & vbCrLf & vbCrLf & _
                "Would you like to wait for server response?", vbYesNo + vbQuestion, _
                "Time out")
    If intRetVal = vbYes Then
        m_objTimeOut.Reset
        m_objTimeOut.StartTimer
        ShowTimeOut = True
    End If
End Function

Public Property Let Timeout(ByRef NewValue As Integer)
    m_intTimeout = NewValue
    m_objTimeOut.TimeoutValue = NewValue
End Property

Public Property Get Timeout() As Integer
    Timeout = m_intTimeout
End Property

Public Property Get Busy() As Boolean
    Busy = m_bBusy
End Property

Private Function ProcessWinsockError(ByRef iError As ErrorConstants, ByRef sDesc As String) As Boolean
    m_strLastErrorDesc = sDesc
    Select Case iError
        Case sckAddressInUse
            m_LastError = ERROR_FTP_WINSOCK_AddressInUse
        Case sckAddressNotAvailable
            m_LastError = ERROR_FTP_WINSOCK_AddressNotAvailable
        Case sckAlreadyComplete
            m_LastError = ERROR_FTP_WINSOCK_AlreadyComplete
        Case sckAlreadyConnected
            m_LastError = ERROR_FTP_WINSOCK_AlreadyConnected
        Case sckBadState
            m_LastError = ERROR_FTP_WINSOCK_BadState
        Case sckConnectAborted
            m_LastError = ERROR_FTP_WINSOCK_ConnectAborted
        Case sckConnectionRefused
            m_LastError = ERROR_FTP_WINSOCK_ConnectionRefused
        Case sckConnectionReset
            m_LastError = ERROR_FTP_WINSOCK_ConnectionReset
        Case sckGetNotSupported
            m_LastError = ERROR_FTP_WINSOCK_GetNotSupported
        Case sckHostNotFound
            m_LastError = ERROR_FTP_WINSOCK_HostNotFound
        Case sckHostNotFoundTryAgain
            m_LastError = ERROR_FTP_WINSOCK_HostNotFoundTryAgain
        Case sckInProgress
            m_LastError = ERROR_FTP_WINSOCK_InProgress
        Case sckInvalidArg
            m_LastError = ERROR_FTP_WINSOCK_InvalidArg
        Case sckInvalidArgument
            m_LastError = ERROR_FTP_WINSOCK_InvalidArgument
        Case sckInvalidOp
            m_LastError = ERROR_FTP_WINSOCK_InvalidOp
        Case sckInvalidPropertyValue
            m_LastError = ERROR_FTP_WINSOCK_InvalidPropertyValue
        Case sckMsgTooBig
            m_LastError = ERROR_FTP_WINSOCK_MsgTooBig
        Case sckNetReset
            m_LastError = ERROR_FTP_WINSOCK_NetReset
        Case sckNetworkSubsystemFailed
            m_LastError = ERROR_FTP_WINSOCK_NetworkSubsystemFailed
        Case sckNetworkUnreachable
            m_LastError = ERROR_FTP_WINSOCK_NetworkUnreachable
        Case sckNoBufferSpace
            m_LastError = ERROR_FTP_WINSOCK_NoBufferSpace
        Case sckNoData
            m_LastError = ERROR_FTP_WINSOCK_NoData
        Case sckNonRecoverableError
            m_LastError = ERROR_FTP_WINSOCK_NonRecoverableError
        Case sckNotConnected
            m_LastError = ERROR_FTP_WINSOCK_NotConnected
        Case sckNotInitialized
            m_LastError = ERROR_FTP_WINSOCK_NotInitialized
        Case sckNotSocket
            m_LastError = ERROR_FTP_WINSOCK_NotSocket
        Case sckOpCanceled
            m_LastError = ERROR_FTP_WINSOCK_OpCanceled
        Case sckOutOfMemory
            m_LastError = ERROR_FTP_WINSOCK_OutOfMemory
        Case sckOutOfRange
            m_LastError = ERROR_FTP_WINSOCK_OutOfRange
        Case sckPortNotSupported
            m_LastError = ERROR_FTP_WINSOCK_PortNotSupported
        Case sckSetNotSupported
            m_LastError = ERROR_FTP_WINSOCK_SetNotSupported
        Case sckSocketShutdown
            m_LastError = ERROR_FTP_WINSOCK_SocketShutdown
        Case sckSuccess
            m_LastError = ERROR_FTP_WINSOCK_Success
        Case sckTimedout
            m_LastError = ERROR_FTP_WINSOCK_Timedout
        Case sckUnsupported
            m_LastError = ERROR_FTP_WINSOCK_Unsupported
        Case sckWouldBlock
            m_LastError = ERROR_FTP_WINSOCK_WouldBlock
        Case sckWrongProtocol
            m_LastError = ERROR_FTP_WINSOCK_WrongProtocol
        Case Else
            ProcessWinsockError = False
            Exit Function
    End Select
    ProcessWinsockError = True
End Function

Private Function ProcessFtpResponse(ByRef iCode As FTP_RESPONSE_CODES) As Boolean
    Select Case iCode
        Case FTP_RESPONSE_RESTART_MARKER_REPLY
        Case FTP_RESPONSE_SERVICE_READY_IN_MINUTES
        Case FTP_RESPONSE_DATA_CONNECTION_ALREADY_OPEN
        Case FTP_RESPONSE_FILE_STATUS_OK
        Case FTP_RESPONSE_COMMAND_OK
        Case FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_SUPERFLUOUS_AT_THIS_SITE
        Case FTP_RESPONSE_SYSTEM_STATUS_OR_SYSTEM_HELP_REPLY
        Case FTP_RESPONSE_DIRECTORY_STATUS
        Case FTP_RESPONSE_FILE_STATUS
        Case FTP_RESPONSE_HELP_MESSAGE
        Case FTP_RESPONSE_NAME_SYSTEM_TYPE
        Case FTP_RESPONSE_SERVICE_READY_FOR_NEW_USER
        Case FTP_RESPONSE_SERVICE_CLOSING_CONTROL_CONNECTION
        Case FTP_RESPONSE_DATA_CONNECTION_OPEN
        Case FTP_RESPONSE_CLOSING_DATA_CONNECTION
        Case FTP_RESPONSE_ENTERING_PASSIVE_MODE
        Case FTP_RESPONSE_USER_LOGGED_IN
        Case FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED
        Case FTP_RESPONSE_PATHNAME_CREATED
        Case FTP_RESPONSE_USER_NAME_OK_NEED_PASSWORD
            m_LastError = ERROR_FTP_PROTOCOL_USER_NAME_OK_NEED_PASSWORD
        Case FTP_RESPONSE_NEED_ACCOUNT_FOR_LOGIN
            m_LastError = ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_LOGIN
        Case FTP_RESPONSE_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_PENDING_FURTHER_INFO
        Case FTP_RESPONSE_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
            m_LastError = ERROR_FTP_PROTOCOL_SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION
            m_strLastErrorDesc = "Service not available, closing control connection."
        Case FTP_RESPONSE_CANNOT_OPEN_DATA_CONNECTION
            m_strLastErrorDesc = "Can't open data connection."
            m_LastError = ERROR_FTP_PROTOCOL_CANNOT_OPEN_DATA_CONNECTION
        Case FTP_RESPONSE_CONNECTION_CLOSED_TRANSFER_ABORTED
            m_strLastErrorDesc = "Connection closed; transfer aborted."
            m_LastError = ERROR_FTP_PROTOCOL_CONNECTION_CLOSED_TRANSFER_ABORTED
        Case FTP_RESPONSE_REQUESTED_FILE_ACTION_NOT_TAKEN
            m_strLastErrorDesc = "Requested file action not taken."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN
        Case FTP_RESPONSE_REQUESTED_ACTION_ABORTED
            m_strLastErrorDesc = "Requested action aborted: local error in processing."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED
        Case FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN
            m_strLastErrorDesc = "Requested action not taken. Insufficient storage space in system."
        Case FTP_RESPONSE_SYNTAX_ERROR_COMMAND_UNRECOGNIZED
            m_strLastErrorDesc = "Syntax error, command unrecognized."
            m_LastError = ERROR_FTP_PROTOCOL_SYNTAX_ERROR_COMMAND_UNRECOGNIZED
        Case FTP_RESPONSE_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS
            m_strLastErrorDesc = "Syntax error in parameters or arguments."
            m_LastError = ERROR_FTP_PROTOCOL_SYNTAX_ERROR_IN_PARAMETERS_OR_ARGUMENTS
        Case FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED
            m_strLastErrorDesc = "Command not implemented."
            m_LastError = ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED
        Case FTP_RESPONSE_BAD_SEQUENCE_OF_COMMANDS
            m_strLastErrorDesc = "Bad sequence of commands."
            m_LastError = ERROR_FTP_PROTOCOL_BAD_SEQUENCE_OF_COMMANDS
        Case FTP_RESPONSE_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER
            m_strLastErrorDesc = "Command not implemented for that parameter."
            m_LastError = ERROR_FTP_PROTOCOL_COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER
        Case FTP_RESPONSE_NOT_LOGGED_IN
            m_strLastErrorDesc = "Not logged in."
            m_LastError = ERROR_FTP_PROTOCOL_NOT_LOGGED_IN
        Case FTP_RESPONSE_NEED_ACCOUNT_FOR_STORING_FILES
            m_strLastErrorDesc = "Need account for storing files."
            m_LastError = ERROR_FTP_PROTOCOL_NEED_ACCOUNT_FOR_STORING_FILES
        Case FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE
            m_strLastErrorDesc = "Requested action not taken. File unavailable (e.g., file not found, no access)."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_UNAVAILABLE
        Case FTP_RESPONSE_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN
            m_strLastErrorDesc = "Requested action aborted: page type unknown."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_ABORTED_PAGE_TYPE_UNKNOWN
        Case FTP_RESPONSE_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION
            m_strLastErrorDesc = "Requested file action aborted. Exceeded storage allocation (for current directory or dataset)."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_FILE_ACTION_ABORTED_EXCEEDED_STORAGE_ALLOCATION
        Case FTP_RESPONSE_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED
            m_strLastErrorDesc = "Requested action not taken. File name not allowed."
            m_LastError = ERROR_FTP_PROTOCOL_REQUESTED_ACTION_NOT_TAKEN_FILE_NAME_NOT_ALLOWED
        Case Else
            ProcessFtpResponse = False
            Exit Function
    End Select
    ProcessFtpResponse = True
End Function

Public Function GetCurrentDirectory() As String
    m_bBusy = True
    If ProcessPWDCommand Then
        GetCurrentDirectory = m_strCurrentDirectory
    End If
    m_bBusy = False
End Function

Private Function ProcessQUITCommand() As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessQUITCommand_Err_Handler
    wscControl.SendData "QUIT" & vbCrLf
    'Debug.Print "QUIT"
    RaiseEvent ReplyMessage("QUIT" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        '
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        '
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_SERVICE_CLOSING_CONTROL_CONNECTION Then
        ProcessQUITCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessQUITCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessQUITCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessABORCommand() As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessABORCommand_Err_Handler
    wscControl.SendData "ABOR" & vbCrLf
    'Debug.Print "ABOR"
    RaiseEvent ReplyMessage("ABOR" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        '
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        '
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = "226" & vbCrLf
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = 426 Then
        ProcessABORCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessABORCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessABORCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Function CancelTransfer() As Boolean
    m_bBusy = True
    If ProcessABORCommand Then
        CancelTransfer = True
    End If
    If m_bTransferInProgress Or m_bUploadFile Then
        Close m_intLocalFileID
        m_strDataBuffer = ""
        m_lDownloadedBytes = 0
        m_lUploadedBytes = 0
        m_bTransferInProgress = False
        m_bUploadFile = False
    End If
    m_bFileIsOpened = False
    m_objTimeOut.StopTimer
    '    wscData.Close
    m_bBusy = False
End Function

Public Function SetParentAsCurrentDirectory() As Boolean
    m_bBusy = True
    SetParentAsCurrentDirectory = ProcessCDUPCommand
    m_bBusy = False
End Function

Private Function ProcessCDUPCommand() As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessCDUPCommand_Err_Handler
    wscControl.SendData "CDUP" & vbCrLf
    'Debug.Print "CDUP"
    RaiseEvent ReplyMessage("CDUP" & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        '
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        '
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessCDUPCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessCDUPCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessCDUPCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Private Function ProcessCWDCommand(ByRef sNewDir As String) As Boolean
    Dim strResponse                     As String
    Dim strData                         As String
    On Error GoTo ProcessCWDCommand_Err_Handler
    wscControl.SendData "CWD " & sNewDir & vbCrLf
    RaiseEvent ReplyMessage("CWD " & sNewDir & vbCrLf)
    m_objTimeOut.StartTimer
    Do
        DoEvents2
        If m_objTimeOut.Timeout Then
            m_LastError = ERROR_FTP_USER_TIMEOUT
            Exit Do
        End If
        If InStr(1, m_strWinsockBuffer, vbCrLf) > 0 Then
            strData = m_strWinsockBuffer
            m_strWinsockBuffer = ""
            Exit Do
        End If
    Loop
    m_objTimeOut.StopTimer
    If GetResponseCode(strData) = FTP_RESPONSE_REQUESTED_FILE_ACTION_OK_COMPLETED Then
        ProcessCWDCommand = True
    Else
        ProcessFtpResponse GetResponseCode(strData)
    End If
Exit_Label:
    Exit Function
ProcessCWDCommand_Err_Handler:
    If Not ProcessWinsockError(Err.Number, Err.Description) Then
        Err.Raise vbObjectError + 1000 + Err.Number, "CFtpConnection.ProcessCWDCommand", Err.Description
    End If
    GoTo Exit_Label
End Function

Public Function GetFtpErrorDescription() As String
    GetFtpErrorDescription = m_strLastErrorDesc
End Function

Public Function CloseConnection() As Boolean
    m_bBusy = True
    If m_bTransferInProgress Or m_bUploadFile Then
        m_LastError = ERROR_FTP_USER_TRANSFER_IN_PROGRESS
        m_strLastErrorDesc = "Can't close control connection. Transfer in progress."
    Else
        CloseConnection = ProcessQUITCommand
        wscData.Close
        wscControl.Close
    End If
    m_bBusy = False
End Function
