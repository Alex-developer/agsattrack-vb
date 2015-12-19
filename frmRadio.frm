VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5003B5C3-1891-11D1-B6CF-0000C02DDDED}#1.0#0"; "CBOTKNOB.OCX"
Begin VB.Form frmRadio 
   Caption         =   "FT 847"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2700
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2190
      Top             =   3810
   End
   Begin VB.Frame Frame3 
      Caption         =   " VF0 B "
      Height          =   2295
      Left            =   990
      TabIndex        =   10
      Top             =   1950
      Width           =   915
      Begin VB.OptionButton optModeB 
         Caption         =   "LSB"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optModeB 
         Caption         =   "USB"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   564
         Width           =   675
      End
      Begin VB.OptionButton optModeB 
         Caption         =   "CW"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   888
         Width           =   615
      End
      Begin VB.OptionButton optModeB 
         Caption         =   "FM"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1212
         Width           =   615
      End
      Begin VB.OptionButton optModeB 
         Caption         =   "AM"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1536
         Width           =   615
      End
      Begin VB.OptionButton optModeB 
         Caption         =   "CWR"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1860
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " VFO "
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   3195
      Begin MSComctlLib.ProgressBar pSMeter 
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1410
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Max             =   31
      End
      Begin VB.TextBox txtVFO2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   60
         TabIndex        =   3
         Text            =   "000.000.000"
         Top             =   810
         Width           =   2085
      End
      Begin VB.TextBox txtVFO1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   2
         Text            =   "000.000.000"
         Top             =   180
         Width           =   2085
      End
      Begin CBOTKNOBLib.CbotKnob VFOKnob 
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   17
         Top             =   180
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   873
         _StockProps     =   29
         FineAdjustWidth =   6
      End
      Begin CBOTKNOBLib.CbotKnob VFOKnob 
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   18
         Top             =   810
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   873
         _StockProps     =   29
         FineAdjustWidth =   6
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " VF0 A "
      Height          =   2295
      Left            =   30
      TabIndex        =   0
      Top             =   1950
      Width           =   915
      Begin VB.OptionButton optModeA 
         Caption         =   "CWR"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   735
      End
      Begin VB.OptionButton optModeA 
         Caption         =   "AM"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1536
         Width           =   615
      End
      Begin VB.OptionButton optModeA 
         Caption         =   "FM"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1212
         Width           =   615
      End
      Begin VB.OptionButton optModeA 
         Caption         =   "CW"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   888
         Width           =   615
      End
      Begin VB.OptionButton optModeA 
         Caption         =   "USB"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   564
         Width           =   675
      End
      Begin VB.OptionButton optModeA 
         Caption         =   "LSB"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim bSatMode As Boolean

Dim DecimalsOn As Boolean            ' true=show decimals on display
Dim ShowFrequency As String          ' displayed frequency
Dim Frequency As String              ' decoded frequency with leading zeros
Dim Ret As Long                      ' general return code variable
Dim Cmd(4) As Byte                   ' 5 byte command string for FT847
Dim FreqString As String             ' work area for frequency
Dim VFO1FreqLong As Long             ' frequency in long format
Dim VFO2FreqLong As Long             ' frequency in long format
Dim FreqLongTemp As Long             ' temporary frequency in long format
Dim FreqArray(7) As Integer          ' work array for frequency
Dim RxData() As Byte                 ' returned data from FT847
Dim StopAction As Boolean            ' true=stop action
Dim Mode As String                   ' work area for mode
Dim CommSpeed As String              ' speed setting for Comm port
Dim CommPort As Integer              ' Comm port number
Dim FreqLowChange As Boolean         ' flags a change of value to sldLow
Dim FreqMidChange As Boolean         ' flags a change of value to sldMid
Dim FreqHighChange As Boolean        ' flags a change of value to sldHigh
Dim FreqLowBeingChanged As Boolean   ' flags sldLow being changed
Dim FreqMidBeingChanged As Boolean   ' flags sldMid being changed
Dim FreqHighBeingChanged As Boolean  ' flags sldHigh being changed
Dim FreqMhzPlus As Boolean           ' flags +1 Mhz
Dim FreqMhzMinus As Boolean          ' flags -1 Mhz
Dim FrequencyError As Boolean        ' flags a frequency out of band
Dim MemName() As String              ' memory element name
Dim MemFreq() As String              ' memory frequency in Khz
Dim MemMode() As String              ' memory mode
Dim MemNumberOf As Integer           ' number of memories
Dim SettingsRequested As Boolean     ' flag settings form requested

Private Sub Form_Load()
    SETtopmostwindow Me, True
    SetCommPort 1, "57600,N,8,1"

    MSComm1.RThreshold = 0
    MSComm1.SThreshold = 0

    ClearCmd                            ' clear last command
    SendCmd Set_CAT_On                  ' take control
    Timer1.Enabled = True               ' start the timer
    bSatMode = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SETtopmostwindow Me, False
      Timer1.Enabled = False              ' stop the timer
      ClearCmd                            ' clear last command
      SendCmd Set_CAT_Off                 ' release control
      MSComm1.PortOpen = False
End Sub

Public Function SetSatelliteMode(bFlag As Boolean)
  
  Timer1.Enabled = False
  
  If Not bSatMode Then
    If MSComm1.PortOpen = True Then
      ClearCmd                            ' clear last command
      SendCmd Set_Satellite_On            ' Turn Satellite mode on
      bSatMode = True
    End If
  Else
    If MSComm1.PortOpen = True Then
      ClearCmd                            ' clear last command
      SendCmd Set_Satellite_Off           ' Turn Satellite mode on
      bSatMode = False
    End If
  End If

  Timer1.Enabled = True
  
End Function
Public Sub SetCommPort(Port As Integer, Speed As String)
  '***********************************************************
  ' set up com2 for talking to the radio
  '***********************************************************

  ' close comm port if already open
  If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
  End If

  ' Communications port settings.
  MSComm1.CommPort = Port             ' port number
  MSComm1.Settings = Speed            ' match the speed on your transceiver
  ' by using Menu 37, CAT function
  MSComm1.InputMode = comInputModeBinary
  MSComm1.InputLen = 0

  ' Open the communications port.
  'On Error Resume Next
  MSComm1.PortOpen = True
  If Err Then
    MsgBox "COM Port not available. Change the CommPort to another port."
    Exit Sub
  End If

  ' Clear the thresholds
  MSComm1.RThreshold = 0
  MSComm1.SThreshold = 0

End Sub

Public Sub MakeFreqKhz(FreqWanted As Long)
  '***********************************************************
  ' Convert a 'long' frequency in Khz to the Cmd() byte array to send to
  ' the transceiver.
  ' Usage
  '       MakeFreqKhz <frequency in Khz>
  '***********************************************************

  ' validate proposed frequency
  If FrequencyValidate(FreqWanted * 100) = False Then Exit Sub

  ' convert long value to a string of 8 chars with leading zeros
  FreqString = Format(FreqWanted * 100, "00000000")

  ' turn into array of numbers wanted by subtracting 48
  ' from the ASCII value of each byte
  For i = 0 To 7
    FreqArray(i) = Asc(Mid(FreqString, i + 1, 1)) - 48
  Next

  ' now add each pair of numbers for each byte of Cmd()
  For i = 0 To 3
    Cmd(i) = CByte((FreqArray(i * 2) * 16) + FreqArray(i * 2 + 1))
  Next

End Sub

Public Sub MakeFreqHz(FreqWanted As Long)
  '***********************************************************
  ' Convert a 'long' frequency in Hz to the Cmd() byte array to send to
  ' the transceiver.
  ' Usage
  '       MakeFreqHz <frequency in 10Hz>
  '***********************************************************

  ' validate proposed frequency
  If FrequencyValidate(FreqWanted) = False Then Exit Sub

  ' convert long value to a string of 8 chars with leading zeros
  FreqString = Format(FreqWanted, "00000000")

  ' turn into array of numbers wanted by subtracting 48
  ' from the ASCII value of each byte
  For i = 0 To 7
    FreqArray(i) = Asc(Mid(FreqString, i + 1, 1)) - 48
  Next

  ' now add each pair of numbers for each byte of Cmd()
  For i = 0 To 3
    Cmd(i) = CByte((FreqArray(i * 2) * 16) + FreqArray(i * 2 + 1))
  Next

End Sub

Public Function FrequencyValidate(FrequencyToCheck As Long)
  '***********************************************************
  ' validate a frequency in 10Hz is within transceiver range
  '***********************************************************

  Select Case FrequencyToCheck
    Case 10000 To 7600000
    Case 10800000 To 17400000
    Case 42000000 To 51200000
    Case Else
      FrequencyValidate = False           ' return False
      FrequencyError = True               ' flag bad frequency
      Exit Function
  End Select

  FrequencyValidate = True                    ' return True
  FrequencyError = False                      ' flag not bad frequenncy

End Function

Public Sub ReadSmeter()
  '***********************************************************
  ' Read the S meter procedure
  '***********************************************************

  Dim Reading As Integer

  ClearCmd                                    ' clear last command
  SendCmd Read_Receiver_Status                ' request S-meter reading

  Receive 1                                   ' read response
  If UBound(RxData) = 0 Then                  ' if we got something back
    Reading = RxData(0) Mod 32              ' extract S-meter part
    Me.pSMeter.Value = Reading        ' show the reading
  End If

End Sub
Private Sub ReadVFO()

  Call ErrorHandlerStartProcedure("frmRadio", "ReadVFO")
  On Error GoTo ERROR_ReadVFO
  Call ErrorHandlerParameter("Entry frmRadio | ReadVFO", vbAbortRetryIgnore + vbCritical + vbDefaultButton1, gbViewMessage)

  Dim strFrequency As Variant
  
  If bSatMode Then
    strFrequency = ReadFreq(Read_Freq_Mode_Sat_RX)
    If UBound(strFrequency) > 0 Then
      frmRadio.txtVFO1.Text = Trim(strFrequency(0))
      frmRadio.optModeA(strFrequency(1)).Value = True
    Else
      frmRadio.txtVFO1.Text = "OO.OOO.OO"
    End If
    strFrequency = ReadFreq(Read_Freq_Mode_Sat_TX)
    If UBound(strFrequency) > 0 Then
      frmRadio.txtVFO2.Text = Trim(strFrequency(0))
      frmRadio.optModeB(strFrequency(1)).Value = True
    Else
      frmRadio.txtVFO2.Text = "OO.OOO.OO"
    End If
  Else
    strFrequency = ReadFreq(Read_Freq_Mode_Main)
    If UBound(strFrequency) > 0 Then
      frmRadio.txtVFO1.Text = Trim(strFrequency(0))
      frmRadio.optModeA(strFrequency(1)).Value = True
      frmRadio.txtVFO2.Text = "OO.OOO.OO"
    Else
      frmRadio.txtVFO1.Text = "OO.OOO.OO"
    End If
  End If

  GoTo END_ReadVFO
ERROR_ReadVFO:
  Select Case TreatErrorHandler()
      Case 0: Resume
      Case 1: Resume Next
      Case 2: Resume END_ReadVFO ' *** Abort
      Case 3: Resume END_ReadVFO ' *** Abort
      Case 4: Resume           ' *** Retry
      Case Else: Resume Next   ' *** Ignore
  End Select
END_ReadVFO:
  Call ErrorHandlerEnd("ReadVFO")

End Sub

Private Function ReadFreq(bType As Byte) As Variant
  Dim vResults(10) As Variant
  '***********************************************************
  ' Read frequency and mode procedure - result returned should be
  ' 4 bytes of frequency and 1 byte of mode information
  '***********************************************************

  ' request data
  ClearCmd                                    ' clear last command
  SendCmd bType               ' request the frequency
  
  ' look for response
  Receive 5                                   ' read response
  
  vResults(0) = ""
  
  If UBound(RxData) = 4 Then                  ' make sure got some data
    ' decode the result
    Frequency = Format(Hex(RxData(0)), "00")
    Frequency = Frequency + Format(Hex(RxData(1)), "00")
    Frequency = Frequency + Format(Hex(RxData(2)), "00")
    Frequency = Frequency + Format(Hex(RxData(3)), "00")
    Select Case bType
      Case Read_Freq_Mode_Main
        VFO1FreqLong = Val(Frequency)               ' save frequency in long format
      Case Read_Freq_Mode_Sat_RX
        VFO1FreqLong = Val(Frequency)               ' save frequency in long format
      Case Read_Freq_Mode_Sat_TX
        VFO2FreqLong = Val(Frequency)               ' save frequency in long format
    End Select

    ' format for display
    ShowFrequency = Left(Frequency, 3) + "." + Mid(Frequency, 4, 3)
  '  If DecimalsOn Then
      ShowFrequency = ShowFrequency + "." + Right(Frequency, 2)
    'End If

    ' clear out leading zeros
    For i = 1 To 3
      If Mid(ShowFrequency, i, 1) = "0" Then
        Mid(ShowFrequency, i, 1) = " "
      End If
      If Mid(ShowFrequency, i + 1, 1) <> "0" Then
        Exit For
      End If
    Next

    ' show on screem
    'txtFreq.Text = Trim(ShowFrequency)
 '   SatDetails.txtVFO1.Text = Trim(ShowFrequency)
  vResults(0) = ShowFrequency
    ' decode mode
    Select Case RxData(4)                   ' final byte is the mode
      Case 0
        Mode = 0
      Case 1
        Mode = 1
      Case 2
        Mode = 2
      Case 3
        Mode = 5
      Case 4
        Mode = 4
      Case 8
        Mode = 3
      Case 82
        Mode = "CW(N)"
      Case 83
        Mode = "CW(N)-R"
      Case 84
        Mode = "AM(N)"
      Case 88
        Mode = "FM(N)"
      Case Else
        Mode = "???"                    ' unknown mode
    End Select
vResults(1) = Mode
'vResults(2) = Mode
  ReadFreq = vResults
    ' show mode
    'txtMode.Text = Trim(Mode)

  End If

End Function

Public Sub ClearCmd()
  '***********************************************************
  ' this procedure clears the first 4 bytes of the command
  '***********************************************************

  For i = 0 To 3
    Cmd(i) = 0
  Next

End Sub

Public Sub SendCmd(Command As Byte)
  Dim bInTx As Boolean
  '***********************************************************
  ' this is a generalised procedure to build a 5 byte command
  ' string and send it to the serial port.
  ' Usage - SendCmd <command constant>
  '***********************************************************
  ' clear the input buffer
  MSComm1.InBufferCount = 0

  If Not bInTx Then
    ' if com port open -
    If MSComm1.PortOpen = True Then
      ' get the command
      Cmd(4) = Command

      ' Send command to transceiver
      MSComm1.Output = Cmd()            ' send the command

      ' Wait for all the data to be sent.
      bInTx = True
      Do
        Ret = DoEvents()
      Loop Until MSComm1.OutBufferCount = 0
      bInTx = False
    End If
  End If
End Sub

Public Sub Receive(Bytes As Long)
  '***********************************************************
  ' This procedure receives 'n' bytes from the serial port and
  ' places it into the variant RxData.
  ' Usage - Receive <n> where <n> is the number of bytes
  '***********************************************************
  Dim start As Variant
  Dim bInRx As Boolean

  If Not bInRx Then
    ' if com port open -
    If MSComm1.PortOpen = True Then
      MSComm1.InputLen = Bytes          ' set no. of bytes
      start = Now
      bInRx = True
      Do
        Ret = DoEvents()
      Loop Until MSComm1.InBufferCount = Bytes Or DateDiff("s", start, Now) > 1
      bInRx = False
      ' the SettingsRequested check is to give the system a 'kick' if
      ' we have got out of sync with commands sent and responses received.
      ' A click on the Settings button will take the program out of a
      ' receive loop

      RxData = MSComm1.Input            ' read <n> bytes
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  ReadVFO
  ReadSmeter
End Sub

Private Sub VFOKnob_PositionChanged(Index As Integer, ByVal delta As Integer)
  Dim nInc As Integer

  Timer1.Enabled = False
  
  If delta > 0 Then
    nInc = 1
  Else
    nInc = -1
  End If

  If Index = 0 Then
    MakeFreqHz VFO1FreqLong + nInc
    If FrequencyError Then Exit Sub     ' ignore bad frequencies
    VFO1FreqLong = VFO1FreqLong + nInc
    If bSatMode Then
      SendCmd Set_Freq_Sat_RX             ' change frequency
    Else
      SendCmd Set_Freq_Main               ' change frequency
    End If
  Else
    If bSatMode Then
      MakeFreqHz VFO2FreqLong + nInc
      If FrequencyError Then Exit Sub     ' ignore bad frequencies
      VFO2FreqLong = VFO2FreqLong + nInc
      SendCmd Set_Freq_Sat_TX           ' change frequency
    End If
  End If
  Timer1.Enabled = True
End Sub

Public Sub SetVFO(bMainVFO As Boolean, sFrequency As Double)
If bMainVFO Then
    MakeFreqHz CLng(sFrequency)
    If FrequencyError Then Exit Sub     ' ignore bad frequencies
    'VFO1FreqLong = VFO1FreqLong + nInc
    If bSatMode Then
      SendCmd Set_Freq_Sat_RX             ' change frequency
    Else
      SendCmd Set_Freq_Main               ' change frequency
    End If
Else
End If
End Sub

