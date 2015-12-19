Attribute VB_Name = "ft847_Control"
Option Explicit

Public Sub SetCommPort(Port As Integer, Speed As String)
  '***********************************************************
  ' set up com2 for talking to the radio
  '***********************************************************

  ' close comm port if already open
  If UserControl.MSComm1.PortOpen = True Then
    UserControl.MSComm1.PortOpen = False
  End If

  ' Communications port settings.
  UserControl.MSComm1.CommPort = Port             ' port number
  UserControl.MSComm1.Settings = Speed            ' match the speed on your transceiver
  ' by using Menu 37, CAT function
  UserControl.MSComm1.InputMode = comInputModeBinary
  UserControl.MSComm1.InputLen = 0

  ' Open the communications port.
  'On Error Resume Next
  UserControl.MSComm1.PortOpen = True
  If Err Then
    MsgBox "COM Port not available. Change the CommPort to another port."
    Exit Sub
  End If

  ' Clear the thresholds
  UserControl.MSComm1.RThreshold = 0
  UserControl.MSComm1.SThreshold = 0

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
    UserControl.pbarSmeter.Value = Reading        ' show the reading
  End If

End Sub

Public Sub ReadFreq()
  '***********************************************************
  ' Read frequency and mode procedure - result returned should be
  ' 4 bytes of frequency and 1 byte of mode information
  '***********************************************************

  ' request data
  ClearCmd                                    ' clear last command
  SendCmd Read_Freq_Mode_Main                 ' request the frequency

  ' look for response
  Receive 5                                   ' read response

  If UBound(RxData) = 4 Then                  ' make sure got some data
    ' decode the result
    Frequency = Format(Hex(RxData(0)), "00")
    Frequency = Frequency + Format(Hex(RxData(1)), "00")
    Frequency = Frequency + Format(Hex(RxData(2)), "00")
    Frequency = Frequency + Format(Hex(RxData(3)), "00")
    FreqLong = Val(Frequency)               ' save frequency in long format

    ' format for display
    ShowFrequency = Left(Frequency, 3) + "." + Mid(Frequency, 4, 3)
    If DecimalsOn Then
      ShowFrequency = ShowFrequency + "." + Right(Frequency, 2)
    End If

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
    UserControl.txtFreq.Text = Trim(ShowFrequency)

    ' decode mode
    Select Case RxData(4)                   ' final byte is the mode
      Case 0
        Mode = "LSB"
      Case 1
        Mode = "USB"
      Case 2
        Mode = "CW"
      Case 3
        Mode = "CW-R"
      Case 4
        Mode = "AM"
      Case 8
        Mode = "FM"
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

    ' show mode
    UserControl.txtMode.Text = Trim(Mode)

  End If

End Sub

Public Sub ClearCmd()
  '***********************************************************
  ' this procedure clears the first 4 bytes of the command
  '***********************************************************

  For i = 0 To 3
    Cmd(i) = 0
  Next

End Sub

Public Sub SendCmd(Command As Byte)
  '***********************************************************
  ' this is a generalised procedure to build a 5 byte command
  ' string and send it to the serial port.
  ' Usage - SendCmd <command constant>
  '***********************************************************

  ' clear the input buffer
  UserControl.MSComm1.InBufferCount = 0

  ' if com port open -
  If UserControl.MSComm1.PortOpen = True Then
    ' get the command
    Cmd(4) = Command

    ' Send command to transceiver
    UserControl.MSComm1.Output = Cmd()            ' send the command

    ' Wait for all the data to be sent.
    Do
      Ret = DoEvents()
    Loop Until UserControl.MSComm1.OutBufferCount = 0
  End If

End Sub

Public Sub Receive(Bytes As Long)
  '***********************************************************
  ' This procedure receives 'n' bytes from the serial port and
  ' places it into the variant RxData.
  ' Usage - Receive <n> where <n> is the number of bytes
  '***********************************************************

  ' if com port open -
  If UserControl.MSComm1.PortOpen = True Then
    UserControl.MSComm1.InputLen = Bytes          ' set no. of bytes
    Do
      Ret = DoEvents()
      Loop Until UserControl.MSComm1.InBufferCount = Bytes Or _
      SettingsRequested            ' wait for them or a SettingsRequested click

      ' the SettingsRequested check is to give the system a 'kick' if
      ' we have got out of sync with commands sent and responses received.
      ' A click on the Settings button will take the program out of a
      ' receive loop

      RxData = UserControl.MSComm1.Input            ' read <n> bytes
    End If

  End Sub

