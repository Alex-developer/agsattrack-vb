Attribute VB_Name = "ft847_Commands"
Option Explicit

'***********************************************************
' Command Codes for the FT847
' Check your manual for other models and adapt accordingly
'***********************************************************

' Set CAT Mode On/Off
Public Const Set_CAT_On = &H0              ' Set CAT mode On
Public Const Set_CAT_Off = &H80            ' Set CAT mode Off

' Set PTT On/Off
Public Const Set_PTT_On = &H8              ' Set PTT On
Public Const Set_PTT_Off = &H88            ' Set PTT Off

' Set Satellite On/Off
Public Const Set_Satellite_On = &H4E       ' Set Satellite mode On
Public Const Set_Satellite_Off = &H8E      ' Set Satellite mode Off

' Set Frequency
Public Const Set_Freq_Main = &H1           ' Set Main VFO frequency
Public Const Set_Freq_Sat_RX = &H11        ' Set SAT RX frequency
Public Const Set_Freq_Sat_TX = &H21        ' Set SAT TX frequency

' Set Operating Mode
Public Const Set_Mode_LSB = &H0            ' Set Mode LSB
Public Const Set_Mode_USB = &H1            ' Set Mode USB
Public Const Set_Mode_CW = &H2             ' Set Mode CW
Public Const Set_Mode_CWR = &H3            ' Set Mode CW-R
Public Const Set_Mode_AM = &H4             ' Set Mode AM
Public Const Set_Mode_FM = &H8             ' Set Mode FM
Public Const Set_Mode_CWN = &H82           ' Set Mode CW(N)
Public Const Set_Mode_CWRN = &H83          ' Set Mode CW-R(N)
Public Const Set_Mode_AMN = &H84           ' Set Mode AM(N)
Public Const Set_Mode_FMN = &H88           ' Set Mode FM(N)
Public Const Set_Mode_Main = &H7           ' Set Mode on Main VFO
Public Const Set_Mode_Sat_RX = &H17        ' Set Mode on SAT RX
Public Const Set_Mode_Sat_TX = &H27        ' Set Mode on SAT TX

' Read Frequency & Mode
Public Const Read_Freq_Mode_Main = &H3     ' Read Frequency & mode on main VFO
Public Const Read_Freq_Mode_Sat_RX = &H13  ' Read Frequency & Mode on Sat RX VFO
Public Const Read_Freq_Mode_Sat_TX = &H23  ' Read Frequency & Mode on Sat TX VFO

' Set CTCSS/DCS mode
Public Const Set_DCS_On = &HA              ' Set DCS On
Public Const Set_CTSS_ENC_DEC_On = &H2A    ' Set CTCSS ENC/DEC On
Public Const Set_CTCSS_ENC_On = &H4A       ' Set CTCSS ENC On
Public Const Set_CTCSS_DCS_On = &H8A       ' Set CTCSS/DCS On
Public Const Set_CTCSS_Main = &HA          ' Set Main VFO
Public Const Set_CTCSS_SAT_RX = &H1A       ' Set SAT RX VFO
Public Const Set_CTCSS_SAR_TX = &H2A       ' Set SAT TX VFO

' CTCSS Frequency
Public Const Set_CTCSS_Frequency_Main = &HB ' Set Main VFO
Public Const Set_CTCSS_Frequency_Sat_RX = &H1B ' Set SAT RX VFO
Public Const Set_CTCSS_Frequency_Sat_TX = &H2B ' Set SAT TX VFO

' DCS Code
Public Const Set_DCS_Main = &HC            ' Set Main VFO
Public Const Set_DCS_Sat_RX = &H1C         ' Set SAT RX VFO
Public Const Set_DCS_Sat_TX = &H2C         ' Set SAT TX VFO

' Repeater Shift
Public Const Set_Repeater_Shift = &H9          ' Set  repeater shift opcode
Public Const Set_Repeater_Shift_Minus = &H9    ' Set repeater shift minus
Public Const Set_Repeater_Shift_Plus = &H49    ' Set repeater shift plus
Public Const Set_Repeater_Shift_Simplex = &H89 ' Set repeater shift simplex

' Repeater Offset
Public Const Set_Repeater_Offset = &HF9    ' Set  repeater shift opcode

' Receiver Status
Public Const Read_Receiver_Status = &HE7   ' Read Receiver Status

' Transmit Status
Public Const Read_Transmit_Status = &HF7   ' Read Transmit Status

' end of constants for FT847 command codes
'***********************************************************

