Attribute VB_Name = "MFtpSupport"
Option Explicit
Public p_intCounter                     As Integer
Public g_intPort                        As Long

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    p_intCounter = p_intCounter + 1
End Sub
