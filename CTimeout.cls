VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private m_lngTimerID                    As Long
Private m_intTimeoutValue               As Integer

Public Sub StartTimer()
    If m_lngTimerID > 0 Then
        KillTimer 0, m_lngTimerID
    End If
    MFtpSupport.p_intCounter = 0
    m_lngTimerID = SetTimer(0, 0, 1000, AddressOf MFtpSupport.TimerProc)
End Sub

Public Sub StopTimer()
    MFtpSupport.p_intCounter = 0
    KillTimer 0, m_lngTimerID
End Sub

Public Property Let TimeoutValue(NewValue As Integer)
    m_intTimeoutValue = NewValue
End Property

Public Property Get TimeoutValue() As Integer
    TimeoutValue = m_intTimeoutValue
End Property

Public Property Get Timeout() As Boolean
    If MFtpSupport.p_intCounter > m_intTimeoutValue Then
        Timeout = True
        Call StopTimer
    Else
        Timeout = False
    End If
End Property

Public Sub Reset()
    MFtpSupport.p_intCounter = 0
End Sub

Private Sub Class_Terminate()
    Call StopTimer
End Sub
