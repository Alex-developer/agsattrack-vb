VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWaitCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private m_nPointer As MousePointerConstants

Public Sub SetCursor(Optional nPointer As MousePointerConstants = vbHourglass)
  Screen.MousePointer = nPointer
End Sub

Public Sub Restore()
  Screen.MousePointer = m_nPointer
End Sub

Private Sub Class_Initialize()
  m_nPointer = Screen.MousePointer
End Sub

Private Sub Class_Terminate()
  Restore
End Sub
