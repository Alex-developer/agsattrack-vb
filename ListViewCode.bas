Attribute VB_Name = "ListViewCode"
Option Explicit

Public Declare Function SendMessageLong Lib "user32" _
    Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function SendMessageAny Lib "user32" _
    Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
       
Public Const LVM_FIRST = &H1000
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT = (LVM_FIRST + 45)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)

Public Const LVS_EX_GRIDLINES = &H1
Public Const LVS_EX_CHECKBOXES = &H4
Public Const LVS_EX_FULLROWSELECT = &H20 'applies to report mode only

Public Const LVIF_STATE = &H8
Public Const LVIS_STATEIMAGEMASK As Long = &HF000
 
Public Const SW_NORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Public Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

    
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal Hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Dim fPath As String

Sub DisplayCheckBoxes(lstListView As ListView, bState As Boolean)

  Call SendMessageLong(lstListView.Hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, bState)

End Sub

Sub DisplayGridLines(lstListView As ListView, bState As Boolean)
  
  Call SendMessageLong(lstListView.Hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, bState)

End Sub

Sub FullRowSelect(lstListView As ListView, bState As Boolean)

  Call SendMessageLong(lstListView.Hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, bState)

End Sub

Sub SetCheck(lstListView As ListView, lItemIndex As Long, bState As Boolean)

  Dim LV As LVITEM
  Dim r As Long
  Dim Hwnd As Long
  
  Hwnd = lstListView.Hwnd
  
  With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
  End With
    
  Call SendMessageAny(Hwnd, LVM_SETITEMSTATE, lItemIndex, LV)

End Sub


 Sub SetCheckAllItems(lstListView As ListView, bState As Boolean)

   Dim LV As LVITEM
   Dim lvCount As Long
   Dim lvIndex As Long
   Dim lvState As Long
   Dim r As Long
   
   
   lvState = IIf(bState, &H2000, &H1000)
   
   lvCount = lstListView.ListItems.Count - 1
   
   Do
         
      With LV
         .mask = LVIF_STATE
         .state = lvState
         .stateMask = LVIS_STATEIMAGEMASK
      End With
      
      Call SendMessageAny(lstListView.Hwnd, LVM_SETITEMSTATE, lvIndex, LV)
      lvIndex = lvIndex + 1
   
   Loop Until lvIndex > lvCount
  
  
End Sub


 Sub SetCheckInvertAll(lstListView As ListView, bState As Boolean)


   Dim LV As LVITEM
   Dim r As Long
   Dim lvCount As Long
   Dim lvIndex As Long
   
   lvCount = lstListView.ListItems.Count - 1
   
   Do
         
      r = SendMessageLong(lstListView.Hwnd, LVM_GETITEMSTATE, lvIndex, LVIS_STATEIMAGEMASK)
      
      With LV
         .mask = LVIF_STATE
         .stateMask = LVIS_STATEIMAGEMASK
      
         If r And &H2000& Then
              'its checked, so set the state
              'to 'unchecked'
               .state = &H1000
         Else: .state = &H2000
         End If
         
      End With
      
      Call SendMessageAny(lstListView.Hwnd, LVM_SETITEMSTATE, lvIndex, LV)
      lvIndex = lvIndex + 1
   
   Loop Until lvIndex > lvCount

End Sub

