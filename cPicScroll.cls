VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: SetUpScrollBars
' PURPOSE:        Initializes the Max, smallchange and largechange properties
'                 of the Vertical and Horizontal ScrolBars.
' RETURNS:        Nothing
'
'---------------------------------------------------------------------------
' DATE:  October,22 99
' TIME:  10:10
'
' Procedure Documenter : gh0ul99
'---------------------------------------------------------------------------


Sub SetUpScrollBars(H_SB As HScrollBar, V_SB As VScrollBar, lVer_MaxLen As Long, iVer_SmallChange As Integer, lVer_LargeChange As Long, _
                    lHor_MaxLen As Long, iHor_SmallChange As Integer, lHor_LargeChange As Long)
    '
    ' init Vertical scrollbar
    V_SB.Max = lVer_MaxLen
    V_SB.SmallChange = iVer_SmallChange
    V_SB.LargeChange = lVer_LargeChange
    'init horizintal scrollbar
    H_SB.Max = lHor_MaxLen
    H_SB.SmallChange = iHor_SmallChange
    H_SB.LargeChange = lHor_LargeChange
End Sub

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: MoveV
' PURPOSE:        Moves the inner picturebox vertically according
'                 to the value of the neg vertical scrollbar
' RETURNS:        Nothing
'
'---------------------------------------------------------------------------
' DATE:  October,22 99
' TIME:  10:12
'
' Procedure Documenter : gh0ul99
'---------------------------------------------------------------------------

Sub MoveV(V_SB As VScrollBar, pic As PictureBox)
    pic.Top = -V_SB.Value
End Sub

'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: MoveH
' PURPOSE:        Moves the inner picturebox Horizontally according
'                 to the value of the neg Horizantal scrollbar
' RETURNS:        Nothing
'
'---------------------------------------------------------------------------
' DATE:  October,22 99
' TIME:  10:16
'
' Procedure Documenter : gh0ul99
'---------------------------------------------------------------------------

Sub MoveH(H_SB As HScrollBar, pic As PictureBox)
    pic.Left = -H_SB.Value
End Sub
