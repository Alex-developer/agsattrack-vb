VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SatDetails 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Satellite Details"
   ClientHeight    =   3570
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ListView lstDetails 
      Height          =   2685
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4736
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuChooserMenu 
      Caption         =   "Chooser Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuFCAdd 
         Caption         =   "Add Field"
         Begin VB.Menu menuBlank 
            Caption         =   "Blank Line"
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 1"
            Index           =   1
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 2"
            Index           =   2
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 3"
            Index           =   3
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 4"
            Index           =   4
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 5"
            Index           =   5
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 6"
            Index           =   6
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 7"
            Index           =   7
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 8"
            Index           =   8
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 9"
            Index           =   9
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 10"
            Index           =   10
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 11"
            Index           =   11
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 12"
            Index           =   12
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 13"
            Index           =   13
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 14"
            Index           =   14
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 15"
            Index           =   15
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 16"
            Index           =   16
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 17"
            Index           =   17
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 18"
            Index           =   18
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 19"
            Index           =   19
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field 20"
            Index           =   20
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field21"
            Index           =   21
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field22"
            Index           =   22
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field23"
            Index           =   23
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field 24"
            Index           =   24
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field25"
            Index           =   25
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field26"
            Index           =   26
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field27"
            Index           =   27
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field28"
            Index           =   28
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field29"
            Index           =   29
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field30"
            Index           =   30
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field31"
            Index           =   31
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "Field32"
            Index           =   32
         End
         Begin VB.Menu mnuFCField 
            Caption         =   "field33"
            Index           =   33
         End
      End
      Begin VB.Menu mnuFCDelete 
         Caption         =   "Delete Field"
      End
   End
End
Attribute VB_Name = "SatDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strText(-1 To 100) As String

Private Sub Form_Load()
  Dim i As Integer
  
  SetupText
  
  For i = 1 To 33
    Me.mnuFCField(i).Caption = strText(i)
  Next i
  
  SETtopmostwindow Me, False
  Form_Resize
  Rebuild_List
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  
  Me.lstDetails.Left = 0
  Me.lstDetails.Top = 0
  Me.lstDetails.Width = Me.ScaleWidth
  Me.lstDetails.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If Not bUnloading Then
'    Me.Tag = ""
'    Me.Hide
'    Cancel = 1
'  End If
  nSatDetailsTag = -10
  bSatDetailsVisible = False
End Sub

Private Sub lstDetails_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  If Button And vbRightButton Then
    PopupMenu mnuChooserMenu
  End If
  
End Sub

Sub Rebuild_List()
  Dim i As Integer
  Dim lItem As ListItem
  
  Me.lstDetails.ListItems.Clear
  For i = 1 To 100
    If nFields(i) <> 0 Then
      Set lItem = Me.lstDetails.ListItems.Add
      lItem.Text = strText(nFields(i))
    Else
      Exit For
    End If
  Next i
  
End Sub


Private Sub SetupText()
  Erase strText
  
  strText(-1) = ""
  strText(1) = "NORAD Id"
  strText(2) = "Sat Name"
  strText(3) = "Azimuth       (°)"
  strText(4) = "Elevation     (°)"
  strText(5) = "Lat           (°)"
  strText(6) = "Long          (°)"
  strText(7) = "Range         (Km)"
  strText(8) = "Orbit"
  strText(9) = "Range Rate"
  strText(10) = "Uplink       (MHz)"
  strText(11) = "Downlink     (MHz)"
  strText(12) = "Uplink Tx    (Mhz)"
  strText(13) = "Downlink Rx  (MHz)"
  strText(14) = "Doppler      (KHz)"
  strText(15) = "path Loss    (Db)"
  strText(16) = "Max Range    (Km)"
  strText(17) = "Status"
  strText(18) = "RS"
  strText(19) = "Squint Angle (°)"
  strText(20) = "Drag            "
  strText(21) = "Set             "
  strText(22) = "Epoch           "
  strText(23) = "Orbit           "
  strText(24) = "Mean Motion     "
  strText(25) = "Mean Anomoly    "
  strText(26) = "Inclination     "
  strText(27) = "AOP             "
  strText(28) = "RAAN            "
  strText(29) = "Eccentricity    "
  strText(30) = "Height          "
  strText(31) = "MA              "
  strText(32) = "Next AOS Date   "
  strText(33) = "Next AOS Time   "

End Sub

Private Sub menuBlank_Click()
  Dim i As Integer
  Dim nPos As Integer
  
  If lstDetails.SelectedItem Is Nothing Then
    nPos = 1
  Else
    nPos = lstDetails.SelectedItem.Index
  End If
  
  For i = 99 To nPos + 1 Step -1
    nFields(i) = nFields(i - 1)
  Next i
  
  nFields(nPos) = -1
  Rebuild_List

End Sub

Private Sub mnuFCDelete_Click()
  Dim i As Integer
  Dim nPos As Integer
  
  If Not (lstDetails.SelectedItem Is Nothing) Then
    nPos = lstDetails.SelectedItem.Index
    For i = nPos To 99
      nFields(i) = nFields(i + 1)
    Next i
    Rebuild_List
  End If
'  SelectedControls(0).UpdateDataWindow
End Sub

Private Sub mnuFCField_Click(Index As Integer)
  Dim i As Integer
  Dim nPos As Integer
  
  If lstDetails.SelectedItem Is Nothing Then
    nPos = 1
  Else
    nPos = lstDetails.SelectedItem.Index
  End If
  
  For i = 99 To nPos + 1 Step -1
    nFields(i) = nFields(i - 1)
  Next i
  
  nFields(nPos) = Index
  Rebuild_List
  
End Sub
