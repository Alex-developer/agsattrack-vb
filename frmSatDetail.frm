VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSatDetail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      PersonalizedMenus=   0
      Style           =   0
      Tools           =   "frmSatDetail.frx":0000
      ToolBars        =   "frmSatDetail.frx":2696
   End
   Begin MSComctlLib.ListView lstDetails 
      Height          =   2685
      Left            =   0
      TabIndex        =   0
      Top             =   540
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
End
Attribute VB_Name = "frmSatDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fParent As Form
'Option Explicit
'Dim strText(-1 To 100) As String
'
'Private Sub Form_Load()
'  Dim i As Integer
'
'  SetupText
'
'  For i = 1 To 33
'    Me.mnuFCField(i).Caption = strText(i)
'  Next i
'
'  SETtopmostwindow Me, False
'  Form_Resize
'  Rebuild_List
'End Sub
'
'Private Sub Form_Resize()
'  On Error Resume Next
'
'  Me.lstDetails.Left = 0
'  Me.lstDetails.Top = 0
'  Me.lstDetails.Width = Me.ScaleWidth
'  Me.lstDetails.Height = Me.ScaleHeight
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''  If Not bUnloading Then
''    Me.Tag = ""
''    Me.Hide
''    Cancel = 1
''  End If
'  nSatDetailsTag = -10
'  bSatDetailsVisible = False
'End Sub
'
'Private Sub lstDetails_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'  If Button And vbRightButton Then
'    PopupMenu mnuChooserMenu
'  End If
'
'End Sub
'
'Sub Rebuild_List()
'  Dim i As Integer
'  Dim lItem As ListItem
'
'  Me.lstDetails.ListItems.Clear
'  For i = 1 To 100
'    If nFields(i) <> 0 Then
'      Set lItem = Me.lstDetails.ListItems.Add
'      lItem.Text = strText(nFields(i))
'    Else
'      Exit For
'    End If
'  Next i
'
'End Sub
'
'
'Private Sub SetupText()
'  Erase strText
'
'  strText(-1) = ""
'  strText(1) = "NORAD Id"
'  strText(2) = "Sat Name"
'  strText(3) = "Azimuth       (°)"
'  strText(4) = "Elevation     (°)"
'  strText(5) = "Lat           (°)"
'  strText(6) = "Long          (°)"
'  strText(7) = "Range         (Km)"
'  strText(8) = "Orbit"
'  strText(9) = "Range Rate"
'  strText(10) = "Uplink       (MHz)"
'  strText(11) = "Downlink     (MHz)"
'  strText(12) = "Uplink Tx    (Mhz)"
'  strText(13) = "Downlink Rx  (MHz)"
'  strText(14) = "Doppler      (KHz)"
'  strText(15) = "path Loss    (Db)"
'  strText(16) = "Max Range    (Km)"
'  strText(17) = "Status"
'  strText(18) = "RS"
'  strText(19) = "Squint Angle (°)"
'  strText(20) = "Drag            "
'  strText(21) = "Set             "
'  strText(22) = "Epoch           "
'  strText(23) = "Orbit           "
'  strText(24) = "Mean Motion     "
'  strText(25) = "Mean Anomoly    "
'  strText(26) = "Inclination     "
'  strText(27) = "AOP             "
'  strText(28) = "RAAN            "
'  strText(29) = "Eccentricity    "
'  strText(30) = "Height          "
'  strText(31) = "MA              "
'  strText(32) = "Next AOS Date   "
'  strText(33) = "Next AOS Time   "
'
'End Sub
'
'Private Sub menuBlank_Click()
'  Dim i As Integer
'  Dim nPos As Integer
'
'  If lstDetails.SelectedItem Is Nothing Then
'    nPos = 1
'  Else
'    nPos = lstDetails.SelectedItem.Index
'  End If
'
'  For i = 99 To nPos + 1 Step -1
'    nFields(i) = nFields(i - 1)
'  Next i
'
'  nFields(nPos) = -1
'  Rebuild_List
'
'End Sub
'
'Private Sub mnuFCDelete_Click()
'  Dim i As Integer
'  Dim nPos As Integer
'
'  If Not (lstDetails.SelectedItem Is Nothing) Then
'    nPos = lstDetails.SelectedItem.Index
'    For i = nPos To 99
'      nFields(i) = nFields(i + 1)
'    Next i
'    Rebuild_List
'  End If
''  SelectedControls(0).UpdateDataWindow
'End Sub
'
'Private Sub mnuFCField_Click(Index As Integer)
'  Dim i As Integer
'  Dim nPos As Integer
'
'  If lstDetails.SelectedItem Is Nothing Then
'    nPos = 1
'  Else
'    nPos = lstDetails.SelectedItem.Index
'  End If
'
'  For i = 99 To nPos + 1 Step -1
'    nFields(i) = nFields(i - 1)
'  Next i
'
'  nFields(nPos) = Index
'  Rebuild_List
'
'End Sub
'
Private Sub Form_Load()

End Sub
