VERSION 5.00
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSatDetail 
   HelpContextID   =  1410
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Satellite Details"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   720
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      PersonalizedMenus=   0
      Style           =   0
      Tools           =   "frmSatelliteDetail.frx":0000
      ToolBars        =   "frmSatelliteDetail.frx":2696
   End
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
      HelpContextID   =  2310
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
Attribute VB_Name = "frmSatDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strText(-1 To 100) As String
Public fParent As Form
Dim nFields(100) As Integer


Private Sub Form_Load()
  Dim i As Integer
  
  SetupText
  
  For i = 1 To 33
    Me.mnuFCField(i).Caption = strText(i)
  Next i
  
  Form_Resize
 ' Rebuild_List
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  
  Me.Width = 3100
  Me.lstDetails.Left = 0
  Me.lstDetails.Top = 0
  Me.lstDetails.Width = Me.ScaleWidth
  Me.lstDetails.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  fParent.RemoveDetailWindow Me.Tag
End Sub

Private Sub lstDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button And vbRightButton Then
    PopupMenu mnuChooserMenu
  End If
  
End Sub

Public Sub Rebuild_List()
  Dim i As Integer
  Dim lItem As ListItem
  Dim vData As Variant
  Dim strFields As String
  
  strFields = fParent.ocxSat.DisplayDataFields
  
  vData = StrParse(strFields, ",")
  Erase nFields
  For i = 1 To UBound(vData) + 1
    If vData(i - 1) <> "" Then
      nFields(i) = vData(i - 1)
    End If
  Next i

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

Public Sub Update_List()
  Dim strTemp As String
  Dim i As Integer
  Dim nSat As Integer
  Dim nOldSat As Integer
  Dim dLat As Single
  Dim dLon As Single
  
  With fParent.ocxSat
    nOldSat = .SatelliteIndex
    For i = 1 To .SatelliteCount
      .SatelliteIndex = i
      If .SatelliteName = Me.Tag Then
        nSat = i
        Exit For
      End If
    Next i
    For i = 1 To 100
      If nFields(i) <> 0 Then
          Select Case nFields(i)
            Case 1
              lstDetails.ListItems(i).SubItems(1) = .SatelliteDesignator
            Case 2
              lstDetails.ListItems(i).SubItems(1) = .SatelliteName
            Case 3
              lstDetails.ListItems(i).SubItems(1) = .SatelliteAzimuth
            Case 4
              lstDetails.ListItems(i).SubItems(1) = .SatelliteElevation
            Case 5
              dLat = .satellitelatitude
              If dLat < 0 Then
                strTemp = Format(Str(Abs(dLat)), "##0.00") & "°S  "
              Else
                strTemp = Format(Str(dLat), "##0.00") & "°N  "
              End If
              lstDetails.ListItems(i).SubItems(1) = strTemp
            Case 6
              dLon = 360 - .SatelliteLongitude
              Select Case dLon
                Case 0 To 180
                  dLon = -dLon
                Case 181 To 360
                  dLon = (360 - dLon)
              End Select
              If dLon < 0 Then
                strTemp = Format(Str(Abs(dLon)), "##0.00") & "°W  "
              Else
                strTemp = Format(Str(dLon), "##0.00") & "°E  "
              End If
              lstDetails.ListItems(i).SubItems(1) = strTemp
            Case 7
              lstDetails.ListItems(i).SubItems(1) = .SatelliteRange
            Case 8
              lstDetails.ListItems(i).SubItems(1) = .SatelliteOrbitNumber
            Case 9
              'lstDetails.ListItems(i).SubItems(1) = tangerate
            Case 10
              If .UplinkFrequency <> 0 Then
                lstDetails.ListItems(i).SubItems(1) = .UplinkFrequency
              Else
                lstDetails.ListItems(i).SubItems(1) = "N/A"
              End If
            Case 11
              If .DownLinkFrequency <> 0 Then
                lstDetails.ListItems(i).SubItems(1) = .DownLinkFrequency
              Else
                lstDetails.ListItems(i).SubItems(1) = "N/A"
              End If
            Case 12
              If .SatelliteRxFrequency <> 0 Then
                lstDetails.ListItems(i).SubItems(1) = .SatelliteRxFrequency
              Else
                lstDetails.ListItems(i).SubItems(1) = "N/A"
              End If
            Case 13
              If .SatelliteTXFrequency <> 0 Then
                lstDetails.ListItems(i).SubItems(1) = .SatelliteTXFrequency
              Else
                lstDetails.ListItems(i).SubItems(1) = "N/A"
              End If
            Case 14
              If .SatelliteRxFrequency <> 0 And .UplinkFrequency <> 0 Then
                lstDetails.ListItems(i).SubItems(1) = Format(.SatelliteRxFrequency - .UplinkFrequency, "###.######") * 1000
              End If
            Case 15
       '       lstDetails.ListItems(i).SubItems(1) = Format(.pa, "##.##")
            Case 16
              lstDetails.ListItems(i).SubItems(1) = Format(.SatelliteMaximumDX, "######")
            Case 17
              lstDetails.ListItems(i).SubItems(1) = .SatelliteStatusText
            Case 18
              'lstDetails.ListItems(i).SubItems(1) = RS
            Case 19
              lstDetails.ListItems(i).SubItems(1) = .SquintAngle
            Case 20
              lstDetails.ListItems(i).SubItems(1) = .KepsDecayRate
            Case 21
              lstDetails.ListItems(i).SubItems(1) = .KepsElementSet
            Case 22
              lstDetails.ListItems(i).SubItems(1) = .KepsEpochTime
            Case 23
              lstDetails.ListItems(i).SubItems(1) = .KepsOrbitNumber
            Case 24
              lstDetails.ListItems(i).SubItems(1) = .KepsMeanMotion
            Case 25
              lstDetails.ListItems(i).SubItems(1) = .KepsMeanAnomoly
            Case 26
              lstDetails.ListItems(i).SubItems(1) = .KepsInclination
            Case 27
              lstDetails.ListItems(i).SubItems(1) = .KepsAOP
            Case 28
              lstDetails.ListItems(i).SubItems(1) = .KepsRAAN
            Case 29
              lstDetails.ListItems(i).SubItems(1) = .KepsEccentricity
            Case 30
              lstDetails.ListItems(i).SubItems(1) = .SatelliteAltitude
            Case 31
              lstDetails.ListItems(i).SubItems(1) = .SatelliteMA
            Case 32
              lstDetails.ListItems(i).SubItems(1) = Format(.SatelliteNextAOS, "General Date")
            Case 33
              lstDetails.ListItems(i).SubItems(1) = Format(.SatelliteNextAOS, "HH:nn:ss")
          End Select
      Else
        Exit For
      End If
    Next i
    .SatelliteIndex = nOldSat
  End With
End Sub
