VERSION 5.00
Object = "{9E49D79B-F64C-4517-A416-CCE4417B02CE}#4.9#0"; "SatTrack.ocx"
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLstBar.ocx"
Begin VB.Form frmDocument 
   HelpContextID   =  45
   Caption         =   "frmDocument"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10110
   Begin Listbar.SSListBar SSListBarOptions 
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   11456
      _Version        =   65537
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   2
      GroupCount      =   2
      IconsLargeCount =   6
      IconsSmallCount =   5
      Image(1).Index  =   1
      Image(1).Picture=   "frmDocument.frx":0442
      Image(1).Key    =   "deg0"
      Image(2).Index  =   2
      Image(2).Picture=   "frmDocument.frx":1094
      Image(2).Key    =   "hor"
      Image(3).Index  =   3
      Image(3).Picture=   "frmDocument.frx":1CE6
      Image(3).Key    =   "deg180"
      Image(4).Index  =   4
      Image(4).Picture=   "frmDocument.frx":2938
      Image(4).Key    =   "globe"
      Image(5).Index  =   5
      Image(5).Picture=   "frmDocument.frx":358A
      Image(5).Key    =   "list"
      Image(6).Index  =   6
      Image(6).Picture=   "frmDocument.frx":3950
      Image(6).Key    =   "sat"
      SmallImage(1).Index=   1
      SmallImage(1).Picture=   "frmDocument.frx":3D5A
      SmallImage(1).Key=   "deg0"
      SmallImage(2).Index=   2
      SmallImage(2).Picture=   "frmDocument.frx":40EC
      SmallImage(2).Key=   "deg180"
      SmallImage(3).Index=   3
      SmallImage(3).Picture=   "frmDocument.frx":44B2
      SmallImage(3).Key=   "globe"
      SmallImage(4).Index=   4
      SmallImage(4).Picture=   "frmDocument.frx":4878
      SmallImage(4).Key=   "hor"
      SmallImage(5).Index=   5
      SmallImage(5).Picture=   "frmDocument.frx":4C3E
      SmallImage(5).Key=   "list"
      Groups(1).ItemCount=   5
      Groups(1).Font3D=   5
      Groups(1).PictureBackgroundUseMask=   -1  'True
      Groups(1).CurrentGroup=   -1  'True
      Groups(1).Caption=   "Views"
      Groups(1).Key   =   "Views"
      Groups(1).ListItems(1).Text=   "0 Deg"
      Groups(1).ListItems(1).Key=   "0Deg"
      Groups(1).ListItems(1).IconLarge=   "deg0"
      Groups(1).ListItems(1).IconSmall=   "deg0"
      Groups(1).ListItems(2).Index=   2
      Groups(1).ListItems(2).Text=   "180 Deg"
      Groups(1).ListItems(2).Key=   "180Deg"
      Groups(1).ListItems(2).IconLarge=   "deg180"
      Groups(1).ListItems(2).IconSmall=   "deg180"
      Groups(1).ListItems(3).Index=   3
      Groups(1).ListItems(3).Text=   "Horizon"
      Groups(1).ListItems(3).Key=   "Horizon"
      Groups(1).ListItems(3).IconLarge=   "hor"
      Groups(1).ListItems(3).IconSmall=   "hor"
      Groups(1).ListItems(4).Index=   4
      Groups(1).ListItems(4).Text=   "Globe"
      Groups(1).ListItems(4).Key=   "Globe"
      Groups(1).ListItems(4).IconLarge=   "globe"
      Groups(1).ListItems(4).IconSmall=   "globe"
      Groups(1).ListItems(5).Index=   5
      Groups(1).ListItems(5).Text=   "List"
      Groups(1).ListItems(5).Key=   "List"
      Groups(1).ListItems(5).IconLarge=   "list"
      Groups(1).ListItems(5).IconSmall=   "list"
      Groups(2).Index =   2
      Groups(2).ItemCount=   1
      Groups(2).Font3D=   5
      Groups(2).Caption=   "Data"
      Groups(2).Key   =   "Data"
      Groups(2).ListItems(1).Text=   "ISS"
      Groups(2).ListItems(1).IconLarge=   "sat"
   End
   Begin VB.TextBox txtEl 
      Height          =   285
      Left            =   5010
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtAz 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   540
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Timer tmrRotator 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9120
      Top             =   2190
   End
   Begin VB.TextBox txtRotInterface 
      Height          =   285
      Left            =   5010
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   90
      Visible         =   0   'False
      Width           =   4155
   End
   Begin SatTrack.SatTrackControl ocxSat 
      Height          =   4905
      Left            =   1740
      TabIndex        =   1
      Top             =   1320
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8652
      DisplayDataFields=   ""
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   4620
   End
   Begin VB.CheckBox Check2 
      Caption         =   "List Display"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   7620
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   420
      Top             =   4800
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nSatSpeed As Integer
Public bOnDesktop As Boolean
Public nUpdate As Integer
Public bSpeech As Boolean
Public nSpeechInterval As Integer
Public bIcons As Boolean
Public bViewStatusBar As Boolean
Public bViewSatLabel As Boolean
Public strMap0 As String
Public strMap180 As String
Public strMapHorizon As String
Private strAz As String
Private strEl As String
Private nSpeechCounter As Integer

Const AUTOMATIC = 1
Const MANUAL = 2
Const NONE = 0

Private Type sDetail
  fForm As Form
  strName As String
End Type
Private fDetailForms(100) As sDetail
Private nDetailCounter As Integer
Private bDeactivating As Boolean

Public Sub Resize()
  Form_Resize
End Sub

Private Sub Form_Activate()
  Dim i As Integer
  Dim nOldIndex As Integer
  
 If Not bDeactivating Then
  fMainForm.UpdateToolbar
  fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Clear
  fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.AddItem "All Satellites"
  nOldIndex = ocxSat.SatelliteIndex
  For i = 1 To ocxSat.SatelliteCount
    ocxSat.SatelliteIndex = i
    fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.AddItem ocxSat.SatelliteName
  Next i
  
  UpdateSatelliteData
  
  'ShowDetailWindows
  ocxSat.SatelliteIndex = nOldIndex
  End If
End Sub
Public Sub UpdateSatelliteData()
Dim i As Integer

  Me.SSListBarOptions.Groups("Data").ListItems.Clear
  For i = 1 To ocxSat.SatelliteCount
    ocxSat.SatelliteIndex = i
    With Me.SSListBarOptions.Groups("Data").ListItems
      .Add ocxSat.SatelliteName, ocxSat.SatelliteName, ocxSat.SatelliteName
      .Item(ocxSat.SatelliteName).IconLarge = "sat"
    End With
  Next i
End Sub

Private Sub Form_Deactivate()
  bDeactivating = True
'  HideDetailWindows
  bDeactivating = False
End Sub

Private Sub Form_GotFocus()
'  Dim i As Integer
'  Dim nOldIndex As Integer
'
'  fMainForm.UpdateToolbar
'  fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Clear
'  fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.AddItem "All Satellites"
'  nOldIndex = ocxSat.SatelliteIndex
'  For i = 1 To ocxSat.SatelliteCount
'    ocxSat.SatelliteIndex = i
'    fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.AddItem ocxSat.SatelliteName
'  Next i
'  ocxSat.SatelliteIndex = nOldIndex
End Sub

Private Sub Form_Load()
  Dim i As Integer

  nDetailCounter = 1
  sCurZone = GetRegValueStr("System\CurrentControlSet\Control\TimeZoneInformation", "StandardName")

  For i = 0 To UBound(LocTZI)
    If LocTZI(i).StandardName = sCurZone Then
      CurrentTZI = LocTZI(i)
      Exit For
    End If
  Next

  ocxSat.FrequencyDatabasePath = App.Path
  ocxSat.ObserverLatitude = sProgramOptions.nLatitude
  ocxSat.ObserverLongitude = sProgramOptions.nLongitude
  ocxSat.ObserverHeight = sProgramOptions.nHeight
  ocxSat.ObserverLocation = sProgramOptions.strLocation

  ocxSat.SecondObserverLatitude = sProgramOptions.nSecondLatitude
  ocxSat.SecondObserverLongitude = sProgramOptions.nSecondLongitude
  ocxSat.SecondObserverHeight = sProgramOptions.nSecondHeight
  ocxSat.SecondObserverLocation = sProgramOptions.nSecondName
  ocxSat.SecondObserverEnabled = sProgramOptions.bSecondUsed

  ocxSat.ObserverMapCentre = 0
  ' ocxSat.DisplayTimes = True
  ocxSat.ViewsOrthLocations = sProgramOptions.strOrthLocations
  ' ocxSat.SetIndexOnSelect = True

  ocxSat.Timezone = sProgramOptions.nTimezoneAdjust
  ocxSat.TimeZoneName = sProgramOptions.strTimeZone
  If sProgramOptions.bAutoadjust Then
    If IsDayLight(Now, CurrentTZI) Then
      ocxSat.DaylightSaving = True
      ocxSat.DaylightSavingAdjust = CurrentTZI.DaylightBias / 60
    Else
      ocxSat.DaylightSaving = False
      ocxSat.DaylightSavingAdjust = 0
    End If
  Else
    ocxSat.DaylightSaving = sProgramOptions.bDaylightSaving
    ocxSat.DaylightSavingAdjust = CurrentTZI.DaylightBias / 60
  End If

  'Form_Resize

  Form_GotFocus

  tmrLoad.Enabled = True
  nUpdate = sProgramOptions.nUpdateInterval
  ' ocxSat.AutoInterval = 10
  'ocxSat.AutoMode = True

  Me.nSpeechInterval = sProgramOptions.nSpeechInterval
  Me.ocxSat.EnableSpeech = sProgramOptions.bSpeech

  'Me.ocxSat.DisplayStatusBar = True
  'Me.ocxSat.DisplaySatelliteLabel = False

  Me.ocxSat.EnableSatStatus = True
  'Me.ocxSat.DisplayGroundTrackAsPoints = False
  'Me.ocxSat.GroundTrackInterval = ag15

  '  Me.ocxSat.CalculationModel = agPlan13
  '  Me.ocxSat.CalculationModel = agSGP
  Me.ocxSat.DisplayIcons = sProgramOptions.bIcons
  Me.ocxSat.DisplayAOSCircle = sProgramOptions.bDisplayRangeCircle
  Me.ocxSat.GroundTrackPointSize = sProgramOptions.nGroundTrackPointSize
End Sub

Private Sub Form_Resize()
  Dim lListWidth As Long
  Dim lAdjust As Long
  Dim MaxWidth As Long
  Dim MaxHeight As Long
  Dim nBorders As Long
  
  If Me.WindowState <> vbMinimized Then
    If sProgramOptions.bShowListbar Then
      lListWidth = SSListBarOptions.Width
      lAdjust = 50
      Me.SSListBarOptions.Visible = True
    Else
      lListWidth = 0
      lAdjust = 0
      Me.SSListBarOptions.Visible = False
    End If
    MaxWidth = ocxSat.MaxWidth
    MaxHeight = ocxSat.MaxHeight
    If Me.Width + lAdjust > MaxWidth Then
      nBorders = Me.Width - Me.ScaleWidth
'      Me.Width = MaxWidth + lAdjust + nBorders * 4
'      Exit Sub
    End If
    If Me.Height > MaxHeight Then
      nBorders = Me.Height - Me.ScaleHeight
'      Me.Height = MaxHeight + nBorders
 '     Exit Sub
    End If
    SSListBarOptions.Left = 0
    SSListBarOptions.Top = 0
    SSListBarOptions.Height = Me.ScaleHeight
    ocxSat.Top = 0
    ocxSat.Left = lListWidth + lAdjust
    ocxSat.Width = Me.ScaleWidth - lListWidth - lAdjust
    ocxSat.Height = Me.ScaleHeight
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  lDocumentCount = lDocumentCount - 1
  fMainForm.UpdateToolbar
End Sub

Private Sub ocxSat_GotFocus()
  Form_GotFocus
  'fMainForm.UpdateToolbar
End Sub

Private Sub ocxSat_SatelliteSelected(Index As Integer)
  Dim nTemp As Integer

  If Index > 0 Then
    fMainForm.ActiveForm.ocxSat.SatelliteIndex = Index
    nTemp = fMainForm.ActiveForm.ocxSat.SatelliteTrackOrbits - 1
    If nTemp > -1 Then
      fMainForm.SSActiveToolBars1.Tools("ID_Orbits").ComboBox.ListIndex = nTemp
      Me.ocxSat.DrawFootprints
    End If
  End If
End Sub


Private Sub Timer1_Timer()
  Static nCounter As Integer

  nCounter = nCounter + 1

  If nCounter > nUpdate Or nSatSpeed <> 0 Then
    nCounter = 0
    UpdatePosTable
    If nSatSpeed <> 0 Then
      nCounter = 10
    End If
  End If

End Sub

Public Sub UpdatePosTable()
  Dim ItemX As ListItem
  Dim FormattedDateTime As String
  Dim strTemp As String
  Dim nOldIndex As Integer
  Dim sBearing As Single

  nOldIndex = ocxSat.SatelliteIndex
  For i = -1 To ocxSat.SatelliteCount
    ocxSat.SatelliteIndex = i
    If Not ocxSat.SatelliteBusy Then
      If nSatSpeed = 0 Then
        FormattedDateTime$ = Format$(Now, "yyyymmddhhmmss")
        ocxSat.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
        ocxSat.DisplayYear = Val(Left$(FormattedDateTime$, 4))
        ocxSat.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
        ocxSat.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
        ocxSat.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
        ocxSat.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
        ocxSat.DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
      Else
        OffsetTime
      End If
    Else
      If nSatSpeed <> 0 Then
        OffsetTime
      End If
    End If
  Next i

  UpdateToolBarTime

  ocxSat.CalculateALLPositions
  ocxSat.DrawFootprints

  UpdateDetailWindows
  
  sBearing = ocxSat.SatelliteBearing
  If bSpeech Then
    nSpeechCounter = nSpeechCounter + 1
    If nSpeechCounter >= nSpeechInterval Then
      nSpeechCounter = 0
      ocxSat.SatelliteIndex = ocxSat.SetSelectedSatellite
      ocxSat.SpeakPosition
    End If
  End If
  ocxSat.SatelliteIndex = nOldIndex

  If Not (frmRotatorForm Is Nothing) Then
    If frmRotatorForm.hwnd = Me.hwnd And bMoveRotator Then
      MoveRotator
    End If
  End If

  'Debug.Print ocxSat.SatelliteNextAOS
  'Debug.Print ocxSat.SatelliteNextAOSMaxElevation
  'Debug.Print ocxSat.SatelliteNextAOSMaxElevationTime
End Sub

Private Sub OffsetTime()
  If nSatSpeed <> 0 Then
    ocxSat.AddTimeToSatellite nSatSpeed * 60
  End If
End Sub

Sub UpdateToolBarTime()
  strTemp = Format(ocxSat.DisplayHour, "00") & ":" & Format(ocxSat.DisplayMinute, "00")
  fMainForm.SSActiveToolBars1.Tools("ID_SatelliteTime").Edit.Text = strTemp
  strTemp = Format(ocxSat.DisplayDay, "00") & "/" & Format(ocxSat.DisplayMonth, "00") & "/" & ocxSat.DisplayYear
  fMainForm.SSActiveToolBars1.Tools("ID_SatelliteDate").Edit.Text = strTemp

End Sub

Private Sub tmrLoad_Timer()
  If Me.Tag = "Select" Then
    Me.Tag = ""
    frmSelect.Show vbModal
    Form_GotFocus
  End If
  tmrLoad.Enabled = False
  Me.Timer1.Enabled = True
End Sub

Public Function OpenRotatorLink() As Boolean

  On Error GoTo ERROR_OpenRotatorLink

  Me.txtRotInterface.LinkMode = NONE

  txtRotInterface.LinkTopic = "ARSWIN|RCI"    'Sets up link with VB source.
  txtRotInterface.LinkItem = "AZIMUTH"    'Set link to text box on source.
  txtRotInterface.LinkMode = MANUAL 'Establish a manual DDE link.
  txtRotInterface.LinkMode = AUTOMATIC  'Reestablish new LinkMode.
  txtAz.LinkTopic = "ARSWIN|RCI"    'Sets up link with VB source.
  txtAz.LinkItem = "AZIMUTH"    'Set link to text box on source.
  txtAz.LinkMode = MANUAL 'Establish a manual DDE link.
  txtAz.LinkMode = AUTOMATIC  'Reestablish new LinkMode.
  txtEl.LinkTopic = "ARSWIN|RCI"    'Sets up link with VB source.
  txtEl.LinkItem = "AZIMUTH"    'Set link to text box on source.
  txtEl.LinkMode = MANUAL 'Establish a manual DDE link.
  txtEl.LinkMode = AUTOMATIC  'Reestablish new LinkMode.
  bMoveRotator = True
  Me.tmrRotator.Enabled = True

EXIT_OpenRotatorLink:
  Exit Function

ERROR_OpenRotatorLink:
  Call MsgBox("Unable to start communications with rotator interface. Please ensure that it is running.", vbExclamation + vbOKOnly + vbDefaultButton1, "Rotator error")
  bMoveRotator = False
  Set frmRotatorForm = Nothing
  fMainForm.SSActiveToolBars1.Tools("ID_Rotator").State = ssUnchecked
  Resume EXIT_OpenRotatorLink

End Function

Public Function CloseRotatorLink() As Boolean
  Me.tmrRotator.Enabled = False
  Me.txtRotInterface.LinkMode = NONE
  Me.txtAz.LinkMode = NONE
  Me.txtEl.LinkMode = NONE
  Me.Tag = ""
  Me.ocxSat.UserStatusPanelText = ""
  strAz = ""
  strEl = ""
End Function

Public Function MoveRotator() As Boolean
  Dim nOldIndex As Integer
  Dim i As Integer
  Dim bFoundIt As Boolean
  Dim bMove As Boolean
  
  On Error GoTo ERROR_MoveRotator

  nOldIndex = Me.ocxSat.SatelliteIndex
  bFoundIt = False
  For i = 1 To Me.ocxSat.SatelliteCount
    Me.ocxSat.SatelliteIndex = i
    If Me.ocxSat.SatelliteDesignator = Me.Tag Then
      bFoundIt = True
      Exit For
    End If
  Next i
  If bFoundIt Then
    bMove = True
    If Not sProgramOptions.bRotatorAlwaysTrack Then
      If Me.ocxSat.SatelliteElevation < 0 Then
        bMove = False
      End If
    End If
    If bMove Then
      txtRotInterface.Text = "GA:" & Me.ocxSat.SatelliteAzimuth
      txtRotInterface.LinkPoke
      If Me.ocxSat.SatelliteElevation > -1 Then
        txtRotInterface.Text = "GE:" & Me.ocxSat.SatelliteElevation
        txtRotInterface.LinkPoke
      End If
      Me.txtAz.Text = "RA:"
      Me.txtAz.LinkPoke
      Me.txtEl.Text = "RE:"
      Me.txtEl.LinkPoke
    End If
  Else
    CloseRotatorLink
    fMainForm.SSActiveToolBars1.Tools("ID_Rotator").State = ssUnchecked
    MsgBox "The satellite you were tracking with the antennas has been removed from this form. Antenna tracking has been disabled", vbCritical + vbOKOnly, "Rotator Move Error"
    bMoveRotator = False
    Set frmRotatorForm = Nothing
    strAz = ""
    strEl = ""
  End If

EXIT_MoveRotator:
  Me.ocxSat.SatelliteIndex = nOldIndex
  Exit Function

ERROR_MoveRotator:
  MsgBox "Link with rotator interface lost. Interface will be disabled", vbCritical + vbOKOnly, "Rotator Interface"
  bMoveRotator = False
  Set frmRotatorForm = Nothing
  Me.tmrRotator.Enabled = False
  Me.txtRotInterface.LinkMode = NONE
  Me.txtAz.LinkMode = NONE
  Me.txtEl.LinkMode = NONE
  strAz = ""
  strEl = ""
  Resume EXIT_MoveRotator

End Function

Private Sub txtAz_Change()
  Dim rt As String
  
  rt = Mid$(txtAz.Text, 4, 3)
    
  If Left$(txtAz.Text, 3) = "RA:" Then
    If rt <> "" Then
      strAz = rt
    End If
  End If
  UpdateAntennaPos

End Sub

Private Sub txtEl_Change()
  Dim rt As String
  
  rt = Mid$(txtEl.Text, 4, 3)
  If Left$(txtEl.Text, 3) = "RE:" Then
    If rt <> "" Then
      strEl = rt
    End If
  End If
  UpdateAntennaPos
End Sub

Private Sub UpdateAntennaPos()
  Dim strTempAz As String
  Dim strTempEl As String

  If strAz <> "" Then
    strTempAz = "Azimuth: " & strAz & "°"
  Else
    strTempAz = "Azimuth: ???"
  End If
  If strEl <> "" Then
    strTempEl = "Elevation: " & strEl & "°"
  Else
    strTempEl = "Elevation: ???"
  End If

  Me.ocxSat.UserStatusPanelText = strTempAz & "  " & strTempEl
End Sub
Private Sub SSListBarOptions_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
  Dim fForm As Form
  Dim nPos As Integer
  
  Select Case ItemClicked.Key
    Case "0Deg"
      fMainForm.SSActiveToolBars1.Tools("ID_0DegreeCentre").State = ssChecked
    Case "180Deg"
      fMainForm.SSActiveToolBars1.Tools("ID_180DegreeCentre").State = ssChecked
    Case "Horizon"
      fMainForm.SSActiveToolBars1.Tools("ID_HorizonView").State = ssChecked
    Case "Globe"
      fMainForm.SSActiveToolBars1.Tools("ID_Globe").State = ssChecked
    Case "List"
      fMainForm.SSActiveToolBars1.Tools("ID_TableView").State = ssChecked
    Case Else
      nPos = IsDisplayed(ItemClicked.Key)
      If nPos = 0 Then
        Set fForm = New frmSatDetail
        fForm.Caption = ItemClicked.Key
        fForm.Tag = ItemClicked.Key
        fForm.Show
        Set fForm.fParent = Me
        fDetailForms(nDetailCounter).strName = ItemClicked.Key
        Set fDetailForms(nDetailCounter).fForm = fForm
        nDetailCounter = nDetailCounter + 1
        DoEvents
        fForm.Rebuild_List
        fForm.Update_List
      Else
        fDetailForms(nPos).fForm.SetFocus
      End If
  End Select
End Sub

Private Sub ShowDetailWindows()
  Dim i As Integer
  
  For i = 1 To nDetailCounter
    If Not (fDetailForms(i).fForm Is Nothing) Then
      fDetailForms(i).fForm.Visible = True
    End If
  Next i
End Sub
Private Sub HideDetailWindows()
  Dim i As Integer
  
  For i = 1 To nDetailCounter
    If Not (fDetailForms(i).fForm Is Nothing) Then
      fDetailForms(i).fForm.Visible = False
    End If
  Next i
End Sub
Private Sub UpdateDetailWindows()
  Dim i As Integer
  
  For i = 1 To nDetailCounter
    If Not (fDetailForms(i).fForm Is Nothing) Then
      fDetailForms(i).fForm.Update_List
    End If
  Next i
End Sub

Private Function IsDisplayed(strName As String) As Integer
  Dim i As Integer
  
  IsDisplayed = 0
  For i = 1 To nDetailCounter
    If Not (fDetailForms(i).fForm Is Nothing) Then
      If fDetailForms(i).strName = strName Then
        IsDisplayed = i
        Exit For
      End If
    End If
  Next i
End Function

Public Sub RemoveDetailWindow(strName As String)
  Dim nPos As Integer
  Dim i As Integer
  
  nPos = IsDisplayed(strName)
  If nPos > 0 Then
    For i = nPos To 99
      fDetailForms(i) = fDetailForms(i + 1)
    Next i
    nDetailCounter = nDetailCounter - 1
  End If
End Sub
