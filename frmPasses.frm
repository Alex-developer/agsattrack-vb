VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{9E49D79B-F64C-4517-A416-CCE4417B02CE}#4.4#0"; "SatTrack.ocx"
Begin VB.Form frmPasses 
   HelpContextID   =  105
   Caption         =   "Predictions"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmPasses.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   9705
   Begin SatTrack.SatTrackControl SatTrackControl1 
      Height          =   1245
      Left            =   8580
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   2196
   End
   Begin VB.Frame frmOptions 
      Caption         =   " Options "
      Height          =   1065
      Left            =   90
      TabIndex        =   4
      Top             =   4410
      Width           =   9525
      Begin VB.CommandButton Command1 
         Caption         =   "Whats up"
         Height          =   375
         Left            =   2430
         TabIndex        =   0
         Top             =   240
         Width           =   945
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   6870
         TabIndex        =   12
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDays"
         BuddyDispid     =   196611
         OrigLeft        =   6120
         OrigTop         =   480
         OrigRight       =   6360
         OrigBottom      =   855
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtDays 
         Height          =   315
         Left            =   6450
         TabIndex        =   11
         Text            =   "1"
         ToolTipText     =   "Number of days to run prediction for"
         Top             =   450
         Width           =   405
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   4170
         TabIndex        =   8
         ToolTipText     =   "Start date for prediction"
         Top             =   450
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59113472
         CurrentDate     =   36866
      End
      Begin VB.CheckBox chkAOSLOSOnly 
         Caption         =   "Only Show AOS/LOS"
         Height          =   255
         Left            =   3690
         TabIndex        =   7
         ToolTipText     =   "Only display AOS/LOS details"
         Top             =   180
         Width           =   1935
      End
      Begin VB.CommandButton cmdDetails 
         Caption         =   "&Details"
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmdToday 
         Caption         =   "&Display"
         Height          =   375
         Left            =   1290
         TabIndex        =   5
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Days"
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "for"
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   225
         Left            =   3690
         TabIndex        =   9
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   2490
      TabIndex        =   3
      Top             =   690
      Width           =   5775
   End
   Begin VB.ListBox lstSats 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Available satellites"
      Top             =   660
      Width           =   2325
   End
   Begin VB.ComboBox cmbFiles 
      Height          =   315
      Left            =   60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   2325
   End
   Begin VB.Label lblCaption 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Top             =   330
      Width           =   5805
   End
End
Attribute VB_Name = "frmPasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bForceVis As Boolean

Private Sub cmbFiles_Click()
  Dim i As Integer
  Dim nSats As Integer
  Dim strVis As String

  nSats = ReadKeps(App.Path & "\Elements\" & Me.cmbFiles.Text)
  Me.lstSats.Clear
  For i = 0 To nSats - 1
    If sProgramOptions.bIndicateVis Or bForceVis Then
      With Me.SatTrackControl1
        .EraseSatellites
        .AddSatellite
        .SatelliteIndex = 1
        SetupKeps i
        FormattedDateTime$ = Format$(Now, "yyyymmddhhmmss")
        .DisplayCentury = Val(Left$(FormattedDateTime$, 2))
        .DisplayYear = Val(Left$(FormattedDateTime$, 4))
        .DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
        .DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
        .DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
        .DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
        .DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
        .CalculateSatellitePosition False, 1
        strVis = IIf(.SatelliteElevation > -1, "Yes", "No")
      End With
    End If
    Me.lstSats.AddItem Left$(sKeps(i).strName & Space(20), 15) & " " & strVis
  Next i

  Me.cmdDetails.Enabled = False
  Me.cmdToday.Enabled = False

End Sub

Private Sub cmdDetails_Click()
  frmSatDetails.Tag = Me.lstSats.ListIndex
  frmSatDetails.Show vbModal
End Sub

Private Sub cmdToday_Click()

  Dim nPos As Integer
  
  nPos = Me.lstSats.ListIndex
  Me.SatTrackControl1.EraseSatellites
  Me.SatTrackControl1.AddSatellite
  Me.SatTrackControl1.SatelliteIndex = 1
  SetupKeps nPos
  Me.SatTrackControl1.UpdateDerived
  CalculateDisplay Me.dtpDate.Value, 0, Me.txtDays.Text
End Sub

Private Sub SetupKeps(nPos As Integer)

  With Me.SatTrackControl1
'    .SatelliteDesignator = sKeps(nPos).lDesignator
'    .SatelliteName = Trim(sKeps(nPos).strName)
'    .KepsEpochTime = sKeps(nPos).strEpoch
'    .KepsDecayRate = sKeps(nPos).dDrag
'    .KepsOrbitNumber = sKeps(nPos).lRevolutionnumber
'    .KepsInclination = sKeps(nPos).dInclination
'    .KepsRAAN = sKeps(nPos).dRAAN
'    .KepsEccentricity = sKeps(nPos).dEccentricity
'    .KepsAOP = sKeps(nPos).dAOP
'    .KepsMeanAnomoly = sKeps(nPos).dMeanAnomoly
'    .KepsMeanMotion = sKeps(nPos).dMeanMotion

    .UpdateKeps sKeps(nPos).strLine1, sKeps(nPos).strLine2, sKeps(nPos).strLine3, 1, agplan13

    .DisplayDataFields = "1,2,-1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19"
  End With

End Sub
Private Sub CalculateDisplay(vStartDate As Variant, vStartTime As Variant, nDays)
  Dim vDate As Variant
  Dim vTime As Variant
  Dim i As Integer
  Dim j As Integer
  Dim strLine As String
  Dim bGotAOS As Boolean
  Dim bDidSomething As Boolean
  
  Dim vAOSTime As Variant
  Dim vAOSDate As Variant
  Dim nAOSaz As Integer
  Dim nMaxEle As Integer
  Dim nLOSaz As Integer
  
  Me.lstData.Clear
  Me.lstData.Visible = False
  If Me.chkAOSLOSOnly.Value = 0 Then
    Me.lblCaption = "Date        Time    Ele  Az   Range    Lat   Lon  Uplink         Downlink"
  Else
    Me.lblCaption = "AOS Date    Time    LOS Date    Time   Duration  Ele  AOSAz LOSAz"
  End If
  
  frmProgress.ProgressBar1.Min = 0
  frmProgress.ProgressBar1.Max = (nDays * 1440) / 10
  frmProgress.Caption = "Calculating pass information"
  frmProgress.Show
  
  With Me.SatTrackControl1
    vDate = vStartDate
    vTime = vStartTime
    For j = 1 To nDays
      bDidSomething = False
      For i = 1 To 1440
        FormattedDateTime$ = Format$(vDate + vTime, "yyyymmddhhmmss")
        .DisplayCentury = Val(Left$(FormattedDateTime$, 2))
        .DisplayYear = Val(Left$(FormattedDateTime$, 4))
        .DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
        .DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
        .DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
        .DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
        .DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
        .CalculateSatellitePosition False, 1
        .SatelliteIndex = 1
        If .SatelliteElevation > -1 Then
          If Me.chkAOSLOSOnly.Value = 0 Then
            strLine = Format(vDate, "dd mmm yyyy") & " " & Format(vTime, "HH:MM") & " "
            strLine = strLine & Format(.SatelliteElevation, "@@@") & "  " & Format(.SatelliteAzimuth, "@@@") & "  " & Format(Int(.SatelliteRange), "@@@@@@") & "Km  " & Format(.SatelliteLatitude, "@@@") & "  " & Format(.SatelliteLongitude, "@@@")
            Me.lstData.AddItem strLine
            bDidSomething = True
          End If
          If Not bGotAOS Then
            vAOSTime = vTime
            vAOSDate = vDate
            nAOSaz = .SatelliteAzimuth
          End If
          nMaxEle = IIf(.SatelliteElevation > nMaxEle, .SatelliteElevation, nMaxEle)
          bGotAOS = True
        End If
        If .SatelliteElevation < -1 And bGotAOS Then
          If Me.chkAOSLOSOnly.Value = 1 Then
            strLine = Format(vAOSDate, "dd mmm yyyy") & " " & Format(vAOSTime, "HH:MM") & "   "
            strLine = strLine & Format(vDate, "dd mmm yyyy") & " " & Format(vTime, "HH:MM") & "  " & Format(vTime - vAOSTime, "hh:mm:ss") & "  "
            strLine = strLine & Format(nMaxEle, "@@@") & "  " & Format(nAOSaz, "@@@") & "  " & Format(.SatelliteAzimuth, "@@@")
            Me.lstData.AddItem strLine
            nMaxEle = -99
            bDidSomething = True
          Else
            Me.lstData.AddItem "---------------------------------------------------------"
            bDidSomething = True
          End If
          bGotAOS = False
        End If
        nLastElevation = .SatelliteElevation
        vTime = DateAdd("n", 1, vTime)
        frmProgress.ProgressBar1.Value = Int((((j - 1) * 1440) + i) / 10)
        DoEvents
        If gbCancel Then
          Exit For
        End If
      Next i
      If Left(Me.lstData.List(Me.lstData.ListCount - 1), 1) = "-" Then
        Me.lstData.RemoveItem (Me.lstData.ListCount - 1)
      End If
      If bDidSomething And Not gbCancel Then
        Me.lstData.AddItem "============= End of Passes for " & Format(vDate, "DD mmm YYYY") & " ============="
      End If
      If gbCancel Then
        Me.lstData.AddItem "====================== cancelled ========================"
        Exit For
      End If
      vDate = DateAdd("d", 1, vDate)
    Next j
    Unload frmProgress
  End With
  Me.lstData.Visible = True
End Sub

Private Sub Command1_Click()
  bForceVis = True
  cmbFiles_Click
  bForceVis = False
End Sub

Private Sub Form_GotFocus()
fMainForm.UpdateToolbar

End Sub

Private Sub Form_Load()
  Dim strPath As String
  Dim strFilename As String

  With Me.SatTrackControl1
    .FrequencyDatabasePath = App.Path
    .ObserverLatitude = sProgramOptions.nLatitude
    .ObserverLongitude = sProgramOptions.nLongitude
    .ObserverHeight = sProgramOptions.nHeight
    .ObserverLocation = sProgramOptions.strLocation
    .ObserverMapCentre = 0
    .DisplayTimes = False
    .OutputStyle = 4
    .CalculationModel = agplan13
    .CalculationModel = agSGP
    .Visible = False

    .Timezone = sProgramOptions.nTimezoneAdjust
    .TimeZoneName = sProgramOptions.strTimeZone
    If sProgramOptions.bAutoadjust Then
      If IsDayLight(Now, CurrentTZI) Then
        .DaylightSaving = True
        .DaylightSavingAdjust = CurrentTZI.DaylightBias / 60
      Else
        .DaylightSaving = False
        .DaylightSavingAdjust = 0
      End If
    Else
      .DaylightSaving = sProgramOptions.bDaylightSaving
      .DaylightSavingAdjust = CurrentTZI.DaylightBias / 60
    End If

    .AllowDoEvents = False
    .UseHourglass = False
  End With

  strPath = App.Path & "\Elements\*.txt"

  strFilename = Dir(strPath)
  If strFilename <> "" Then
    Do
      Me.cmbFiles.AddItem strFilename
      strFilename = Dir()
    Loop Until strFilename = ""
    Me.cmbFiles.ListIndex = 0
  End If
  Me.cmdDetails.Enabled = False
  Me.cmdToday.Enabled = False
  Me.dtpDate.Value = Date

End Sub

Private Sub Form_Resize()

On Error Resume Next

  Me.cmbFiles.Left = 0
  Me.cmbFiles.Top = 0
  
  Me.frmOptions.Top = Me.ScaleHeight - Me.frmOptions.Height
  Me.frmOptions.Left = 0
  Me.frmOptions.Width = Me.ScaleWidth
  
  Me.lstSats.Left = 0
  Me.lstSats.Top = Me.cmbFiles.Height
  Me.lstSats.Height = Me.ScaleHeight - Me.cmbFiles.Height - Me.frmOptions.Height
  Me.lstSats.Width = Me.cmbFiles.Width
  
  Me.lstData.Left = Me.cmbFiles.Width
  Me.lstData.Top = Me.cmbFiles.Height
  Me.lstData.Height = Me.ScaleHeight - Me.cmbFiles.Height - Me.frmOptions.Height
  Me.lstData.Width = Me.ScaleWidth - Me.cmbFiles.Width
  
  Me.lblCaption.Left = Me.cmbFiles.Width
  Me.lblCaption.Top = 0
  Me.lblCaption.Width = Me.ScaleWidth - Me.cmbFiles.Width
End Sub

Private Sub lstSats_Click()
  Me.cmdDetails.Enabled = True
  Me.cmdToday.Enabled = True
  Me.lstData.Clear
  Me.lblCaption = ""
End Sub

Private Sub lstSats_DblClick()
  cmdDetails_Click
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
  If Not bCheckNumNeg(KeyAscii, False) And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
End Sub
Function bCheckNumNeg(nChar As Integer, bAllowneg As Boolean) As Boolean
If nChar = 8 Or nChar = 9 Or nChar = 13 Or nChar = 3 Or nChar = 22 Or (nChar >= Asc("0") And nChar <= Asc("9")) Or (nChar = Asc("-") And bAllowneg) Then bCheckNumNeg = True
End Function

