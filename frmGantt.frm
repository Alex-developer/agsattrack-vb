VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9E49D79B-F64C-4517-A416-CCE4417B02CE}#4.9#0"; "SatTrack.ocx"
Begin VB.Form frmGantt 
   HelpContextID   =  135
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9420
   Icon            =   "frmGantt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   9420
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   4140
   End
   Begin VB.PictureBox picGantt 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   180
      ScaleHeight     =   3915
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   120
      Width           =   8955
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   4470
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1170
      Top             =   3450
   End
   Begin SatTrack.SatTrackControl ocxSat 
      Height          =   645
      Left            =   300
      TabIndex        =   0
      Top             =   2880
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1138
      OutputStyle     =   4
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2670
      TabIndex        =   1
      Top             =   4170
      Width           =   4755
   End
   Begin VB.Menu mnuGantt 
      HelpContextID   =  1310
      Caption         =   "Gantt Menu"
      Begin VB.Menu mnuGanttSat 
         Caption         =   "Details Here"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGanttSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGanttDetails 
         Caption         =   "Display Pass Details"
      End
   End
End
Attribute VB_Name = "frmGantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tSatPos
  vStart As Variant
  vEnd As Variant
  nSatNum As Integer
End Type
Private Type tRect
  lLeft As Long
  lTop As Long
  lWidth As Long
  lHeight As Long
  nPos As Integer
End Type

Dim vDisplay(1000) As tSatPos
Dim vRects(1000) As tRect
Dim nVPos As Integer
Dim nLineHeight As Integer
Dim nHourSize As Integer
Dim nMinSize As Integer
Dim bUpdating As Boolean
Dim bAOSLOSOnly As Boolean
Dim nSatMouseOver As Integer
Dim vLastTime As Variant

Private Sub Form_Activate()
  fMainForm.UpdateToolbar
End Sub

Private Sub Form_Load()
  With Me.ocxSat
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
  
  vLastTime = Now
  
  Me.tmrLoad.Enabled = True
  
End Sub

Private Sub Form_Paint()
  If Me.ocxSat.SatelliteCount > 0 Then
    Gantt_Paint
  End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  
  If Not bUpdating Then
    Me.ProgressBar1.Visible = False
    Me.Label1.Visible = False
    Me.picGantt.Left = 0
    Me.picGantt.Top = 0
    Me.picGantt.Width = Me.ScaleWidth
    Me.picGantt.Height = Me.ScaleHeight
    Form_Paint
  Else
    Me.ProgressBar1.Top = Me.ScaleHeight - Me.ProgressBar1.Height
    Me.ProgressBar1.Left = 0
    Me.ProgressBar1.Width = Me.ScaleWidth
    
    Me.Label1.Top = Me.ProgressBar1.Top - Me.Label1.Height
    Me.Label1.Left = 0
    Me.Label1.Width = Me.ScaleWidth
    
    Me.picGantt.Left = 0
    Me.picGantt.Top = 0
    Me.picGantt.Width = Me.ScaleWidth
    Me.picGantt.Height = Me.Label1.Top
    Me.ProgressBar1.Visible = True
    Me.Label1.Visible = True
  End If
End Sub

Private Sub mnuGanttDetails_Click()
  Dim strText As String
  
  Me.ocxSat.SatelliteIndex = nSatMouseOver
  strText = "Pass details for " & Me.ocxSat.SatelliteName
  CalculateDisplay Date, "00:00", nSatMouseOver
  frmGanttDetails.Caption = strText
  frmGanttDetails.Show vbModal
End Sub

Private Sub picGantt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim bOver As Boolean
  Dim strText As String
  
  For i = 0 To 100
    If vRects(i).lHeight = 0 Then Exit For
    
    If X >= vRects(i).lLeft And X <= vRects(i).lWidth Then
      If Y >= vRects(i).lTop And Y <= vRects(i).lHeight Then
        Me.ocxSat.SatelliteIndex = vDisplay(vRects(i).nPos).nSatNum
        strText = Me.ocxSat.SatelliteName
        strText = strText & " AOS = " & vDisplay(vRects(i).nPos).vStart
        strText = strText & " LOS = " & vDisplay(vRects(i).nPos).vEnd
        Me.picGantt.ToolTipText = strText
        bOver = True
        nSatMouseOver = vDisplay(vRects(i).nPos).nSatNum
      End If
    End If
  Next i
  If Not bOver Then
    Me.picGantt.ToolTipText = ""
    nSatMouseOver = -1
  End If
End Sub

Private Sub picGantt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim strText As String
  
  For i = 0 To 100
    If vRects(i).lHeight = 0 Then Exit For
    
    If X >= vRects(i).lLeft And X <= vRects(i).lWidth Then
      If Y >= vRects(i).lTop And Y <= vRects(i).lHeight Then
        Me.ocxSat.SatelliteIndex = vDisplay(vRects(i).nPos).nSatNum
        strText = Me.ocxSat.SatelliteName
        strText = strText & " AOS = " & vDisplay(vRects(i).nPos).vStart
        strText = strText & " LOS = " & vDisplay(vRects(i).nPos).vEnd
        Me.mnuGanttSat.Caption = strText
        PopupMenu Me.mnuGantt
      End If
    End If
  Next i
  
End Sub

Private Sub tmrLoad_Timer()
  Me.tmrLoad.Enabled = False
  frmSelect.Show vbModal
  If Me.ocxSat.SatelliteCount > 0 Then
    UpdateGannt
    Me.tmrUpdate.Enabled = True
  End If
End Sub

Private Sub Gantt_Paint()
  Dim nXStart As Integer
  Dim nXEnd As Integer
  Dim nY As Integer
  Dim i As Integer
  Dim j As Integer
  Dim vSunRise As Variant
  Dim vSunSet As Variant
  Dim vTemp As Variant
  Dim nYpos As Integer
  Dim vNow As Variant
  
  Me.picGantt.Cls
  nLineHeight = Me.picGantt.Height / (Me.ocxSat.SatelliteCount + 3)
  nHourSize = (Me.picGantt.Width - 1000) / 24
  nMinSize = nHourSize / 60

  vSunRise = Format(Me.ocxSat.SunRise, "Short Time")
  vSunSet = Format(Me.ocxSat.SunSet, "Short Time")
  
  Me.picGantt.FillStyle = vbSolid
  For i = 0 To 24
    Me.picGantt.CurrentX = (i * nHourSize) + 1000
    Me.picGantt.CurrentY = 0
    Me.picGantt.ForeColor = RGB(255, 255, 255)
    Me.picGantt.Print i
    Me.picGantt.ForeColor = RGB(128, 128, 128)
    Me.picGantt.Line (((i * nHourSize) + 1000), 200)-(((i * nHourSize) + 1000), Me.ocxSat.SatelliteCount * nLineHeight)
  Next i
    
  For i = 1 To Me.ocxSat.SatelliteCount
    nYpos = i * nLineHeight
    Me.ocxSat.SatelliteIndex = i
    Me.picGantt.CurrentX = 0
    Me.picGantt.CurrentY = nYpos
    Me.picGantt.ForeColor = RGB(255, 255, 255)
    Me.picGantt.Print Me.ocxSat.SatelliteName
    Me.picGantt.ForeColor = RGB(192, 192, 192)
    Me.picGantt.FillColor = RGB(128, 128, 128)
    Me.picGantt.Line (1000, nYpos)-(Me.ScaleWidth, (nYpos) + nLineHeight / 4), , B
    'draw sunrise
    nXStart = GetX(DateAdd("n", -30, vSunRise))
    nXEnd = GetX(vSunRise)
    Me.picGantt.FillColor = RGB(255, 211, 33)
    Me.picGantt.ForeColor = RGB(255, 211, 33)
    Me.picGantt.Line (nXStart, nYpos)-(nXEnd, nYpos + nLineHeight / 4), , B
    'draw daylight
    nXStart = GetX(vSunRise)
    nXEnd = GetX(DateAdd("n", -30, vSunSet))
    Me.picGantt.FillColor = RGB(255, 101, 33)
    Me.picGantt.ForeColor = RGB(255, 101, 33)
    Me.picGantt.Line (nXStart, nYpos)-(nXEnd, nYpos + nLineHeight / 4), , B
    'draw sunset
    nXStart = GetX(DateAdd("n", -30, vSunSet))
    nXEnd = GetX(vSunSet)
    Me.picGantt.FillColor = RGB(255, 211, 33)
    Me.picGantt.ForeColor = RGB(255, 211, 33)
    Me.picGantt.Line (nXStart, nYpos)-(nXEnd, nYpos + nLineHeight / 4), , B
  Next i
  
  vNow = Format(Now, "Short Time")
  nXStart = GetX(vNow)
  Me.picGantt.ForeColor = RGB(255, 0, 0)
  Me.picGantt.Line (nXStart, 10)-(nXStart, nYpos)
  
  Me.picGantt.FillColor = RGB(33, 255, 38)
  Me.picGantt.ForeColor = RGB(33, 255, 38)
  Me.picGantt.FillStyle = vbSolid
  Erase vRects
  For i = 0 To 1000
    If vDisplay(i).nSatNum = 0 Then Exit For
    nY = vDisplay(i).nSatNum * nLineHeight + 40
    nXStart = GetX(vDisplay(i).vStart)
    nXEnd = GetX(vDisplay(i).vEnd)
    Me.picGantt.Line (nXStart, nY)-(nXEnd, nY + nLineHeight / 4 - 80), , B
    vRects(i).lLeft = nXStart
    vRects(i).lTop = nY + 40
    vRects(i).lWidth = nXEnd
    vRects(i).lHeight = nY + nLineHeight / 4 - 80
    vRects(i).nPos = i
  Next i
End Sub

Private Function GetX(vTime As Variant) As Integer
  Dim nHour As Integer
  Dim nMin As Integer

  nHour = Val(Left(vTime, 2))
  nMin = Val(Mid(vTime, 4, 2))

  GetX = ((nHour * nHourSize) + 1000) + (nMin * nMinSize)
End Function
Private Sub CalculatePasses()
  Dim i As Integer
  
  For i = 0 To 24
  Next i
End Sub

Public Sub UpdateGannt()
  Dim nCurrSat As Integer
  Dim nCurrTime As Variant
  Dim FormattedDateTime As String
  Dim nMins As Integer
  Dim bGotAOS As Boolean
  
  Erase vDisplay
  If Me.ocxSat.SatelliteCount > 0 Then
    bUpdating = True
    Form_Resize
    DoEvents
    ProgressBar1.Min = 0
    ProgressBar1.Max = (Me.ocxSat.SatelliteCount * 1440) / 10
    nVPos = 0
    For nCurrSat = 1 To Me.ocxSat.SatelliteCount
      Me.ocxSat.SatelliteIndex = nCurrSat
      Me.ocxSat.SetSelectedSatellite = nCurrSat
      Me.Label1.Caption = "Calculating " & Me.ocxSat.SatelliteName
      Me.Label1.Refresh
      nCurrTime = "00:00"
      bGotAOS = False
      For nMins = 0 To 1440
        FormattedDateTime$ = Format$(Date & " " & nCurrTime, "yyyymmddhhmmss")
        ocxSat.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
        ocxSat.DisplayYear = Val(Left$(FormattedDateTime$, 4))
        ocxSat.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
        ocxSat.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
        ocxSat.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
        ocxSat.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
        ocxSat.DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
        ocxSat.CalculateSatellitePosition False, nCurrSat
        If ocxSat.SatelliteElevation > -1 Then
          If Not (bGotAOS) Then
            bGotAOS = True
            vDisplay(nVPos).nSatNum = nCurrSat
            vDisplay(nVPos).vStart = nCurrTime
          End If
        End If
        If ocxSat.SatelliteElevation < -1 And bGotAOS Then
          bGotAOS = False
          vDisplay(nVPos).nSatNum = nCurrSat
          vDisplay(nVPos).vEnd = nCurrTime
          nVPos = nVPos + 1
          If nVPos > 1000 Then Exit For
        End If
        nCurrTime = DateAdd("n", 1, nCurrTime)
        Me.ProgressBar1.Value = Int((((nCurrSat - 1) * 1440) + nMins) / 10)
      Next nMins
      If nVPos > 1000 Then Exit For
    Next nCurrSat
    bUpdating = False
  End If
  Form_Resize
  DoEvents
End Sub

Private Sub CalculateDisplay(vStartDate As Variant, vStartTime As Variant, nSat As Integer)
  Dim vDate As Variant
  Dim vTime As Variant
  Dim i As Integer
  Dim j As Integer
  Dim strLine As String
  Dim bGotAOS As Boolean
  Dim bDidSomething As Boolean
  Dim FormattedDateTime As String
  Dim nLastElevation As Integer
  
  Dim vAOSTime As Variant
  Dim vAOSDate As Variant
  Dim nAOSaz As Integer
  Dim nMaxEle As Integer
  Dim nLOSaz As Integer
  
  Dim nSatIndex As Integer
  
  nSatIndex = nSat
  
  frmGanttDetails.lstData.Clear
  frmGanttDetails.lstData.Visible = False
  If Not bAOSLOSOnly Then
    frmGanttDetails.lblCaption = "Date        Time    Ele  Az   Range    Lat   Lon  Uplink         Downlink"
  Else
    frmGanttDetails.lblCaption = "AOS Date    Time    LOS Date    Time   Duration  Ele  AOSAz LOSAz"
  End If
  
  frmProgress.ProgressBar1.Min = 0
  frmProgress.ProgressBar1.Max = (1 * 1440) / 10
  frmProgress.Caption = "Calculating pass information"
  frmProgress.Show
  
  With Me.ocxSat
    vDate = vStartDate
    vTime = vStartTime
    For j = 1 To 1
      bDidSomething = False
      For i = 1 To 1440
        FormattedDateTime$ = Format$(vDate & " " & vTime, "yyyymmddhhmmss")
        .DisplayCentury = Val(Left$(FormattedDateTime$, 2))
        .DisplayYear = Val(Left$(FormattedDateTime$, 4))
        .DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
        .DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
        .DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
        .DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
        .DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
        .CalculateSatellitePosition False, nSatIndex
        .SatelliteIndex = nSatIndex
        If .SatelliteElevation > -1 Then
          If Not bAOSLOSOnly Then
            strLine = Format(vDate, "dd mmm yyyy") & " " & Format(vTime, "HH:MM") & " "
            strLine = strLine & Format(.SatelliteElevation, "@@@") & "  " & Format(.SatelliteAzimuth, "@@@") & "  " & Format(Int(.SatelliteRange), "@@@@@@") & "Km  " & Format(.satellitelatitude, "@@@") & "  " & Format(.SatelliteLongitude, "@@@")
            frmGanttDetails.lstData.AddItem strLine
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
          If bAOSLOSOnly Then
            strLine = Format(vAOSDate, "dd mmm yyyy") & " " & Format(vAOSTime, "HH:MM") & "   "
            strLine = strLine & Format(vDate, "dd mmm yyyy") & " " & Format(vTime, "HH:MM") & "  " & Format(vTime - vAOSTime, "hh:mm:ss") & "  "
            strLine = strLine & Format(nMaxEle, "@@@") & "  " & Format(nAOSaz, "@@@") & "  " & Format(.SatelliteAzimuth, "@@@")
            frmGanttDetails.lstData.AddItem strLine
            nMaxEle = -99
            bDidSomething = True
          Else
            frmGanttDetails.lstData.AddItem "---------------------------------------------------------"
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
      If Left(frmGanttDetails.lstData.List(frmGanttDetails.lstData.ListCount - 1), 1) = "-" Then
        frmGanttDetails.lstData.RemoveItem (frmGanttDetails.lstData.ListCount - 1)
      End If
      If bDidSomething And Not gbCancel Then
        frmGanttDetails.lstData.AddItem "============= End of Passes for " & Format(vDate, "DD mmm YYYY") & " ============="
      End If
      If gbCancel Then
        frmGanttDetails.lstData.AddItem "====================== cancelled ========================"
        Exit For
      End If
      vDate = DateAdd("d", 1, vDate)
    Next j
    Unload frmProgress
  End With
  frmGanttDetails.lstData.Visible = True
End Sub
Public Sub UpdateReport()
  UpdateGannt
End Sub

Private Sub tmrUpdate_Timer()
Dim vNow As Variant
Dim vMidnight As Variant

  If Me.ocxSat.SatelliteCount > 0 Then
    vMidnight = Format("00:00:00", "Short Time")
    vNow = Format(Now, "Short Time")
    If vNow > vMidnight And vLastTime < vMidnight Then
      UpdateGannt
      Gantt_Paint
    Else
      Gantt_Paint
    End If
    vLastTime = vNow
  End If
End Sub
