VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form frmWindowOptions 
   HelpContextID   =  65
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Options"
   ClientHeight    =   5010
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6030
   Icon            =   "frmWindowOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   4185
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7382
      _Version        =   131082
      TabCount        =   4
      Tabs            =   "frmWindowOptions.frx":000C
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   3795
         Left            =   30
         TabIndex        =   33
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6694
         _Version        =   131082
         TabGuid         =   "frmWindowOptions.frx":00EB
         Begin VB.Frame Frame14 
            Caption         =   " Speech "
            Height          =   1095
            Left            =   120
            TabIndex        =   46
            Top             =   2370
            Width           =   3405
            Begin VB.CheckBox chkSpeech 
               Caption         =   "Announce time until AOS"
               Height          =   285
               Left            =   150
               TabIndex        =   48
               Top             =   240
               Width           =   2535
            End
            Begin VB.TextBox txtSpeech 
               Height          =   315
               Left            =   1410
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   600
               Width           =   330
            End
            Begin MSComCtl2.UpDown UpDown3 
               Height          =   315
               Left            =   1770
               TabIndex        =   49
               Top             =   600
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Value           =   5
               BuddyControl    =   "txtSpeech"
               BuddyDispid     =   196611
               OrigLeft        =   1756
               OrigTop         =   600
               OrigRight       =   1996
               OrigBottom      =   915
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label11 
               Caption         =   "Announce every"
               Height          =   255
               Left            =   150
               TabIndex        =   51
               Top             =   660
               Width           =   1245
            End
            Begin VB.Label Label12 
               Caption         =   "updates"
               Height          =   255
               Left            =   2100
               TabIndex        =   50
               Top             =   630
               Width           =   825
            End
         End
         Begin VB.Frame fraSample3 
            Caption         =   " AOS / LOS Value "
            Height          =   735
            Left            =   2520
            TabIndex        =   42
            Top             =   1590
            Width           =   2235
            Begin VB.TextBox txtAOSLOS 
               Height          =   285
               Left            =   180
               TabIndex        =   43
               Text            =   "Text1"
               Top             =   300
               Width           =   390
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   285
               Left            =   600
               TabIndex        =   44
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtAOSLOS"
               BuddyDispid     =   196615
               OrigLeft        =   1080
               OrigTop         =   240
               OrigRight       =   1320
               OrigBottom      =   555
               Min             =   -5
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label7 
               Caption         =   "Degrees"
               Height          =   225
               Left            =   930
               TabIndex        =   45
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   " General "
            Height          =   1365
            Left            =   60
            TabIndex        =   37
            Top             =   90
            Width           =   5325
            Begin VB.CheckBox chkRangeCircles 
               Caption         =   "Display Range Circle"
               Height          =   315
               Left            =   120
               TabIndex        =   53
               Top             =   990
               Width           =   1875
            End
            Begin VB.CheckBox chkIcons 
               Caption         =   "Display Satellites/Sun/Moon as pictures"
               Height          =   285
               Left            =   120
               TabIndex        =   52
               Top             =   720
               Width           =   3225
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   285
               Left            =   1680
               TabIndex        =   41
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtUpdate"
               BuddyDispid     =   196620
               OrigLeft        =   1890
               OrigTop         =   330
               OrigRight       =   2130
               OrigBottom      =   645
               Max             =   120
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtUpdate 
               Height          =   285
               Left            =   1260
               TabIndex        =   39
               Top             =   300
               Width           =   405
            End
            Begin VB.Label Label6 
               Caption         =   "seconds"
               Height          =   255
               Left            =   2040
               TabIndex        =   40
               Top             =   330
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Update every"
               Height          =   255
               Left            =   150
               TabIndex        =   38
               Top             =   330
               Width           =   1005
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Orthographic "
            Height          =   735
            Left            =   90
            TabIndex        =   34
            Top             =   1590
            Width           =   2325
            Begin VB.CheckBox chkShade 
               Caption         =   "Show Day and Night"
               Height          =   285
               Left            =   210
               TabIndex        =   35
               Top             =   300
               Width           =   1905
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   3795
         Left            =   30
         TabIndex        =   27
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6694
         _Version        =   131082
         TabGuid         =   "frmWindowOptions.frx":0113
         Begin VB.PictureBox Picture1 
            Height          =   2415
            Left            =   0
            ScaleHeight     =   2355
            ScaleWidth      =   5415
            TabIndex        =   64
            Top             =   1200
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.Frame Frame4 
            Caption         =   " Timezone"
            Height          =   945
            Left            =   0
            TabIndex        =   28
            Top             =   210
            Width           =   5445
            Begin VB.CheckBox chkAutoAdjust 
               Caption         =   "Automatically adjust for daylight saving"
               Height          =   285
               Left            =   2010
               TabIndex        =   32
               Top             =   600
               Width           =   3195
            End
            Begin VB.CheckBox chkDaylightSaving 
               Caption         =   "Daylight Saving"
               Height          =   255
               Left            =   180
               TabIndex        =   31
               Top             =   630
               Width           =   1575
            End
            Begin VB.ComboBox cmbTimezones 
               Height          =   315
               Left            =   180
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   5115
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3795
         Left            =   30
         TabIndex        =   5
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6694
         _Version        =   131082
         TabGuid         =   "frmWindowOptions.frx":013B
         Begin VB.Frame Frame7 
            Caption         =   " Second Observer "
            Height          =   1875
            Left            =   30
            TabIndex        =   54
            Top             =   1800
            Width           =   5265
            Begin VB.CheckBox chkEnableSecond 
               Caption         =   "Enable Mutual Calculations"
               Height          =   225
               Left            =   120
               TabIndex        =   30
               ToolTipText     =   "Select to enable mutual observer calculations"
               Top             =   240
               Width           =   2355
            End
            Begin VB.TextBox txtSecondHeight 
               Height          =   285
               Left            =   4500
               TabIndex        =   59
               ToolTipText     =   "Enter the second observer height in meters"
               Top             =   960
               Width           =   555
            End
            Begin VB.TextBox txtSecondName 
               Height          =   285
               Left            =   1080
               TabIndex        =   58
               ToolTipText     =   "Enter the second observer name to plot on the map"
               Top             =   1365
               Width           =   2205
            End
            Begin VB.TextBox txtSecondLong 
               Height          =   285
               Left            =   2760
               TabIndex        =   57
               ToolTipText     =   "Enter the second observer Longitude in the format HHH.MM.SS (E/W)"
               Top             =   960
               Width           =   1035
            End
            Begin VB.TextBox txtSecondLat 
               Height          =   285
               Left            =   900
               TabIndex        =   56
               ToolTipText     =   "Enter the second observer latitude in the format HH.MM.SS (N/S)"
               Top             =   960
               Width           =   945
            End
            Begin VB.ComboBox cmbSecondObs 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   55
               ToolTipText     =   "Select the mutual observer from the list"
               Top             =   540
               Width           =   5055
            End
            Begin VB.Label Label13 
               Caption         =   "Height"
               Height          =   195
               Left            =   3900
               TabIndex        =   63
               Top             =   990
               Width           =   555
            End
            Begin VB.Label Label10 
               Caption         =   "Location"
               Height          =   195
               Left            =   180
               TabIndex        =   62
               Top             =   1395
               Width           =   675
            End
            Begin VB.Label Label9 
               Caption         =   "Longitude"
               Height          =   195
               Left            =   1950
               TabIndex        =   61
               Top             =   1035
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Latitude"
               Height          =   195
               Left            =   180
               TabIndex        =   60
               Top             =   1035
               Width           =   735
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Position "
            Height          =   1395
            Left            =   30
            TabIndex        =   6
            Top             =   120
            Width           =   5265
            Begin VB.TextBox txtLatitude 
               Height          =   285
               Left            =   900
               TabIndex        =   10
               Text            =   "Text1"
               ToolTipText     =   "Enter the main station latitude in the format HH.MM.SS (N/S)"
               Top             =   255
               Width           =   885
            End
            Begin VB.TextBox txtLongitude 
               Height          =   285
               Left            =   2760
               TabIndex        =   9
               Text            =   "Text1"
               ToolTipText     =   "Enter the main station Longitude in the format HHH.MM.SS (E/W)"
               Top             =   255
               Width           =   1035
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Left            =   900
               TabIndex        =   8
               Text            =   "Text2"
               ToolTipText     =   "Enter the main station name to plot on the map"
               Top             =   660
               Width           =   2205
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Left            =   4500
               TabIndex        =   7
               Text            =   "Text1"
               ToolTipText     =   "Enter the main station height in meters"
               Top             =   255
               Width           =   555
            End
            Begin VB.Label Label2 
               Caption         =   "Latitude"
               Height          =   195
               Left            =   180
               TabIndex        =   14
               Top             =   330
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Longitude"
               Height          =   195
               Left            =   1950
               TabIndex        =   13
               Top             =   330
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Location"
               Height          =   195
               Left            =   180
               TabIndex        =   12
               Top             =   690
               Width           =   675
            End
            Begin VB.Label Label5 
               Caption         =   "Height"
               Height          =   195
               Left            =   3900
               TabIndex        =   11
               Top             =   300
               Width           =   495
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3795
         Left            =   30
         TabIndex        =   4
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6694
         _Version        =   131082
         TabGuid         =   "frmWindowOptions.frx":0163
         Begin VB.Frame Frame2 
            Caption         =   "Display Type"
            Height          =   1515
            Left            =   3330
            TabIndex        =   23
            Top             =   1440
            Width           =   1755
            Begin VB.OptionButton optDisplayType 
               Caption         =   "Orthographic"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   36
               Top             =   1050
               Width           =   1395
            End
            Begin VB.OptionButton optDisplayType 
               Caption         =   "Horizon"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   780
               Width           =   1395
            End
            Begin VB.OptionButton optDisplayType 
               Caption         =   "Map"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   25
               Top             =   540
               Width           =   1455
            End
            Begin VB.OptionButton optDisplayType 
               Caption         =   "Table"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Map Center"
            Height          =   975
            Left            =   3300
            TabIndex        =   20
            Top             =   300
            Width           =   1755
            Begin VB.OptionButton optMapCenter 
               Caption         =   "180 Degrees"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optMapCenter 
               Caption         =   "0 Degrees"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   300
               Width           =   1215
            End
         End
         Begin VB.Frame fraSample1 
            Caption         =   "Images"
            Height          =   3225
            Left            =   90
            TabIndex        =   15
            Tag             =   "Sample 1"
            Top             =   210
            Width           =   3120
            Begin VB.CommandButton cmdLoadNew 
               Caption         =   "Horizon"
               Height          =   435
               Index           =   2
               Left            =   180
               TabIndex        =   19
               Top             =   1620
               Width           =   795
            End
            Begin VB.CommandButton cmdLoadNew 
               Caption         =   "Load 180"
               Height          =   435
               Index           =   1
               Left            =   180
               TabIndex        =   18
               Top             =   960
               Width           =   795
            End
            Begin VB.CommandButton cmdLoadNew 
               Caption         =   "Reset"
               Height          =   435
               Index           =   3
               Left            =   180
               TabIndex        =   17
               Top             =   2520
               Width           =   795
            End
            Begin VB.CommandButton cmdLoadNew 
               Caption         =   "Load 0"
               Height          =   435
               Index           =   0
               Left            =   180
               TabIndex        =   16
               Top             =   300
               Width           =   795
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      Top             =   4410
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3390
      TabIndex        =   1
      Top             =   4395
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   4395
      Width           =   1095
   End
End
Attribute VB_Name = "frmWindowOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMapFile(10) As String
Dim bError As Boolean

Private Sub chkAutoAdjust_Click()
  cmdApply.Enabled = True
End Sub

Private Sub chkDaylightSaving_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkEnableSecond_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkIcons_Click()
  cmdApply.Enabled = True
End Sub

Private Sub chkRangeCircles_Click()
  cmdApply.Enabled = True
End Sub

Private Sub chkShade_Click()
  cmdApply.Enabled = True
End Sub

Private Sub chkSpeech_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub cmbSecondObs_Click()
  Dim nPos As Integer
  Dim strLat As String
  Dim strLon As String
  
  nPos = Me.cmbSecondObs.ListIndex
  
  frmProgramOptions.FormatLatLon sDxDetails(nPos).strLat * 60, -sDxDetails(nPos).strLon * 60, strLat, strLon, True

  Me.txtSecondLat = strLat
  Me.txtSecondLong = strLon
  Me.txtSecondName = sDxDetails(nPos).strName
  
  Me.cmdApply.Enabled = True

End Sub

Private Sub cmbTimezones_Click()
  CurrentTZI = LocTZI(Me.cmbTimezones.ListIndex)
  UpdateTZInfo
  Me.cmdApply.Enabled = True
End Sub
Private Sub UpdateTZInfo()
  Picture1.Cls
  With CurrentTZI
       Picture1.Print "Normal Bias ", , SignStr(.Bias / 60) & " hour(s) to convert local time to UTC"
       Picture1.Print
       Picture1.Print "Standard Name ", .StandardName
       Picture1.Print "Standard Bias ", SignStr(.StandardBias / 60) & " hour(s) to add to Normal Bias"
       Picture1.Print
       If .DaylightDate.wMonth = 0 Then
          Picture1.Print "No DayLight difference at this zone"
          Exit Sub
       End If
       Picture1.Print "Daylight Name  ", .DaylightName
       Picture1.Print "DayLight Bias ", SignStr(.DaylightBias / 60) & " hour(s) to add to Normal Bias"
  End With
  With CurrentTZI.DaylightDate
       If .wYear Then
          Picture1.Print "Daylight Begins On ", DateSerial(.wYear, .wMonth, .wDay) & " At "; TimeSerial(.wHour, .wMinute, .wSecond)
       Else
          Picture1.Print "Daylight Begins On ", TranslateDay(.wDayOfWeek, .wDay) & " " & GetMonth(.wMonth) & " At " & TimeSerial(.wHour, .wMinute, .wSecond)
          Picture1.Print , , "(This year - " & Format(WeekDayToDate(Year(Now), .wMonth, .wDayOfWeek, .wDay), "dd mmmm yyyy)")
       End If
  End With
  With CurrentTZI.StandardDate
       If .wYear Then
          Picture1.Print "Daylight Ends On ", DateSerial(.wYear, .wMonth, .wDay) & " At "; TimeSerial(.wHour, .wMinute, .wSecond)
       Else
          Picture1.Print "Daylight Ends On ", TranslateDay(.wDayOfWeek, .wDay) & " " & GetMonth(.wMonth) & " At " & TimeSerial(.wHour, .wMinute, .wSecond)
          Picture1.Print , , "(This year - " & Format(WeekDayToDate(Year(Now), .wMonth, .wDayOfWeek, .wDay), "dd mmmm yyyy)")
       End If
  End With
  If IsDayLight(Now, CurrentTZI) Then
    Me.chkDaylightSaving.Value = 1
  Else
    Me.chkDaylightSaving.Value = 0
  End If
End Sub
 Private Function SignStr(ByVal sng As Single) As String
  Dim s As String
  s = CStr(sng)
  If Left$(s, 1) <> "-" Then s = "+" & s
  SignStr = s
End Function

 Private Function TranslateDay(ByVal nDayOfWeek&, ByVal nDay&) As String
  Dim sReturn$
  sReturn = "The "
  Select Case nDay
    Case 1: sReturn = sReturn & "First "
    Case 2: sReturn = sReturn & "Second "
    Case 3: sReturn = sReturn & "Third "
    Case 4: sReturn = sReturn & "Fourth "
    Case 5: sReturn = sReturn & "Last "
  End Select
  Select Case nDayOfWeek
    Case 0: sReturn = sReturn & "Sunday"
    Case 1: sReturn = sReturn & "Monday"
    Case 2: sReturn = sReturn & "Tuesday"
    Case 3: sReturn = sReturn & "Wednesday"
    Case 4: sReturn = sReturn & "Thursday"
    Case 5: sReturn = sReturn & "Friday"
    Case 6: sReturn = sReturn & "Saturday"
  End Select
  TranslateDay = sReturn & " In"
End Function

 Private Function GetMonth(ByVal nMonth&) As String
  Select Case nMonth
    Case 1: GetMonth = "January"
    Case 2: GetMonth = "February"
    Case 3: GetMonth = "March"
    Case 4: GetMonth = "April"
    Case 5: GetMonth = "May"
    Case 6: GetMonth = "June"
    Case 7: GetMonth = "July"
    Case 8: GetMonth = "August"
    Case 9: GetMonth = "September"
    Case 10: GetMonth = "October"
    Case 11: GetMonth = "November"
    Case 12: GetMonth = "December"
  End Select
End Function


 Private Function SysDate() As Date
   Dim st As SYSTEMTIME
   Call GetSystemTime(st)
   SysDate = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

Private Sub cmdApply_Click()
  Dim i As Integer

  bError = False
  If frmProgramOptions.strValidateLatLon(Me, Me.txtLongitude, Me.txtLatitude, True) = "Ok" Then
    If frmProgramOptions.strValidateLatLon(Me, Me.txtLatitude, Me.txtLongitude, False) = "Ok" Then
  If frmProgramOptions.strValidateLatLon(Me, Me.txtSecondLong, Me.txtSecondLat, True) = "Ok" Then
    If frmProgramOptions.strValidateLatLon(Me, Me.txtSecondLat, Me.txtSecondLong, False) = "Ok" Then

  If optMapCenter(0).Value = True Then
    fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0
  Else
    fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 180
  End If

  For i = 0 To 2
    If strMapFile(0) = "Reset" Then
      fMainForm.ActiveForm.ocxSat.SetMap i, "Reset"
      Exit For
    Else
      If strMapFile(i) <> "" Then
        fMainForm.ActiveForm.ocxSat.SetMap i, strMapFile(i)
      End If
    End If
  Next i
    
  If strMapFile(0) = "Reset" Then
    fMainForm.ActiveForm.strMap0 = ""
    fMainForm.ActiveForm.strMap180 = ""
    fMainForm.ActiveForm.strMapHorizon = ""
  Else
    fMainForm.ActiveForm.strMap0 = strMapFile(0)
    fMainForm.ActiveForm.strMap180 = strMapFile(1)
    fMainForm.ActiveForm.strMapHorizon = strMapFile(2)
  End If
  
  If optDisplayType(0).Value = True Then
    fMainForm.ActiveForm.ocxSat.OutputStyle = 0
  Else
    If optDisplayType(1).Value = True Then
      fMainForm.ActiveForm.ocxSat.OutputStyle = 1
    Else
      If optDisplayType(2).Value = True Then
        fMainForm.ActiveForm.ocxSat.OutputStyle = 2
      Else
        fMainForm.ActiveForm.ocxSat.OutputStyle = 3
      End If
    End If
  End If
  
  fMainForm.ActiveForm.ocxSat.DisplayAOSCircle = IIf(Me.chkRangeCircles.Value = 1, True, False)

  ' Tab 2 LOcation
'  fMainForm.ActiveForm.ocxSat.ObserverLatitude = Me.txtLatitude
'  fMainForm.ActiveForm.ocxSat.ObserverLongitude = Me.txtLongitude
  fMainForm.ActiveForm.ocxSat.ObserverLatitude = frmProgramOptions.sConvertPos(Me.txtLatitude) / 60
  fMainForm.ActiveForm.ocxSat.ObserverLongitude = frmProgramOptions.sConvertPos(Me.txtLongitude) / 60
  fMainForm.ActiveForm.ocxSat.ObserverHeight = Me.txtHeight
  fMainForm.ActiveForm.ocxSat.ObserverLocation = Me.txtLocation
  
  
'  fMainForm.ActiveForm.ocxSat.SecondObserverLatitude = Me.txtSecondLat
'  fMainForm.ActiveForm.ocxSat.SecondObserverLongitude = Me.txtSecondLong
  fMainForm.ActiveForm.ocxSat.SecondObserverLatitude = frmProgramOptions.sConvertPos(Me.txtSecondLat) / 60
  fMainForm.ActiveForm.ocxSat.SecondObserverLongitude = frmProgramOptions.sConvertPos(Me.txtSecondLong) / 60
  fMainForm.ActiveForm.ocxSat.SecondObserverHeight = Me.txtSecondHeight
  fMainForm.ActiveForm.ocxSat.SecondObserverLocation = Me.txtSecondName
  fMainForm.ActiveForm.ocxSat.SecondObserverEnabled = IIf(Me.chkEnableSecond.Value = 1, True, False)
  fMainForm.ActiveForm.ocxSat.PlotObserver

  fMainForm.ActiveForm.ocxSat.TimeZoneName = LocTZI(Me.cmbTimezones.ListIndex).StandardName
  fMainForm.ActiveForm.ocxSat.Timezone = LocTZI(Me.cmbTimezones.ListIndex).Bias / 60

  fMainForm.ActiveForm.ocxSat.DisplayIcons = IIf(Me.chkIcons.Value = 1, True, False)

  ' fMainForm.ActiveForm.ocxSat.Timezone = CalculateTimeDifference(Me.txtLongitude)
  If Me.chkDaylightSaving.Value = 1 Then
    fMainForm.ActiveForm.ocxSat.DaylightSaving = True
    fMainForm.ActiveForm.ocxSat.DaylightSavingAdjust = LocTZI(Me.cmbTimezones.ListIndex).DaylightBias / 60
  Else
    fMainForm.ActiveForm.ocxSat.DaylightSaving = False
    fMainForm.ActiveForm.ocxSat.DaylightSavingAdjust = 0
  End If
  sProgramOptions.bAutoadjust = IIf(Me.chkAutoAdjust.Value = 1, True, False)

  fMainForm.ActiveForm.nSpeechInterval = Me.txtSpeech
  fMainForm.ActiveForm.bSpeech = IIf(Me.chkSpeech.Value = 1, True, False)

  ' Tab 3 General Satellite details
  fMainForm.ActiveForm.ocxSat.SetAOSLOS = Val(Me.txtAOSLOS)

  ' tab 4 - views
  fMainForm.ActiveForm.ocxSat.ViewsOrthShade = IIf(Me.chkShade.Value = 1, True, False)
  fMainForm.ActiveForm.nUpdate = Val(Me.txtUpdate)
  cmdApply.Enabled = False

  fMainForm.ActiveForm.UpdatePosTable
  fMainForm.ActiveForm.ocxSat.DrawFootprints
    
    Else
      MsgBox "The second Observer Latitude you have entered is invalid", vbCritical + vbOKOnly, "Error"
      bError = True
    End If
  Else
    MsgBox "The second Observer Longitude you have entered is invalid", vbCritical + vbOKOnly, "Error"
    bError = True
  End If
    
    Else
      MsgBox "The Default Observer Latitude you have entered is invalid", vbCritical + vbOKOnly, "Error"
      bError = True
    End If
  Else
    MsgBox "The Default Observer Longitude you have entered is invalid", vbCritical + vbOKOnly, "Error"
    bError = True
  End If

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdLoadNew_Click(Index As Integer)

  On Error GoTo ERROR_cmdLoadNew_Click

  Dim sFile As String

  If Index <> 3 Then
    With CommonDialog1
      .FileName = App.Path & "\Images\*.gif"
      .Filter = "Images *.gif|*.gif|All Files (*.*)|*.*"
      .CancelError = True
      .ShowOpen
      If Len(.FileName) = 0 Then
        Exit Sub
      End If
      sFile = .FileName
      strMapFile(Index) = .FileName
      cmdApply.Enabled = True
    End With
  Else
    strMapFile(0) = "Reset"
    nMapType = 0
    cmdApply.Enabled = True
  End If


EXIT_cmdLoadNew_Click:
  Exit Sub

ERROR_cmdLoadNew_Click:
  Select Case Err
    Case cdlCancel
    
    Case Else
      MsgBox "Error in ERROR_cmdLoadNew_Click : " & Error
      Resume EXIT_cmdLoadNew_Click
  End Select
End Sub

Private Sub cmdOk_Click()
  bError = False
  If cmdApply.Enabled = True Then
    cmdApply_Click
  End If
  If Not bError Then
    Unload Me
  End If
End Sub


Private Sub Form_Load()
  Dim sCurZone As String
  Dim n As Integer
  Dim i As Integer
  Dim strTemp As String
  Dim strLat As String
  Dim strLon As String
  
  If fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0 Then
    optMapCenter(0).Value = True
  Else
    optMapCenter(1).Value = True
  End If

  optDisplayType(fMainForm.ActiveForm.ocxSat.OutputStyle).Value = True

  Erase strMapFile
  strMapFile(0) = fMainForm.ActiveForm.strMap0
  strMapFile(1) = fMainForm.ActiveForm.strMap180
  strMapFile(2) = fMainForm.ActiveForm.strMapHorizon

  Me.chkRangeCircles.Value = IIf(fMainForm.ActiveForm.ocxSat.DisplayAOSCircle, 1, 0)
' Tab 2 Location
  
  frmProgramOptions.FormatLatLon fMainForm.ActiveForm.ocxSat.ObserverLatitude * 60, fMainForm.ActiveForm.ocxSat.ObserverLongitude * 60, strLat, strLon, True
'  Me.txtLatitude = fMainForm.ActiveForm.ocxSat.ObserverLatitude
'  Me.txtLongitude = fMainForm.ActiveForm.ocxSat.ObserverLongitude
  Me.txtLatitude = strLat
  Me.txtLongitude = strLon
  Me.txtHeight = fMainForm.ActiveForm.ocxSat.ObserverHeight
  Me.txtLocation = fMainForm.ActiveForm.ocxSat.ObserverLocation
  
  frmProgramOptions.FormatLatLon fMainForm.ActiveForm.ocxSat.SecondObserverLatitude * 60, fMainForm.ActiveForm.ocxSat.SecondObserverLongitude * 60, strLat, strLon, True
'  Me.txtSecondLat = fMainForm.ActiveForm.ocxSat.SecondObserverLatitude
'  Me.txtSecondLong = fMainForm.ActiveForm.ocxSat.SecondObserverLongitude
  Me.txtSecondLat = strLat
  Me.txtSecondLong = strLon
  Me.txtSecondHeight = fMainForm.ActiveForm.ocxSat.SecondObserverHeight
  Me.txtSecondName = fMainForm.ActiveForm.ocxSat.SecondObserverLocation
  Me.chkEnableSecond.Value = IIf(fMainForm.ActiveForm.ocxSat.SecondObserverEnabled, 1, 0)
  
  
  If sProgramOptions.bDaylightSaving Then
    Me.chkDaylightSaving = 1
  End If
  
  Me.txtSpeech = fMainForm.ActiveForm.nSpeechInterval
  Me.chkSpeech.Value = IIf(fMainForm.ActiveForm.bSpeech, 1, 0)

  For i = 0 To 2000
    If sDxDetails(i).strName = "" And sDxDetails(i).Callsign = "" Then Exit For
    strTemp = sDxDetails(i).Callsign & "  " & sDxDetails(i).strName
    Me.cmbSecondObs.AddItem strTemp
  Next i
  

' Tab 3 General satellite details
  Me.txtAOSLOS = fMainForm.ActiveForm.ocxSat.SetAOSLOS
  Me.chkIcons.Value = IIf(fMainForm.ActiveForm.ocxSat.DisplayIcons, 1, 0)
  
' Tab 4 - Timezone
  sCurZone = fMainForm.ActiveForm.ocxSat.TimeZoneName
  If sCurZone = "Not Set" Then
    sCurZone = GetRegValueStr("System\CurrentControlSet\Control\TimeZoneInformation", "StandardName")
  End If
  For i = 0 To UBound(LocTZI)
    Me.cmbTimezones.AddItem LocTZI(i).DisplayName
    If LocTZI(i).StandardName = sCurZone Then n = i
  Next
  Me.cmbTimezones.ListIndex = n
  DoEvents
  cmbTimezones_Click
  If fMainForm.ActiveForm.ocxSat.DaylightSaving Then
    Me.chkDaylightSaving.Value = 1
  Else
    Me.chkDaylightSaving.Value = 0
  End If
  
  If sProgramOptions.bAutoadjust Then
    Me.chkAutoAdjust.Value = 1
   Else
    Me.chkAutoAdjust.Value = 0
  End If

' Tab5 - views
  Me.chkShade.Value = IIf(fMainForm.ActiveForm.ocxSat.ViewsOrthShade = True, 1, 0)
  Me.txtUpdate = fMainForm.ActiveForm.nUpdate
  cmdApply.Enabled = False
End Sub

Private Sub optDisplayType_Click(Index As Integer)
  cmdApply.Enabled = True
End Sub

Private Sub optMapCenter_Click(Index As Integer)
  cmdApply.Enabled = True
End Sub

Private Sub txtAOSLOS_Change()
  cmdApply.Enabled = True
End Sub

Private Sub txtHeight_Click()
  cmdApply.Enabled = True
End Sub

Private Sub txtLatitude_Click()
  cmdApply.Enabled = True
End Sub

Private Sub txtLatitude_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) And KeyAscii <> cSPACE And KeyAscii <> cPoint And KeyAscii <> Asc("N") And KeyAscii <> Asc("S") Then KeyAscii = 0
End Sub

Private Sub txtLocation_Click()
  cmdApply.Enabled = True

End Sub

Private Sub txtLongitude_Click()
  cmdApply.Enabled = True
End Sub

Private Sub txtLongitude_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) And KeyAscii <> cSPACE And KeyAscii <> cPoint And KeyAscii <> Asc("E") And KeyAscii <> Asc("W") Then KeyAscii = 0
End Sub

Private Sub txtSecondHeight_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtSecondLat_Click()
  Me.cmdApply.Enabled = True
End Sub


Private Sub txtSecondLong_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtSecondName_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtSpeech_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtUpdate_Change()
  cmdApply.Enabled = True
End Sub

Private Sub txtUpdate_Click()
  cmdApply.Enabled = True
End Sub

Private Sub UpDown2_Change()
  cmdApply.Enabled = True
End Sub
