VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{9E49D79B-F64C-4517-A416-CCE4417B02CE}#4.9#0"; "SatTrack.ocx"
Begin VB.Form frmSatDetails 
   HelpContextID   =  115
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmSatDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SatTrack.SatTrackControl SatTrackControl1 
      Height          =   405
      Left            =   570
      TabIndex        =   55
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   714
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5850
      TabIndex        =   39
      Top             =   5250
      Width           =   1005
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4710
      TabIndex        =   38
      Top             =   5250
      Width           =   1005
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   5025
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   8864
      _Version        =   131082
      TabCount        =   3
      Tabs            =   "frmSatDetails.frx":0442
      Begin ActiveTabs.SSActiveTabPanel frmMain 
         Height          =   4635
         Left            =   -99969
         TabIndex        =   52
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8176
         _Version        =   131082
         TabGuid         =   "frmSatDetails.frx":04EF
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000007&
            Height          =   4455
            Left            =   60
            ScaleHeight     =   4395
            ScaleWidth      =   6555
            TabIndex        =   57
            Top             =   60
            Width           =   6615
            Begin VB.PictureBox picOrbit 
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000007&
               BorderStyle     =   0  'None
               Height          =   4395
               Left            =   780
               ScaleHeight     =   4395
               ScaleWidth      =   5115
               TabIndex        =   58
               Top             =   0
               Width           =   5115
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4635
         Left            =   -99969
         TabIndex        =   37
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8176
         _Version        =   131082
         TabGuid         =   "frmSatDetails.frx":0517
         Begin VB.Frame Frame2 
            Caption         =   " Selected Frequency "
            Height          =   2265
            Left            =   150
            TabIndex        =   41
            Top             =   2070
            Width           =   5685
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Height          =   375
               Left            =   4680
               TabIndex        =   51
               Top             =   1770
               Width           =   915
            End
            Begin VB.CommandButton Delete 
               Caption         =   "Delete"
               Height          =   375
               Left            =   3690
               TabIndex        =   50
               Top             =   1770
               Width           =   915
            End
            Begin VB.CommandButton cmdNew 
               Caption         =   "New"
               Height          =   375
               Left            =   2700
               TabIndex        =   49
               Top             =   1770
               Width           =   915
            End
            Begin VB.TextBox txtDownlink 
               Height          =   315
               Left            =   1110
               TabIndex        =   47
               Top             =   1080
               Width           =   1635
            End
            Begin VB.TextBox txtFreqDesc 
               Height          =   315
               Left            =   1110
               TabIndex        =   46
               Top             =   300
               Width           =   4365
            End
            Begin VB.TextBox txtUplink 
               Height          =   315
               Left            =   1110
               TabIndex        =   44
               Top             =   690
               Width           =   1635
            End
            Begin VB.CheckBox chkInverting 
               Caption         =   "Inverting transponder"
               Height          =   285
               Left            =   1110
               TabIndex        =   42
               Top             =   1470
               Width           =   1935
            End
            Begin VB.Label Label9 
               Caption         =   "Downlink"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   1110
               Width           =   795
            End
            Begin VB.Label Label8 
               Caption         =   "Description"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   330
               Width           =   975
            End
            Begin VB.Label Label7 
               Caption         =   "Uplink"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   585
            End
         End
         Begin MSComctlLib.ListView lsvFreq 
            Height          =   1635
            Left            =   120
            TabIndex        =   40
            Top             =   210
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   2884
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Inverting"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Uplink"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Downlink"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4635
         Left            =   30
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8176
         _Version        =   131082
         TabGuid         =   "frmSatDetails.frx":053F
         Begin VB.Frame Frame1 
            Caption         =   " Derived Values "
            Height          =   2835
            Left            =   3300
            TabIndex        =   26
            Top             =   180
            Width           =   3165
            Begin VB.TextBox txtSemiMInorAxis 
               Height          =   315
               Left            =   1320
               TabIndex        =   56
               Top             =   1140
               Width           =   1605
            End
            Begin VB.TextBox txtPeriod 
               Height          =   315
               Left            =   1320
               TabIndex        =   33
               Top             =   2340
               Width           =   1605
            End
            Begin VB.TextBox txtAltApogee 
               Height          =   315
               Left            =   1320
               TabIndex        =   32
               Top             =   1935
               Width           =   1605
            End
            Begin VB.TextBox txtAltPerigee 
               Height          =   315
               Left            =   1320
               TabIndex        =   31
               Top             =   1530
               Width           =   1605
            End
            Begin VB.TextBox txtLonOfNode 
               Height          =   315
               Left            =   1320
               TabIndex        =   29
               Top             =   330
               Width           =   1605
            End
            Begin VB.TextBox txtSMA 
               Height          =   315
               Left            =   1320
               TabIndex        =   27
               Top             =   735
               Width           =   1605
            End
            Begin VB.Label Label11 
               Caption         =   "Semi minor axis"
               Height          =   315
               Left            =   90
               TabIndex        =   53
               Top             =   1170
               Width           =   1155
            End
            Begin VB.Label Label10 
               Caption         =   "Label10"
               Height          =   30
               Left            =   1680
               TabIndex        =   54
               Top             =   1380
               Width           =   30
            End
            Begin VB.Label Label6 
               Caption         =   "Period"
               Height          =   315
               Left            =   60
               TabIndex        =   36
               Top             =   2400
               Width           =   1065
            End
            Begin VB.Label Label5 
               Caption         =   "Alt at Apogee"
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   1995
               Width           =   1065
            End
            Begin VB.Label Label4 
               Caption         =   "Alt at Perigee"
               Height          =   315
               Left            =   60
               TabIndex        =   34
               Top             =   1590
               Width           =   1065
            End
            Begin VB.Label Label3 
               Caption         =   "Lng. of node"
               Height          =   315
               Left            =   90
               TabIndex        =   30
               Top             =   390
               Width           =   1065
            End
            Begin VB.Label Label2 
               Caption         =   "Semi major axis"
               Height          =   315
               Left            =   90
               TabIndex        =   28
               Top             =   795
               Width           =   1155
            End
         End
         Begin VB.TextBox txtEpochOrbit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   25
            Top             =   4050
            Width           =   1455
         End
         Begin VB.TextBox txtDecayRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   24
            Top             =   3690
            Width           =   1455
         End
         Begin VB.TextBox txtMeanMotion 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   23
            Top             =   3333
            Width           =   1455
         End
         Begin VB.TextBox txtMeanAnomoly 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   22
            Top             =   2976
            Width           =   1455
         End
         Begin VB.TextBox txtAOP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   21
            Top             =   2619
            Width           =   1455
         End
         Begin VB.TextBox txtEccentricity 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   20
            Top             =   2262
            Width           =   1455
         End
         Begin VB.TextBox txtRAAN 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   19
            Top             =   1905
            Width           =   1455
         End
         Begin VB.TextBox txtInclination 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   18
            Top             =   1548
            Width           =   1455
         End
         Begin VB.TextBox txtElementSet 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   17
            Top             =   1191
            Width           =   1455
         End
         Begin VB.TextBox txtEpochTime 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1620
            TabIndex        =   16
            Top             =   834
            Width           =   1455
         End
         Begin VB.TextBox txtCat 
            Height          =   285
            Left            =   1620
            TabIndex        =   15
            Top             =   477
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1620
            TabIndex        =   14
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Epoch orbit"
            Height          =   225
            Index           =   11
            Left            =   150
            TabIndex        =   13
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Decay rate"
            Height          =   225
            Index           =   10
            Left            =   150
            TabIndex        =   12
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Mean motion"
            Height          =   225
            Index           =   9
            Left            =   150
            TabIndex        =   11
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Mean anomoly"
            Height          =   225
            Index           =   8
            Left            =   150
            TabIndex        =   10
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Arg. of perigee"
            Height          =   225
            Index           =   7
            Left            =   150
            TabIndex        =   9
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Eccentricity"
            Height          =   225
            Index           =   6
            Left            =   150
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "RA Asc. Node"
            Height          =   225
            Index           =   5
            Left            =   150
            TabIndex        =   7
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Inclination"
            Height          =   225
            Index           =   4
            Left            =   150
            TabIndex        =   6
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Element set"
            Height          =   225
            Index           =   3
            Left            =   150
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Epoch time"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   4
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Catalog number"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Satellite name"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   120
            Width           =   1095
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   510
      Top             =   3390
   End
End
Attribute VB_Name = "frmSatDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nPos As Integer

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  
  With Me.SatTrackControl1
    .FrequencyDatabasePath = App.Path
    .ObserverLatitude = sProgramOptions.nLatitude
    .ObserverLongitude = sProgramOptions.nLongitude
    .ObserverHeight = sProgramOptions.nHeight
    .ObserverLocation = sProgramOptions.strLocation
    .ObserverMapCentre = 0
    .DisplayTimes = False
    .OutputStyle = 4
    
    .Timezone = 0
    .DaylightSaving = sProgramOptions.bDaylightSaving
    .AllowDoEvents = False
    .UseHourglass = False
  End With
End Sub


Private Sub lsvFreq_BeforeLabelEdit(Cancel As Integer)
  Cancel = 1
End Sub

Private Sub lsvFreq_Click()
  If Not (Me.lsvFreq.SelectedItem Is Nothing) Then
    Me.chkInverting.Value = IIf(Me.lsvFreq.SelectedItem.Text = "Yes", 1, 0)
    Me.txtUplink = Me.lsvFreq.SelectedItem.SubItems(1)
    Me.txtDownlink = Me.lsvFreq.SelectedItem.SubItems(2)
    Me.txtFreqDesc = Me.lsvFreq.SelectedItem.SubItems(3)
  End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
  Me.Timer1.Enabled = False
  nPos = Me.Tag
  SetupDisplay
  ReadFrequencies nPos
End Sub

Private Sub SetupDisplay()
  Dim dTemp As Double
  
  Me.Caption = "Details for " & sKeps(nPos).strName

  Me.txtName = sKeps(nPos).strName
  Me.txtCat = sKeps(nPos).lDesignator
  Me.txtEpochTime = sKeps(nPos).strEpoch
  Me.txtElementSet = sKeps(nPos).nElementSet
  Me.txtInclination = sKeps(nPos).dInclination
  Me.txtRAAN = sKeps(nPos).dRAAN
  Me.txtEccentricity = sKeps(nPos).dEccentricity
  Me.txtAOP = sKeps(nPos).dAOP
  Me.txtMeanAnomoly = sKeps(nPos).dMeanAnomoly
  Me.txtMeanMotion = sKeps(nPos).dMeanMotion
  Me.txtDecayRate = -sKeps(nPos).dDrag
  Me.txtEpochOrbit = sKeps(nPos).lOrbitNUmber

  UpdateDerived nPos
End Sub

Private Sub UpdateDerived(nPos As Integer)
  Dim sAltPer As Single
  Dim sAltApo As Single
  Dim dSMA As Double
  Dim dTemp As Double
  
  SetupKeps nPos
  
  Me.SatTrackControl1.SatelliteIndex = 1
  Me.SatTrackControl1.UpdateDerived 1
  
  Me.txtSMA = Format$(Me.SatTrackControl1.SatelliteSemiMajorAxis, "#####0.000")
  Me.txtLonOfNode = Me.SatTrackControl1.SatelliteLonOfNode
  
  Me.txtAltPerigee = Format$(Me.SatTrackControl1.SatelliteAltAtPerigee, "#########.0")
  Me.txtAltApogee = Format$(Me.SatTrackControl1.SatelliteAltAtApogee, "#########.0")
  Me.txtPeriod = Format$(Me.SatTrackControl1.SatellitePeriod, "####0.000")
  Me.txtSemiMInorAxis = Format$(Me.SatTrackControl1.SatelliteSemiMinorAxis, "#####0.000")

  ec = Me.SatTrackControl1.SatelliteSemiMajorAxis - (Me.SatTrackControl1.SatelliteSemiMajorAxis * sKeps(nPos).dEccentricity)
  sRe = ((cEarthRadius / Me.SatTrackControl1.SatelliteSemiMajorAxis) * Me.picOrbit.Width)
  With Me.picOrbit
    .Cls
    .ScaleTop = 0
    .ScaleLeft = 0
    .ScaleWidth = Int(Me.SatTrackControl1.SatelliteSemiMajorAxis + 50) * 2
    .ScaleHeight = Int(Me.SatTrackControl1.SatelliteSemiMinorAxis + 50) * 2
  
    .FillStyle = vbFSTransparent
    Me.picOrbit.Circle (Me.SatTrackControl1.SatelliteSemiMajorAxis, Me.SatTrackControl1.SatelliteSemiMinorAxis), Me.SatTrackControl1.SatelliteSemiMajorAxis, RGB(255, 0, 0), , , Me.SatTrackControl1.SatelliteSemiMinorAxis / Me.SatTrackControl1.SatelliteSemiMajorAxis
    .FillColor = RGB(0, 255, 0)
    .FillStyle = vbFSSolid
    Me.picOrbit.Circle (ec, Me.SatTrackControl1.SatelliteSemiMinorAxis), sRe
  End With

End Sub

Private Sub SetupKeps(nPos As Integer)

  With Me.SatTrackControl1
    .AddSatellite
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
    .DisplayTracks = False
  End With

End Sub

Private Sub ReadFrequencies(nPos)
  Dim nFile As Integer
  Dim strLine As String
  Dim vData() As Variant
  Dim sNode As ListItem
  Dim lCount As Long
  
  nFile = FreeFile
  Open App.Path & "\Frequencies\Frequencies.txt" For Input As #nFile
  While Not EOF(1)
    Line Input #nFile, strLine
    vData = StrParse(strLine, ",")
    If vData(0) = sKeps(nPos).lDesignator Then
      Set sNode = Me.lsvFreq.ListItems.Add(, "A" & lCount, IIf(Left(vData(3), 1) = "0", "Yes", "no"))
      sNode.SubItems(1) = vData(1)
      sNode.SubItems(2) = vData(2)
      sNode.SubItems(3) = Mid(vData(3), 3)
      lCount = lCount + 1
    End If
  Wend
  Close #nFile
End Sub

