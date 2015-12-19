VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVOICE.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SatTrackControl 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8640
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "AgTrack.ctx":0000
   ScaleHeight     =   6165
   ScaleWidth      =   8640
   ToolboxBitmap   =   "AgTrack.ctx":002B
   Begin VB.PictureBox picPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   90
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   16
      Top             =   4020
      Visible         =   0   'False
      Width           =   645
   End
   Begin SatTrack.ScrollingViewPort ViewPort 
      Height          =   2565
      Left            =   210
      TabIndex        =   13
      Top             =   270
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4524
      BackColor       =   14737632
      Begin VB.PictureBox picInner 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   6750
         Left            =   0
         ScaleHeight     =   450
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   790
         TabIndex        =   14
         Top             =   0
         Width           =   11850
         Begin VB.Label lblMousePos 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   5415
         End
      End
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5790
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   10
      Top             =   4380
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   3030
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   12
      Top             =   4350
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComctlLib.ImageList ImgMasks 
      Left            =   7290
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":033D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":3391
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":63E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":9439
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":C48D
            Key             =   "Iss"
            Object.Tag             =   "25544"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":F4E1
            Key             =   "Hubble"
            Object.Tag             =   "20580"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":12535
            Key             =   "Mir"
            Object.Tag             =   "16609"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgObjects 
      Left            =   6600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":15589
            Key             =   "Sun"
            Object.Tag             =   "-1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":185DD
            Key             =   "Moon"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":1B631
            Key             =   "Shuttle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":1E685
            Key             =   "Default"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":216D9
            Key             =   "Iss"
            Object.Tag             =   "25544"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":2472D
            Key             =   "Hubble"
            Object.Tag             =   "20580"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AgTrack.ctx":27781
            Key             =   "Mir"
            Object.Tag             =   "16609"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4290
      ScaleHeight     =   192
      ScaleMode       =   0  'User
      ScaleWidth      =   118.044
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1245
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS TXSpeech 
      Height          =   315
      Left            =   8220
      OleObjectBlob   =   "AgTrack.ctx":2A7D5
      TabIndex        =   9
      Top             =   3420
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Timer tmrAutoUpdate 
      Enabled         =   0   'False
      Left            =   1290
      Top             =   3570
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   5850
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Local time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "UTC/GMT"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView TabList 
      Height          =   2415
      Left            =   150
      TabIndex        =   7
      Top             =   4950
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4260
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Designator"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Orbit"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Latitude"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Longitude"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Elevation"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Azimuth"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Range"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Height"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Doppler"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Next AOS"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Model"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox DefMapHorizon 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   885
      Left            =   6540
      Picture         =   "AgTrack.ctx":2A82D
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Map3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   4
      Top             =   3030
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Map180 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5250
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   3
      Top             =   1740
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Map0 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5160
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox DefMap180 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   1305
      Left            =   6570
      Picture         =   "AgTrack.ctx":3E333
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox DefMap0 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   1515
      Left            =   6540
      Picture         =   "AgTrack.ctx":45494
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblSatPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   3675
   End
   Begin VB.Menu mnuPopupSat 
      Caption         =   "Sat Name Here"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSatName 
         Caption         =   "Sat Name"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupNextAOS 
         Caption         =   "Display next AOS"
      End
      Begin VB.Menu mnuPopupPreviousAOS 

         Caption         =   "Display previous AOS"
      End
      Begin VB.Menu mnuPopupReset 
         Caption         =   "Reset Satellite"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Set Mode"
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModelS 
         Caption         =   "Models"
         Begin VB.Menu mnuModelSGP 
            Caption         =   "SGP4/SDP4"
         End
         Begin VB.Menu mnuModelPlan13 
            Caption         =   "Plan13"
         End
         Begin VB.Menu mnuModelDTSGP 
            Caption         =   "Alt SGP"
         End
      End
   End
End
Attribute VB_Name = "SatTrackControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Private mDx7 As DirectX7
'
'Private Const VIEWPORT_WIDTH = 640
'Private Const VIEWPORT_HEIGHT = 480
'
'Public D3DRM As IDirect3DRM
'Public D3DRM As Direct3DRM3
'
'Private mDrw As DirectDraw7
'Private mDrm As Direct3DRM3
'Private mFrS As Direct3DRMFrame3
'Private mFrC As Direct3DRMFrame3
'Private mFrO As Direct3DRMFrame3
'Private mFrL As Direct3DRMFrame3
'Private mDev As Direct3DRMDevice3
'Private mVpt As Direct3DRMViewport2
'
'Private mDownX As Single
'Private mDownY As Single
'Private oMX As Single
'Private oMY As Single
'Private mStopFlag As Boolean
'Private mMouseDown As Boolean
'
'Private Type dxPTM
'    dX As Single
'    dY As Single
'    Distance As Single
'End Type

Private Const MaxPoints = 2000

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private nStartX As Long
Private nStartY As Long
Private nLastx As Long
Private nLasty As Long
Private bStarted As Boolean

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
     Const SRCAND = &H8800C6 'Combines pixels of the destination and source
    'bitmap using the Boolean AND operator.
     Const SRCCOPY = &HCC0020 'Copies the source bitmap to the destination
    'bitmap.
     Const SRCPAINT = &HEE0086 'Combines pixels of the destination and source
    'bitmap using the Boolean OR operator.
    
Private Type tagSATELLITE
  cSatelliteName As String * 23
  iSecondMeanMotion As Long
  iSatelliteNumber As Long
  iLaunchYear As Long
  iLaunchNumber As Long
  cLaunchPiece  As String * 3
  iEpochYear As Long
  fEpochDay As Double
  iEpochDay As Long
  fEpochFraction As Double
  fBalisticCoefficient As Double
  fSecondMeanMotion As Double
  fRadiationCoefficient As Double
  EmphemeristType As String * 2
  iElementNumber As Long

  fInclination As Double
  fRightAscending As Double
  fEccentricity As Double
  fPeregee As Double
  fMeanAnomaly As Double
  fMeanMotion As Double
  iRevAtEpoch As Long
  fJulianEpoch As Double
End Type
  
Type tagMisc
  dOrbit As Double
  nDeep As Long
End Type

Type tagVECTOR
  x As Double
  y As Double
  z As Double
  w As Double
End Type

Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type


  Dim sgpSat As tagSATELLITE
  Dim vOPos As tagVECTOR
  Dim vPos As tagVECTOR
  Dim vVel As tagVECTOR
  Dim vObs As tagVECTOR
  Dim vSatPos As tagVECTOR
  Dim strTemp As String
  Dim sTime As SYSTEMTIME
  Dim sSat As tagSATELLITE



'Private Declare Function AGCalcSGP Lib "AGSGP.dll" (sgpSat As tagSATELLITE, sTime As SYSTEMTIME, a As tagVECTOR, b As tagVECTOR, c As tagVECTOR, d As tagVECTOR, e As tagVECTOR) As Boolean
Private Declare Function AGCalcSGP Lib "AGSGP.dll" (ByVal strLine0 As String, ByVal strLine1 As String, ByVal strLine2 As String, sTime As SYSTEMTIME, a As tagVECTOR, b As tagVECTOR, c As tagVECTOR, d As tagVECTOR, e As tagVECTOR, m As tagMisc) As Boolean
Private Declare Function AGSGPGetVersion Lib "AGSGP.dll" () As Integer

Public Enum Models
  agplan13
  agSGP
  agdtsgp
End Enum

Public Enum ModelTypes
  agTypeSGP
  agTypeSGPD
  agTypeDTSGP
End Enum

Public Enum OS
  agTabList = 0
  agMercator = 1
  agHorizon = 2
  agGlobe = 3
  agInvisible = 4
End Enum

Public Enum TI
  ag60 = 60
  ag30 = 30
  ag15 = 15
  agAuto = 1
End Enum

Private Declare Function GenerateGlobe Lib "AgGlobe.dll" (ByVal dLat As Double, ByVal dLon As Double, ByVal nHeight As Long, ByVal nWidth As Long, ByVal bLabel As Boolean, ByVal bLocations As Boolean, ByVal strLocsFile As String, ByVal dDay As Long, ByVal dNight As Long, ByVal dTerminator As Long, ByVal bShade As Boolean, ByVal bStars As Boolean, ByVal bGrid As Boolean, ByVal lTime As Long, ByVal lSatLat As Long, ByVal lSatLon As Long, ByVal strSatName As String, TrackLat() As Double, TrackLon() As Double, lColours() As Long) As Boolean

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20

Private strTempPath As String

Private Const gc0528 = 75369793000000#

Private cIntDeg As Variant
Private Sunx As Variant
Private Suny As Variant
Private Sunz As Variant
Private bGotSun As Boolean
  
Private SatMeanMotion As Double
Private SatMeanMotionMinute As Double
  Private SatLinearDrag As Double
  Private DR As Double
  Private EA As Double
  Private c As Double
  Private S As Double
  Private DNOM As Double
  Private d As Double
  Private a As Double
  Private b As Double
  Private cw As Double
  Private SW As Double
  Private CQ As Double
  Private VELx As Double
  Private VELy As Double
  Private VELz As Double
  Private U As Double
  Private e As Double
  Private N As Double
  Private ElapsedTimeSinceEpoch As Double

Private cSunRise As clsSunrise
Private cWaitCur As CWaitCursor

Private Const ONEPPM = 0.000001
Private Const CVAC = 299792.458

Private HalfMapWidth   As Integer
Private HalfMapHeight As Integer
Private PixelsPerDegLon As Single
Private PixelsPerDegLat As Single
Private MapWidth As Integer
Private MapHeight As Integer

' SUN Stuff
Private MAS0 As Double
Private MASD As Double

Dim SatMouseOver As Integer
Private PI As Variant
Private MeanYear As Double
Private TropicalYear As Double
Private EarthRotationRate As Double
Private EarthRotationRateDay As Double
Private EarthRotationRateSeconds As Double
Private SatDragCoeff As Double
Private SatKepsMeanAnomoly As Double
Private SatKepsMeanMotion As Double
Private SatKepsRAAN As Double
Private SatKepsArgOfPerigee As Double

Private GravitationalConstant As Double
Private ZonalCoeff As Double
Private YG As Double
Private G0 As Double
Private LA As Double
Private LO As Double
Private HT As Double
Private OLDRN As Integer
Private SatEpochDayNumber As Long
Private SatelliteTimeRequired As Double

Private RS() As Double
Private CL As Double
Private SL As Double
Private CO As Double
Private SO As Double
Private RE As Double
Private FL As Double
Private RP As Double
Private Rx As Double
Private Ry As Double
Private Rz As Double
Private Ex As Double
Private Ey As Double
Private Ez As Double
Private sX As Double
Private sY As Double
Private Sz As Double
Private Ux As Double
Private Uy As Double
Private Uz As Double
Private oX As Double
Private oY As Double
Private Oz As Double
Private Ax As Double
Private Ay As Double
Private Az As Double
Private Nx As Double
Private Ny As Double
Private Nz As Double
Private VOx As Double
Private VOy As Double
Private M2 As Double
Private CI As Double
Private SI As Double
Private B0 As Double
Private QD As Double
Private WD As Double
Private TEG As Double
Private GHAE As Double
Private N0 As Double

Private lItem As ListItem

Private LastCursorX As Integer
Private LastCursorY As Integer

Private PointsToDraw As Integer
Private TwoPI As Variant
Private FootprintLON() As Single
Private FootprintLAT() As Single

Private TempSatAlt As Double
Private TempSatMA As Double
Private TempSatLon As Double
Private TempSatlat As Double
Private TempSatAz As Double
Private TempSatUplinkDoppler As Double
Private TempSatDownlinkDoppler As Double
Private TempSatElev As Double
Private TempSatOrbit As Double
Private TempSatDayNum As Double
Private TempSatTimeReq As Double
Private TempSatRange As Double
Private TempRS As Double
Private TempDisplayYear As Double
Private TempDisplayMonth As Double
Private TempDisplayDay As Double
Private TempDisplayHour As Double
Private TempDisplayMin As Double
Private TempDisplaySecond As Double
Private TempSatmaxDx As Double
Private TempSatPathLoss As Double
Private TempObsLat As Double
Private TempObsLon As Double
Private TempSquintAngle As Double
Private TempRangeRate As Double
Private TempSatStatusText As String

Private mvarObserverLatitude As Double 'local copy
Private mvarObserverLongitude As Double 'local copy
Private mvarObserverHeight As Double 'local copy
Private mvarSecondObserverLatitude As Double 'local copy
Private mvarSecondObserverLongitude As Double 'local copy
Private mvarSecondObserverHeight As Double 'local copy
Private mvarKepsYearEpochTime() As Double
Private mvarSatEpochYear() As Double
Private mvarKepsYearEpochTimeFraction() As Double
Private m_SecondObserverEnabled As Boolean

Private nDaylightSaving As Integer


'Default Property Values:
Const m_def_SunRise = 0
Const m_def_SunNoon = 0
Const m_def_SunSet = 0
Const m_def_GroundTrackPointSize = 0
Const m_def_UserStatusPanelText = "0"
Const m_def_SatelliteBearing = 0
Const m_def_SecondObserverEnabled = False
Const m_def_DisplayAOSCircle = False
Const m_def_DisplayIcons = 0
Const m_def_GroundTrackInterval = 1
Const m_def_CalculationModel = 1
Const m_def_DisplayGroundTrackAsPoints = 1
Const m_def_EnableSatStatus = 0
Const m_def_EnableSpeech = 0
Const m_def_DisplayStatusBar = True
Const m_def_DisplaySatelliteLabel = True
Const m_def_SetActiveWindowAsWallpaper = False
Const m_def_SatPosLabelAlign = 2
Const m_def_FT847CATSettings = "1,57600,n,8,1"
Const m_def_MaxWidth = 0
Const m_def_MaxHeight = 0
Const m_def_ViewsOrthLocations = ""
Const m_def_ViewsOrthShade = True
Const m_def_DaylightSavingAdjust = 0
Const m_def_TimeZoneName = "0"
Const m_def_UseHourglass = 0
Const m_def_SatelliteInAOS = 0
'Const m_def_SatelliteInAOS = 0
Const m_def_AllowDoEvents = False
Const m_def_SetSelectedSatellite = 0
Const m_def_SetIndexOnSelect = 0
Const m_def_FrequencyDatabasePath = ""
Const m_def_DisplayDataFields = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,31,32"
Const m_def_AutoMode = 0
'Const m_def_AutoMode = 0
Const m_def_AutoInterval = 0
Const m_def_UplinkFrequency = 0
Const m_def_DownLinkFrequency = 0
Private Const m_def_Timezone = 0
Const m_def_DaylightSaving = 0
Const m_def_Enable847Sat = 0
Const m_def_Enable847 = 0
Const m_def_PortSettings = "57600,N,8,1"
Const m_def_SatelliteStatusText = 0
Const m_def_SquintAngle = 0
Const m_def_RangeRate = 0
Const m_def_Busy = 0
Const m_def_ocxBusy = 0
Const m_def_DatabasePath = ""
Const m_def_SetAOSLOS = -1
Const m_def_CurrentSelectedSatellite = 0
Const m_def_SelectedSatellite = 0
Const m_def_IsSatLoaded = 0
Const m_def_SatelliteMaximumDX = 0
Const m_def_SelectedSatelliteName = ""
Const m_def_DisplaySecond = 0
Const m_def_DisplayTracks = True
Const m_def_DisplayFootprints = True
Const m_def_DisplayTimes = False
Const m_def_DisplaySun = False
Const m_def_DisplayMoon = False
Const m_def_DisplaySunFootprint = False
Const m_def_DisplayMoonFootprint = False
Const m_def_SatelliteBusy = 0
Const m_def_SatelliteTXFrequency = 0
Const m_def_SatelliteRxFrequency = 0
Const m_def_DataValid = 0
Const m_def_SatelliteCount = 0
Const m_def_DisplayTimeRequired = 0
Const m_def_SatelliteDayNumber = 0
Const m_def_DisplayMinute = 0
Const m_def_DisplayHour = 0
Const m_def_DisplayDay = 0
Const m_def_DisplayMonth = 0
Const m_def_DisplayYear = 0
Const m_def_DisplayCentury = 0
Const m_def_SatelliteIndex = 0
Const m_def_SatelliteOrbitNumber = 0
Const m_def_SatelliteLongitude = 0
Const m_def_SatelliteLatitude = 0
Const m_def_KepsChecksum = 0
Const m_def_KepsDecayRate = 0
Const m_def_KepsElementSet = 0
Const m_def_KepsEpochTime = 0
Const m_def_KepsOrbitNumber = 0
Const m_def_KepsMeanMotion = 0
Const m_def_KepsMeanAnomoly = 0
Const m_def_KepsInclination = 0
Const m_def_KepsAOP = 0
Const m_def_KepsRAAN = 0
Const m_def_KepsEccentricity = 0
Const m_def_SatelliteElevation = 0
Const m_def_SatelliteAzimuth = 0
Const m_def_SatelliteRange = 0
Const m_def_SatelliteDesignator = 0
Const m_def_SatelliteName = 0
Const m_def_ObserverLocation = "London"
Const m_def_SecondObserverLocation = "London"
Const m_def_OutputStyle = 1
Const m_def_ObserverMapCentre = 0

Const m_def_ObserverLatitude = 0.872664625997165
Const m_def_ObserverLongitude = 0
Const m_def_ObserverHeight = 0
Const m_def_SecondObserverLatitude = 0.872664625997165
Const m_def_SecondObserverLongitude = 0
Const m_def_SecondObserverHeight = 0

'Property Variables:
Dim m_SunRise As Variant
Dim m_SunNoon As Variant
Dim m_SunSet As Variant
Dim m_GroundTrackPointSize As Integer
Dim m_UserStatusPanelText As String
Dim m_SatelliteBearing As Single
Private m_CalculationModel As Integer
Private m_DisplayAOSCircle As Boolean
Private m_DisplayIcons As Boolean
Private m_GroundTrackInterval As Integer
Private m_DisplayGroundTrackAsPoints As Boolean
Private m_EnableSatStatus As Boolean
Private m_EnableSpeech As Boolean
Private m_DisplayStatusBar As Boolean
Private m_DisplaySatelliteLabel As Boolean
Private m_SetActiveWindowAsWallpaper As Boolean
Private m_SatPosLabelAlign As Integer
Private m_FT847CATSettings As String
Private m_MaxWidth As Integer
Private m_MaxHeight As Integer
Private m_ViewsOrthLocations As String
Private m_ViewsOrthShade As Boolean
Private m_SatellitePeriod() As Double
Private m_SatelliteSemiMajorAxis() As Double
Private m_SatelliteSemiMinorAxis() As Double
Private m_SatelliteLonOfNode() As Double
Private m_SatelliteAltAtPerigee() As Double
Private m_SatelliteAltAtApogee() As Double
Private m_DaylightSavingAdjust As Integer
Private m_TimeZoneName As String
Private m_UseHourglass As Boolean
Private m_SatelliteInAOS() As Boolean
Private m_AllowDoEvents As Boolean
Private m_SetSelectedSatellite As Integer
Private m_SetIndexOnSelect As Boolean
Private m_FrequencyDatabasePath As String
Private m_DisplayDataFields As String
Private m_AutoMode As Boolean
Private m_SatelliteTrackOrbits() As Integer
Private m_SatMaxEleTime() As Variant
'private m_AutoMode As Variant
Private m_AutoInterval As Variant
Private m_UplinkFrequency() As Double
Private m_DownLinkFrequency() As Double
Private m_Timezone As Variant
Private m_DaylightSaving As Boolean
Private m_Enable847Sat As Boolean
Private m_Enable847 As Boolean
Private m_PortSettings As Variant
Private m_SatelliteStatusText() As String
Private m_SquintAngle() As Variant
Private m_RangeRate() As Variant
Private m_Busy As Boolean
Private m_ocxBusy As Variant
Private m_DatabasePath As Variant
Private m_SetAOSLOS As Variant
Private m_CurrentSelectedSatellite As Variant
Private m_SelectedSatellite As Variant
Private m_IsSatLoaded As Variant
Private m_SatelliteMaximumDX() As Variant
Private m_SelectedSatelliteName As String
Private m_DisplaySecond() As Integer
Private m_DisplayTracks As Boolean
Private m_DisplayFootprints As Boolean
Private m_DisplayTimes As Boolean
Private m_DisplaySun As Boolean
Private m_DisplayMoon As Boolean
Private m_DisplaySunFootprint As Boolean
Private m_DisplayMoonFootprint As Boolean
Private m_SatelliteBusy() As Boolean
Private m_SatelliteTXFrequency() As Double
Private m_SatelliteRxFrequency() As Double
Private m_SatellitePathLoss() As Double
Private m_DataValid() As Boolean
Private m_SatelliteCount As Variant
Private m_DisplayTimeRequired() As Variant
Private m_SatelliteDayNumber() As Variant
Private m_DisplayMinute() As Integer
Private m_DisplayHour() As Integer
Private m_DisplayDay() As Integer
Private m_DisplayMonth() As Integer
Private m_DisplayYear() As Integer
Private m_DisplayCentury() As Integer
Private m_SatelliteIndex As Integer
Private m_SatelliteOrbitNumber() As Double
Private m_SatelliteLongitude() As Single
Private m_SatelliteLatitude() As Single
Private m_SatelliteAltitude() As Single
Private m_KepsChecksum() As Variant
Private m_KepsDecayRate() As Single
Private m_KepsElementSet() As Variant
Private m_KepsEpochTime() As Variant
Private m_KepsOrbitNumber() As Variant
Private m_KepsMeanMotion() As Double
Private m_KepsMeanAnomoly() As Single
Private m_KepsInclination() As Single
Private m_KepsAOP() As Single
Private m_KepsRAAN() As Single
Private m_KepsEccentricity() As Single
Private m_fRadiationCoefficient() As Double
Private m_SatelliteElevation() As Long
Private m_SatelliteAzimuth() As Long
Private m_SatelliteRange() As Long
Private m_SatelliteDesignator() As String
Private m_SatelliteName() As String
Private m_SatelliteMA() As Double
Private m_OrbitalModel() As Models
Private m_OrbitalModelType() As ModelTypes
Private m_ObserverLocation As String
Private m_SecondObserverLocation As String
Private m_OutputStyle As Integer
Private m_ObserverMapCentre As Integer

Private SatScreenX() As Integer
Private SatScreenY() As Integer
Private SatTrackTime() As Variant
Private SatTrackLon() As Double
Private SatTrackLat() As Double
Private SatTrackElev() As Integer
Private SatTrackAzim() As Integer
Private SatTrackMutual() As Long
Private SatTrackPoints() As Integer
Private satTrackNextAOS() As String
Private satTrackMaxEle() As Integer
Private satTrackMaxEleTime() As String
Private m_strLine0() As String
Private m_strLine1() As String
Private m_strLine2() As String

Private SelectedSatellite As Integer

Private DaysInMonth(13) As Integer

Type dX
  Callsign As String
  strLat As String
  strLon As String
  strName As String
End Type

Dim sDxDetails(1100) As dX

Private nSatCount As Integer

Private cDJTSGP As Object

'Event Declarations:
Event SatelliteSelected(nIndex As Integer)
Attribute SatelliteSelected.VB_Description = "Fires when a satellite is selected."
Event SatelliteAtAOS(SatelliteIndex As Integer)
Event SatelliteUpdated(nSatIndex As Integer)
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Event MouseOverSatellite(Index As Integer, Name As String)
Event Resize()

Public Property Let ObserverHeight(ByVal vData As Variant)
  mvarObserverHeight = vData / 1000
  PropertyChanged "ObserverHeight"
  vOPos.z = mvarObserverHeight
End Property

Public Property Get ObserverHeight() As Variant
Attribute ObserverHeight.VB_ProcData.VB_Invoke_Property = "Observer"
  ObserverHeight = mvarObserverHeight * 1000
End Property

Public Property Let ObserverLongitude(ByVal vData As Variant)
  mvarObserverLongitude = FNRAD(vData)
  PropertyChanged "ObserverLongitude"
  CalculateSunRise
'  vOPos.y = FNDEG(mvarObserverLongitude)
  vOPos.y = mvarObserverLongitude
End Property

Public Property Get ObserverLongitude() As Variant
Attribute ObserverLongitude.VB_ProcData.VB_Invoke_Property = "Observer"
  ObserverLongitude = Round(FNDEG(mvarObserverLongitude), 2)
End Property

Public Property Let ObserverLatitude(ByVal vData As Variant)
  mvarObserverLatitude = FNRAD(vData)
  PropertyChanged "ObserverLatitude"
  CalculateSunRise
'  vOPos.x = FNDEG(mvarObserverLatitude)
  vOPos.x = mvarObserverLatitude
End Property

Public Property Get ObserverLatitude() As Variant
Attribute ObserverLatitude.VB_ProcData.VB_Invoke_Property = "Observer"
  ObserverLatitude = Round(FNDEG(mvarObserverLatitude), 2)
End Property
Public Property Get ObserverMapCentre() As Integer
Attribute ObserverMapCentre.VB_ProcData.VB_Invoke_Property = "Observer"
    ObserverMapCentre = m_ObserverMapCentre
End Property

Public Property Let ObserverMapCentre(ByVal New_ObserverMapCentre As Integer)

    m_ObserverMapCentre = New_ObserverMapCentre
    If m_ObserverMapCentre = 180 Then
        picInner.Picture = Map180.Picture
    Else
        picInner.Picture = Map0.Picture
    End If
    PropertyChanged "ObserverMapCentre"
    DrawFootprints

End Property

Public Property Get SatelliteTXFrequency() As Double
    SatelliteTXFrequency = m_SatelliteTXFrequency(m_SatelliteIndex)
End Property

Public Property Let SatelliteTXFrequency(ByVal New_SatelliteTXFrequency As Double)
    m_SatelliteTXFrequency(m_SatelliteIndex) = New_SatelliteTXFrequency
    PropertyChanged "SatelliteTXFrequency"
End Property

Public Property Get SatelliteRxFrequency() As Double
    SatelliteRxFrequency = m_SatelliteRxFrequency(m_SatelliteIndex)
End Property

Public Property Let SatelliteRxFrequency(ByVal New_SatelliteRxFrequency As Double)
    If Ambient.UserMode = False Then Err.Raise 382
    m_SatelliteRxFrequency(m_SatelliteIndex) = New_SatelliteRxFrequency
    PropertyChanged "SatelliteRxFrequency"
End Property


Private Sub lblSatPos_DblClick()
  m_SatPosLabelAlign = m_SatPosLabelAlign + 1
  If m_SatPosLabelAlign > 2 Then
    m_SatPosLabelAlign = 0
  End If
End Sub
Private Sub mnuModelSGP_Click()
  m_OrbitalModel(SatMouseOver) = agSGP
End Sub
Private Sub mnuModelPlan13_Click()
  m_OrbitalModel(SatMouseOver) = agplan13
End Sub
Private Sub mnuModelDTSGP_Click()
  m_OrbitalModel(SatMouseOver) = agdtsgp
End Sub

Private Sub mnuMode_Click()
  On Error GoTo ERROR_mnuMode_Click

  Dim nFile As Integer
  Dim nDownlink(100) As Double
  Dim nUplink(100) As Double
  Dim strMode(100) As String
  Dim nCount As Integer
  Dim strData As String
  Dim vData() As Variant
  Dim vdata1() As Variant
  Dim strDesignator As String
  Dim nSatPos As Integer
  Dim strPath As String
  
  If m_FrequencyDatabasePath <> "" Then
    strPath = m_FrequencyDatabasePath & "\Frequencies\Frequencies.txt"
  Else
    strPath = App.Path & "\Frequencies.txt"
  End If
  
  nSatPos = SatMouseOver
  nFile = FreeFile
  strDesignator = m_SatelliteDesignator(nSatPos)
  Open strPath For Input As #nFile
  While Not EOF(nFile)
    Line Input #nFile, strData
    vData = StrParse(strData, ",")
    If vData(0) = strDesignator Then
      nDownlink(nCount) = vData(2)
      nUplink(nCount) = vData(1)
      strData = vData(3)
      vdata1 = StrParse(strData, ";")
      frmModeSelect.lstModes.AddItem Left(vdata1(1), 15) & " " & Format(nDownlink(nCount), "0000.00000") & " " & Format(nUplink(nCount), "0000.00000")
      nCount = nCount + 1
    End If
  Wend
  Close #nFile

  If nCount = 0 Then
    Call MsgBox("There are no details in the frequencies.txt file for this satellite.", vbInformation + vbOKOnly + vbDefaultButton1, "Satellite mode select error")
  Else
    frmModeSelect.Show vbModal
    If bOpt <> -1 Then
      m_DownLinkFrequency(nSatPos) = nDownlink(bOpt)
      m_UplinkFrequency(nSatPos) = nUplink(bOpt)
      CalculateSatellitePosition True, nSatPos
      UpdateDataWindow
    End If
  End If
EXIT_mnuMode_Click:
  Exit Sub

ERROR_mnuMode_Click:
  Select Case Err
    Case 52
      Call MsgBox("There is no Frequencies.txt file in the main program directory. Please refer to the help file for detais on creating this file. Mode selection cannot be made until this file is created.", vbExclamation + vbOKOnly + vbDefaultButton1, "Mode Selection error")
    Case Else
      MsgBox "Error in ERROR_mnuMode_Click : " & Error
      Resume EXIT_mnuMode_Click
  End Select
  Close #nFile
End Sub

Private Sub mnuPopupData_Click()
'  bSatDetailsVisible = True
'  SETtopmostwindow SatDetails, True
'  UpdateDataWindow
'  SatDetails.Show
End Sub

Private Sub mnuPopupNextAOS_Click()
  DisplayAOS 2, SatMouseOver
End Sub

Private Sub mnuPopupPreviousAOS_Click()
  DisplayAOS 1, SatMouseOver
End Sub

Private Sub mnuPopupReset_Click()
  ResetSatellite SatMouseOver
End Sub

Private Sub picInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  If (Button And vbLeftButton) And SatMouseOver = -2 And m_OutputStyle = 1 Then
    nStartX = x
    nStartY = y
    nLastx = x
    nLasty = y
    bStarted = True
    picInner.ForeColor = vbBlack
    picInner.DrawMode = vbNotXorPen
    picInner.Line (nStartX, nStartY)-(nLastx, nLasty)
    picInner.DrawMode = vbCopyPen
  End If
  If m_OutputStyle = 3 Then
    mMouseDown = True
    mDownX = x
    mDownY = y
    'HitTest X, Y
    oMX = x
    oMY = y
End If
End Sub

Private Sub picInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If m_OutputStyle = 3 Then
      If Not mMouseDown = True Then Exit Sub
    If Button = 1 Then
   '   Rotate x, y, False
    ElseIf Button = 2 Then
'        mFrO.AddTranslation D3DRMCOMBINE_BEFORE, (X - omX), 0, (Y - omY)
'        mFrO.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, (y - oMY) / 1000
'        RefreshGlobe
    End If
    oX = x
    oY = y

  Else
  If bStarted Then
    picInner.DrawMode = vbNotXorPen
    picInner.Line (nStartX, nStartY)-(nLastx, nLasty)
    picInner.Line (nStartX, nStartY)-(x, y)
    picInner.DrawMode = vbCopyPen
    nLastx = x
    nLasty = y
    MousePos False, x, y
  Else
    Select Case m_OutputStyle
      Case 1, 2
        MousePos False, x, y
      Case Else
        lblMousePos.Caption = ""
    End Select
  End If
End If
End Sub
Private Sub MousePos(Op As Boolean, x As Single, y As Single)
  Dim tLat As Single
  Dim tLon As Single
  Dim SatX As Integer
  Dim SatY As Integer
  Dim i As Integer
  Dim SatPos As Integer
  Dim SatName As String
  Dim bFound As Boolean
  Dim strTemp As String
  
  tLat = ConvertScreenToLAT(y)
  tLon = ConvertScreenToLON(x)
  
  strTemp = ConvertMousePos(tLat, tLon)
  If x <> LastCursorX Or y <> LastCursorY Then
    lblMousePos.Caption = strTemp
    LastCursorX = x
    LastCursorY = y
  End If
  
  SatMouseOver = -2
  Screen.MousePointer = vbDefault
  For i = 1 To nSatCount
    If m_SatelliteName(i) = "" Then Exit For

    SatX = SatScreenX(i)
    SatY = SatScreenY(i)

    If SatX > x - 4 And SatX < x + 4 Then
      If SatY > y - 4 And SatY < y + 4 Then
        SatMouseOver = i
        lblMousePos.Caption = m_SatelliteName(i) & strTemp
        RaiseEvent MouseOverSatellite(i, m_SatelliteName(i))
        bFound = True
        If Op Then
          Screen.MousePointer = vbDefault
          nSatDetailsTag = i
          mnuPopupSatName.Caption = m_SatelliteName(i)
          Select Case m_OrbitalModel(i)
            Case agSGP
              mnuModelSGP.Checked = True
              mnuModelPlan13.Checked = False
              mnuModelDTSGP.Checked = False
            Case agplan13
              mnuModelSGP.Checked = False
              mnuModelPlan13.Checked = True
              mnuModelDTSGP.Checked = False
            Case agdtsgp
              mnuModelSGP.Checked = False
              mnuModelPlan13.Checked = False
              mnuModelDTSGP.Checked = True
          End Select
          If m_SatelliteBusy(i) Then
            mnuPopupReset.Enabled = True
          Else
            mnuPopupReset.Enabled = False
          End If
          PopupMenu mnuPopupSat
        Else
          Screen.MousePointer = vbCrosshair
        End If
        Exit For
      End If
    End If
  Next i

End Sub
Private Function ConvertMousePos(sLat As Single, sLon As Single) As String
  Dim strTemp As String
  Dim strTemp1 As String
  Dim dDist As Double
  Dim sStartLat As Single
  Dim sStartLong As Single
  Dim sEndLat As Single
  Dim sEndLon As Single
  Dim sBearing As Single

  If sLat < 0 Then
    strTemp = "Lat: " & Format(Str(Abs(sLat)), "##0.00") & "°S  "
  Else
    strTemp = "Lat: " & Format(Str(sLat), "##0.00") & "°N  "
  End If
  If sLon < 0 Then
    strTemp = strTemp & "Lon: " & Format(Str(Abs(sLon)), "##0.00") & "°W  "
  Else
    strTemp = strTemp & "Lon: " & Format(Str(sLon), "##0.00") & "°E  "
  End If

  If bStarted Then
    sStartLat = FNRAD(ConvertScreenToLAT(CSng(nStartY)))
    sStartLong = FNRAD(ConvertScreenToLON(CSng(nStartX)))
    If sLat < 0 Then
      strTemp1 = "Lat: " & Format(Str(Abs(sStartLat)), "##0.00") & "°S  "
    Else
      strTemp1 = "Lat: " & Format(Str(sStartLat), "##0.00") & "°N  "
    End If
    If sLon < 0 Then
      strTemp1 = strTemp1 & "Lon: " & Format(Str(Abs(sStartLong)), "##0.00") & "°W  "
    Else
      strTemp1 = strTemp1 & "Lon: " & Format(Str(sStartLong), "##0.00") & "°E  "
    End If
    sEndLat = FNRAD(sLat)
    sEndLon = FNRAD(sLon)
    sStartLat = ConvertScreenToLAT(CSng(nStartY))
    sStartLong = ConvertScreenToLON(CSng(nStartX))
    sEndLat = sLat
    sEndLon = sLon
    dDist = CalculateDistAndBearing(CDbl(sStartLat), CDbl(sStartLong), CDbl(sEndLat), CDbl(sEndLon), sBearing)
    strTemp = strTemp1 & " -> " & strTemp & " Dist " & Int(dDist) & "Km"
  End If

  ConvertMousePos = strTemp
End Function
Private Function CalculateDistAndBearing(dStartY As Double, dStartX As Double, dEndY As Double, dEndX As Double, sBearing As Single) As Single
Dim d1 As Double
Dim d2 As Double
Dim d3 As Double
Dim d4 As Double
Dim d5 As Double
Dim d6 As Double
Dim d7 As Double
Dim d8 As Double
Dim d9 As Double
Dim d10 As Double
Dim d11 As Double
Dim gvAtnFour As Double
Dim gvAtnFour180 As Double
Dim gvAtn180Four As Double
Dim d12 As Double
Dim d13 As Double
Dim md1 As Double
Dim md2 As Double

d12 = 6378
d13 = 6356
md1 = (d12 + d13) / 2
md2 = d13 * d13 / (d12 * d12)

gvAtnFour = 4 * Atn(1)
gvAtnFour180 = gvAtnFour / 180#
gvAtn180Four = 180# / gvAtnFour


If dStartY = dEndY And dStartX = dEndX Then Exit Function
d1 = dStartY * gvAtnFour180
d2 = dStartX * gvAtnFour180
d3 = dEndY * gvAtnFour180
d4 = dEndX * gvAtnFour180
d5 = Atn(md2 * Tan(d1))
d6 = Atn(md2 * Tan(d3))
d7 = Cos(d2 - d4) * Cos(d5) * Cos(d6) + Sin(d5) * Sin(d6)
d8 = Atn(Abs(Sqr(1 - d7 * d7) / d7))
If d7 < 0 Then d8 = gvAtnFour - d8
d9 = md1 * d8
d11 = Sin(d4 - d2) * Cos(d6) * Cos(d5)
d7 = Sin(d6) - Sin(d5) * Cos(d8)
If d7 = 0 Then d10 = gvAtnFour / 2 Else d10 = Atn(Abs(d11 / d7))
If d7 < 0 Then d10 = gvAtnFour - d10
If d11 < 0 Then d10 = -d10
If d10 < 0 Then d10 = d10 + 2 * gvAtnFour
sBearing = d10 * gvAtn180Four
CalculateDistAndBearing = d9
End Function
Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMouseDown = False
  '  Rotate x, y, True

  If bStarted Then
    picInner.DrawMode = vbNotXorPen
    picInner.Line (nStartX, nStartY)-(nLastx, nLasty)
    picInner.DrawMode = vbCopyPen
    bStarted = False
  End If
  If Button And vbRightButton Then
    MousePos True, x, y
  End If
  If Button And vbLeftButton Then
    If SatMouseOver <> -2 Then
      If m_SetIndexOnSelect Then
        m_SatelliteIndex = SatMouseOver
      End If
      SelectedSatellite = SatMouseOver
      RaiseEvent SatelliteSelected(SelectedSatellite)
      UpdateSatelliteLabel
    End If
  End If

End Sub

Private Sub picInner_Paint()
 If m_OutputStyle = 3 Then
'  RefreshGlobe
 End If
End Sub

Private Sub TabList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  TabList.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub TabList_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim i As Integer

  For i = 0 To nSatCount
    If Item.Text = m_SatelliteName(i) Then
      If m_SetIndexOnSelect Then
        m_SatelliteIndex = i
      End If
      SelectedSatellite = i
      RaiseEvent SatelliteSelected(SelectedSatellite)
      UpdateSatelliteLabel
      Exit For
    End If
  Next i

End Sub

Private Sub TabList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim SourceNode As ListItem
  Dim i As Integer

  Set SourceNode = TabList.HitTest(x, y)
  If Not (SourceNode Is Nothing) Then
    For i = 1 To nSatCount
      If m_SatelliteName(i) = SourceNode.Text Then
        If Button And vbLeftButton Then
          SelectedSatellite = i
          UpdateSatelliteLabel SelectedSatellite
          Exit For
        End If
        If Button And vbRightButton Then
          Screen.MousePointer = vbDefault
          nSatDetailsTag = i
          SatMouseOver = i
          mnuPopupSatName.Caption = m_SatelliteName(i)
          If m_OrbitalModel(i) = agSGP Then
            mnuModelSGP.Checked = True
            mnuModelPlan13.Checked = False
          Else
            mnuModelSGP.Checked = False
            mnuModelPlan13.Checked = True
          End If
          If m_SatelliteBusy(i) Then
            mnuPopupReset.Enabled = True
          Else
            mnuPopupReset.Enabled = False
          End If
          PopupMenu mnuPopupSat
        End If
      End If
    Next i
    Set SourceNode = Nothing
  End If
End Sub

Private Sub tmrAutoUpdate_Timer()

  Dim vDate As Variant
  Dim nSatNumber As Integer
  Dim i As Integer
  
  For i = -1 To nSatCount
    If m_SatelliteName(i) = "" Then Exit For

    'vDate = DateValue(m_DisplayDay(i) & " " & m_DisplayMonth(i) & " " & m_DisplayYear(i)) & " " & TimeValue(m_DisplayHour(i) & ":" & m_DisplayMinute(i) & ":" & m_DisplaySecond(i))

    vDate = Now
 '   vDate = DateAdd("s", tmrAutoUpdate.Interval / 1000, vDate)

    m_DisplayDay(i) = Day(vDate)
    m_DisplayMonth(i) = Month(vDate)
    m_DisplayYear(i) = Year(vDate)

    m_DisplayHour(i) = hour(vDate)
    m_DisplayMinute(i) = Minute(vDate)
    m_DisplaySecond(i) = Second(vDate)
    CalculateSatellitePosition True, i
  
    RaiseEvent SatelliteUpdated(i)
  Next i
  
  DrawFootprints

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ObserverMapCentre = m_def_ObserverMapCentre
  m_OutputStyle = m_def_OutputStyle
  m_ObserverLocation = m_def_ObserverLocation

  mvarObserverLatitude = m_def_ObserverLatitude
  mvarObserverLongitude = m_def_ObserverLongitude
  mvarObserverHeight = m_def_ObserverHeight
  
  m_SecondObserverEnabled = m_def_SecondObserverEnabled
  
  m_DisplaySun = m_def_DisplaySun
  m_DisplayMoon = m_def_DisplayMoon
  m_DisplaySunFootprint = m_def_DisplaySunFootprint
  m_DisplayMoonFootprint = m_def_DisplayMoonFootprint
  m_DisplayTimes = m_def_DisplayTimes
  m_DisplayFootprints = m_def_DisplayFootprints
  m_DisplayTracks = m_def_DisplayTracks
  m_SelectedSatelliteName = m_def_SelectedSatelliteName
  m_IsSatLoaded = m_def_IsSatLoaded
  m_SelectedSatellite = m_def_SelectedSatellite
  m_SetSatelliteTime = m_def_SetSatelliteTime
  m_CurrentSelectedSatellite = m_def_CurrentSelectedSatellite
  m_SetAOSLOS = m_def_SetAOSLOS
  m_DatabasePath = m_def_DatabasePath
  m_Busy = m_def_Busy
  m_ocxBusy = m_def_ocxBusy
  m_Enable847 = m_def_Enable847
  m_PortSettings = m_def_PortSettings
  m_Enable847Sat = m_def_Enable847Sat
  m_Timezone = m_def_Timezone
  m_DaylightSaving = m_def_DaylightSaving
'  m_AutoMode = m_def_AutoMode
  m_AutoInterval = m_def_AutoInterval
  m_AutoMode = m_def_AutoMode
  m_DisplayDataFields = m_def_DisplayDataFields
  m_FrequencyDatabasePath = m_def_FrequencyDatabasePath
  m_SetIndexOnSelect = m_def_SetIndexOnSelect
  m_SetSelectedSatellite = m_def_SetSelectedSatellite
  m_AllowDoEvents = m_def_AllowDoEvents
  

  m_UseHourglass = m_def_UseHourglass
  m_TimeZoneName = m_def_TimeZoneName
  m_DaylightSavingAdjust = m_def_DaylightSavingAdjust
  m_ViewsOrthShade = m_def_ViewsOrthShade
  m_ViewsOrthLocations = m_def_ViewsOrthLocations
  m_MaxWidth = m_def_MaxWidth
  m_MaxHeight = m_def_MaxHeight
  m_FT847CATSettings = m_def_FT847CATSettings
  m_SatPosLabelAlign = m_def_SatPosLabelAlign
  m_SetActiveWindowAsWallpaper = m_def_SetActiveWindowAsWallpaper
  m_DisplayStatusBar = m_def_DisplayStatusBar
  m_DisplaySatelliteLabel = m_def_DisplaySatelliteLabel
  m_EnableSpeech = m_def_EnableSpeech
  m_EnableSatStatus = m_def_EnableSatStatus
  m_DisplayGroundTrackAsPoints = m_def_DisplayGroundTrackAsPoints
  m_CalculationModel = m_def_CalculationModel
  m_GroundTrackInterval = m_def_GroundTrackInterval
  m_DisplayIcons = m_def_DisplayIcons
  m_DisplayAOSCircle = m_def_DisplayAOSCircle
  m_SatelliteBearing = m_def_SatelliteBearing
  m_UserStatusPanelText = m_def_UserStatusPanelText
  m_GroundTrackPointSize = m_def_GroundTrackPointSize
  m_SunRise = m_def_SunRise
  m_SunNoon = m_def_SunNoon
  m_SunSet = m_def_SunSet
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_ObserverMapCentre = PropBag.ReadProperty("ObserverMapCentre", m_def_ObserverMapCentre)
  m_OutputStyle = PropBag.ReadProperty("OutputStyle", m_def_OutputStyle)
  m_ObserverLocation = PropBag.ReadProperty("ObserverLocation", m_def_ObserverLocation)

  mvarObserverLatitude = PropBag.ReadProperty("ObserverLatitude", m_def_ObserverLatitude)
  mvarObserverLongitude = PropBag.ReadProperty("ObserverLongitude", m_def_ObserverLongitude)
  mvarObserverHeight = PropBag.ReadProperty("ObserverHeight", m_def_ObserverHeight)

  m_DisplaySun = PropBag.ReadProperty("DisplaySun", m_def_DisplaySun)
  m_DisplayMoon = PropBag.ReadProperty("DisplayMoon", m_def_DisplayMoon)
  m_DisplaySunFootprint = PropBag.ReadProperty("DisplaySunFootprint", m_def_DisplaySunFootprint)
  m_DisplayMoonFootprint = PropBag.ReadProperty("DisplayMoonFootprint", m_def_DisplayMoonFootprint)
  m_DisplayTimes = PropBag.ReadProperty("DisplayTimes", m_def_DisplayTimes)
  m_DisplayFootprints = PropBag.ReadProperty("DisplayFootprints", m_def_DisplayFootprints)
  m_DisplayTracks = PropBag.ReadProperty("DisplayTracks", m_def_DisplayTracks)
  m_SelectedSatelliteName = PropBag.ReadProperty("SelectedSatelliteName", m_def_SelectedSatelliteName)
  m_IsSatLoaded = PropBag.ReadProperty("IsSatLoaded", m_def_IsSatLoaded)
  m_SelectedSatellite = PropBag.ReadProperty("SelectedSatellite", m_def_SelectedSatellite)
  m_CurrentSelectedSatellite = PropBag.ReadProperty("CurrentSelectedSatellite", m_def_CurrentSelectedSatellite)
  m_SetAOSLOS = PropBag.ReadProperty("SetAOSLOS", m_def_SetAOSLOS)
  m_DatabasePath = PropBag.ReadProperty("DatabasePath", m_def_DatabasePath)
  m_Busy = PropBag.ReadProperty("Busy", m_def_Busy)
  m_ocxBusy = PropBag.ReadProperty("ocxBusy", m_def_ocxBusy)
  m_Enable847 = PropBag.ReadProperty("Enable847", m_def_Enable847)
  m_PortSettings = PropBag.ReadProperty("PortSettings", m_def_PortSettings)
  m_Enable847Sat = PropBag.ReadProperty("Enable847Sat", m_def_Enable847Sat)
  m_Timezone = PropBag.ReadProperty("Timezone", m_def_Timezone)
  m_DaylightSaving = PropBag.ReadProperty("DaylightSaving", m_def_DaylightSaving)
'  m_AutoMode = PropBag.ReadProperty("AutoMode", m_def_AutoMode)
  m_AutoInterval = PropBag.ReadProperty("AutoInterval", m_def_AutoInterval)
  m_AutoMode = PropBag.ReadProperty("AutoMode", m_def_AutoMode)
  m_DisplayDataFields = PropBag.ReadProperty("DisplayDataFields", m_def_DisplayDataFields)
  m_FrequencyDatabasePath = PropBag.ReadProperty("FrequencyDatabasePath", m_def_FrequencyDatabasePath)
  m_SetIndexOnSelect = PropBag.ReadProperty("SetIndexOnSelect", m_def_SetIndexOnSelect)
  SelectedSatellite = PropBag.ReadProperty("SetSelectedSatellite", m_def_SetSelectedSatellite)
  m_AllowDoEvents = PropBag.ReadProperty("AllowDoEvents", m_def_AllowDoEvents)

  m_UseHourglass = PropBag.ReadProperty("UseHourglass", m_def_UseHourglass)
  m_TimeZoneName = PropBag.ReadProperty("TimeZoneName", m_def_TimeZoneName)
  m_DaylightSavingAdjust = PropBag.ReadProperty("DaylightSavingAdjust", m_def_DaylightSavingAdjust)
  m_ViewsOrthShade = PropBag.ReadProperty("ViewsOrthShade", m_def_ViewsOrthShade)
  m_ViewsOrthLocations = PropBag.ReadProperty("ViewsOrthLocations", m_def_ViewsOrthLocations)
  m_MaxWidth = PropBag.ReadProperty("MaxWidth", m_def_MaxWidth)
  m_MaxHeight = PropBag.ReadProperty("MaxHeight", m_def_MaxHeight)
  m_FT847CATSettings = PropBag.ReadProperty("FT847CATSettings", m_def_FT847CATSettings)
  m_SatPosLabelAlign = PropBag.ReadProperty("SatPosLabelAlign", m_def_SatPosLabelAlign)
  m_SetActiveWindowAsWallpaper = PropBag.ReadProperty("SetActiveWindowAsWallpaper", m_def_SetActiveWindowAsWallpaper)
  m_DisplayStatusBar = PropBag.ReadProperty("DisplayStatusBar", m_def_DisplayStatusBar)
  m_DisplaySatelliteLabel = PropBag.ReadProperty("DisplaySatelliteLabel", m_def_DisplaySatelliteLabel)
  m_EnableSpeech = PropBag.ReadProperty("EnableSpeech", m_def_EnableSpeech)
  m_EnableSatStatus = PropBag.ReadProperty("EnableSatStatus", m_def_EnableSatStatus)
  m_DisplayGroundTrackAsPoints = PropBag.ReadProperty("DisplayGroundTrackAsPoints", m_def_DisplayGroundTrackAsPoints)
  m_CalculationModel = PropBag.ReadProperty("CalculationModel", m_def_CalculationModel)
  m_GroundTrackInterval = PropBag.ReadProperty("GroundTrackInterval", m_def_GroundTrackInterval)
  m_DisplayIcons = PropBag.ReadProperty("DisplayIcons", m_def_DisplayIcons)
  m_DisplayAOSCircle = PropBag.ReadProperty("DisplayAOSCircle", m_def_DisplayAOSCircle)
 ' Set Picture = PropBag.ReadProperty("Picture", Nothing)
 ' Set Picture = PropBag.ReadProperty("Picture", Nothing)
  m_SatelliteBearing = PropBag.ReadProperty("SatelliteBearing", m_def_SatelliteBearing)
  m_UserStatusPanelText = PropBag.ReadProperty("UserStatusPanelText", m_def_UserStatusPanelText)
  m_GroundTrackPointSize = PropBag.ReadProperty("GroundTrackPointSize", m_def_GroundTrackPointSize)
  m_SunRise = PropBag.ReadProperty("SunRise", m_def_SunRise)
  m_SunNoon = PropBag.ReadProperty("SunNoon", m_def_SunNoon)
  m_SunSet = PropBag.ReadProperty("SunSet", m_def_SunSet)
End Sub

Private Sub UserControl_Show()
  If Ambient.UserMode Then
    UserControl_Resize
    ReadDX
  End If
End Sub

Private Sub UserControl_Terminate()
  Set cDJTSGP = Nothing
  Set cSunRise = Nothing
  Set cWaitCur = Nothing
   
'    Set mVpt = Nothing
'    Set mDev = Nothing
'    Set mFrL = Nothing
'    Set mFrO = Nothing
'    Set mFrC = Nothing
'    Set mFrS = Nothing
'    Set mDrm = Nothing
'    Set mDx7 = Nothing
 
   
 ' Unload SatDetails
  
  TXSpeech.AudioReset
        
  If m_SetActiveWindowAsWallpaper Then
    vblSetDesktopWallpaper ""
  End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("ObserverMapCentre", m_ObserverMapCentre, m_def_ObserverMapCentre)
  Call PropBag.WriteProperty("OutputStyle", m_OutputStyle, m_def_OutputStyle)
  Call PropBag.WriteProperty("ObserverLocation", m_ObserverLocation, m_def_ObserverLocation)

  Call PropBag.WriteProperty("ObserverLatitude", mvarObserverLatitude, m_def_ObserverLatitude)
  Call PropBag.WriteProperty("ObserverLongitude", mvarObserverLongitude, m_def_ObserverLongitude)
  Call PropBag.WriteProperty("ObserverHeight", mvarObserverHeight, m_def_ObserverHeight)
  Call PropBag.WriteProperty("DisplaySun", m_DisplaySun, m_def_DisplaySun)
  Call PropBag.WriteProperty("DisplayMoon", m_DisplayMoon, m_def_DisplayMoon)
  Call PropBag.WriteProperty("DisplaySunFootprint", m_DisplaySunFootprint, m_def_DisplaySunFootprint)
  Call PropBag.WriteProperty("DisplayMoonFootprint", m_DisplayMoonFootprint, m_def_DisplayMoonFootprint)
  Call PropBag.WriteProperty("DisplayTimes", m_DisplayTimes, m_def_DisplayTimes)
  Call PropBag.WriteProperty("DisplayFootprints", m_DisplayFootprints, m_def_DisplayFootprints)
  Call PropBag.WriteProperty("DisplayTracks", m_DisplayTracks, m_def_DisplayTracks)
  Call PropBag.WriteProperty("SelectedSatelliteName", m_SelectedSatelliteName, m_def_SelectedSatelliteName)
  Call PropBag.WriteProperty("IsSatLoaded", m_IsSatLoaded, m_def_IsSatLoaded)
  Call PropBag.WriteProperty("SelectedSatellite", m_SelectedSatellite, m_def_SelectedSatellite)
  Call PropBag.WriteProperty("CurrentSelectedSatellite", m_CurrentSelectedSatellite, m_def_CurrentSelectedSatellite)
  Call PropBag.WriteProperty("SetAOSLOS", m_SetAOSLOS, m_def_SetAOSLOS)
  Call PropBag.WriteProperty("DatabasePath", m_DatabasePath, m_def_DatabasePath)
  Call PropBag.WriteProperty("Busy", m_Busy, m_def_Busy)
  Call PropBag.WriteProperty("ocxBusy", m_ocxBusy, m_def_ocxBusy)
  Call PropBag.WriteProperty("Enable847", m_Enable847, m_def_Enable847)
  Call PropBag.WriteProperty("PortSettings", m_PortSettings, m_def_PortSettings)
  Call PropBag.WriteProperty("Enable847Sat", m_Enable847Sat, m_def_Enable847Sat)
  Call PropBag.WriteProperty("Timezone", m_Timezone, m_def_Timezone)
  Call PropBag.WriteProperty("DaylightSaving", m_DaylightSaving, m_def_DaylightSaving)
'  Call PropBag.WriteProperty("AutoMode", m_AutoMode, m_def_AutoMode)
  Call PropBag.WriteProperty("AutoInterval", m_AutoInterval, m_def_AutoInterval)
  Call PropBag.WriteProperty("AutoMode", m_AutoMode, m_def_AutoMode)
  Call PropBag.WriteProperty("DisplayDataFields", m_DisplayDataFields, m_def_DisplayDataFields)
  Call PropBag.WriteProperty("FrequencyDatabasePath", m_FrequencyDatabasePath, m_def_FrequencyDatabasePath)
  Call PropBag.WriteProperty("SetIndexOnSelect", m_SetIndexOnSelect, m_def_SetIndexOnSelect)
  Call PropBag.WriteProperty("SetSelectedSatellite", SelectedSatellite, m_def_SetSelectedSatellite)
  Call PropBag.WriteProperty("AllowDoEvents", m_AllowDoEvents, m_def_AllowDoEvents)
  Call PropBag.WriteProperty("UseHourglass", m_UseHourglass, m_def_UseHourglass)
  Call PropBag.WriteProperty("TimeZoneName", m_TimeZoneName, m_def_TimeZoneName)
  Call PropBag.WriteProperty("DaylightSavingAdjust", m_DaylightSavingAdjust, m_def_DaylightSavingAdjust)
  Call PropBag.WriteProperty("ViewsOrthShade", m_ViewsOrthShade, m_def_ViewsOrthShade)
  Call PropBag.WriteProperty("ViewsOrthLocations", m_ViewsOrthLocations, m_def_ViewsOrthLocations)
  Call PropBag.WriteProperty("MaxWidth", m_MaxWidth, m_def_MaxWidth)
  Call PropBag.WriteProperty("MaxHeight", m_MaxHeight, m_def_MaxHeight)
  Call PropBag.WriteProperty("FT847CATSettings", m_FT847CATSettings, m_def_FT847CATSettings)
  Call PropBag.WriteProperty("SatPosLabelAlign", m_SatPosLabelAlign, m_def_SatPosLabelAlign)
  Call PropBag.WriteProperty("SetActiveWindowAsWallpaper", m_SetActiveWindowAsWallpaper, m_def_SetActiveWindowAsWallpaper)
  Call PropBag.WriteProperty("DisplayStatusBar", m_DisplayStatusBar, m_def_DisplayStatusBar)
  Call PropBag.WriteProperty("DisplaySatelliteLabel", m_DisplaySatelliteLabel, m_def_DisplaySatelliteLabel)
  Call PropBag.WriteProperty("EnableSpeech", m_EnableSpeech, m_def_EnableSpeech)
  Call PropBag.WriteProperty("EnableSatStatus", m_EnableSatStatus, m_def_EnableSatStatus)
  Call PropBag.WriteProperty("DisplayGroundTrackAsPoints", m_DisplayGroundTrackAsPoints, m_def_DisplayGroundTrackAsPoints)
  Call PropBag.WriteProperty("CalculationModel", m_CalculationModel, m_def_CalculationModel)
  Call PropBag.WriteProperty("CalculationModel", m_GroundTrackInterval, m_def_GroundTrackInterval)
  Call PropBag.WriteProperty("DisplayIcons", m_DisplayIcons, m_def_DisplayIcons)
  Call PropBag.WriteProperty("DisplayAOSCircle", m_DisplayAOSCircle, m_def_DisplayAOSCircle)
'  Call PropBag.WriteProperty("Picture", Picture, Nothing)
'  Call PropBag.WriteProperty("Picture", Picture, Nothing)
  Call PropBag.WriteProperty("SatelliteBearing", m_SatelliteBearing, m_def_SatelliteBearing)
  Call PropBag.WriteProperty("UserStatusPanelText", m_UserStatusPanelText, m_def_UserStatusPanelText)
  Call PropBag.WriteProperty("GroundTrackPointSize", m_GroundTrackPointSize, m_def_GroundTrackPointSize)
  Call PropBag.WriteProperty("SunRise", m_SunRise, m_def_SunRise)
  Call PropBag.WriteProperty("SunNoon", m_SunNoon, m_def_SunNoon)
  Call PropBag.WriteProperty("SunSet", m_SunSet, m_def_SunSet)
End Sub

Public Property Get OutputStyle() As OS
Attribute OutputStyle.VB_ProcData.VB_Invoke_Property = "General"
    OutputStyle = m_OutputStyle
End Property

Public Property Let OutputStyle(ByVal New_OutputStyle As OS)
    Dim bDummy As Boolean
        
    m_OutputStyle = New_OutputStyle
    Select Case m_OutputStyle
      Case 0
        picInner.Visible = False
        ViewPort.Visible = False
        TabList.Top = 0
        TabList.Left = 0
        TabList.Visible = True
        UserControl_Resize
      Case 1
        If m_ObserverMapCentre = 0 Then
          picInner.Picture = Map0.Picture
        Else

          picInner.Picture = Map180.Picture
        End If
        picInner.Visible = True
        ViewPort.Visible = True
        TabList.Top = 0
        TabList.Left = 0
        TabList.Visible = False
        UserControl_Resize
      Case 2
        picInner.Picture = Map3.Picture
        picInner.Visible = True
        ViewPort.Visible = True
        TabList.Top = 0
        TabList.Left = 0
        TabList.Visible = False
        UserControl_Resize
      Case 3
        picInner.Visible = True
        picInner.BackColor = RGB(0, 0, 0)
        ViewPort.Visible = True
        picInner.Picture = LoadPicture()
        TabList.Top = 0
        TabList.Left = 0
        TabList.Visible = False
        UserControl_Resize
      Case 4
        picInner.Visible = False
        ViewPort.Visible = False
        TabList.Visible = False
        StatusBar1.Visible = False
        lblSatPos.Visible = False
        UserControl_Resize
    End Select
    
    
    UserControl.lblSatPos.Visible = IIf(m_DisplaySatelliteLabel, True, False)
    UserControl.StatusBar1.Visible = IIf(m_DisplayStatusBar, True, False)
    
    DrawFootprints
    PropertyChanged "OutputStyle"
End Property

Public Property Get ObserverLocation() As String
Attribute ObserverLocation.VB_ProcData.VB_Invoke_Property = "Observer"
    ObserverLocation = m_ObserverLocation
End Property

Public Property Let ObserverLocation(ByVal New_ObserverLocation As String)
    m_ObserverLocation = New_ObserverLocation
    PropertyChanged "ObserverLocation"
End Property

Public Property Get DataValid() As Boolean
    DataValid = m_DataValid(m_SatelliteIndex)
End Property

Public Property Let DataValid(ByVal New_DataValid As Boolean)
    If Ambient.UserMode = False Then Err.Raise 382
    m_DataValid(m_SatelliteIndex) = New_DataValid
    PropertyChanged "DataValid"
End Property

Public Property Get SatelliteCount() As Variant
 ' Dim i As Integer

 ' For i = 1 To 20
 '   If m_SatelliteName(i) = "" Then Exit For
 ' Next i
  SatelliteCount = nSatCount
    
   ' SatelliteCount = m_SatelliteCount
End Property

Public Property Let SatelliteCount(ByVal New_SatelliteCount As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteCount = New_SatelliteCount
    PropertyChanged "SatelliteCount"
End Property

Public Property Get DisplayTimeRequired() As Variant
    DisplayTimeRequired = m_DisplayTimeRequired(m_SatelliteIndex)
End Property

Public Property Let DisplayTimeRequired(ByVal New_DisplayTimeRequired As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayTimeRequired(m_SatelliteIndex) = New_DisplayTimeRequired
    PropertyChanged "DisplayTimeRequired"
End Property

Public Property Get SatelliteDayNumber() As Variant
    SatelliteDayNumber = m_SatelliteDayNumber(m_SatelliteIndex)
End Property

Public Property Let SatelliteDayNumber(ByVal New_SatelliteDayNumber As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteDayNumber(m_SatelliteIndex) = New_SatelliteDayNumber
    PropertyChanged "SatelliteDayNumber"
End Property

Public Property Get DisplayMinute() As Integer
    DisplayMinute = m_DisplayMinute(m_SatelliteIndex)
End Property

Public Property Let DisplayMinute(ByVal New_DisplayMinute As Integer)
'    If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayMinute(m_SatelliteIndex) = New_DisplayMinute
    PropertyChanged "DisplayMinute"
End Property

Public Property Get DisplayHour() As Integer
    DisplayHour = m_DisplayHour(m_SatelliteIndex)
End Property

Public Property Let DisplayHour(ByVal New_DisplayHour As Integer)
   ' If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayHour(m_SatelliteIndex) = New_DisplayHour
    PropertyChanged "DisplayHour"
End Property

Public Property Get DisplayDay() As Integer
    DisplayDay = m_DisplayDay(m_SatelliteIndex)
End Property

Public Property Let DisplayDay(ByVal New_DisplayDay As Integer)
'    If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayDay(m_SatelliteIndex) = New_DisplayDay
    PropertyChanged "DisplayDay"
End Property

Public Property Get DisplayMonth() As Integer
    DisplayMonth = m_DisplayMonth(m_SatelliteIndex)
End Property

Public Property Let DisplayMonth(ByVal New_DisplayMonth As Integer)
    'If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayMonth(m_SatelliteIndex) = New_DisplayMonth
    PropertyChanged "DisplayMonth"
End Property

Public Property Get DisplayYear() As Integer
    DisplayYear = m_DisplayYear(m_SatelliteIndex)
End Property

Public Property Let DisplayYear(ByVal New_DisplayYear As Integer)
    'If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayYear(m_SatelliteIndex) = New_DisplayYear
    PropertyChanged "DisplayYear"
End Property

Public Property Get DisplayCentury() As Integer
    DisplayCentury = m_DisplayCentury(m_SatelliteIndex)
End Property

Public Property Let DisplayCentury(ByVal New_DisplayCentury As Integer)
    'If Ambient.UserMode = False Then Err.Raise 382
    m_DisplayCentury(m_SatelliteIndex) = New_DisplayCentury
    PropertyChanged "DisplayCentury"
End Property

Public Property Get SatelliteIndex() As Integer
    SatelliteIndex = m_SatelliteIndex
End Property

Public Property Let SatelliteIndex(ByVal New_SatelliteIndex As Integer)
    If Ambient.UserMode = False Then Err.Raise 382
    m_SatelliteIndex = New_SatelliteIndex
    PropertyChanged "SatelliteIndex"
End Property

Public Property Get SatelliteOrbitNumber() As Long
    SatelliteOrbitNumber = m_SatelliteOrbitNumber(m_SatelliteIndex)
End Property

Public Property Let SatelliteOrbitNumber(ByVal New_SatelliteOrbitNumber As Long)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteOrbitNumber(m_SatelliteIndex) = New_SatelliteOrbitNumber
    PropertyChanged "SatelliteOrbitNumber"
End Property

Public Property Get SatelliteLongitude() As Single
    SatelliteLongitude = m_SatelliteLongitude(m_SatelliteIndex)
End Property

Public Property Let SatelliteLongitude(ByVal New_SatelliteLongitude As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteLongitude(m_SatelliteIndex) = New_SatelliteLongitude
    PropertyChanged "SatelliteLongitude"
End Property

Public Property Get SatelliteLatitude() As Single
    SatelliteLatitude = m_SatelliteLatitude(m_SatelliteIndex)
End Property

Public Property Let SatelliteLatitude(ByVal New_SatelliteLatitude As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteLatitude(m_SatelliteIndex) = New_SatelliteLatitude
    PropertyChanged "SatelliteLatitude"
End Property
Public Property Get SatelliteAltitude() As Single
    SatelliteAltitude = m_SatelliteAltitude(m_SatelliteIndex)
End Property

Public Property Let SatelliteAltitude(ByVal New_SatelliteAltitude As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteAltitude(m_SatelliteIndex) = New_SatelliteAltitude
    PropertyChanged "SatelliteAltitude"
End Property

Public Property Get KepsChecksum() As Variant
    KepsChecksum = m_KepsChecksum(m_SatelliteIndex)
End Property

Public Property Let KepsChecksum(ByVal New_KepsChecksum As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsChecksum(m_SatelliteIndex) = New_KepsChecksum
    PropertyChanged "KepsChecksum"
End Property

Public Property Get KepsDecayRate() As Single
    KepsDecayRate = m_KepsDecayRate(m_SatelliteIndex)
End Property

Public Property Let KepsDecayRate(ByVal New_KepsDecayRate As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsDecayRate(m_SatelliteIndex) = New_KepsDecayRate
    PropertyChanged "KepsDecayRate"
End Property

Public Property Get KepsElementSet() As Variant
    KepsElementSet = m_KepsElementSet(m_SatelliteIndex)
End Property

Public Property Let KepsElementSet(ByVal New_KepsElementSet As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsElementSet(m_SatelliteIndex) = New_KepsElementSet
    PropertyChanged "KepsElementSet"
End Property

Public Property Get KepsEpochTime() As Variant
    KepsEpochTime = m_KepsEpochTime(m_SatelliteIndex)
End Property

Public Property Let KepsEpochTime(ByVal New_KepsEpochTime As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsEpochTime(m_SatelliteIndex) = New_KepsEpochTime

  If Val(Left$(New_KepsEpochTime, 2)) < 50 Then
    century% = 20
  Else
    century% = 19
  End If
  mvarKepsYearEpochTime(m_SatelliteIndex) = m_KepsEpochTime(m_SatelliteIndex) - 1000 * Int(m_KepsEpochTime(m_SatelliteIndex) / 1000)
  mvarSatEpochYear(m_SatelliteIndex) = 100 * century% + Int(m_KepsEpochTime(m_SatelliteIndex) / 1000)

    PropertyChanged "KepsEpochTime"
End Property

Public Property Get KepsOrbitNumber() As Variant
    KepsOrbitNumber = m_KepsOrbitNumber(m_SatelliteIndex)
End Property

Public Property Let KepsOrbitNumber(ByVal New_KepsOrbitNumber As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsOrbitNumber(m_SatelliteIndex) = New_KepsOrbitNumber
    PropertyChanged "KepsOrbitNumber"
End Property

Public Property Get KepsMeanMotion() As Double
    KepsMeanMotion = m_KepsMeanMotion(m_SatelliteIndex)
End Property

Public Property Let KepsMeanMotion(ByVal New_KepsMeanMotion As Double)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsMeanMotion(m_SatelliteIndex) = New_KepsMeanMotion
    PropertyChanged "KepsMeanMotion"
End Property

Public Property Get KepsMeanAnomoly() As Single
    KepsMeanAnomoly = m_KepsMeanAnomoly(m_SatelliteIndex)
End Property

Public Property Let KepsMeanAnomoly(ByVal New_KepsMeanAnomoly As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsMeanAnomoly(m_SatelliteIndex) = New_KepsMeanAnomoly
    PropertyChanged "KepsMeanAnomoly"
End Property

Public Property Get KepsInclination() As Single
    KepsInclination = m_KepsInclination(m_SatelliteIndex)
End Property

Public Property Let KepsInclination(ByVal New_KepsInclination As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsInclination(m_SatelliteIndex) = New_KepsInclination
    PropertyChanged "KepsInclination"
End Property

Public Property Get KepsAOP() As Single
    KepsAOP = m_KepsAOP(m_SatelliteIndex)
End Property

Public Property Let KepsAOP(ByVal New_KepsAOP As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsAOP(m_SatelliteIndex) = New_KepsAOP
    PropertyChanged "KepsAOP"
End Property

Public Property Get KepsRAAN() As Single
    KepsRAAN = m_KepsRAAN(m_SatelliteIndex)
End Property

Public Property Let KepsRAAN(ByVal New_KepsRAAN As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsRAAN(m_SatelliteIndex) = New_KepsRAAN
    PropertyChanged "KepsRAAN"
End Property

Public Property Get KepsEccentricity() As Single
    KepsEccentricity = m_KepsEccentricity(m_SatelliteIndex)
End Property

Public Property Let KepsEccentricity(ByVal New_KepsEccentricity As Single)
    If Ambient.UserMode = False Then Err.Raise 382
    m_KepsEccentricity(m_SatelliteIndex) = New_KepsEccentricity
    PropertyChanged "KepsEccentricity"
End Property

Public Property Get SatelliteElevation() As Long
    SatelliteElevation = m_SatelliteElevation(m_SatelliteIndex)
End Property

Public Property Let SatelliteElevation(ByVal New_SatelliteElevation As Long)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteElevation(m_SatelliteIndex) = New_SatelliteElevation
    PropertyChanged "SatelliteElevation"
End Property

Public Property Get SatelliteAzimuth() As Long
    SatelliteAzimuth = m_SatelliteAzimuth(m_SatelliteIndex)
End Property

Public Property Let SatelliteAzimuth(ByVal New_SatelliteAzimuth As Long)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteAzimuth(m_SatelliteIndex) = New_SatelliteAzimuth
    PropertyChanged "SatelliteAzimuth"
End Property

Public Property Get SatelliteRange() As Long
    SatelliteRange = m_SatelliteRange(m_SatelliteIndex)
End Property

Public Property Let SatelliteRange(ByVal New_SatelliteRange As Long)
    If Ambient.UserMode = False Then Err.Raise 382
    If Ambient.UserMode Then Err.Raise 393
    m_SatelliteRange(m_SatelliteIndex) = New_SatelliteRange
    PropertyChanged "SatelliteRange"
End Property

Public Property Get SatelliteDesignator() As Variant
    SatelliteDesignator = m_SatelliteDesignator(m_SatelliteIndex)
End Property

Public Property Let SatelliteDesignator(ByVal New_SatelliteDesignator As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_SatelliteDesignator(m_SatelliteIndex) = New_SatelliteDesignator
    PropertyChanged "SatelliteDesignator"
End Property

Public Property Get SatelliteName() As Variant
      SatelliteName = m_SatelliteName(m_SatelliteIndex)
End Property

Public Property Let SatelliteName(ByVal New_SatelliteName As Variant)
    If Ambient.UserMode = False Then Err.Raise 382
    m_SatelliteName(m_SatelliteIndex) = New_SatelliteName
    PropertyChanged "SatelliteName"
End Property
Public Property Get SatelliteBusy() As Boolean
  SatelliteBusy = m_SatelliteBusy(m_SatelliteIndex)
End Property

Public Property Let SatelliteBusy(ByVal New_SatelliteBusy As Boolean)
  If Ambient.UserMode = False Then Err.Raise 382
  m_SatelliteBusy(m_SatelliteIndex) = New_SatelliteBusy
  PropertyChanged "SatelliteBusy"
End Property

Public Property Get DisplaySun() As Boolean
Attribute DisplaySun.VB_ProcData.VB_Invoke_Property = "Display"
    DisplaySun = m_DisplaySun
End Property

Public Property Let DisplaySun(ByVal New_DisplaySun As Boolean)
    m_DisplaySun = New_DisplaySun
    PropertyChanged "DisplaySun"
    
End Property

Public Property Get DisplayMoon() As Boolean
Attribute DisplayMoon.VB_ProcData.VB_Invoke_Property = "Display"
    DisplayMoon = m_DisplayMoon
End Property

Public Property Let DisplayMoon(ByVal New_DisplayMoon As Boolean)
    m_DisplayMoon = New_DisplayMoon
    PropertyChanged "DisplayMoon"
End Property

Public Property Get DisplaySunFootprint() As Boolean
Attribute DisplaySunFootprint.VB_ProcData.VB_Invoke_Property = "Display"
    DisplaySunFootprint = m_DisplaySunFootprint
End Property

Public Property Let DisplaySunFootprint(ByVal New_DisplaySunFootprint As Boolean)
    m_DisplaySunFootprint = New_DisplaySunFootprint
    PropertyChanged "DisplaySunFootprint"
End Property

Public Property Get DisplayMoonFootprint() As Boolean
Attribute DisplayMoonFootprint.VB_ProcData.VB_Invoke_Property = "Display"
    DisplayMoonFootprint = m_DisplayMoonFootprint
End Property

Public Property Let DisplayMoonFootprint(ByVal New_DisplayMoonFootprint As Boolean)
    m_DisplayMoonFootprint = New_DisplayMoonFootprint
    PropertyChanged "DisplayMoonFootprint"
End Property

Public Property Get DisplayTracks() As Boolean
Attribute DisplayTracks.VB_ProcData.VB_Invoke_Property = "Display"
  DisplayTracks = m_DisplayTracks
End Property

Public Property Let DisplayTracks(ByVal New_DisplayTracks As Boolean)
  m_DisplayTracks = New_DisplayTracks
  
  If Not m_DisplayTracks Then
'    Erase SatTrackLon
'    Erase SatTrackLat
'    Erase SatTrackElev
'    Erase SatTrackAzim
'    Erase SatTrackMutual
  End If
  PropertyChanged "DisplayTracks"
End Property

Public Property Get DisplaySecond() As Integer
  DisplaySecond = m_DisplaySecond(m_SatelliteIndex)
End Property

Public Property Let DisplaySecond(ByVal New_DisplaySecond As Integer)
  m_DisplaySecond(m_SatelliteIndex) = New_DisplaySecond
  PropertyChanged "DisplaySecond"
End Property


Private Sub UserControl_Initialize()
  On Error Resume Next
  Dim lLen As Long
  
  nSatCount = 0
  ResizeArrays nSatCount
    
  Set cSunRise = New clsSunrise
  Set cDJTSGP = CreateObject("DJTSatLib.Satellites")

  bSatDetailsVisible = False

  'SatDetails.Hide

  ResetMaps

  picInner.Picture = Map0.Picture

  DaysInMonth(1) = 31
  DaysInMonth(2) = 28
  DaysInMonth(3) = 31
  DaysInMonth(4) = 30
  DaysInMonth(5) = 31
  DaysInMonth(6) = 30
  DaysInMonth(7) = 31
  DaysInMonth(8) = 31
  DaysInMonth(9) = 30
  DaysInMonth(10) = 31
  DaysInMonth(11) = 30
  DaysInMonth(12) = 31

  PI = 4 * Atn(1)
  MeanYear = 365.25
  TropicalYear = 365.242197

  EarthRotationRate = 2 * PI / TropicalYear
  EarthRotationRateDay = 2 * PI + EarthRotationRate
  EarthRotationRateSeconds = EarthRotationRateDay / 86400
  RE = 6378.14
  FL = 1 / 298.257
  GravitationalConstant = 398600!
  ZonalCoeff = 0.00108263
  YG = 1990
  G0 = 99.4033
  SetupMoon
  SetupSun
  SatMouseOver = 0
  PointsToDraw = 60
  TwoPI = 2 * PI
  cIntDeg = 180 / PI
  bGotSun = False

  With picInner
    HalfMapWidth = .ScaleX(.Width, vbTwips, vbPixels) / 2
    HalfMapHeight = .ScaleY(.Height, vbTwips, vbPixels) / 2
    MapWidth = .ScaleX(.Width, vbTwips, vbPixels)
    MapHeight = .ScaleY(.Height, vbTwips, vbPixels)
    PixelsPerDegLon = MapWidth / 360
    PixelsPerDegLat = (MapHeight / 360) * 2
  End With

  ReDim Preserve FootprintLON(PointsToDraw)
  ReDim Preserve FootprintLAT(PointsToDraw)

  Set cWaitCur = New CWaitCursor

  strTempPath = String$(255, " ")
  lLen = GetTempPath(Len(strTempPath), strTempPath)
  strTempPath = Left(strTempPath, lLen)
  
'  Set mDx7 = New DirectX7
'  Set mDrm = mDx7.Direct3DRMCreate
'  Set mDrw = mDx7.DirectDrawCreate("")
'  CreateSceneGraph
'  CreateDisplay
'  LoadMesh
'  mDev.SetQuality D3DRMRENDER_GOURAUD
 ' AGCalcSGP sSat, sTime, vOPos, vPos, vVel, vObs, vSatPos

End Sub


Sub UpdateSatelliteLabel(Optional SatelliteIndex As Integer)
  Dim SatNumber As Integer
  Dim strTemp As String
  Dim vHours As Variant
  Dim vMins As Variant
  Dim dLon As Double
  Dim dLat As Double
  
  If m_DisplaySatelliteLabel Then
    If IsMissing(SatelliteIndex) Then
      SatNumber = SelectedSatellite
    Else
      SatNumber = SatelliteIndex
    End If

    If SelectedSatellite = 0 Then
      If Me.SatelliteCount > 0 Then
        SelectedSatellite = 1
      End If
    End If

    If SelectedSatellite <> 0 Then
      SatNumber = SelectedSatellite

      strTemp = "NAME: " & m_SatelliteName(SatNumber) & "  "
      strTemp = strTemp & "Rev:" & Str(m_SatelliteOrbitNumber(SatNumber)) & "  "
      strTemp = strTemp & "Az:" & Str(m_SatelliteAzimuth(SatNumber)) & "  "
      strTemp = strTemp & "El:" & Str(m_SatelliteElevation(SatNumber)) & "  "
      strTemp = strTemp & "Rg:" & Str(m_SatelliteRange(SatNumber)) & "  "
      strTemp = strTemp & "Alt:" & Str(m_SatelliteAltitude(SatNumber)) & "  "
      dLat = m_SatelliteLatitude(SatNumber)
      If dLat < 0 Then
        strTemp = strTemp & "Lat: " & Format(Str(Abs(dLat)), "##0.00") & "°S  "
      Else
        strTemp = strTemp & "Lat: " & Format(Str(dLat), "##0.00") & "°N  "
      End If
      dLon = 360 - m_SatelliteLongitude(SatNumber)
      Select Case dLon
        Case 0 To 180
          dLon = -dLon
        Case 181 To 360
          dLon = (360 - dLon)
      End Select
      If dLon < 0 Then
        strTemp = strTemp & "Lon:" & Format(Str(Abs(dLon)), "##0.00") & "°W  "
      Else
        strTemp = strTemp & "Lon:" & Format(Str(dLon), "##0.00") & "°E  "
      End If

'      strTemp = strTemp & "Keps: " & Format(mvarKepsYearEpochTime(SatNumber), "dd/mm/") & mvarSatEpochYear(SatNumber) & "  "
      strTemp = strTemp & " " & m_SatelliteStatusText(SatNumber) & " "
      If satTrackNextAOS(SatNumber) <> "" Then
        strTemp = strTemp & "Next AOS: " & satTrackNextAOS(SatNumber)
        vMins = DateDiff("n", Now, satTrackNextAOS(SatNumber))
        vHours = Int(vMins / 60)
        vMins = vMins - vHours * 60
        strTemp = strTemp & " " & vHours & "hr " & vMins & "mins until AOS"
      Else
        strTemp = strTemp & "Next AOS: <Unavailable>"
      End If

      lblSatPos.Alignment = m_SatPosLabelAlign

      If m_SatelliteElevation(SatNumber) > m_SetAOSLOS Then
        lblSatPos.BackColor = RGB(255, 0, 0)
      Else
        lblSatPos.BackColor = RGB(255, 255, 255)
      End If
    Else
      strTemp = "No Satellite Selected"
    End If

    lblSatPos.Caption = strTemp
  End If
End Sub
Private Sub UserControl_Resize()
  Dim nLabelHeight As Integer
  Dim nStatusBarHeight As Integer
  On Error Resume Next
  
  nLabelHeight = IIf(m_DisplaySatelliteLabel, UserControl.lblSatPos.Height, 0)
  nStatusBarHeight = IIf(m_DisplayStatusBar, UserControl.StatusBar1.Height, 0)
  Select Case m_OutputStyle
    Case 0
      TabList.Top = 0
      TabList.Left = 0
      TabList.Width = UserControl.ScaleWidth

      lblSatPos.Left = 0
      lblSatPos.Top = UserControl.TabList.Height
      lblSatPos.Width = UserControl.ScaleWidth

    Case 1, 2, 3
      ViewPort.Left = 0
      ViewPort.Top = 0
      ViewPort.Width = ScaleWidth
      ViewPort.Height = ScaleHeight - nLabelHeight - nStatusBarHeight


      lblSatPos.Left = 0
      lblSatPos.Top = ViewPort.Height
      lblSatPos.Width = UserControl.ScaleWidth
      lblMousePos.Width = UserControl.ScaleWidth
    Case 4
      Size 640, 640
  End Select
    Select Case m_OutputStyle
      Case 1, 2
 '       m_MaxHeight = picInner.ScaleY(picInner.Height, vbPixels, vbTwips) + nLabelHeight + nStatusBarHeight
        m_MaxHeight = picInner.Height + nLabelHeight + nStatusBarHeight
        m_MaxHeight = m_MaxHeight + HScroll1.Height
        m_MaxWidth = picInner.ScaleY(picInner.Width, vbPixels, vbTwips)
        m_MaxWidth = picInner.Width
      Case 3
        m_MaxHeight = picInner.ScaleY(picInner.Height, vbPixels, vbTwips) + nLabelHeight + nStatusBarHeight
        m_MaxHeight = m_MaxHeight + HScroll1.Height
        m_MaxWidth = picInner.ScaleY(picInner.Width, vbPixels, vbTwips)
      Case 4
        m_MaxHeight = 0
        m_MaxWidth = 0
    End Select
    RaiseEvent Resize

End Sub

Public Function CalculateSatellitePosition(bCalculateTrack As Boolean, Optional SatelliteIndex As Integer) As Boolean
  
  Dim SatNumber As Integer
  Dim nLastEle As Integer
  
If Not m_Busy And Not bStarted Then
  If IsMissing(SatelliteIndex) Then
    SatNumber = m_SatelliteIndex
  Else
    SatNumber = SatelliteIndex
  End If

  m_Busy = True
  
  If m_UseHourglass Then
    cWaitCur.SetCursor vbHourglass
  End If
  
  TempDisplayYear = m_DisplayYear(SatNumber)
  TempDisplayMonth = m_DisplayMonth(SatNumber)
  TempDisplayDay = m_DisplayDay(SatNumber)
  TempDisplayHour = m_DisplayHour(SatNumber)
  TempDisplayMin = m_DisplayMinute(SatNumber)
  TempDisplaySecond = m_DisplaySecond(SatNumber)
    
  LocalDateToUTC = ConvertToGMT(TempDisplayYear, TempDisplayMonth, TempDisplayDay, TempDisplayHour, TempDisplayMin, TempDisplaySecond)
  
  TempDisplayYear = Year(LocalDateToUTC)
  TempDisplayMonth = Month(LocalDateToUTC)
  TempDisplayDay = Day(LocalDateToUTC)
  TempDisplayHour = hour(LocalDateToUTC)
  TempDisplayMin = Minute(LocalDateToUTC)
  TempDisplaySecond = Second(LocalDateToUTC)

  TempSatDayNum = FNDAy(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
  TempSatTimeReq = (TempDisplayHour + (TempDisplayMin / 60) + (TempDisplaySecond / 60 / 60)) / 24

  m_DataValid(SatNumber) = False
  
  TempObsLat = mvarObserverLatitude
  TempObsLon = mvarObserverLongitude

  nLastEle = m_SatelliteElevation(SatNumber)

  PositionEngine SatNumber, True, False
    
  m_SatelliteOrbitNumber(SatNumber) = TempSatOrbit
  m_SatelliteAzimuth(SatNumber) = TempSatAz
  m_SatelliteElevation(SatNumber) = TempSatElev
  m_SatelliteLongitude(SatNumber) = TempSatLon
  m_SatelliteLatitude(SatNumber) = TempSatlat
  m_SatelliteRxFrequency(SatNumber) = TempSatUplinkDoppler
  m_SatelliteTXFrequency(SatNumber) = TempSatDownlinkDoppler
  m_SatelliteRange(SatNumber) = TempSatRange
  m_SatellitePathLoss(SatNumber) = TempSatPathLoss
  m_DataValid(SatNumber) = True
  m_SatelliteStatusText(SatNumber) = TempSatStatusText
  m_SatelliteMA(SatNumber) = TempSatMA

  m_SquintAngle(SatNumber) = TempSquintAngle
  m_RangeRate(SatNumber) = TempRangeRate
   
  RS(SatNumber) = TempRS
  m_SatelliteAltitude(SatNumber) = Round(Abs(TempRS - RE), 1)
  
  If SatNumber > 0 Then
    If bCalculateTrack Then
      If m_DisplayTracks Then
        CalculateTrack SatNumber
      End If
    End If
  End If
    
  If m_SatelliteElevation(SatNumber) < m_SetAOSLOS Then
'    If m_SatelliteElevation(SatNumber) Then
'      RaiseEvent SatelliteAtAOS(SatNumber)
'    End If
    m_SatelliteInAOS(SatNumber) = False
  End If
  
  If m_SatelliteElevation(SatNumber) >= m_SetAOSLOS And Not m_SatelliteInAOS(SatNumber) Then
    RaiseEvent SatelliteAtAOS(SatNumber)
    m_SatelliteInAOS(SatNumber) = True
  End If
  
  m_Busy = False
  If m_UseHourglass Then
    cWaitCur.Restore
  End If
  If m_AllowDoEvents Then
    DoEvents
  End If
End If
End Function
Private Sub PositionEngine(SatNumber As Integer, Setup As Boolean, bPosOnly As Boolean)
 ' Dim lStart As Long
  'Dim lEnd As Long
  
If SatNumber < 1 Then
  Plan13 SatNumber, Setup, bPosOnly
Else
  'lStart = GetTickCount
  Select Case m_OrbitalModel(SatNumber)
    Case Models.agplan13
      Plan13 SatNumber, Setup, bPosOnly
    Case Models.agSGP
      sgp SatNumber, Setup, bPosOnly
    Case Models.agdtsgp
      DTSGP
    Case Else
      Plan13 SatNumber, Setup, bPosOnly
  End Select
  
  If Not bPosOnly Then
    'Calculate Path loss and Doppler Frequency
    If m_DownLinkFrequency(SatNumber) <> 0 Then
      TempSatPathLoss = Format(20# * (Log((PI * 4) * TempSatRange / CVAC * m_DownLinkFrequency(SatNumber)) / Log(10)), "##.##")
    Else
      TempSatPathLoss = 0
    End If
    If m_UplinkFrequency(SatNumber) <> 0 Then
      TempSatUplinkDoppler = Format(m_UplinkFrequency(SatNumber) + (m_UplinkFrequency(SatNumber) * TempRangeRate / CVAC), "####.######")
    Else
      TempSatUplinkDoppler = 0
    End If
    If m_DownLinkFrequency(SatNumber) <> 0 Then
      TempSatDownlinkDoppler = Format(m_DownLinkFrequency(SatNumber) + (-m_DownLinkFrequency(SatNumber) * TempRangeRate / CVAC), "####.######")
    Else
      TempSatDownlinkDoppler = 0
    End If
  End If
  
  
  'DoEvents
  'lEnd = GetTickCount
  'UserControl.StatusBar1.Panels(6) = Str((lStart - lEnd))
End If

End Sub

Private Sub DTSGP()
  Dim bResult As Boolean
  Dim bSouthbound As Boolean
  Dim dPhase As Double
  Dim strName As String
  Dim strTemp As String
  Dim dTime As Double
  
  cDJTSGP.KeplerPath = App.Path & "\Elements1"
  cDJTSGP.ObsLat = ObserverLatitude
  cDJTSGP.ObsLon = ObserverLongitude
  cDJTSGP.ObsHeight = ObserverHeight
  
  TempSatAz = 0
  TempSatElev = 0
  TempSatLon = 0
  TempSatlat = 0
  TempSatRange = 0
  TempSatAlt = 0
  dTime = DateSerial(TempDisplayYear, TempDisplayMonth, TempDisplayDay) + TimeSerial(TempDisplayHour, TempDisplayMin, TempDisplaySecond)
  
  strName = "AO-10"
'  bResult = cDJTSGP.IsVisible("ISS (ZARYA)", dTime, m_SetAOSLOS, bSouthbound, TempSatAz, TempSatElev, TempSatLon, TempSatlat, TempSatRange, TempRangeRate, TempSatAlt, dPhase)
'  bResult = cDJTSGP.IsVisible("AO-10", dTime, -1, bSouthbound, TempSatAz, TempSatElev, TempSatLon, TempSatlat, TempSatRange, TempRangeRate, TempSatAlt, dPhase)
  bResult = cDJTSGP.IsVisible(strName, dTime, m_SetAOSLOS, bSouthbound, TempSatAz, TempSatElev, TempSatLon, TempSatlat, TempSatRange, TempRangeRate, TempSatAlt, dPhase)

strTemp = cDJTSGP.PassListForPeriod("ISS (ZARYA)", dTime, 1, 1, -1)
End Sub
Private Sub sgp(SatNumber As Integer, Setup As Boolean, bPosOnly As Boolean)

  Dim nTemp As Integer
  Dim bResult As Boolean
  Dim sMisc As tagMisc
  
  sTime.wHour = TempDisplayHour
  sTime.wMinute = TempDisplayMin
  sTime.wSecond = TempDisplaySecond
  sTime.wYear = TempDisplayYear
  sTime.wMonth = TempDisplayMonth
  sTime.wDay = TempDisplayDay
  
  bResult = AGCalcSGP(m_strLine0(SatNumber), m_strLine1(SatNumber), m_strLine2(SatNumber), sTime, vOPos, vPos, vVel, vObs, vSatPos, sMisc)

If bResult = False Then
  Beep
Else
  TempSatAz = FNDEG(vObs.x)
  TempSatElev = FNDEG(vObs.y)
  TempRS = RE + vSatPos.z
  
  TempSquintAngle = 0

  TempSatLon = vSatPos.y
  TempSatlat = vSatPos.x
  TempSatRange = vObs.z
  TempRangeRate = vObs.w
  TempSatOrbit = Round(sMisc.dOrbit, 2)
End If

End Sub
Private Sub Plan13(SatNumber As Integer, Setup As Boolean, bPosOnly As Boolean)

  If Setup Then
    SetupSatellite SatNumber
  End If

  ElapsedTimeSinceEpoch = (TempSatDayNum - SatEpochDayNumber) + (TempSatTimeReq - mvarKepsYearEpochTimeFraction(SatNumber))
  SatMeanMotion = SatDragCoeff * ElapsedTimeSinceEpoch / 2
  SatMeanMotionMinute = 1 + 4 * SatMeanMotion
  SatLinearDrag = 1 - 7 * SatMeanMotion
  m = SatKepsMeanAnomoly + SatKepsMeanMotion * ElapsedTimeSinceEpoch * (1 - 3 * SatMeanMotion)
  DR = Int(m / (2 * PI))
  TempSatOrbit = m_KepsOrbitNumber(SatNumber) + (m / (2 * PI))
  If m_KepsOrbitNumber(SatNumber) > 0 Then
    TempSatMA = TempSatOrbit - Int(TempSatOrbit)
    TempSatMA = Int(TempSatMA * 256#)
  Else
    TempSatMA = 0
  End If
  TempSatOrbit = Round(TempSatOrbit, 2)
  m = m - DR * 2 * PI  ' Mean anomoly
'  TempSatOrbit = m_KepsOrbitNumber(SatNumber) + DR

  EA = m
  Do
    c = Cos(EA)
    S = Sin(EA)
    DNOM = 1 - m_KepsEccentricity(SatNumber) * c
    d = (EA - m_KepsEccentricity(SatNumber) * S - m) / DNOM
    EA = EA - d
  Loop Until Abs(d) < 0.00001

  a = EarthRotationRateSeconds * SatMeanMotionMinute
  b = B0 * SatMeanMotionMinute
  TempRS = a * DNOM
  sX = a * (c - m_KepsEccentricity(SatNumber))
  vx = -a * S / DNOM * N0
  sY = b * S
  vy = b * c / DNOM * N0
  AP = SatKepsArgOfPerigee + WD * ElapsedTimeSinceEpoch * SatLinearDrag
  cw = Cos(AP)
  SW = Sin(AP)
  RAAN = SatKepsRAAN + QD * ElapsedTimeSinceEpoch * SatLinearDrag
  CQ = Cos(RAAN)
  SQ = Sin(RAAN)
  CXx = cw * CQ - SW * CI * SQ
  CXy = -SW * CQ - cw * CI * SQ
  CXz = SI * SQ
  CYx = cw * SQ + SW * CI * CQ
  CYy = -SW * SQ + cw * CI * CQ
  CYz = -SI * CQ
  CZx = SW * SI
  CZy = cw * SI
  CZz = CI
  SatX = sX * CXx + sY * CXy
  ANTx = Ax * CXx + Ay * CXy + Az * CXz
  VELx = vx * CXx + vy * CXy
  SatY = sX * CYx + sY * CYy
  ANTy = Ax * CYx + Ay * CYy + Az * CYz
  VELy = vx * CYx + vy * CYy
  SatZ = sX * CZx + sY * CZy
  ANTz = Ax * CZx + Ay * CZy + Az * CZz
  VELz = vx * CZx + vy * CZy

  GHAA = GHAE + EarthRotationRateDay * ElapsedTimeSinceEpoch
  c = Cos(-GHAA)
  S = Sin(-GHAA)
  sX = SatX * c - SatY * S
  Ax = ANTx * c - ANTy * S
  vx = VELx * c - VELy * S
  sY = SatX * S + SatY * c
  Ay = ANTx * S + ANTy * c
  vy = VELx * S + VELy * c
  Sz = SatZ
  Az = ANTz
  vz = VELz

  Rx = sX - oX
  Ry = sY - oY
  Rz = Sz - Oz
  TempSatRange = Sqr(Rx ^ 2 + Ry ^ 2 + Rz ^ 2)

  Rx = Rx / TempSatRange
  Ry = Ry / TempSatRange
  Rz = Rz / TempSatRange

  U = Rx * Ux + Ry * Uy + Rz * Uz
  e = Rx * Ex + Ry * Ey
  N = Rx * Nx + Ry * Ny + Rz * Nz

  TempSatAz = FNIntDEG(FNAtn(e, N))
  TempSatElev = FNIntDEG(FNASN(U))

  TempSquintAngle = FNIntDEG(FNACS(-(Ax * Rx + Ay * Ry + Az * Rz)))

  'TempSquintAngle = 0

  TempSatLon = FNDEG(FNAtn(sY, sX))
  TempSatlat = FNDEG(FNASN(Sz / TempRS))
  
  TempRangeRate = (vx - VOx) * Rx + (vy - VOy) * Ry + vz * Rz

'  If Not bPosOnly Then
'    TempRangeRate = (vx - VOx) * Rx + (vy - VOy) * Ry + vz * Rz
'    'Calculate Path loss and Doppler Frequency
'    If m_DownLinkFrequency(SatNumber) <> 0 Then
'      TempSatPathLoss = Format(20# * (Log((PI * 4) * TempSatRange / CVAC * m_DownLinkFrequency(SatNumber)) / Log(10)), "##.##")
'    Else
'      TempSatPathLoss = 0
'    End If
'    If m_UplinkFrequency(SatNumber) <> 0 Then
'      TempSatUplinkDoppler = Format(m_UplinkFrequency(SatNumber) + (m_UplinkFrequency(SatNumber) * TempRangeRate / CVAC), "####.######")
'    Else
'      TempSatUplinkDoppler = 0
'    End If
'    If m_DownLinkFrequency(SatNumber) <> 0 Then
'      TempSatDownlinkDoppler = Format(m_DownLinkFrequency(SatNumber) + (-m_DownLinkFrequency(SatNumber) * TempRangeRate / CVAC), "####.######")
'    Else
'      TempSatDownlinkDoppler = 0
'    End If
'  End If

  If m_EnableSatStatus Then
    If SatNumber = -1 Then
      YG = 1990
      G0 = 99.4033
      MAS0 = 356.6349
      MASD = 0.98560027
      INS = FNRAD(23.4406)
      CNS = Cos(INS)
      SNS = Sin(INS)
      EQC1 = 0.03343
      EQC2 = 0.00034


      TEG = (TempSatDayNum - FNDAy(YG, 1, 0)) + mvarKepsYearEpochTimeFraction(SatNumber)
      GHAE = FNRAD(G0) + TEG * EarthRotationRateDay
      MRSE = FNRAD(G0) + TEG * EarthRotationRate + PI
      MASE = FNRAD(MAS0 + MASD * TEG)

      MAS = MASE + FNRAD(MASD * ElapsedTimeSinceEpoch)
      TAS = MRSE + EarthRotationRate * ElapsedTimeSinceEpoch + EQC1 * Sin(MAS) + EQC2 * Sin(2 * MAS)
      c = Cos(TAS)
      S = Sin(TAS)
      Sunx = c
      Suny = S * CNS
      Sunz = S * SNS

      bGotSun = True
    End If

    If bGotSun And SatNumber > 0 Then
      SSA = -(Ax * Sunx + Ay * Suny + ANTz * Sunz)
      ILL = Sqr(1 - SSA * SSA)
      CUA = -(sX * Sunx + sY * Suny + Sz * Sunz) / TempRS
      UMD = TempRS * Sqr(1 - CUA * CUA) / RE
      If CUA >= 0 Then
        TempSatStatusText = "Ill"
      Else
        TempSatStatusText = "Dck"
      End If
      If UMD <= 1 And CUA >= 0 Then
        TempSatStatusText = "Ecl"
      End If
      
      c = Cos(-GHAA)
      S = Sin(-GHAA)
      Hx = Sunx * c - Suny * S
      Hy = Sunx * S + Suny * c
      Hz = Sunz
      If TempSatElev > m_SetAOSLOS Then
        If (Hx * Ux + Hy * Uy + Hz * Uz < -0.17) And (TempSatStatusText <> "Ecl") Then
          TempSatStatusText = "Vis"
        End If
      End If
    End If
  End If
End Sub
Public Function CalculateALLPositions() As Boolean
  Dim i As Integer

  For i = -1 To nSatCount
    If m_SatelliteName(i) = "" Then Exit For
    CalculateSatellitePosition True, i
  Next i

End Function

Public Function DrawFootprints(Optional SatelliteIndex As Integer) As Boolean
  Dim i As Integer
  Dim vMins As Variant
  Dim vHours As Variant
  Dim strTemp As String
  Dim bDummy As Boolean
  Dim dLon As Double
  Dim dLat As Double
  Dim lTime As Long
  Dim nWidth As Integer
  Dim nHeight As Integer
  Dim nLats() As Double
  Dim nLons() As Double
  Dim lColours() As Long
  Dim strLocPath As String

  If Not bStarted Then
    If m_UseHourglass Then
      cWaitCur.SetCursor vbHourglass
    End If
    Select Case m_OutputStyle
      Case 0
        TabList.ListItems.Clear
        
        For i = 1 To nSatCount
          If m_SatelliteName(i) = "" Then Exit For
          Set lItem = TabList.ListItems.Add(, , m_SatelliteName(i))
          lItem.SubItems(1) = m_SatelliteDesignator(i)
          lItem.SubItems(2) = Format(m_SatelliteOrbitNumber(i), "######0.00")
          dLat = m_SatelliteLatitude(i)
          If dLat < 0 Then
            strTemp = Format(Str(Abs(dLat)), "##0.00") & "°S  "
          Else
            strTemp = Format(Str(dLat), "##0.00") & "°N  "
          End If
          lItem.SubItems(3) = strTemp
          dLon = 360 - m_SatelliteLongitude(i)
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
          lItem.SubItems(4) = strTemp
          lItem.SubItems(5) = m_SatelliteElevation(i)
          lItem.SubItems(6) = m_SatelliteAzimuth(i)
          lItem.SubItems(7) = m_SatelliteRange(i)
          lItem.SubItems(8) = m_SatelliteAltitude(i)
          If m_SatelliteElevation(i) > m_SetAOSLOS Then
            lItem.SubItems(9) = m_SatelliteRxFrequency(i) + m_SatelliteTXFrequency(i)
          Else
            lItem.SubItems(9) = "<Unavailable>"
          End If
          If satTrackNextAOS(i) <> "" Then
            strTemp = Format(satTrackNextAOS(i), "General Date")
            vMins = DateDiff("n", Now, satTrackNextAOS(i))
            vHours = Int(vMins / 60)
            vMins = vMins - vHours * 60
            strTemp = strTemp & " in " & vHours & "hr " & vMins & "mins"
          Else
            strTemp = "Next AOS: <Unavailable>"
          End If
          lItem.SubItems(10) = strTemp
          Select Case m_OrbitalModel(i)
            Case agplan13
              lItem.SubItems(11) = "Plan 13"
            Case agSGP
              lItem.SubItems(11) = "SGP4/SDP4"
            Case agdtsgp
              lItem.SubItems(11) = "Alt SGP"
          End Select
          If lItem.Text = m_SatelliteName(SelectedSatellite) Then
            Set TabList.SelectedItem = lItem
          End If
        Next i
        UpdateSatelliteLabel
      Case 1
        picInner.Cls

        PlotObserver
        UpdateDataWindow
        UpdateSatelliteLabel

        For i = 1 To nSatCount
          If m_SatelliteName(i) = "" Then Exit For
          DrawFootprintForObject i
          PlotSatellite i
          If m_DisplayTracks Then
            PlotTrack i
          End If
        Next i

        If m_DisplayMoon Then
          PlotSatellite 0
          If m_DisplayMoonFootprint Then
            DrawFootprintForObject 0
          End If
        End If
        If m_DisplaySun Then
          PlotSatellite -1
          If m_DisplaySunFootprint Then
            DrawFootprintForObject -1
          End If
        End If

        If SelectedSatellite > 0 And m_DisplayAOSCircle Then
          DrawFootprintForObject SelectedSatellite, ObserverLatitude, ObserverLongitude
        End If
      Case 2
        picInner.Cls

        UpdateDataWindow
        UpdateSatelliteLabel

        For i = 1 To nSatCount
          If m_SatelliteName(i) = "" Then Exit For
          PlotSatellite i
          If m_DisplayTracks Then
            PlotTrack i
          End If
        Next i

        If m_DisplayMoon Then
          PlotSatellite 0
        End If
        If m_DisplaySun Then
          PlotSatellite -1
        End If
      Case 3
        nWidth = UserControl.ScaleX(UserControl.Width, vbTwips, vbPixels)
        nHeight = UserControl.ScaleY(UserControl.Height, vbTwips, vbPixels)
        nWidth = 500
        ReDim Preserve nLats(SatTrackPoints(SelectedSatellite))
        ReDim Preserve nLons(SatTrackPoints(SelectedSatellite))
        ReDim Preserve lColours(SatTrackPoints(SelectedSatellite))
        For i = 0 To SatTrackPoints(SelectedSatellite)
          nLats(i) = SatTrackLat(SelectedSatellite, i)
          dLon = 360 - SatTrackLon(SelectedSatellite, i)
          Select Case dLon
            Case 0 To 180
              dLon = -dLon
            Case 181 To 360
              dLon = (360 - dLon)
          End Select
          nLons(i) = dLon
          lColours(i) = SatTrackMutual(SelectedSatellite, i)
        Next i
        dLon = 360 - m_SatelliteLongitude(SelectedSatellite)
        Select Case dLon
          Case 0 To 180
            dLon = -dLon
          Case 181 To 360
            dLon = (360 - dLon)
        End Select
        'strTemp = m_DisplayDay(SatNumber) & "/" & m_DisplayMonth(SatNumber) & "/" & m_DisplayYear(SatNumber) & " " & m_DisplayHour(SatNumber) & ":" & m_DisplayMinute(SatNumber) & ":" & m_DisplaySecond(SatNumber)
        strTemp = ConvertToGMT(m_DisplayYear(SatNumber), m_DisplayMonth(SatNumber), m_DisplayDay(SatNumber), m_DisplayHour(SatNumber), m_DisplayMinute(SatNumber), m_DisplaySecond(SatNumber))
        lTime = DateDiff("s", "01/01/70", strTemp)
        If m_ViewsOrthLocations = "" Then
          strLocPath = "built-in"
        Else
          strLocPath = m_ViewsOrthLocations
        End If
        bDummy = GenerateGlobe(m_SatelliteLatitude(SelectedSatellite), dLon, 600, nWidth, False, True, strLocPath, 100, 30, 1, m_ViewsOrthShade, False, True, lTime, m_SatelliteLatitude(SelectedSatellite), dLon, m_SatelliteName(SelectedSatellite), nLats, nLons, lColours)
        If bDummy Then
          picInner.Picture = LoadPicture(strTempPath & "AGGlobe.bmp")
        Else
          Beep
        End If
'         CreateDisplay
      '  RefreshGlobe
        UpdateSatelliteLabel
    End Select

    If m_SetActiveWindowAsWallpaper Then
      Select Case m_OutputStyle
        Case 1, 2
          picPic.Picture = picInner.Image
          picPic.Refresh
          SavePicture UserControl.picPic.Picture, strTempPath & "AGMerc.bmp"
          vblSetDesktopWallpaper strTempPath & "AGMerc.bmp"
          picPic.Picture = LoadPicture
        Case 3
          vblSetDesktopWallpaper strTempPath & "AGGlobe.bmp"
      End Select
    End If

    If m_Enable847 Then
      If m_SatelliteTXFrequency(SelectedSatellite) <> 0 Then
        If m_SatelliteRxFrequency(SelectedSatellite) <> 0 Then
          frmRadio.SetVFO True, m_SatelliteTXFrequency(SelectedSatellite)
          frmRadio.SetVFO False, m_SatelliteRxFrequency(SelectedSatellite)
        End If
      End If
    End If
  End If
  'picInner.Refresh
  If m_UseHourglass Then
    cWaitCur.Restore
  End If
End Function
Private Function vblSetDesktopWallpaper(sFileName As String) As Boolean

  vblSetDesktopWallpaper = _
  CBool(SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, ByVal sFileName, True))

End Function

Private Sub DrawFootprintForObject(i As Integer, Optional oLat As Single = -999, Optional oLon As Single = -999)
  Dim clo As Double
  Dim cla As Double
  Dim TempX As Double
  Dim TempY As Double
  Dim TempZ As Double
  Dim LineColour As Single
  Dim ScreenLON As Integer
  Dim ScreenLAT As Integer
  Dim SatX As Double
  Dim SatY As Double
  Dim SatZ As Double
  Dim LastScreenLON As Integer
  Dim LastScreenLAT As Integer
  Dim ScrollCheck As Integer
  Dim srad As Double
  Dim a As Double
  Dim Counter As Integer
  Dim csrad As Double
  Dim ssrad As Double
  Dim ccla As Double
  Dim scla As Double
  Dim cclo As Double
  Dim sclo As Double
  Dim sLat As Single
  Dim sLon As Single
  
  If oLat = -999 Then
    sLon = m_SatelliteLatitude(i)
    sLat = m_SatelliteLongitude(i)
  Else
    sLon = oLat
    sLat = oLon
  End If
  
  If m_DisplayFootprints Then
      
      Counter = 0
      LastScreenLON = 0
      LastScreenLAT = 0
      If m_DataValid(i) Then

        srad = FNACS(RE / RS(i))
        clo = FNRAD(sLat)
        cla = FNRAD(sLon)

        Select Case i
          Case -1
              LineColour = RGB(255, 255, 0)
          Case 0
              LineColour = RGB(192, 192, 192)
          Case Else
            If oLat = -999 Then
              LineColour = RGB(255, 255, 0)
            Else
              LineColour = RGB(255, 64, 255)
            End If
        End Select
        
        csrad = Cos(srad)
        ssrad = Sin(srad)
        ccla = Cos(cla)
        scla = Sin(cla)
        cclo = Cos(clo)
        sclo = Sin(clo)
        
        For a = 0 To TwoPI Step TwoPI / PointsToDraw
          TempX = csrad
          TempY = ssrad * Sin(a)
          TempZ = ssrad * Cos(a)
          SatX = TempX * ccla - TempZ * scla
          SatY = TempY
          SatZ = TempX * scla + TempZ * ccla
          TempX = SatX * cclo - SatY * sclo
          TempY = SatX * sclo + SatY * cclo
          TempZ = SatZ
          FootprintLON(Counter) = (FNAtn(TempY, TempX))
          FootprintLAT(Counter) = (FNASN(TempZ))
          ScreenLON = ConvertLONToScreen(FootprintLON(Counter) * cIntDeg)
          ScreenLAT = ConvertLATToScreen(FootprintLAT(Counter) * cIntDeg)

          ScrollCheck = (ScreenLON - LastScreenLON) * Sgn(ScreenLON - LastScreenLON)
          
          Select Case Counter
            Case 0
              picInner.PSet (ScreenLON, ScreenLAT)
            Case Else
              Select Case ScrollCheck
                Case Is < 400
                  picInner.Line (LastScreenLON, LastScreenLAT)-(ScreenLON, ScreenLAT), LineColour!
                Case Is >= 400
                  picInner.PSet (ScreenLON, ScreenLAT)
              End Select
          End Select
          LastScreenLON = ScreenLON
          LastScreenLAT = ScreenLAT
          Counter = Counter + 1
        Next

      End If
    End If
End Sub
Private Sub PlotSatellite(i As Integer)
  Dim ScreenLON As Integer
  Dim ScreenLAT As Integer
  Dim strTemp As String
  Dim nTempFontSize As Integer
  Dim hdcImageSource As Long
  Dim hdcMaskSource As Long
  Dim j As Integer
  Dim nPos As Integer
  
  Select Case m_OutputStyle
    Case 1

      ScreenLON = m_SatelliteLongitude(i)
      ScreenLAT = m_SatelliteLatitude(i)
      ScreenLON = ConvertLONToScreen((ScreenLON))
      ScreenLAT = ConvertLATToScreen((ScreenLAT))
      SatScreenX(i) = ScreenLON
      SatScreenY(i) = ScreenLAT
    Case 2

      ScreenLON = m_SatelliteAzimuth(i)
      ScreenLAT = m_SatelliteElevation(i)
      ScreenLON = ScreenLON * PixelsPerDegLon
      ScreenLAT = MapHeight - ScreenLAT * (PixelsPerDegLat * 2)
      SatScreenX(i) = ScreenLON
      SatScreenY(i) = ScreenLAT
  End Select
  
  nPos = 7

  Select Case m_OutputStyle
    Case 1, 2
      Select Case i
        Case -1
          picInner.FillColor = RGB(255, 255, 0)
          If m_DisplayIcons Then
            nPos = 32
            picImage.Picture = UserControl.ImgObjects.ListImages(1).Picture
            picMask.Picture = UserControl.ImgMasks.ListImages(1).Picture
          End If
        Case 0
          picInner.FillColor = RGB(128, 128, 128)
          If m_DisplayIcons Then
            nPos = 32
            picImage.Picture = UserControl.ImgObjects.ListImages(2).Picture
            picMask.Picture = UserControl.ImgMasks.ListImages(2).Picture
          End If
        Case Else
          picInner.FillColor = RGB(255, 0, 0)
          If m_DisplayIcons Then
            nPos = 32
            If InStr(m_SatelliteName(i), "STS") <> 0 Then
              picImage.Picture = UserControl.ImgObjects.ListImages(3).Picture
              picMask.Picture = UserControl.ImgMasks.ListImages(3).Picture
            Else
              picImage.Picture = UserControl.ImgObjects.ListImages(4).Picture
              picMask.Picture = UserControl.ImgMasks.ListImages(4).Picture
              For j = 5 To UserControl.ImgObjects.ListImages.Count
                If m_SatelliteDesignator(i) = UserControl.ImgObjects.ListImages(j).Tag Then
                  picImage.Picture = UserControl.ImgObjects.ListImages(j).Picture
                  picMask.Picture = UserControl.ImgMasks.ListImages(j).Picture
                  Exit For
                End If
              Next j
            End If
          End If
      End Select

      If m_DisplayIcons Then
   '     UserControl.picBuffer.ScaleHeight = 64
   '     UserControl.picBuffer.ScaleWidth = 64
        rc% = BitBlt(UserControl.picBuffer.hDC, 0, 0, 64, 64, picInner.hDC, ScreenLON - 32, ScreenLAT - 32, SRCCOPY) 'copies bg to buffer
        rc% = BitBlt(UserControl.picBuffer.hDC, 0, 0, 64, 64, picMask.hDC, 0, 0, SRCPAINT) 'masks
        rc% = BitBlt(UserControl.picBuffer.hDC, 0, 0, 64, 64, picImage.hDC, 0, 0, SRCAND) 'masks
        rc% = BitBlt(picInner.hDC, ScreenLON - 32, ScreenLAT - 32, 64, 64, UserControl.picBuffer.hDC, 0, 0, SRCCOPY) 'copies product to form1
      Else
        picInner.Circle (ScreenLON, ScreenLAT), 2, 0
      End If

      If (m_DisplayIcons And i > 0) Or Not (m_DisplayIcons) Then
        picInner.ForeColor = RGB(255, 255, 255)
        picInner.CurrentX = ScreenLON + nPos
        picInner.CurrentY = ScreenLAT - 7
        If picInner.CurrentX + (7 * Len(m_SatelliteName(i))) > picInner.ScaleWidth Then
          picInner.CurrentX = (picInner.CurrentX - (nPos * 2)) - 7 * Len(m_SatelliteName(i))
        End If
        picInner.Print m_SatelliteName(i)
      End If
      If m_DisplayTimes Then
        nTempFontSize = picInner.Font.Size
        picInner.Font.Size = 6
        picInner.CurrentX = ScreenLON
        picInner.CurrentY = ScreenLAT + 6
        picInner.ForeColor = RGB(255, 0, 0)
        strTemp = Format(m_DisplayHour(i), "00") & ":" & Format(m_DisplayMinute(i), "00") & " " & Format(m_DisplayDay(i), "00") & "/" & Format(m_DisplayMonth(i), "00") & "/" & Format(m_DisplayYear(i), "00")
        picInner.Print strTemp
        picInner.Font.Size = nTempFontSize
      End If
      UpdateSatelliteLabel
    Case 4
      '      Viewport.Clear
      '      Viewport.Render Scene
      '      Device.Update
  End Select
End Sub
 Private Function ConvertScreenToLAT(y As Single) As Single
    ConvertScreenToLAT = 90 - y / PixelsPerDegLat
End Function
Private Function ConvertScreenToLON(x As Single) As Single
    If m_ObserverMapCentre = 0 Then
        ConvertScreenToLON = x / PixelsPerDegLon - 180
    Else
        ConvertScreenToLON = x / PixelsPerDegLon
    End If
End Function
Private Function ConvertLATToScreen(sLat As Single) As Variant
  ConvertLATToScreen = HalfMapHeight% - sLat * PixelsPerDegLat
End Function

Private Function ConvertLONToScreen(sLon As Integer) As Variant
  sLon% = 360 - sLon%
  Select Case m_ObserverMapCentre
    Case 0
      Select Case sLon%
        Case 0 To 180
          ConvertLONToScreen = HalfMapWidth% - sLon% * PixelsPerDegLon
        Case 181 To 360
          ConvertLONToScreen = HalfMapWidth% + (360 - sLon%) * PixelsPerDegLon
      End Select
    Case 180
      ConvertLONToScreen = MapWidth% - sLon% * PixelsPerDegLon
  End Select
End Function
Private Sub SetupSatellite(nSatNum As Integer)

  Dim xx As Variant
  Dim ZZ As Variant
  Dim d As Variant

  sLon = 180
  sLat = 0
  CL = Cos(TempObsLat)
  SL = Sin(TempObsLat)
  CO = Cos(TempObsLon)
  SO = Sin(TempObsLon)

  RP = RE * (1 - FL)
  xx = RE * RE
  ZZ = RP * RP

  d = Sqr(xx * CL * CL + ZZ * SL * SL)
  Rx = xx / d + mvarObserverHeight
  Rz = ZZ / d + mvarObserverHeight

  Ux = CL * CO
  Ex = -SO
  Nx = -SL * CO
  Uy = CL * SO
  Ey = CO
  Ny = -SL * SO
  Uz = SL
  Ez = 0
  Nz = CL
  oX = Rx * Ux
  oY = Rx * Uy
  Oz = Rz * Uz

  SatKepsRAAN = FNRAD(m_KepsRAAN(nSatNum))
  SatKepsInclination = FNRAD(m_KepsInclination(nSatNum))
  SatKepsArgOfPerigee = FNRAD(m_KepsAOP(nSatNum))
  SatKepsMeanAnomoly = FNRAD(m_KepsMeanAnomoly(nSatNum))
  SatKepsMeanMotion = m_KepsMeanMotion(nSatNum) * 2 * PI
  
  EarthRotationRateSeconds = EarthRotationRateDay / 86400

  M2 = m_KepsDecayRate(nSatNum) * 2 * PI
  VOx = -oY * EarthRotationRateSeconds
  VOy = oX * EarthRotationRateSeconds

  SatEpochDayNumber = FNDAy(mvarSatEpochYear(nSatNum), 1, 0) + Int(mvarKepsYearEpochTime(nSatNum))
  mvarKepsYearEpochTimeFraction(nSatNum) = mvarKepsYearEpochTime(nSatNum) - Int(mvarKepsYearEpochTime(nSatNum))
  N0 = SatKepsMeanMotion / 86400
  EarthRotationRateSeconds = (GravitationalConstant / N0 / N0) ^ (1 / 3)
  B0 = EarthRotationRateSeconds * Sqr(1 - m_KepsEccentricity(nSatNum) * m_KepsEccentricity(nSatNum))
  SI = Sin(SatKepsInclination)
  CI = Cos(SatKepsInclination)
  PC = RE * EarthRotationRateSeconds / (B0 * B0)
  PC = 1.5 * ZonalCoeff * PC * PC * SatKepsMeanMotion
  QD = -PC * CI
  WD = PC * (5 * CI * CI - 1) / 2
  SatDragCoeff = -2 * M2 / SatKepsMeanMotion / 3
  TEG = (SatEpochDayNumber - FNDAy(YG, 1, 0)) + mvarKepsYearEpochTimeFraction(nSatNum)
  GHAE = FNRAD(G0) + TEG * EarthRotationRateDay
  CO = Cos(FNRAD(sLon))
  SO = Sin(FNRAD(sLon))
  CL = Cos(FNRAD(sLat))
  SL = Sin(FNRAD(sLat))
  Ax = -CL * CO
  Ay = -CL * SO
  Az = -SL

' Calculate Maximum communications distance

  T = 1440 / m_KepsMeanMotion(nSatNum)
  SMA = 331.25 * T ^ (2 / 3)
  HP = SMA * (1 - m_KepsEccentricity(nSatNum)) - RE
  HA = SMA * (1 + m_KepsEccentricity(nSatNum)) - RE
  z = RE / (HA + RE)
  m_SatelliteMaximumDX(nSatNum) = Int(RE * (PI / 2 - Atn(z / Sqr(1 - z * z))))

End Sub
Private Function FNACS(x As Variant) As Variant
  On Error GoTo ERROR_FNACS

  FNACS = PI / 2 - Atn(x / Sqr(1 - x ^ 2))

EXIT_FNACS:
  Exit Function

ERROR_FNACS:
'  MsgBox "Error in ERROR_FNACS : " & Error
  Resume EXIT_FNACS

End Function

Private Function FNASN(vValue As Variant) As Variant
On Error Resume Next
  FNASN = Atn(vValue / Sqr(1 - vValue ^ 2))
End Function

Private Function FNAtn(y As Variant, x As Variant) As Variant

  Dim Result As Variant

  If x <> 0 Then
    Result = Atn(y / x)
  Else
    Result = PI / 2 * Sgn(y)
  End If
  If x < 0 Then
    Result = Result + PI
  End If
  If Result < 0 Then
    Result = Result + 2 * PI
  End If
  FNAtn = Result
End Function
Private Function FNIntDEG(x As Variant) As Variant
  FNIntDEG = Int(x * 180 / PI)
End Function
Private Function FNDEG(x As Variant) As Variant
  FNDEG = x * 180 / PI
End Function

Private Function FNRAD(x As Variant) As Variant
  FNRAD = x * PI / 180
End Function
Private Function FNDAy(y As Variant, m As Variant, d As Variant) As Variant
  Dim TempY As Variant
  Dim TempM As Variant
  Dim TempD As Variant

  TempY = y
  TempM = m
  TempD = d

  If TempM <= 2 Then TempY = TempY - 1: TempM = TempM + 12
  FNDAy = Int(TempY * MeanYear) + Int((TempM + 1) * 30.6) + TempD - 428
End Function
Public Sub PlotObserver()

DrawObserver Me.ObserverLatitude, Me.ObserverLongitude, Me.ObserverLocation, RGB(255, 255, 0)

If Me.SecondObserverEnabled Then
  DrawObserver Me.SecondObserverLatitude, Me.SecondObserverLongitude, Me.SecondObserverLocation, RGB(192, 192, 192)
End If

End Sub
Private Sub DrawObserver(sLat As Single, sLon As Single, strName As String, lColour As Long)
  Dim ScreenLON As Integer
  Dim ScreenLAT As Integer

  ScreenLAT = ConvertLATToScreen(sLat)
    
  Select Case m_ObserverMapCentre
    Case 0
      ScreenLON = HalfMapWidth% + sLon * PixelsPerDegLon
    Case 180
      ScreenLON = sLon * PixelsPerDegLon
      If ScreenLON < 0 Then
        ScreenLON = MapWidth% + ScreenLON
      End If
  End Select

  picInner.Line (ScreenLON - 5, ScreenLAT)-Step(10, 0), lColour
  picInner.Line (ScreenLON, ScreenLAT - 5)-Step(0, 10), lColour

  picInner.ForeColor = RGB(255, 255, 255)
  picInner.CurrentX = ScreenLON + 8
  picInner.CurrentY = ScreenLAT - 6
  picInner.Print strName

End Sub
Private Sub CalculateTrack(Optional SatIndex As Integer)
  Dim SatNumber As Integer
  Dim i As Integer
  Dim nCounter As Integer
  Dim SetupFlag As Boolean
  Dim TempHour As Integer
  Dim TempMin As Integer
  Dim TempDay As Integer
  Dim bDone As Boolean
  Dim LastLat As Single
  Dim LastLon As Single
  Dim LastElev As Single
  Dim LastAz As Single
  Dim Period As Long
  Dim EndTime As Variant
  Dim TimerCounter As Long
  Dim Increment As Long
  Dim vDate As Variant
  Dim nLastEle As Integer
  Dim bGotEle As Boolean

  If IsMissing(SatIndex) Then
    SatNumber = m_SatelliteIndex
  Else
    SatNumber = SatIndex
  End If

  If m_DataValid(SatNumber) Then
    TempDisplayYear = m_DisplayYear(SatNumber)
    TempDisplayMonth = m_DisplayMonth(SatNumber)
    TempDisplayDay = m_DisplayDay(SatNumber)
    TempDisplayHour = m_DisplayHour(SatNumber)
    TempDisplayMin = m_DisplayMinute(SatNumber)
    TempDisplaySecond = m_DisplaySecond(SatNumber)
    SetupFlag = True
    bDone = False

    LocalDateToUTC = ConvertToGMT(TempDisplayYear, TempDisplayMonth, TempDisplayDay, TempDisplayHour, TempDisplayMin, TempDisplaySecond)

    TempDisplayYear = Year(LocalDateToUTC)
    TempDisplayMonth = Month(LocalDateToUTC)
    TempDisplayDay = Day(LocalDateToUTC)
    TempDisplayHour = hour(LocalDateToUTC)
    TempDisplayMin = Minute(LocalDateToUTC)
    TempDisplaySecond = Second(LocalDateToUTC)

    TempHour = TempDisplayHour
    TempMin = TempDisplayMin
    TempDay = TempDisplayDay
    nCounter = 0

    LastLat = 0
    LastLon = 0
    LastElev = 0
    LastAz = 0

    satTrackNextAOS(SatNumber) = ""
    satTrackMaxEle(SatNumber) = 0
    satTrackMaxEleTime(SatNumber) = ""
    If m_SatelliteTrackOrbits(SatNumber) <= 0 Then
      m_SatelliteTrackOrbits(SatNumber) = 1
    End If
    If m_SatelliteTrackOrbits(SatNumber) > 10 Then
      m_SatelliteTrackOrbits(SatNumber) = 10
    End If

    Period = ((1440# / m_KepsMeanMotion(SatNumber)) + 0.5) * 60#
    If SatNumber = -1 Or SatNumber = 0 Then
      Period = 86400
    End If

    TimerCounter = 0
    If m_GroundTrackInterval = 1 Then
      Increment = Period / MaxPoints
      If Increment < 60 Then
        Increment = 60
      End If
    Else
      Increment = m_GroundTrackInterval
    End If

    TempObsLat = mvarObserverLatitude
    TempObsLon = mvarObserverLongitude

    nLastEle = 999
    bGotEle = False

    Do
      TempSatDayNum = FNDAy(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
      TempSatTimeReq = (TempDisplayHour + (TempDisplayMin / 60) + (TempDisplaySecond / 60 / 60)) / 24
      PositionEngine SatNumber, SetupFlag, True
      SetupFlag = False
      If LastLat <> TempSatlat Or LastLon <> TempSatLon Then
        SatTrackLon(SatNumber, nCounter) = TempSatLon
        SatTrackLat(SatNumber, nCounter) = TempSatlat
        SatTrackElev(SatNumber, nCounter) = TempSatElev
        SatTrackAzim(SatNumber, nCounter) = TempSatAz
       ' SatTrackTime(SatNumber, nCounter) = DateSerial(TempDisplayYear, TempDisplayMonth, TempDisplayDay) + TimeSerial(TempDisplayHour, TempDisplayMin, TempDisplaySecond)
        If TempSatElev > m_SetAOSLOS And satTrackNextAOS(SatNumber) = "" Then
          vDate = DateSerial(TempDisplayYear, TempDisplayMonth, TempDisplayDay) + TimeSerial(TempDisplayHour, TempDisplayMin, TempDisplaySecond)
          If m_Timezone <> 0 Then
            vDate = DateAdd("h", -m_Timezone, vDate)
          End If
          If m_DaylightSavingAdjust <> 0 Then
            vDate = DateAdd("h", -m_DaylightSavingAdjust, vDate)
          End If
          satTrackNextAOS(SatNumber) = Format(vDate, "General Date")
        End If
        If TempSatElev > satTrackMaxEle(SatNumber) And Not bGotEle Then
          vDate = DateValue(TempDisplayDay & " " & TempDisplayMonth & " " & TempDisplayYear) & " " & TimeValue(TempDisplayHour & ":" & TempDisplayMin)
          If m_DaylightSaving Then
            vDate = DateAdd("h", -m_DaylightSavingAdjust, vDate)
          End If
          satTrackMaxEle(SatNumber) = TempSatElev
          satTrackMaxEleTime(SatNumber) = Trim(Str(TempSatElev)) & "deg @ " & vDate
          m_SatMaxEleTime(SatNumber) = vDate
        End If
        SatTrackMutual(SatNumber, nCounter) = RGB(0, 255, 0)
        nLastEle = TempSatElev
        If TempSatElev > m_SetAOSLOS Then
          If TempSatElev < nLastEle And Not bGotEle Then
            bGotEle = True
          End If
          If Me.SecondObserverEnabled Then
            TempObsLat = mvarSecondObserverLatitude
            TempObsLon = mvarSecondObserverLongitude
            SetupFlag = True
            TempSatDayNum = FNDAy(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
            TempSatTimeReq = (TempDisplayHour + (TempDisplayMin / 60) + (TempDisplaySecond / 60 / 60)) / 24
            PositionEngine SatNumber, SetupFlag, True
            If TempSatElev > m_SetAOSLOS Then
              SatTrackMutual(SatNumber, nCounter) = RGB(255, 255, 0)
            Else
              SatTrackMutual(SatNumber, nCounter) = RGB(255, 0, 0)
            End If
            TempObsLat = mvarObserverLatitude
            TempObsLon = mvarObserverLongitude
          Else
            SatTrackMutual(SatNumber, nCounter) = RGB(255, 0, 0)
          End If
        End If
        LastLon = TempSatLon
        LastLat = TempSatlat
        nCounter = nCounter + 1
        If nCounter > MaxPoints Then
          bDone = True
        End If
      End If
      AddTime Increment
      TimerCounter = TimerCounter + Increment
      If m_AllowDoEvents Then
        DoEvents
      End If
    Loop Until TimerCounter > (Period * m_SatelliteTrackOrbits(SatNumber)) Or bDone
    SatTrackPoints(SatNumber) = nCounter - 1
  End If
End Sub


Private Sub AddTime(nSeconds As Long)
Attribute AddTime.VB_Description = "Adds a number of seconds to the current satellite time."
    
  Dim vDate As Variant
  Dim vlDate As Variant
  Dim vlTime As Variant
  
 ' vDate = DateValue(TempDisplayDay & " " & TempDisplayMonth & " " & TempDisplayYear) & " " & TimeValue(TempDisplayHour & ":" & TempDisplayMin & ":" & TempDisplaySecond)

  vlDate = DateSerial(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
  
  vlTime = TimeSerial(TempDisplayHour, TempDisplayMin, TempDisplaySecond)

  vDate = vlDate + vlTime

  vDate = DateAdd("s", nSeconds, vDate)

  TempDisplayDay = Day(vDate)
  TempDisplayMonth = Month(vDate)
  TempDisplayYear = Year(vDate)

  TempDisplayHour = hour(vDate)
  TempDisplayMin = Minute(vDate)
  TempDisplaySecond = Second(vDate)

End Sub

Private Sub SubTime(nSeconds As Long)

  Dim vDate As Variant

  vDate = DateValue(TempDisplayDay & " " & TempDisplayMonth & " " & TempDisplayYear) & " " & TimeValue(TempDisplayHour & ":" & TempDisplayMin)

  vDate = DateAdd("s", -nSeconds, vDate)

  TempDisplayDay = Day(vDate)
  TempDisplayMonth = Month(vDate)
  TempDisplayYear = Year(vDate)

  TempDisplayHour = hour(vDate)
  TempDisplayMin = Minute(vDate)
End Sub

Sub PlotTrack(Optional SatIndex As Integer)
  Dim i As Integer
  Dim ScreenLON As Integer
  Dim ScreenLAT As Integer
  Dim SatNumber As Integer
  Dim bPlotted As Boolean
  Dim nTempFontSize As Integer
  Dim ScrollCheck As Integer
  Dim LastScreenLAT As Integer
  Dim LastScreenLON As Integer

  If IsMissing(SatIndex) Then
    SatNumber = m_SatelliteIndex
  Else
    SatNumber = SatIndex
  End If

  Select Case m_OutputStyle
    Case 0

    Case 1
      For i = 0 To SatTrackPoints(SatNumber)
        ScreenLON = SatTrackLon(SatNumber, i)
        ScreenLAT = SatTrackLat(SatNumber, i)
        ScreenLON = ConvertLONToScreen((ScreenLON))
        ScreenLAT = ConvertLATToScreen((ScreenLAT))
        '        If SatTrackElev(SatNumber, i) > m_SetAOSLOS Then
        '          picInner.FillColor = RGB(255, 0, 0)
        '        Else
        '         picInner.FillColor = RGB(0, 255, 0)
        '        End If
        'picInner.FillColor = IIf(SatTrackElev(SatNumber, i) > m_SetAOSLOS, RGB(255, 0, 0), RGB(0, 255, 0))
        picInner.FillColor = SatTrackMutual(SatNumber, i)

        If m_DisplayGroundTrackAsPoints Then
          If m_GroundTrackPointSize > 0 Then
            picInner.Circle (ScreenLON, ScreenLAT), m_GroundTrackPointSize, picInner.FillColor
          Else
            picInner.PSet (ScreenLON, ScreenLAT), picInner.FillColor
          End If
          'If SatTrackElev(SatNumber, i) > -1 Then picInner.Print SatTrackElev(SatNumber, i)
'If SatTrackElev(SatNumber, i) > -1 Then
'picInner.Print SatTrackTime(SatNumber, i)
'End If
        Else
          ScrollCheck = (ScreenLON - LastScreenLON) * Sgn(ScreenLON - LastScreenLON)

          If i = 0 Then
            picInner.PSet (ScreenLON, ScreenLAT)
          Else
            Select Case ScrollCheck
              Case Is < 400
                picInner.Line (LastScreenLON, LastScreenLAT)-(ScreenLON, ScreenLAT), picInner.FillColor
              Case Is >= 400
                picInner.PSet (ScreenLON, ScreenLAT)
            End Select
          End If
          LastScreenLON = ScreenLON
          LastScreenLAT = ScreenLAT
        End If
        '       picInner.PSet (ScreenLON, ScreenLAT), picInner.FillColor
      Next i

    Case 2
      For i = 0 To SatTrackPoints(SatNumber)
        ScreenLON = SatTrackAzim(SatNumber, i)
        ScreenLAT = SatTrackElev(SatNumber, i)
        ScreenLON = ScreenLON * PixelsPerDegLon
        ScreenLAT = MapHeight - ScreenLAT * (PixelsPerDegLat * 2)

        '        If SatTrackElev(SatNumber, i) > m_SetAOSLOS Then
        '          picInner.FillColor = RGB(0, 255, 0)
        '        Else
        '          picInner.FillColor = RGB(255, 0, 0)
        '        End If
'        picInner.FillColor = IIf(SatTrackElev(SatNumber, i) > m_SetAOSLOS, RGB(255, 0, 0), RGB(0, 255, 0))
        picInner.FillColor = SatTrackMutual(SatNumber, i)
        ScrollCheck = (ScreenLON - LastScreenLON) * Sgn(ScreenLON - LastScreenLON)
        If m_DisplayGroundTrackAsPoints Then
          If m_GroundTrackPointSize > 0 Then
            picInner.Circle (ScreenLON, ScreenLAT), m_GroundTrackPointSize, picInner.FillColor
          Else
            picInner.PSet (ScreenLON, ScreenLAT), picInner.FillColor
          End If
   '       If SatTrackElev(SatNumber, i) > -1 Then picInner.Print SatTrackElev(SatNumber, i)
        Else

          If i = 0 Then
            picInner.PSet (ScreenLON, ScreenLAT)
          Else
            Select Case ScrollCheck
              Case Is < 400
                picInner.Line (LastScreenLON, LastScreenLAT)-(ScreenLON, ScreenLAT), picInner.FillColor
              Case Is >= 400
                picInner.PSet (ScreenLON, ScreenLAT)
            End Select
          End If
          LastScreenLON = ScreenLON
          LastScreenLAT = ScreenLAT
        End If
        '        picInner.PSet (ScreenLON, ScreenLAT), picInner.FillColor
        If SatTrackElev(SatNumber, i) = satTrackMaxEle(SatNumber) And Not (bPlotted) Then
          picInner.ForeColor = RGB(255, 255, 255)
          picInner.Print m_SatelliteName(SatNumber)
          bPlotted = True

          nTempFontSize = picInner.Font.Size
          picInner.Font.Size = 6
          picInner.CurrentX = ScreenLON
          picInner.CurrentY = ScreenLAT + 15
          picInner.ForeColor = RGB(255, 0, 0)
          picInner.Print satTrackMaxEleTime(SatNumber)
          picInner.Font.Size = nTempFontSize
        End If

      Next i
  End Select
End Sub
Sub DisplayAOS(AOSType As Integer, Optional SatIndex As Integer)
Dim LastElevation As Variant
Dim GotAOS As Integer
Dim SatNumber As Integer

If IsMissing(SatIndex) Then
  SatNumber = m_SatelliteIndex
Else
  SatNumber = SatIndex
End If

LastElevation = m_SatelliteElevation(SatNumber)
GotAOS% = 0

Do
  Select Case AOSType
    Case 1
      m_DisplayMinute(SatNumber) = m_DisplayMinute(SatNumber) - 1
      CalculateSatellitePosition False, SatNumber
      If m_SatelliteElevation(SatNumber) > m_SetAOSLOS And Sgn(m_SatelliteElevation(SatNumber)) <> Sgn(LastElevation) Then
        GotAOS% = 1
      End If
      LastElevation = m_SatelliteElevation(SatNumber)
      If m_DisplayMinute(SatNumber) < 0 Then
        m_DisplayMinute(SatNumber) = m_DisplayMinute(SatNumber) + 60
        m_DisplayHour(SatNumber) = m_DisplayHour(SatNumber) - 1
      End If
      If m_DisplayHour(SatNumber) < 0 Then
        m_DisplayHour(SatNumber) = m_DisplayHour(SatNumber) + 24
        m_DisplayDay(SatNumber) = m_DisplayDay(SatNumber) - 1
      End If
    Case 2
      CalculateSatellitePosition False, SatNumber
      m_DisplayMinute(SatNumber) = m_DisplayMinute(SatNumber) + 1
      If m_SatelliteElevation(SatNumber) > m_SetAOSLOS And Sgn(m_SatelliteElevation(SatNumber)) <> Sgn(LastElevation) Then
        GotAOS% = 1
      End If
      LastElevation = m_SatelliteElevation(SatNumber)
      If m_DisplayMinute(SatNumber) > 59 Then
        m_DisplayMinute(SatNumber) = m_DisplayMinute(SatNumber) - 60
        m_DisplayHour(SatNumber) = m_DisplayHour(SatNumber) + 1
      End If
      If m_DisplayHour(SatNumber) > 23 Then
        m_DisplayHour(SatNumber) = m_DisplayHour(SatNumber) - 24
        m_DisplayDay(SatNumber) = m_DisplayDay(SatNumber) + 1
      End If
  End Select
Loop Until GotAOS% = 1
m_SatelliteBusy(SatNumber) = True
CalculateTrack SatNumber
DrawFootprints
End Sub
'
Sub ResetSatellite(Optional SatIndex)
  Dim FormattedDateTime As String
  Dim OldIndex As Integer

  OldIndex = m_SatelliteIndex
  If IsMissing(SatIndex) Then
    m_SatelliteBusy(m_SatelliteIndex) = False
  Else
    m_SatelliteBusy(SatIndex) = False
    m_SatelliteIndex = SatIndex
  End If

  FormattedDateTime$ = Format$(Now, "yyyymmddhhmm")
  Me.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
  Me.DisplayYear = Val(Left$(FormattedDateTime$, 4))
  Me.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
  Me.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
  Me.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
  Me.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
  CalculateSatellitePosition True
  DrawFootprints
  m_SatelliteIndex = OldIndex
End Sub
Sub SetupMoon()
  Dim FormattedDateTime As Variant

 m_SatelliteName(0) = "Moon"
 m_SatelliteDesignator(0) = ""

 m_KepsChecksum(0) = 0
 m_KepsDecayRate(0) = 0
 m_KepsElementSet(0) = 0
 m_KepsEpochTime(0) = 99269.0279492428
 m_KepsOrbitNumber(0) = 0
 m_KepsMeanMotion(0) = 0.036600996
 m_KepsMeanAnomoly(0) = 301.4974
 m_KepsInclination(0) = 20.1356
 m_KepsAOP(0) = 63.3759
 m_KepsRAAN(0) = 10.9754
 m_KepsEccentricity(0) = 0.0524
  
  If Val(Left$(Trim(Str$(m_KepsEpochTime(0))), 2)) < 50 Then
    century% = 20
  Else
    century% = 19
  End If
  mvarKepsYearEpochTime(0) = m_KepsEpochTime(0) - 1000 * Int(m_KepsEpochTime(0) / 1000)
  mvarSatEpochYear(0) = 100 * century% + Int(m_KepsEpochTime(0) / 1000)

    FormattedDateTime = Format$(Now, "yyyymmddhhmmss")
  m_DisplayCentury(0) = Val(Left$(FormattedDateTime, 2))
  m_DisplayYear(0) = Val(Left$(FormattedDateTime, 4))
  m_DisplayMonth(0) = Val(Mid$(FormattedDateTime, 5, 2))
  m_DisplayDay(0) = Val(Mid$(FormattedDateTime, 7, 2))
  m_DisplayHour(0) = Val(Mid$(FormattedDateTime, 9, 2))
  m_DisplayMinute(0) = Val(Mid$(FormattedDateTime, 11, 2))
  m_DisplaySecond(0) = 0

End Sub

Sub SetupSun()
  Dim FormattedDateTime As Variant
 
 m_SatelliteName(-1) = "Sun"
 m_SatelliteDesignator(-1) = ""

 m_KepsChecksum(-1) = 0
 m_KepsDecayRate(-1) = 0
 m_KepsElementSet(-1) = 0
 m_KepsEpochTime(-1) = 95080.092361111
 m_KepsOrbitNumber(-1) = 0
 m_KepsMeanMotion(-1) = 0.0027379093
 m_KepsMeanAnomoly(-1) = 75.2803
 m_KepsInclination(-1) = 23.44
 m_KepsAOP(-1) = 282.87
 m_KepsRAAN(-1) = 0
 m_KepsEccentricity(-1) = 0.0167
 
 MAS0 = 356.6349: MASD = 0.98560027 ' REM MA Sun and rate, deg, deg/day
  
  If Val(Left$(Trim(Str$(m_KepsEpochTime(-1))), 2)) < 50 Then
    century% = 20
  Else
    century% = 19
  End If
  mvarKepsYearEpochTime(-1) = m_KepsEpochTime(-1) - 1000 * Int(m_KepsEpochTime(-1) / 1000)
  mvarSatEpochYear(-1) = 100 * century% + Int(m_KepsEpochTime(-1) / 1000)

  
  FormattedDateTime = Format$(Now, "yyyymmddhhmmss")
  m_DisplayCentury(-1) = Val(Left$(FormattedDateTime, 2))
  m_DisplayYear(-1) = Val(Left$(FormattedDateTime, 4))
  m_DisplayMonth(-1) = Val(Mid$(FormattedDateTime, 5, 2))
  m_DisplayDay(-1) = Val(Mid$(FormattedDateTime, 7, 2))
  m_DisplayHour(-1) = Val(Mid$(FormattedDateTime, 9, 2))
  m_DisplayMinute(-1) = Val(Mid$(FormattedDateTime, 11, 2))
  m_DisplaySecond(-1) = 0

End Sub
Private Sub CalculateSunRise()
  cSunRise.Latitude = FNIntDEG(mvarObserverLatitude)
  cSunRise.Longitude = FNIntDEG(mvarObserverLongitude)
  cSunRise.DateDay = Now
  cSunRise.DaySavings = m_DaylightSaving
  cSunRise.TimeZone = -m_Timezone
  cSunRise.CalculateSun
  
  m_SunSet = cSunRise.Sunset
  m_SunNoon = cSunRise.SolarNoon
  m_SunRise = cSunRise.Sunrise

  StatusBar1.Panels(1).Text = "Sun Rise " & Format(cSunRise.Sunrise, "Short Time")
  StatusBar1.Panels(2).Text = "Noon " & Format(cSunRise.SolarNoon, "Short Time")
  StatusBar1.Panels(3).Text = "Sun Set " & Format(cSunRise.Sunset, "Short Time")
End Sub

Public Property Get DisplayTimes() As Boolean
Attribute DisplayTimes.VB_ProcData.VB_Invoke_Property = "Display"
  DisplayTimes = m_DisplayTimes
End Property

Public Property Let DisplayTimes(ByVal New_DisplayTimes As Boolean)
  m_DisplayTimes = New_DisplayTimes
  PropertyChanged "DisplayTimes"
End Property

Public Function SetMap(MapView As Integer, FileName As String) As Boolean
  If FileName = "Reset" Then
    ResetMaps
  Else
    Select Case MapView
      Case 0
        Map0.Picture = LoadPicture(FileName)
      Case 1
        Map180.Picture = LoadPicture(FileName)
      Case 2
        Map3.Picture = LoadPicture(FileName)
    End Select
    If Me.ObserverMapCentre = 0 Then
      picInner.Picture = Map0.Picture
    Else
      picInner.Picture = Map180.Picture
    End If
  End If
  
  With picInner
    HalfMapWidth = .ScaleX(.Width, vbTwips, vbPixels) / 2
    HalfMapHeight = .ScaleY(.Height, vbTwips, vbPixels) / 2
    MapWidth = .ScaleX(.Width, vbTwips, vbPixels)
    MapHeight = .ScaleY(.Height, vbTwips, vbPixels)
    PixelsPerDegLon = MapWidth / 360
    PixelsPerDegLat = (MapHeight / 360) * 2
    'PixelsPerDegLat = 2.5
  End With

  
End Function
Private Sub ResetMaps()
  Map0.Picture = DefMap0.Picture
  Map180.Picture = DefMap180.Picture
  Map3.Picture = DefMapHorizon.Picture
End Sub

Public Property Get DisplayFootprints() As Boolean
Attribute DisplayFootprints.VB_ProcData.VB_Invoke_Property = "Display"
  DisplayFootprints = m_DisplayFootprints
End Property

Public Property Let DisplayFootprints(ByVal New_DisplayFootprints As Boolean)
  m_DisplayFootprints = New_DisplayFootprints
  PropertyChanged "DisplayFootprints"
End Property

Public Function EraseSatellites()


'Erase m_SatelliteBusy
'Erase m_SatelliteTXFrequency
'Erase m_SatelliteRxFrequency
'Erase m_DataValid
'Erase m_DisplayTimeRequired
'Erase m_SatelliteDayNumber
'Erase m_DisplaySecond
'Erase m_DisplayMinute
'Erase m_DisplayHour
'Erase m_DisplayDay
'Erase m_DisplayMonth
'Erase m_DisplayYear
'Erase m_DisplayCentury
'Erase m_SatelliteOrbitNumber
'Erase m_SatelliteLongitude
'Erase m_SatelliteLatitude
'Erase m_KepsChecksum
'Erase m_KepsDecayRate
'Erase m_KepsElementSet
'Erase m_KepsEpochTime
'Erase m_KepsOrbitNumber
'Erase m_KepsMeanMotion
'Erase m_KepsMeanAnomoly
'Erase m_KepsInclination
'Erase m_KepsAOP
'Erase m_KepsRAAN
'Erase m_KepsEccentricity
'Erase m_SatelliteAzimuth
'Erase m_SatelliteRange
'Erase m_SatelliteDesignator
'Erase m_SatelliteName
'Erase m_DownLinkFrequency
'Erase m_UplinkFrequency
'
'Erase SatScreenX
'Erase SatScreenY
'Erase SatTrackLon
'Erase SatTrackLat
'Erase SatTrackElev
'Erase SatTrackAzim
'Erase SatTrackPoints
'Erase satTrackNextAOS
'Erase satTrackMaxEle
'Erase satTrackMaxEleTime
'
'Erase m_SatelliteMaximumDX
'Erase m_SatMaxEleTime
'
'Erase m_SatelliteTrackOrbits

nSatCount = 0
ResizeArrays nSatCount
SetupSun
SetupMoon

SelectedSatellite = 0
UpdateSatelliteLabel

End Function
Public Function DisplayAOSReport(AOSType As Integer, Optional SatelliteIndex As Integer) As String
End Function


Private Sub ReadDX()
    On Error GoTo Error_Handler
    
    Dim nFile As Integer
    Dim strPath As String
    Dim strData As String
    Dim strTemp As String
    Dim i As Integer
    
    If m_FrequencyDatabasePath <> "" Then
      strPath = m_FrequencyDatabasePath & "\Observer\ObserverLocations.txt"
    Else
      strPath = App.Path & "\Observer\ObserverLocations.txt"
    End If

    nFile = FreeFile
      
    Open strPath For Input As #nFile
    
    While Not EOF(1)
      Line Input #nFile, strTemp
      strData = strTemp
      sDxDetails(i).Callsign = csvParser(strData, 2)
      strData = strTemp
      sDxDetails(i).strName = csvParser(strData, 3)
      strData = strTemp
      sDxDetails(i).strLat = csvParser(strData, 4)
      strData = strTemp
      sDxDetails(i).strLon = csvParser(strData, 5)
      i = i + 1
    Wend
    Close #nFile
    Exit Sub
    
Error_Handler:
  'MsgBox "Unable to read the DX locations database" & vbCrLf & strPath, vbCritical, "Open file error"
End Sub
Public Function displayDX() As String
  Dim strLoc As String

  Dim oldHeight As Variant
  Dim nCounter As Integer
  Dim nTotal As Integer
  Dim TempTempSatDayNumber As Variant
  Dim TempTempSatTimeReq As Variant
  Dim i As Integer
  Dim strTemp As String
  
  oldHeight = mvarObserverHeight
  mvarObserverHeight = 0


  displayDX = sFormat("Callsign", 10) & " " & sFormat("Location", 40) & " " & sRFormat("Elev", 5) & " " & sRFormat("Lat", 9) & " " & sRFormat("Lon", 10) & vbCrLf

  TempDisplayYear = m_DisplayYear(m_SatelliteIndex)
  TempDisplayMonth = m_DisplayMonth(m_SatelliteIndex)
  TempDisplayDay = m_DisplayDay(m_SatelliteIndex)
  TempDisplayHour = m_DisplayHour(m_SatelliteIndex)
  TempDisplayMin = m_DisplayMinute(m_SatelliteIndex)
  TempDisplaySecond = m_DisplaySecond(m_SatelliteIndex)

  LocalDateToUTC = ConvertToGMT(TempDisplayYear, TempDisplayMonth, TempDisplayDay, TempDisplayHour, TempDisplayMin, TempDisplaySecond)
  
  TempDisplayYear = Year(LocalDateToUTC)
  TempDisplayMonth = Month(LocalDateToUTC)
  TempDisplayDay = Day(LocalDateToUTC)
  TempDisplayHour = hour(LocalDateToUTC)
  TempDisplayMin = Minute(LocalDateToUTC)
  TempDisplaySecond = Second(LocalDateToUTC)

  TempTempSatDayNum = FNDAy(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
  TempTempSatTimeReq = (TempDisplayHour + (TempDisplayMin / 60) + (TempDisplaySecond / 60 / 60)) / 24

  For i = 0 To UBound(sDxDetails)
    If sDxDetails(i).Callsign = "" And sDxDetails(i).strName = "" Then Exit For
      TempObsLat = FNRAD(sDxDetails(i).strLat)
      TempObsLon = FNRAD(-sDxDetails(i).strLon)

      TempSatDayNum = TempTempSatDayNum
      TempSatTimeReq = TempTempSatTimeReq

      PositionEngine m_SatelliteIndex, True, True

      If TempSatElev > m_SetAOSLOS Then
        strTemp = ConvertDXPos(FNDEG(TempObsLat), FNDEG(TempObsLon))
        displayDX = displayDX & sFormat(sDxDetails(i).Callsign, 10) & " " & sFormat(sDxDetails(i).strName, 40) & " " & sRFormat(Str(TempSatElev), 5) & " " & strTemp & vbCrLf
        nTotal = nTotal + 1
      End If
      nCounter = nCounter + 1
    Next i
  
  If nTotal = 0 Then
    displayDX = displayDX & vbCrLf & "There is NO DX visible"
  Else
    displayDX = displayDX & vbCrLf & Str(nTotal) & " DX locations visible out of " & Str(nCounter) & " in the database"
  End If
  mvarObserverHeight = oldHeight

End Function

Private Function ConvertDXPos(sLat As Single, sLon As Single) As String
  Dim strTemp As String
  
  If sLat < 0 Then
    strTemp = sRFormat(Format(Str(Abs(sLat)), "##0.00") & "°S  ", 11)
  Else
    strTemp = sRFormat(Format(Str(sLat), "##0.00") & "°N  ", 11)
  End If
  If sLon < 0 Then
    strTemp = strTemp & sRFormat(Format(Str(Abs(sLon)), "##0.00") & "°W  ", 11)
  Else
    strTemp = strTemp & sRFormat(Format(Str(sLon), "##0.00") & "°E  ", 11)
  End If
  ConvertDXPos = strTemp
End Function

Private Function UpdateAOSPos(SatNumber As Integer, nInc As Long, bSetupFlag As Boolean) As Boolean
    If nInc > 0 Then
      AddTime nInc * 60
    Else
      SubTime nInc * 60
    End If
        
    TempSatDayNum = FNDAy(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
    TempSatTimeReq = (TempDisplayHour + TempDisplayMin / 60) / 24
    
    TempObsLat = mvarObserverLatitude
    TempObsLon = mvarObserverLongitude
    
    PositionEngine SatNumber, bSetupFlag, False

    If TempSatElev > m_SetAOSLOS Then
      UpdateAOSPos = True
    Else
      UpdateAOSPos = False
    End If
    
'Open "C:\aos.txt" For Append As #1
'Print #1, Format(TempDisplayHour, "@@") & ":" & Format(TempDisplayMin, "@@") & "  " & TempSatElev
'Close #1
End Function

Private Function TempDate() As String
  TempDate = Format(TempDisplayDay, "00") & "/" & Format(TempDisplayMonth, "00") & "/" & Format(TempDisplayYear, "0000")
End Function
Private Function TempTime() As String
  Dim nTemp As Integer
  
  If nDaylightSaving <> 0 Then
    nTemp = TempDisplayHour + 1
    If nTemp > 23 Then nTemp = 0
    TempTime = Format(DateAdd("h", 1, nTemp), "00") & ":" & Format(TempDisplayMin, "00")
  Else
    TempTime = Format(TempDisplayHour, "00") & ":" & Format(TempDisplayMin, "00")
  End If
End Function

Private Function sFormat(strString As String, nLen As Integer) As String
  sFormat = Left$(strString & String(nLen, " "), nLen)
End Function
Private Function sRFormat(strString As String, nLen As Integer) As String
  sRFormat = Right$(String(nLen, " ") & strString, nLen)
End Function

Public Property Get SelectedSatelliteName() As String
Attribute SelectedSatelliteName.VB_MemberFlags = "400"
  SelectedSatelliteName = m_SatelliteName(SelectedSatellite)
End Property

Public Property Let SelectedSatelliteName(ByVal New_SelectedSatelliteName As String)
  If Ambient.UserMode = False Then Err.Raise 382
  If Ambient.UserMode Then Err.Raise 393
  m_SelectedSatelliteName = New_SelectedSatelliteName
  PropertyChanged "SelectedSatelliteName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SatelliteMaximumDX() As Variant
Attribute SatelliteMaximumDX.VB_MemberFlags = "400"
  SatelliteMaximumDX = m_SatelliteMaximumDX(m_SatelliteIndex)
End Property

Public Property Let SatelliteMaximumDX(ByVal New_SatelliteMaximumDX As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteMaximumDX(m_SatelliteIndex) = New_SatelliteMaximumDX
  PropertyChanged "SatelliteMaximumDX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get IsSatLoaded(ByVal strDesignator As Variant) As Variant
Attribute IsSatLoaded.VB_MemberFlags = "400"
  Dim i As Integer
  
  IsSatLoaded = False
  
  For i = 1 To nSatCount
    If m_SatelliteDesignator(i) = strDesignator Then
      IsSatLoaded = True
      Exit For
    End If
  Next i
End Property

Public Property Let IsSatLoaded(ByVal strDesignator As Variant, ByVal New_IsSatLoaded As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_IsSatLoaded = New_IsSatLoaded
  PropertyChanged "IsSatLoaded"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SetSatelliteTime() As Variant
Attribute SetSatelliteTime.VB_Description = "Set the time and date for the satellite."
  
  Dim vlDate As Variant
  Dim vlTime As Variant
  Dim vlDateTime As Variant
  
  Dim vgmtDate As Variant
  Dim vgmtTime As Variant
  Dim vgmtDateTime As Variant
  
  vlDate = DateSerial(Me.DisplayYear, Me.DisplayMonth, Me.DisplayDay)
  
  vlTime = TimeSerial(Me.DisplayHour, Me.DisplayMinute, Me.DisplaySecond)

  vlDateTime = vlDate + vlTime
  
  SetSatelliteTime = vlDateTime
End Property

Public Property Let SetSatelliteTime(ByVal New_SetSatelliteTime As Variant)
  Dim FormattedDateTime As Variant
  
    FormattedDateTime = Format$(New_SetSatelliteTime, "yyyymmddhhmmss")
    Me.DisplayCentury = Val(Left$(FormattedDateTime, 2))
    Me.DisplayYear = Val(Left$(FormattedDateTime, 4))
    Me.DisplayMonth = Val(Mid$(FormattedDateTime, 5, 2))
    Me.DisplayDay = Val(Mid$(FormattedDateTime, 7, 2))
    Me.DisplayHour = Val(Mid$(FormattedDateTime, 9, 2))
    Me.DisplayMinute = Val(Mid$(FormattedDateTime, 11, 2))
    Me.DisplaySecond = Val(Mid$(FormattedDateTime, 13, 2))
  PropertyChanged "SetSatelliteTime"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddTimeToSatellite(lSeconds As Long) As Variant
  Dim vDateTime As Variant
  
  vDateTime = Me.SetSatelliteTime
  
  vDateTime = DateAdd("s", lSeconds, vDateTime)

  Me.SetSatelliteTime = vDateTime
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,0
Public Property Get CurrentSelectedSatellite() As Variant
Attribute CurrentSelectedSatellite.VB_MemberFlags = "400"
  CurrentSelectedSatellite = SelectedSatellite

End Property

Public Property Let CurrentSelectedSatellite(ByVal New_CurrentSelectedSatellite As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If New_SelectedSatellite <= Me.SatelliteCount Then
    SelectedSatellite = New_SelectedSatellite
    PropertyChanged "CurrentSelectedSatellite"
    UpdateSatelliteLabel
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SetAOSLOS() As Variant
Attribute SetAOSLOS.VB_Description = "Sets the value to determine AOS and LOS of the satellite"
Attribute SetAOSLOS.VB_ProcData.VB_Invoke_Property = "Observer"
  SetAOSLOS = m_SetAOSLOS
End Property

Public Property Let SetAOSLOS(ByVal New_SetAOSLOS As Variant)
  m_SetAOSLOS = New_SetAOSLOS
  PropertyChanged "SetAOSLOS"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DatabasePath() As Variant
Attribute DatabasePath.VB_Description = "Stores the pathname and database name of the database used to set the keplarian elements."
Attribute DatabasePath.VB_ProcData.VB_Invoke_Property = "General"
  DatabasePath = m_DatabasePath
End Property

Public Property Let DatabasePath(ByVal New_DatabasePath As Variant)
  m_DatabasePath = New_DatabasePath
  PropertyChanged "DatabasePath"
  ReadDX
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,0
Public Property Get Busy() As Boolean
Attribute Busy.VB_MemberFlags = "400"
  Busy = m_Busy
End Property

Public Property Let Busy(ByVal New_Busy As Boolean)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_Busy = New_Busy
  PropertyChanged "Busy"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SquintAngle() As Variant
Attribute SquintAngle.VB_MemberFlags = "400"
  SquintAngle = m_SquintAngle(m_SatelliteIndex)
End Property

Public Property Let SquintAngle(ByVal New_SquintAngle As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SquintAngle(m_SatelliteIndex) = New_SquintAngle
  PropertyChanged "SquintAngle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get RangeRate() As Variant
Attribute RangeRate.VB_MemberFlags = "400"
  RangeRate = m_RangeRate(m_SatelliteIndex)
End Property

Public Property Let RangeRate(ByVal New_RangeRate As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_RangeRate(m_SatelliteIndex) = New_RangeRate
  PropertyChanged "RangeRate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SatelliteStatusText() As Variant
Attribute SatelliteStatusText.VB_MemberFlags = "400"
  SatelliteStatusText = m_SatelliteStatusText(m_SatelliteIndex)
End Property

Public Property Let SatelliteStatusText(ByVal New_SatelliteStatusText As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteStatusText(m_SatelliteIndex) = New_SatelliteStatusText
  PropertyChanged "SatelliteStatusText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enable847() As Boolean
Attribute Enable847.VB_Description = "Determines if the FT847 control is enabled"
Attribute Enable847.VB_ProcData.VB_Invoke_Property = "FT847"
  Enable847 = m_Enable847
End Property

Public Property Let Enable847(ByVal New_Enable847 As Boolean)
  m_Enable847 = New_Enable847
  PropertyChanged "Enable847"
  
  If m_Enable847 Then
    frmRadio.Show
  Else
    Unload frmRadio
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,57600,N,8,1
Public Property Get PortSettings() As Variant
Attribute PortSettings.VB_Description = "Communication port settings for FT847 control. Menu 37 on FT847."
Attribute PortSettings.VB_ProcData.VB_Invoke_Property = "FT847"
  PortSettings = m_PortSettings
End Property

Public Property Let PortSettings(ByVal New_PortSettings As Variant)
  m_PortSettings = New_PortSettings
  PropertyChanged "PortSettings"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get Enable847Sat() As Boolean
Attribute Enable847Sat.VB_ProcData.VB_Invoke_Property = "FT847"
  Enable847Sat = m_Enable847Sat
End Property

Public Property Let Enable847Sat(ByVal New_Enable847Sat As Boolean)
  If Ambient.UserMode = False Then Err.Raise 387
  m_Enable847Sat = New_Enable847Sat
  PropertyChanged "Enable847Sat"

  frmRadio.SetSatelliteMode m_Enable847Sat
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TimeZone() As Variant
Attribute TimeZone.VB_ProcData.VB_Invoke_Property = "Observer"
  TimeZone = m_Timezone
End Property

Public Property Let TimeZone(ByVal New_Timezone As Variant)
  m_Timezone = New_Timezone
  PropertyChanged "Timezone"
  CalculateSunRise
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DaylightSaving() As Boolean
Attribute DaylightSaving.VB_Description = "True for DaylightSaving, False for no DaylightSaving."
Attribute DaylightSaving.VB_ProcData.VB_Invoke_Property = "Observer"
  DaylightSaving = m_DaylightSaving
End Property

Public Property Let DaylightSaving(ByVal New_DaylightSaving As Boolean)
  m_DaylightSaving = New_DaylightSaving
  PropertyChanged "DaylightSaving"
  CalculateSunRise
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get UplinkFrequency() As Variant
Attribute UplinkFrequency.VB_Description = "The satellites uplink frequency"
  UplinkFrequency = m_UplinkFrequency(m_SatelliteIndex)
End Property

Public Property Let UplinkFrequency(ByVal New_UplinkFrequency As Variant)
  m_UplinkFrequency(m_SatelliteIndex) = New_UplinkFrequency
  PropertyChanged "UplinkFrequency"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DownLinkFrequency() As Variant
Attribute DownLinkFrequency.VB_Description = "The Satellites downlink frequency"
  DownLinkFrequency = m_DownLinkFrequency(m_SatelliteIndex)
End Property

Public Property Let DownLinkFrequency(ByVal New_DownLinkFrequency As Variant)
  m_DownLinkFrequency(m_SatelliteIndex) = New_DownLinkFrequency
  PropertyChanged "DownLinkFrequency"
End Property
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get AutoMode() As Variant
'  AutoMode = m_AutoMode
'End Property
'
'Public Property Let AutoMode(ByVal New_AutoMode As Variant)
'  m_AutoMode = New_AutoMode
'  PropertyChanged "AutoMode"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get AutoInterval() As Variant
Attribute AutoInterval.VB_Description = "If in AutoMode then this is the update interval in Seconds 1 - 60."
Attribute AutoInterval.VB_ProcData.VB_Invoke_Property = "General"
  AutoInterval = m_AutoInterval
End Property

Public Property Let AutoInterval(ByVal New_AutoInterval As Variant)
  If New_AutoInterval > -1 And New_AutoInterval < 61 Then
    m_AutoInterval = New_AutoInterval
    PropertyChanged "AutoInterval"
    If m_AutoInterval Then
      UserControl.tmrAutoUpdate.Interval = m_AutoInterval * 1000
      UserControl.tmrAutoUpdate.Enabled = True
    Else
      'UserControl.tmrAutoUpdate.Interval = 0
      UserControl.tmrAutoUpdate.Enabled = False
    End If
  Else
    Err.Raise 380
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SatelliteNextAOS() As Variant
Attribute SatelliteNextAOS.VB_Description = "Returns the next AOS for the selected satellite. Tracks MUST be calculated to enable this."
Attribute SatelliteNextAOS.VB_MemberFlags = "400"
  SatelliteNextAOS = satTrackNextAOS(m_SatelliteIndex)
End Property

Public Property Let SatelliteNextAOS(ByVal New_SatelliteNextAOS As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  'm_SatelliteNextAOS = New_SatelliteNextAOS
  PropertyChanged "SatelliteNextAOS"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SatelliteNextAOSMaxElevation() As Variant
Attribute SatelliteNextAOSMaxElevation.VB_Description = "Max elevation on next Satellite Pass."
Attribute SatelliteNextAOSMaxElevation.VB_MemberFlags = "400"
  SatelliteNextAOSMaxElevation = satTrackMaxEle(m_SatelliteIndex)
End Property

Public Property Let SatelliteNextAOSMaxElevation(ByVal New_SatelliteNextAOSMaxElevation As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
'  m_SatelliteNextAOSMaxElevation = New_SatelliteNextAOSMaxElevation
  PropertyChanged "SatelliteNextAOSMaxElevation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SatelliteNextAOSMaxElevationTime() As Variant
Attribute SatelliteNextAOSMaxElevationTime.VB_Description = "Time of max elevation on next satellite pass."
Attribute SatelliteNextAOSMaxElevationTime.VB_MemberFlags = "400"
  SatelliteNextAOSMaxElevationTime = m_SatMaxEleTime(m_SatelliteIndex)
End Property

Public Property Let SatelliteNextAOSMaxElevationTime(ByVal New_SatelliteNextAOSMaxElevationTime As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
 ' m_SatelliteNextAOSMaxElevationTime = New_SatelliteNextAOSMaxElevationTime
  PropertyChanged "SatelliteNextAOSMaxElevationTime"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SatelliteTrackOrbits() As Variant
Attribute SatelliteTrackOrbits.VB_Description = "NUmber of orbits to plot the ground track for."
  SatelliteTrackOrbits = m_SatelliteTrackOrbits(m_SatelliteIndex)
End Property

Public Property Let SatelliteTrackOrbits(ByVal New_SatelliteTrackOrbits As Variant)
  m_SatelliteTrackOrbits(m_SatelliteIndex) = New_SatelliteTrackOrbits
  PropertyChanged "SatelliteTrackOrbits"
End Property

Public Property Get SatelliteMA() As Double
  SatelliteMA = m_SatelliteMA(m_SatelliteIndex)
End Property

Public Property Let SatelliteMA(ByVal New_SatelliteMA As Double)
  PropertyChanged "SatelliteMA"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoMode() As Boolean
  AutoMode = m_AutoMode
End Property

Public Property Let AutoMode(ByVal New_AutoMode As Boolean)
  m_AutoMode = New_AutoMode
  PropertyChanged "AutoMode"
End Property

Public Sub UpdateDataWindow()
'  Dim strTemp As String
'  Dim i As Integer
'  Dim nSat As Integer
'
'  'If SatDetails.Tag <> "" Then
'If bSatDetailsVisible Then
'    nSat = nSatDetailsTag
'    For i = 1 To 100
'      If nFields(i) <> 0 Then
'        With SatDetails
'          Select Case nFields(i)
'            Case 1
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteDesignator(nSat)
'            Case 2
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteName(nSat)
'            Case 3
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteAzimuth(nSat)
'            Case 4
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteElevation(nSat)
'            Case 5
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteLatitude(nSat)
'            Case 6
'              .lstDetails.ListItems(i).SubItems(1) = 360 - m_SatelliteLongitude(nSat)
'            Case 7
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteRange(nSat)
'            Case 8
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteOrbitNumber(nSat)
'            Case 9
'              .lstDetails.ListItems(i).SubItems(1) = m_RangeRate(nSat)
'            Case 10
'              If m_UplinkFrequency(nSat) <> 0 Then
'                .lstDetails.ListItems(i).SubItems(1) = m_UplinkFrequency(nSat)
'              Else
'                .lstDetails.ListItems(i).SubItems(1) = "N/A"
'              End If
'            Case 11
'              If m_DownLinkFrequency(nSat) <> 0 Then
'                .lstDetails.ListItems(i).SubItems(1) = m_DownLinkFrequency(nSat)
'              Else
'                .lstDetails.ListItems(i).SubItems(1) = "N/A"
'              End If
'            Case 12
'              If m_SatelliteRxFrequency(nSat) <> 0 Then
'                .lstDetails.ListItems(i).SubItems(1) = m_SatelliteRxFrequency(nSat)
'              Else
'                .lstDetails.ListItems(i).SubItems(1) = "N/A"
'              End If
'            Case 13
'              If m_SatelliteTXFrequency(nSat) <> 0 Then
'                .lstDetails.ListItems(i).SubItems(1) = m_SatelliteTXFrequency(nSat)
'              Else
'                .lstDetails.ListItems(i).SubItems(1) = "N/A"
'              End If
'            Case 14
'              If m_SatelliteRxFrequency(nSat) <> 0 And m_UplinkFrequency(nSat) <> 0 Then
'                .lstDetails.ListItems(i).SubItems(1) = Format(m_SatelliteRxFrequency(nSat) - m_UplinkFrequency(nSat), "###.######") * 1000
'              End If
'            Case 15
'              .lstDetails.ListItems(i).SubItems(1) = Format(m_SatellitePathLoss(nSat), "##.##")
'            Case 16
'              .lstDetails.ListItems(i).SubItems(1) = Format(m_SatelliteMaximumDX(nSat), "######")
'            Case 17
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteStatusText(nSat)
'            Case 18
'              .lstDetails.ListItems(i).SubItems(1) = RS(nSat)
'            Case 19
'              .lstDetails.ListItems(i).SubItems(1) = m_SquintAngle(nSat)
'            Case 20
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsDecayRate(nSat)
'            Case 21
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsElementSet(nSat)
'            Case 22
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsEpochTime(nSat)
'            Case 23
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsOrbitNumber(nSat)
'            Case 24
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsMeanMotion(nSat)
'            Case 25
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsMeanAnomoly(nSat)
'            Case 26
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsInclination(nSat)
'            Case 27
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsAOP(nSat)
'            Case 28
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsRAAN(nSat)
'            Case 29
'              .lstDetails.ListItems(i).SubItems(1) = m_KepsEccentricity(nSat)
'            Case 30
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteAltitude(nSat)
'            Case 31
'              .lstDetails.ListItems(i).SubItems(1) = m_SatelliteMA(nSat)
'            Case 32
'              .lstDetails.ListItems(i).SubItems(1) = Format(satTrackNextAOS(nSat), "General Date")
'            Case 33
'              .lstDetails.ListItems(i).SubItems(1) = Format(satTrackNextAOS(nSat), "HH:nn:ss")
'          End Select
'        End With
'      Else
'        Exit For
'      End If
'    Next i
'
'  End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get DisplayDataFields() As String
Attribute DisplayDataFields.VB_Description = "Sets the fields to display in the satellite data window."
Attribute DisplayDataFields.VB_ProcData.VB_Invoke_Property = "Display"
  Dim strTemp As String
  
  strTemp = BuildParseStr(nFields)
  If InStr(strTemp, ",0") <> 0 Then
    strTemp = Left$(strTemp, InStr(strTemp, ",0") + 1)
  End If
  If Left(strTemp, 1) = "," Then
    strTemp = Mid(strTemp, 2)
  End If
  DisplayDataFields = strTemp
End Property

Public Property Let DisplayDataFields(ByVal New_DisplayDataFields As String)
  Dim vData() As Variant
  Dim i As Integer
  
  m_DisplayDataFields = New_DisplayDataFields & ",-1"
  
  vData = StrParse(m_DisplayDataFields, ",")
  Erase nFields
  For i = 1 To UBound(vData) + 1
    nFields(i) = vData(i - 1)
  Next i
 ' SatDetails.Rebuild_List
  PropertyChanged "DisplayDataFields"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function SetupKeps(strLine1 As String, strLine2 As String, strLine3 As String, nIndex As Integer) As Boolean
Attribute SetupKeps.VB_Description = "Sets the keps from three line elements in one call."
  Dim century As Integer
  
  SetupKeps = False
  
  If nIndex > 0 And nIndex < 21 Then
  If CheckKepsChecksum(strLine2) And CheckKepsChecksum(strLine3) Then
      m_SatelliteName(nIndex) = Trim(strLine1)
      m_SatelliteDesignator(nIndex) = Val(Mid$(strLine2, 3, 5))
      m_KepsEpochTime(nIndex) = Mid$(strLine2, 19, 14)
      m_KepsDecayRate(nIndex) = Val(Mid$(strLine2, 35, 9))
      m_KepsOrbitNumber(nIndex) = Val(Mid$(strLine3, 64, 5))
      m_KepsInclination(nIndex) = Val(Mid$(strLine3, 9, 8))
      m_KepsRAAN(nIndex) = Val(Mid$(strLine3, 18, 8))
      m_KepsEccentricity(nIndex) = Val("0." + Mid$(strLine3, 27, 7))
      m_KepsAOP(nIndex) = Val(Mid$(strLine3, 35, 8))
      m_KepsMeanAnomoly(nIndex) = Val(Mid$(strLine3, 44, 8))
      m_KepsMeanMotion(nIndex) = Val(Mid$(strLine3, 53, 11))
      SetupKeps = True
      If Val(Left$(m_KepsEpochTime(nIndex), 2)) < 50 Then
        century = 20
      Else
        century = 19
      End If
      mvarKepsYearEpochTime(nIndex) = m_KepsEpochTime(nIndex) - 1000 * Int(m_KepsEpochTime(nIndex) / 1000)
      mvarSatEpochYear(nIndex) = 100 * century% + Int(m_KepsEpochTime(nIndex) / 1000)
    End If
  End If
End Function

Private Function CheckKepsChecksum(strLine As String) As Boolean
Dim nVal As Integer
Dim i As Integer

  nVal = 0
  If Len(strLine) > 69 Then strLine = Left$(strLine, 69)
  For i = 1 To Len(strLine) - 1
    strChar = Mid$(strLine, i, 1)
    nVal = nVal + Val(strChar)
    If strChar = "-" Then nVal = nVal + 1
    If strChar = "," Then Mid$(strLine, i, 1) = "."
  Next i
  
  If nVal Mod 10 = Right(strLine, 1) Then
    CheckKepsChecksum = True
  Else
    CheckKepsChecksum = False
  End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get FrequencyDatabasePath() As String
Attribute FrequencyDatabasePath.VB_Description = "Path to the Frequencies.txt database. The trailing \\  is not required."
Attribute FrequencyDatabasePath.VB_ProcData.VB_Invoke_Property = "General"
  FrequencyDatabasePath = m_FrequencyDatabasePath
End Property

Public Property Let FrequencyDatabasePath(ByVal New_FrequencyDatabasePath As String)
  m_FrequencyDatabasePath = New_FrequencyDatabasePath
  PropertyChanged "FrequencyDatabasePath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get SetIndexOnSelect() As Boolean
Attribute SetIndexOnSelect.VB_Description = "If True the SatelliteIndex property is updated when a satellite is selected."
Attribute SetIndexOnSelect.VB_ProcData.VB_Invoke_Property = "General"
  SetIndexOnSelect = m_SetIndexOnSelect
End Property

Public Property Let SetIndexOnSelect(ByVal New_SetIndexOnSelect As Boolean)
  m_SetIndexOnSelect = New_SetIndexOnSelect
  PropertyChanged "SetIndexOnSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SetSelectedSatellite() As Integer
Attribute SetSelectedSatellite.VB_Description = "Sets the selected satellite."
Attribute SetSelectedSatellite.VB_ProcData.VB_Invoke_Property = "General"
  SetSelectedSatellite = SelectedSatellite
End Property

Public Property Let SetSelectedSatellite(ByVal New_SetSelectedSatellite As Integer)
  If New_SetSelectedSatellite <= SatelliteCount Then
    SelectedSatellite = New_SetSelectedSatellite
    PropertyChanged "SetSelectedSatellite"
    UpdateDataWindow
    UpdateSatelliteLabel
    RaiseEvent SatelliteSelected(SelectedSatellite)
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,-1
Public Property Get AllowDoEvents() As Boolean
Attribute AllowDoEvents.VB_Description = "Allows control ro be released to other Windows applications. Please Read help file before using this property."
Attribute AllowDoEvents.VB_ProcData.VB_Invoke_Property = "General"
  AllowDoEvents = m_AllowDoEvents
End Property

Public Property Let AllowDoEvents(ByVal New_AllowDoEvents As Boolean)
  m_AllowDoEvents = New_AllowDoEvents
  PropertyChanged "AllowDoEvents"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,0
Public Property Get SatelliteInAOS() As Boolean
Attribute SatelliteInAOS.VB_Description = "Indicates if the satellite is above the horizon."
Attribute SatelliteInAOS.VB_MemberFlags = "400"
  SatelliteInAOS = m_SatelliteInAOS(m_SatelliteIndex)
End Property

Public Property Let SatelliteInAOS(ByVal New_SatelliteInAOS As Boolean)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  'm_SatelliteInAOS = New_SatelliteInAOS
  PropertyChanged "SatelliteInAOS"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseHourglass() As Boolean
Attribute UseHourglass.VB_Description = "If true then the Hourglass cursor will be displayed while calculations are in progress"
Attribute UseHourglass.VB_ProcData.VB_Invoke_Property = "General"
  UseHourglass = m_UseHourglass
End Property

Public Property Let UseHourglass(ByVal New_UseHourglass As Boolean)
  m_UseHourglass = New_UseHourglass
  PropertyChanged "UseHourglass"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get TimeZoneName() As String
Attribute TimeZoneName.VB_Description = "Sets the name of the current timezone"
Attribute TimeZoneName.VB_ProcData.VB_Invoke_Property = "Observer"
  TimeZoneName = m_TimeZoneName
End Property

Public Property Let TimeZoneName(ByVal New_TimeZoneName As String)
  m_TimeZoneName = New_TimeZoneName
  PropertyChanged "TimeZoneName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DaylightSavingAdjust() As Integer
Attribute DaylightSavingAdjust.VB_Description = "Number of hours to adjust for daylight saving"
Attribute DaylightSavingAdjust.VB_ProcData.VB_Invoke_Property = "Observer"
  DaylightSavingAdjust = m_DaylightSavingAdjust
End Property

Public Property Let DaylightSavingAdjust(ByVal New_DaylightSavingAdjust As Integer)
  m_DaylightSavingAdjust = New_DaylightSavingAdjust
  PropertyChanged "DaylightSavingAdjust"
  nDaylightSaving = m_DaylightSaving
  CalculateSunRise
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,2,0
Public Property Get SatelliteSemiMajorAxis() As Double
Attribute SatelliteSemiMajorAxis.VB_Description = "Semi major axis of orbit"
Attribute SatelliteSemiMajorAxis.VB_MemberFlags = "400"
  SatelliteSemiMajorAxis = m_SatelliteSemiMajorAxis(m_SatelliteIndex)
End Property

Public Property Let SatelliteSemiMajorAxis(ByVal New_SatelliteSemiMajorAxis As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteSemiMajorAxis(m_SatelliteIndex) = New_SatelliteSemiMajorAxis
  PropertyChanged "SatelliteSemiMajorAxis"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,2,0
Public Property Get SatelliteSemiMinorAxis() As Double
Attribute SatelliteSemiMinorAxis.VB_Description = "Semi minor axis of orbit"
Attribute SatelliteSemiMinorAxis.VB_MemberFlags = "400"
  SatelliteSemiMinorAxis = m_SatelliteSemiMinorAxis(m_SatelliteIndex)
End Property

Public Property Let SatelliteSemiMinorAxis(ByVal New_SatelliteSemiMinorAxis As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteSemiMinorAxis(m_SatelliteIndex) = New_SatelliteSemiMinorAxis
  PropertyChanged "SatelliteSemiMinorAxis"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,2,0
Public Property Get SatelliteLonOfNode() As Double
Attribute SatelliteLonOfNode.VB_Description = "Longitude of ascending node"
Attribute SatelliteLonOfNode.VB_MemberFlags = "400"
  SatelliteLonOfNode = m_SatelliteLonOfNode(m_SatelliteIndex)
End Property

Public Property Let SatelliteLonOfNode(ByVal New_SatelliteLonOfNode As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteLonOfNode(m_SatelliteIndex) = New_SatelliteLonOfNode
  PropertyChanged "SatelliteLonOfNode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,2,0
Public Property Get SatelliteAltAtPerigee() As Double
Attribute SatelliteAltAtPerigee.VB_Description = "Altitude of satellite at Perigee"
Attribute SatelliteAltAtPerigee.VB_MemberFlags = "400"
  SatelliteAltAtPerigee = m_SatelliteAltAtPerigee(m_SatelliteIndex)
End Property

Public Property Let SatelliteAltAtPerigee(ByVal New_SatelliteAltAtPerigee As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteAltAtPerigee(m_SatelliteIndex) = New_SatelliteAltAtPerigee
  PropertyChanged "SatelliteAltAtPerigee"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,2,0
Public Property Get SatelliteAltAtApogee() As Double
Attribute SatelliteAltAtApogee.VB_Description = "Altitude of satellite at Apogee"
Attribute SatelliteAltAtApogee.VB_MemberFlags = "400"
  SatelliteAltAtApogee = m_SatelliteAltAtApogee(m_SatelliteIndex)
End Property

Public Property Let SatelliteAltAtApogee(ByVal New_SatelliteAltAtApogee As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteAltAtApogee(m_SatelliteIndex) = New_SatelliteAltAtApogee
  PropertyChanged "SatelliteAltAtApogee"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function UpdateDerived(Optional SatelliteIndex As Variant) As Boolean
Attribute UpdateDerived.VB_Description = "Update derived values from Keps"
  Dim SatNumber As Integer
  
  If IsMissing(SatelliteIndex) Then
    SatNumber = m_SatelliteIndex
  Else
    SatNumber = SatelliteIndex
  End If

  
  m_SatelliteSemiMajorAxis(SatNumber) = (gc0528 / (m_KepsMeanMotion(SatNumber) * m_KepsMeanMotion(SatNumber))) ^ (1 / 3)
  m_SatelliteSemiMinorAxis(SatNumber) = (m_SatelliteSemiMajorAxis(SatNumber) * Sqr(1 - m_KepsEccentricity(SatNumber) * m_KepsEccentricity(SatNumber)))

  m_SatelliteAltAtPerigee(SatNumber) = m_SatelliteSemiMajorAxis(SatNumber) * (1 - m_KepsEccentricity(SatNumber)) - RE
  m_SatelliteAltAtApogee(SatNumber) = m_SatelliteSemiMajorAxis(SatNumber) * (1 + m_KepsEccentricity(SatNumber)) - RE

  m_SatellitePeriod(SatNumber) = 1440 / m_KepsMeanMotion(SatNumber)

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,1,2,0
Public Property Get SatellitePeriod() As Double
Attribute SatellitePeriod.VB_Description = "Orbital period of the satellite"
Attribute SatellitePeriod.VB_MemberFlags = "400"
  SatellitePeriod = m_SatellitePeriod(m_SatelliteIndex)
End Property

Public Property Let SatellitePeriod(ByVal New_SatellitePeriod As Double)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatellitePeriod(m_SatelliteIndex) = New_SatellitePeriod
  PropertyChanged "SatellitePeriod"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ViewsOrthShade() As Boolean
Attribute ViewsOrthShade.VB_Description = "Indicated whether Globe view is shaded for day and night."
Attribute ViewsOrthShade.VB_ProcData.VB_Invoke_Property = "Display"
  ViewsOrthShade = m_ViewsOrthShade
End Property

Public Property Let ViewsOrthShade(ByVal New_ViewsOrthShade As Boolean)
  m_ViewsOrthShade = New_ViewsOrthShade
  PropertyChanged "ViewsOrthShade"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ViewsOrthLocations() As String
Attribute ViewsOrthLocations.VB_Description = "Sets the locations file for displaying on the globe."
Attribute ViewsOrthLocations.VB_ProcData.VB_Invoke_Property = "Display"
  ViewsOrthLocations = m_ViewsOrthLocations
End Property

Public Property Let ViewsOrthLocations(ByVal New_ViewsOrthLocations As String)
  If Left(New_ViewsOrthLocations, 1) <> "\" Then
    New_ViewsOrthLocations = New_ViewsOrthLocations
  End If
  m_ViewsOrthLocations = New_ViewsOrthLocations
  PropertyChanged "ViewsOrthLocations"
End Property

  
Private Function ConvertToGMT(TempDisplayYear, TempDisplayMonth, TempDisplayDay, TempDisplayHour, TempDisplayMin, TempDisplaySecond) As Variant
  Dim vlDate As Variant
  Dim vlTime As Variant
  Dim vlDateTime As Variant
  
  Dim vgmtDate As Variant
  Dim vgmtTime As Variant
  Dim vgmtDateTime As Variant
  
  vlDate = DateSerial(TempDisplayYear, TempDisplayMonth, TempDisplayDay)
  
  vlTime = TimeSerial(TempDisplayHour, TempDisplayMin, TempDisplaySecond)

  vlDateTime = vlDate + vlTime

  vgmtDateTime = DateAdd("h", m_Timezone + m_DaylightSavingAdjust, vlDateTime)
  
  ConvertToGMT = vgmtDateTime
  
  UserControl.StatusBar1.Panels(4).Text = Format(vlDateTime, "General Date")
  UserControl.StatusBar1.Panels(5).Text = Format(ConvertToGMT, "General Date")
End Function

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552

  frmAboutOcx.Show vbModal
  Unload frmAboutOcx
  Set frmAboutOcx = Nothing
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,2,0
Public Property Get MaxWidth() As Integer
Attribute MaxWidth.VB_Description = "The maximum width the control can be."
Attribute MaxWidth.VB_MemberFlags = "400"
  MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Integer)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_MaxWidth = New_MaxWidth
  PropertyChanged "MaxWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,2,0
Public Property Get MaxHeight() As Integer
Attribute MaxHeight.VB_Description = "The maximum height the control can be."
Attribute MaxHeight.VB_MemberFlags = "400"
  MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Integer)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_MaxHeight = New_MaxHeight
  PropertyChanged "MaxHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,1,57600,n,8,1
Public Property Get FT847CATSettings() As String
Attribute FT847CATSettings.VB_Description = "Port settings for 847 port,speed,parity,bits,stop"
Attribute FT847CATSettings.VB_ProcData.VB_Invoke_Property = "FT847"
  FT847CATSettings = m_FT847CATSettings
End Property

Public Property Let FT847CATSettings(ByVal New_FT847CATSettings As String)
  m_FT847CATSettings = New_FT847CATSettings
  PropertyChanged "FT847CATSettings"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get SatPosLabelAlign() As Integer
Attribute SatPosLabelAlign.VB_Description = "Sets the alignment of the satellite label."
Attribute SatPosLabelAlign.VB_ProcData.VB_Invoke_Property = "Display"
  SatPosLabelAlign = m_SatPosLabelAlign
End Property

Public Property Let SatPosLabelAlign(ByVal New_SatPosLabelAlign As Integer)
  m_SatPosLabelAlign = New_SatPosLabelAlign
  PropertyChanged "SatPosLabelAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get SetActiveWindowAsWallpaper() As Boolean
Attribute SetActiveWindowAsWallpaper.VB_Description = "Setsthe current window as the Desktop wallpaper"
Attribute SetActiveWindowAsWallpaper.VB_ProcData.VB_Invoke_Property = "Display"
  SetActiveWindowAsWallpaper = m_SetActiveWindowAsWallpaper
End Property

Public Property Let SetActiveWindowAsWallpaper(ByVal New_SetActiveWindowAsWallpaper As Boolean)
  
  If Not New_SetActiveWindowAsWallpaper And m_SetActiveWindowAsWallpaper Then
    vblSetDesktopWallpaper ""
  End If

  m_SetActiveWindowAsWallpaper = New_SetActiveWindowAsWallpaper
  PropertyChanged "SetActiveWindowAsWallpaper"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DisplayStatusBar() As Boolean
Attribute DisplayStatusBar.VB_Description = "True if the status bar is to be displayed."
  DisplayStatusBar = m_DisplayStatusBar
End Property

Public Property Let DisplayStatusBar(ByVal New_DisplayStatusBar As Boolean)
  m_DisplayStatusBar = New_DisplayStatusBar
  PropertyChanged "DisplayStatusBar"
  UserControl.Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DisplaySatelliteLabel() As Boolean
Attribute DisplaySatelliteLabel.VB_Description = "True if the satellite label  is to be displayed."
  DisplaySatelliteLabel = m_DisplaySatelliteLabel
End Property

Public Property Let DisplaySatelliteLabel(ByVal New_DisplaySatelliteLabel As Boolean)
  m_DisplaySatelliteLabel = New_DisplaySatelliteLabel
  PropertyChanged "DisplaySatelliteLabel"
  UserControl.Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EnableSpeech() As Boolean
Attribute EnableSpeech.VB_Description = "Enables speech"
  EnableSpeech = m_EnableSpeech
End Property

Public Property Let EnableSpeech(ByVal New_EnableSpeech As Boolean)
  m_EnableSpeech = New_EnableSpeech
  PropertyChanged "EnableSpeech"
  TXSpeech.Speed = 150
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function SpeakPosition() As Boolean
  Dim vHours As Variant
  Dim vMins As Variant
  Dim strSpeech As String
  Dim nSat As Integer

  TXSpeech.AudioReset

  nSat = m_SatelliteIndex
  If m_SatelliteElevation(nSat) > m_SetAOSLOS Then
    strSpeech = m_SatelliteName(nSat) & " is in range, azimuth " & m_SatelliteAzimuth(nSat) & " degrees,, elevation " & m_SatelliteElevation(nSat) & " degrees"
    TXSpeech.Speak strSpeech
  Else
    strTemp = satTrackNextAOS(nSat)
    If strTemp <> "" Then
      vMins = DateDiff("n", Now, strTemp)
      vHours = Int(vMins / 60)
      vMins = vMins - vHours * 60
      strSpeech = m_SatelliteName(nSat) & ", will be in range in,, "
      If vHours <> 0 Then
        If vHours = 1 Then
          strSpeech = strSpeech & " one hour "
        Else
          strSpeech = strSpeech & vHours & " hours "
        End If
      End If
      If vMins <> 0 Then
        If vHours <> 0 Then
          If vMins <> 1 Then
            strSpeech = strSpeech & " and " & vMins & " minutes"
          Else
            strSpeech = strSpeech & " and " & vMins & " minute"
          End If
        Else
          If vMins <> 1 Then
            strSpeech = strSpeech & vMins & " minutes"
          Else
            strSpeech = strSpeech & vMins & " minute"
          End If
        End If
      End If
      If vMins <> 0 Or vHours <> 0 Then
        TXSpeech.Speak strSpeech
      End If
    End If
  End If

End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get EnableSatStatus() As Boolean
Attribute EnableSatStatus.VB_Description = "Enables calculations if Vis/Ecl"
  EnableSatStatus = m_EnableSatStatus
End Property

Public Property Let EnableSatStatus(ByVal New_EnableSatStatus As Boolean)
  m_EnableSatStatus = New_EnableSatStatus
  PropertyChanged "EnableSatStatus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get DisplayGroundTrackAsPoints() As Boolean
Attribute DisplayGroundTrackAsPoints.VB_Description = "Displays the groundtrack as points if true and as a line if false"
  DisplayGroundTrackAsPoints = m_DisplayGroundTrackAsPoints
End Property

Public Property Let DisplayGroundTrackAsPoints(ByVal New_DisplayGroundTrackAsPoints As Boolean)
  m_DisplayGroundTrackAsPoints = New_DisplayGroundTrackAsPoints
  PropertyChanged "DisplayGroundTrackAsPoints"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get CalculationModel() As Models
Attribute CalculationModel.VB_Description = "Sets the Plan13 or SGP model. SGDP will be automatically selected if required."
  DisplayGroundTrackAsPoints = m_DisplayGroundTrackAsPoints
End Property

Public Property Let CalculationModel(ByVal New_CalculationModel As Models)
  m_CalculationModel = New_CalculationModel
  PropertyChanged "CalculationModel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get GroundTrackInterval() As TI
  GroundTrackInterval = m_GroundTrackInterval
End Property

Public Property Let GroundTrackInterval(ByVal New_GroundTRackInterval As TI)
  m_GroundTrackInterval = New_GroundTRackInterval
  PropertyChanged "GroundTRackInterval"
End Property

Public Property Get DisplayIcons() As Boolean
  DisplayIcons = m_DisplayIcons
End Property

Public Property Let DisplayIcons(ByVal New_DisplayIcons As Boolean)
  m_DisplayIcons = New_DisplayIcons
  PropertyChanged "DisplayIcons"
End Property
Public Property Get DisplayAOSCircle() As Boolean
  DisplayAOSCircle = m_DisplayAOSCircle
End Property

Public Property Let DisplayAOSCircle(ByVal New_DisplayAOSCircle As Boolean)
  m_DisplayAOSCircle = New_DisplayAOSCircle
  PropertyChanged "DisplayAOSCircle"
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Property Let SecondObserverHeight(ByVal vData As Variant)
  mvarSecondObserverHeight = vData / 1000
  PropertyChanged "SecondObserverHeight"
End Property

Public Property Get SecondObserverHeight() As Variant
  SecondObserverHeight = mvarSecondObserverHeight * 1000
End Property

Public Property Let SecondObserverLongitude(ByVal vData As Variant)
  mvarSecondObserverLongitude = FNRAD(vData)
  PropertyChanged "SecondObserverLongitude"
End Property

Public Property Get SecondObserverLongitude() As Variant
  SecondObserverLongitude = Round(FNDEG(mvarSecondObserverLongitude), 2)
End Property

Public Property Let SecondObserverLatitude(ByVal vData As Variant)
  mvarSecondObserverLatitude = FNRAD(vData)
  PropertyChanged "SecondObserverLatitude"
End Property

Public Property Get SecondObserverLatitude() As Variant
  SecondObserverLatitude = Round(FNDEG(mvarSecondObserverLatitude), 2)
End Property
Public Property Get SecondObserverLocation() As String
    SecondObserverLocation = m_SecondObserverLocation
End Property

Public Property Let SecondObserverLocation(ByVal New_SecondObserverLocation As String)
    m_SecondObserverLocation = New_SecondObserverLocation
    PropertyChanged "SecondObserverLocation"
End Property

Public Property Get SecondObserverEnabled() As Boolean
  SecondObserverEnabled = m_SecondObserverEnabled
End Property

Public Property Let SecondObserverEnabled(ByVal New_SecondObserverEnabled As Boolean)
  m_SecondObserverEnabled = New_SecondObserverEnabled
  PropertyChanged "SecondObserverEnabled"
End Property

Private Sub ResizeArrays(nMax)

  ReDim Preserve RS(-1 To nMax) As Double
  ReDim Preserve m_SatellitePeriod(-1 To nMax) As Double
  ReDim Preserve m_SatelliteSemiMajorAxis(-1 To nMax) As Double
  ReDim Preserve m_SatelliteSemiMinorAxis(-1 To nMax) As Double
  ReDim Preserve m_SatelliteLonOfNode(-1 To nMax) As Double
  ReDim Preserve m_SatelliteAltAtPerigee(-1 To nMax) As Double
  ReDim Preserve m_SatelliteAltAtApogee(-1 To nMax) As Double
  ReDim Preserve m_SatelliteInAOS(-1 To nMax) As Boolean
  ReDim Preserve m_SatelliteTrackOrbits(-1 To nMax) As Integer
  ReDim Preserve m_SatMaxEleTime(-1 To nMax) As Variant
  ReDim Preserve m_UplinkFrequency(-1 To nMax) As Double
  ReDim Preserve m_DownLinkFrequency(-1 To nMax) As Double
  ReDim Preserve m_SatelliteStatusText(-1 To nMax) As String
  ReDim Preserve m_SquintAngle(-1 To nMax) As Variant
  ReDim Preserve m_RangeRate(-1 To nMax) As Variant
  ReDim Preserve m_SatelliteMaximumDX(-1 To nMax) As Variant
  ReDim Preserve m_DisplaySecond(-1 To nMax) As Integer
  ReDim Preserve m_SatelliteBusy(-1 To nMax) As Boolean
  ReDim Preserve m_SatelliteTXFrequency(-1 To nMax) As Double
  ReDim Preserve m_SatelliteRxFrequency(-1 To nMax) As Double
  ReDim Preserve m_SatellitePathLoss(-1 To nMax) As Double
  ReDim Preserve m_DataValid(-1 To nMax) As Boolean
  ReDim Preserve m_DisplayTimeRequired(-1 To nMax) As Variant
  ReDim Preserve m_SatelliteDayNumber(-1 To nMax) As Variant
  ReDim Preserve m_DisplayMinute(-1 To nMax) As Integer
  ReDim Preserve m_DisplayHour(-1 To nMax) As Integer
  ReDim Preserve m_DisplayDay(-1 To nMax) As Integer
  ReDim Preserve m_DisplayMonth(-1 To nMax) As Integer
  ReDim Preserve m_DisplayYear(-1 To nMax) As Integer
  ReDim Preserve m_DisplayCentury(-1 To nMax) As Integer
  ReDim Preserve m_SatelliteOrbitNumber(-1 To nMax) As Double
  ReDim Preserve m_SatelliteLongitude(-1 To nMax) As Single
  ReDim Preserve m_SatelliteLatitude(-1 To nMax) As Single
  ReDim Preserve m_SatelliteAltitude(-1 To nMax) As Single
  ReDim Preserve m_KepsChecksum(-1 To nMax) As Variant
  ReDim Preserve m_KepsDecayRate(-1 To nMax) As Single
  ReDim Preserve m_KepsElementSet(-1 To nMax) As Variant
  ReDim Preserve m_KepsEpochTime(-1 To nMax) As Variant
  ReDim Preserve m_KepsOrbitNumber(-1 To nMax) As Variant
  ReDim Preserve m_KepsMeanMotion(-1 To nMax) As Double
  ReDim Preserve m_KepsMeanAnomoly(-1 To nMax) As Single
  ReDim Preserve m_KepsInclination(-1 To nMax) As Single
  ReDim Preserve m_KepsAOP(-1 To nMax) As Single
  ReDim Preserve m_KepsRAAN(-1 To nMax) As Single
  ReDim Preserve m_KepsEccentricity(-1 To nMax) As Single
  ReDim Preserve m_fRadiationCoefficient(-1 To nMax) As Double
  ReDim Preserve m_SatelliteElevation(-1 To nMax) As Long
  ReDim Preserve m_SatelliteAzimuth(-1 To nMax) As Long
  ReDim Preserve m_SatelliteRange(-1 To nMax) As Long
  ReDim Preserve m_SatelliteDesignator(-1 To nMax) As String
  ReDim Preserve m_SatelliteName(-1 To nMax) As String
  ReDim Preserve mvarKepsYearEpochTime(-1 To nMax) As Double
  ReDim Preserve mvarSatEpochYear(-1 To nMax) As Double
  ReDim Preserve mvarKepsYearEpochTimeFraction(-1 To nMax) As Double
  ReDim Preserve m_OrbitalModel(-1 To nMax) As Models
  ReDim Preserve m_OrbitalModelType(-1 To nMax) As ModelTypes
  ReDim Preserve m_SatelliteMA(-1 To nMax) As Double
  
  ReDim Preserve SatScreenX(-1 To nMax) As Integer
  ReDim Preserve SatScreenY(-1 To nMax) As Integer
  ReDim SatTrackLon(-1 To nMax, MaxPoints) As Double
  ReDim SatTrackLat(-1 To nMax, MaxPoints) As Double
  ReDim SatTrackElev(-1 To nMax, MaxPoints) As Integer
  ReDim SatTrackAzim(-1 To nMax, MaxPoints) As Integer
  ReDim SatTrackMutual(-1 To nMax, MaxPoints) As Long
  ReDim Preserve SatTrackPoints(-1 To nMax) As Integer
  ReDim Preserve satTrackNextAOS(-1 To nMax) As String
  ReDim Preserve satTrackMaxEle(-1 To nMax) As Integer
  ReDim Preserve satTrackMaxEleTime(-1 To nMax) As String
  ReDim Preserve m_strLine0(-1 To nMax) As String
  ReDim Preserve m_strLine1(-1 To nMax) As String
  ReDim Preserve m_strLine2(-1 To nMax) As String

  'ReDim SatTrackTime(-1 To nMax, 1000) As Variant
End Sub

Public Function AddSatellite() As Integer
  nSatCount = nSatCount + 1
  ResizeArrays nSatCount
  AddSatellite = nSatCount
End Function

Public Function UpdateKeps(strLine1 As String, strLine2 As String, strLine3 As String, nSatIndex As Integer, mModel As Models) As Boolean
  UpdateKeps = GetKeps(nSatIndex, strLine1, strLine2, strLine3, mModel)
End Function

Private Function GetKeps(nPos As Integer, strLine1 As String, strLine2 As String, strLine3 As String, mModel As Models) As Boolean
  Dim nTemp As Integer
  Dim dTemp As Double
  Dim tothrd As Double
  Dim xke As Double
  Dim ge As Double
  Dim xkmper As Double
  Dim xno As Double
  Dim temp As Double
  Dim cks As Double
  Dim J2 As Double
  Dim xincl As Double
  Dim eo As Double
  Dim del1 As Double
  Dim delo As Double
  Dim xnodp As Double
  Dim xmpda As Double
  
  m_strLine0(nPos) = strLine1
  m_strLine1(nPos) = strLine2
  m_strLine2(nPos) = strLine3
  m_SatelliteName(nPos) = Trim(strLine1)
  m_SatelliteDesignator(nPos) = Val(Mid$(strLine2, 3, 5))
  m_KepsEpochTime(nPos) = Mid$(strLine2, 19, 14)
  m_KepsDecayRate(nPos) = Val(Mid$(strLine2, 35, 9))
  m_KepsInclination(nPos) = Val(Mid$(strLine3, 9, 8))
  m_KepsRAAN(nPos) = Val(Mid$(strLine3, 18, 8))
  m_KepsEccentricity(nPos) = Val("0." + Mid$(strLine3, 27, 7))
  m_KepsAOP(nPos) = Val(Mid$(strLine3, 35, 8))
  m_KepsMeanAnomoly(nPos) = Val(Mid$(strLine3, 44, 8))
  m_KepsMeanMotion(nPos) = Val(Mid$(strLine3, 53, 11))
  m_KepsElementSet(nPos) = Val(Mid$(strLine2, 66, 3))
  m_KepsOrbitNumber(nPos) = Val(Mid$(strLine3, 64, 5))
  nTemp = Val(Mid$(strLine2, 61, 2))
  dTemp = Val(Mid$(strLine2, 55, 6))
  m_fRadiationCoefficient(nPos) = dTemp * (10 ^ (-6 + nTemp))
  
  If Val(Left$(m_KepsEpochTime(nPos), 2)) < 50 Then
    century% = 20
  Else
    century% = 19
  End If
  mvarKepsYearEpochTime(nPos) = m_KepsEpochTime(nPos) - 1000 * Int(m_KepsEpochTime(nPos) / 1000)
  mvarSatEpochYear(nPos) = 100 * century% + Int(m_KepsEpochTime(nPos) / 1000)

'  tothrd = 2 / 3
'  ge = 398600.8
'  xmnpda = 1440#
'
'  xkmper = 6378.135
'  xke = Sqr(3600# * ge / (xkmper * xkmper * xkmper))
'  xno = FNRAD(m_KepsMeanMotion(nPos))
'  xno = 2# * PI / xmnpda
'  J2 = 0.0010826158
'  ck2 = J2 / 2#
'  xincl = FNRAD(m_KepsInclination(nPos))
'  eo = m_KepsEccentricity(nPos)
'  a1 = (xke / xno) ^ tothrd
'  temp = 1.5 * ck2 * (3 * Sqr(Cos(xincl)) - 1) / (1 - eo * eo ^ 1.5)
'  del1 = temp / (a1 * a1)
'  ao = a1 * (1 - del1 * (0.5 * tothrd + del1 * (1 + 134 / 81 * del1)))
'  delo = temp / (ao * ao)
'  xnodp = xno / (1 + delo)
'  If (TwoPI / xnodp >= 225) Then
'    m_OrbitalModelType(nPos) = agTypeSGPD
'  Else
'    m_OrbitalModelType(nPos) = agTypeSGP
'  End If
  m_OrbitalModel(nPos) = mModel
End Function

Public Property Get OrbitModel() As ModelTypes
  OrbitModel = m_OrbitalModel(m_SatelliteIndex)
End Property

Public Property Let OrbitModel(ByVal New_OrbitalModel As ModelTypes)
  m_OrbitalModel(m_SatelliteIndex) = New_OrbitalModel
  PropertyChanged "OrbitModel"
End Property

'Private Sub CreateSceneGraph()
'  Dim DxL1 As Direct3DRMLight
'  Dim DxL2 As Direct3DRMLight
'
'  With mDrm
'    Set mFrS = .CreateFrame(Nothing)
'    Set mFrC = .CreateFrame(mFrS)
'    Set mFrO = .CreateFrame(mFrS)
'    Set mFrL = .CreateFrame(mFrS)
'    Set DxL1 = .CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.8, 0.8, 0.8)
'    Set DxL2 = .CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)
'  End With
'  mFrL.AddLight DxL1
'  mFrL.AddLight DxL2
'  mFrC.SetPosition Nothing, 0, 0, -4
'End Sub
'Private Sub CreateDisplay()
'  Dim DxClipper As DirectDrawClipper
'
'  Set mVpt = Nothing
'  Set mDev = Nothing
'  Set DxClipper = mDrw.CreateClipper(0)
'
'  picInner.ScaleMode = vbPixels
'  DxClipper.SetHWnd picInner.hwnd
'  Set mDev = mDrm.CreateDeviceFromClipper(DxClipper, "", picInner.ScaleWidth, picInner.ScaleHeight)
'  Set mVpt = mDrm.CreateViewport(mDev, mFrC, 0, 0, picInner.ScaleWidth, picInner.ScaleHeight)
'
'End Sub
'
'Private Sub LoadMesh()
'  Dim DxMeshB As Direct3DRMMeshBuilder3
'
'  mDrm.SetSearchPath App.Path
'  Set DxMeshB = mDrm.CreateMeshBuilder()
'  BuildSphere DxMeshB
'  PutSphereTexture mDrm, DxMeshB, App.Path & "/images/globe.bmp"
'
'  mFrO.AddVisual DxMeshB
'
'End Sub
'
'Public Sub BuildSphere(objMeshBuilder As Direct3DRMMeshBuilder3)
'  'Public Sub BuildSphere(objMeshBuilder As IDirect3DRMMeshBuilder)
'  Dim aVertices(1 To 1000) As D3DVECTOR
'  Dim aNormals(0) As D3DVECTOR
'  Dim aFaces(1 To 10000) As Long
'  Dim intVertices As Long
'  Const STEPA = 10
'  Const STEPB = 10
'  Dim axeZ As D3DVECTOR, origine As D3DVECTOR, AxeY As D3DVECTOR
'  origine.x = 0:   origine.y = 1:   origine.z = 0
'  axeZ.x = 0:      axeZ.y = 0:      axeZ.z = 1
'  AxeY.x = 0:      AxeY.y = 1:      AxeY.z = 0
'  intVertices = 1
'  Dim i As Integer, j As Integer
'  Dim tmp As D3DVECTOR
'  For i = STEPA To 180 - STEPA Step STEPA
'    For j = 0 To 360 - STEPB Step STEPB
'      mDx7.VectorRotate tmp, origine, axeZ, i * PI / 180
'      mDx7.VectorRotate aVertices(intVertices), tmp, AxeY, j * PI / 180
'      intVertices = intVertices + 1
'    Next
'  Next
'
'  intVertices = intVertices - 1
'  Dim Index As Integer
'  Index = 1
'  For i = STEPA To 180 - 2 * STEPA Step STEPA
'    Dim FirstIndex As Long
'    FirstIndex = Index
'    For j = 0 To 360 - STEPB Step STEPB
'      aFaces(Index) = 4
'      aFaces(Index + 1) = (Index \ 5) + 1
'      aFaces(Index + 2) = (Index \ 5)
'      aFaces(Index + 3) = ((Index \ 5) + (360 \ STEPB))
'      aFaces(Index + 4) = (Index \ 5) + 1 + (360 \ STEPB)
'      If j = 360 - STEPB Then
'        aFaces(Index + 1) = FirstIndex \ 5  '+ 1
'        aFaces(Index + 4) = FirstIndex \ 5 + (360 \ STEPB)
'      End If
'      Index = Index + 5
'    Next
'  Next
'  aFaces(Index) = (360 / STEPB) - 1
'  Index = Index + 1
'  For i = 1 To (360 / STEPB) - 1
'    aFaces(Index) = i
'    Index = Index + 1
'  Next
'  aFaces(Index) = 360 / STEPB
'  Index = Index + 1
'  For i = 0 To (360 / STEPB) - 1
'    aFaces(Index) = intVertices - i - 1
'    Index = Index + 1
'  Next
'  aFaces(Index) = 0
'  objMeshBuilder.AddFaces intVertices, aVertices, 0, aNormals, aFaces
'End Sub
'Public Sub PutSphereTexture(D3DRM As Direct3DRM3, MeshBuilder As Direct3DRMMeshBuilder3, ByVal strTextureFileName As String)
'  'Public Sub PutSphereTexture(D3DRM As IDirect3DRM, MeshBuilder As IDirect3DRMMeshBuilder, ByVal strTextureFileName As String)
'  Dim Box As D3DRMBOX
'  Dim MaxY As Single, MinY As Single
'  Dim Height As Single
'  Dim Wrap As Direct3DRMWrap
'  'Dim Texture As Direct3DRMTexture
'  ' Bounding box
'  MeshBuilder.GetBox Box
'  MaxY = Box.Max.y
'  MinY = Box.Min.y
'  Height = MaxY - MinY
'  Set Wrap = D3DRM.CreateWrap(D3DRMWRAP_CYLINDER, Nothing, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, MinY / Height, 1, -1 / Height)
'  Wrap.Apply MeshBuilder
'  '    D3DRM.LoadTexture strTextureFileName, Texture
'  '    MeshBuilder.SetTexture Texture
'  MeshBuilder.SetTexture D3DRM.LoadTexture(strTextureFileName)
'End Sub
'
'Private Sub Rotate(x As Single, y As Single, bCancel As Boolean)
'  Dim PTM As dxPTM
'  Dim Theta As Single
'
'  PointToMouse PTM, x, y
'
'  With PTM
'    If bCancel Then
'      mFrO.SetRotation Nothing, .dY, .dX, 0, 0
'    Else
'      Theta = .Distance / 1000
'      mFrO.SetRotation Nothing, .dY, .dX, 0, Theta
'    End If
'  End With
'  RefreshGlobe
'End Sub
''Private Sub PointToMouse(PTM As dxPTM, x As Single, y As Single)
''  Dim sX As Single, sY As Single
''
''  With PTM
''    .dX = mDownX - x
''    .dY = mDownY - y
''    sX = (.dX * .dX)
''    sY = (.dY * .dY)
''    .Distance = Sqr(sX + sY)
''  End With
''End Sub
''Private Sub RefreshGlobe()
''
'' ' picInner.Picture = LoadPicture()
''  mFrS.Move 1
''  With mVpt
''    .Clear D3DRMCLEAR_ALL
''    .Render mFrS
''  End With
''  mDev.Update
''  DoEvents
''
''End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picInner,picInner,-1,Picture

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picInner,picInner,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  picPic.Picture = picInner.Image
  picPic.Refresh
  Set Picture = picPic.Picture
  picPic.Picture = LoadPicture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  Set picInner.Picture = New_Picture
  PropertyChanged "Picture"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,1,2,0
Public Property Get SatelliteBearing() As Single
Attribute SatelliteBearing.VB_Description = "The satellites current bearing"
Attribute SatelliteBearing.VB_MemberFlags = "400"
  Dim sDist As Single
  Dim sBearing As Single
  Dim dStartY As Double
  Dim dStartX As Double
  Dim dEndX As Double
  Dim dEndY As Double

  sDist = CalculateDistAndBearing(SatTrackLat(m_SatelliteIndex, 0), SatTrackLon(m_SatelliteIndex, 0), SatTrackLat(m_SatelliteIndex, 1), SatTrackLon(m_SatelliteIndex, 1), sBearing)
  SatelliteBearing = sBearing
End Property

Public Property Let SatelliteBearing(ByVal New_SatelliteBearing As Single)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SatelliteBearing = New_SatelliteBearing
  PropertyChanged "SatelliteBearing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get UserStatusPanelText() As String
Attribute UserStatusPanelText.VB_Description = "Sets the test in the user definable panel in the status bar"
  UserStatusPanelText = m_UserStatusPanelText
End Property

Public Property Let UserStatusPanelText(ByVal New_UserStatusPanelText As String)
  m_UserStatusPanelText = New_UserStatusPanelText
  UserControl.StatusBar1.Panels(6).Text = m_UserStatusPanelText
  PropertyChanged "UserStatusPanelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get GroundTrackPointSize() As Integer
Attribute GroundTrackPointSize.VB_Description = "Size, in pixels, of each point plotted on the ground track"
  GroundTrackPointSize = m_GroundTrackPointSize
End Property

Public Property Let GroundTrackPointSize(ByVal New_GroundTrackPointSize As Integer)
  m_GroundTrackPointSize = New_GroundTrackPointSize
  PropertyChanged "GroundTrackPointSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get Sunrise() As Variant
Attribute Sunrise.VB_Description = "Sunrise time"
Attribute Sunrise.VB_MemberFlags = "400"
  Sunrise = m_SunRise
End Property

Public Property Let Sunrise(ByVal New_SunRise As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SunRise = New_SunRise
  PropertyChanged "SunRise"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get SunNoon() As Variant
Attribute SunNoon.VB_Description = "Time of solar noon"
Attribute SunNoon.VB_MemberFlags = "400"
  SunNoon = m_SunNoon
End Property

Public Property Let SunNoon(ByVal New_SunNoon As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SunNoon = New_SunNoon
  PropertyChanged "SunNoon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,2,0
Public Property Get Sunset() As Variant
Attribute Sunset.VB_Description = "Sunset time"
Attribute Sunset.VB_MemberFlags = "400"
  Sunset = m_SunSet
End Property

Public Property Let Sunset(ByVal New_SunSet As Variant)
  If Ambient.UserMode = False Then Err.Raise 387
  If Ambient.UserMode Then Err.Raise 382
  m_SunSet = New_SunSet
  PropertyChanged "SunSet"
End Property

