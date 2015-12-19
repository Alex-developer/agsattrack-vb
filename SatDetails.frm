VERSION 5.00
Begin VB.Form SatDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sat Details"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3600
   Icon            =   "SatDetails.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "SatData"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDownlinkDoppler 
      Height          =   195
      Left            =   2010
      TabIndex        =   25
      Top             =   1830
      Width           =   1065
   End
   Begin VB.Label lblDownlinkFrequency 
      Height          =   195
      Left            =   930
      TabIndex        =   24
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label lblUplinkDoppler 
      Height          =   195
      Left            =   2010
      TabIndex        =   23
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblUplinkFreq 
      Height          =   195
      Left            =   930
      TabIndex        =   22
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Downlink"
      Height          =   195
      Left            =   90
      TabIndex        =   21
      Top             =   1830
      Width           =   705
   End
   Begin VB.Label lblPathLoss 
      Caption         =   "Path Loss here"
      Height          =   195
      Left            =   2760
      TabIndex        =   20
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label11 
      Caption         =   "Path Loss"
      Height          =   195
      Left            =   1860
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblMaxDX 
      Caption         =   "Max DX"
      Height          =   195
      Left            =   2580
      TabIndex        =   18
      Top             =   540
      Width           =   795
   End
   Begin VB.Label Label10 
      Caption         =   "Max DX"
      Height          =   195
      Left            =   1860
      TabIndex        =   17
      Top             =   540
      Width           =   615
   End
   Begin VB.Label lblTime 
      Caption         =   "TimeHere"
      Height          =   195
      Left            =   2580
      TabIndex        =   16
      Top             =   300
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "Time"
      Height          =   195
      Left            =   1860
      TabIndex        =   15
      Top             =   300
      Width           =   435
   End
   Begin VB.Label lblDate 
      Caption         =   "DateHere"
      Height          =   195
      Left            =   2580
      TabIndex        =   14
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "Date"
      Height          =   195
      Left            =   1860
      TabIndex        =   13
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label7 
      Caption         =   "Uplink"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label lblRange 
      Caption         =   "RangeHere"
      Height          =   195
      Left            =   900
      TabIndex        =   11
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Range"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1260
      Width           =   675
   End
   Begin VB.Label lblLongitude 
      Caption         =   "LonHere"
      Height          =   195
      Left            =   900
      TabIndex        =   9
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Longitude"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label lblLatitude 
      Caption         =   "LatHere"
      Height          =   195
      Left            =   900
      TabIndex        =   7
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Latitude"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   705
   End
   Begin VB.Label lblAzimuth 
      Caption         =   "AzHere"
      Height          =   195
      Left            =   900
      TabIndex        =   5
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Azimuth"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   675
   End
   Begin VB.Label lblElevation 
      Caption         =   "EleHere"
      Height          =   195
      Left            =   900
      TabIndex        =   3
      Top             =   300
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Elevation"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   765
   End
   Begin VB.Label lblOrbit 
      Caption         =   "OrbitHere"
      Height          =   195
      Left            =   900
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Orbit"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   705
   End
End
Attribute VB_Name = "SatDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    SETtopmostwindow Me, False
End Sub

