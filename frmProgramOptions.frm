VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form frmProgramOptions 
   HelpContextID   =  55
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Options"
   ClientHeight    =   5310
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6345
   Icon            =   "frmProgramOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   510
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   4545
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   8017
      _Version        =   131082
      TabCount        =   7
      TagVariant      =   ""
      Tabs            =   "frmProgramOptions.frx":000C
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   4155
         Left            =   30
         TabIndex        =   95
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":0183
         Begin VB.Frame Frame19 
            Caption         =   " Miscellaneous "
            Height          =   975
            Left            =   2160
            TabIndex        =   99
            Top             =   240
            Width           =   3555
            Begin VB.CheckBox chkAlwaysTrack 
               Caption         =   "Always tack Azmiuth"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               ToolTipText     =   "Track Azimuth even when satellite is below the horizon"
               Top             =   300
               Width           =   3195
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   " Rotator Interface "
            Height          =   3855
            Left            =   120
            TabIndex        =   96
            Top             =   120
            Width           =   1935
            Begin VB.OptionButton optRotator 
               Caption         =   "None"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   94
               ToolTipText     =   "download from www.ea4tx.com"
               Top             =   300
               Width           =   1095
            End
            Begin VB.CommandButton cmdRotatorConfigure 
               Caption         =   "Configure"
               Height          =   375
               Left            =   1020
               TabIndex        =   98
               Top             =   3420
               Width           =   795
            End
            Begin VB.OptionButton optRotator 
               Caption         =   "ARSWIN"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   97
               ToolTipText     =   "download from www.ea4tx.com"
               Top             =   600
               Width           =   1095
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   4155
         Left            =   -99969
         TabIndex        =   82
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":01AB
         Begin VB.Frame Frame20 
            Caption         =   " Image Upload "
            Height          =   1935
            Left            =   120
            TabIndex        =   102
            Top             =   2100
            Width           =   5655
            Begin VB.TextBox txtFTPImagesDir 
               Height          =   285
               Left            =   1380
               TabIndex        =   101
               ToolTipText     =   "Enter the directory on the FTP Server to store the images"
               Top             =   1320
               Width           =   2835
            End
            Begin VB.TextBox txtFTPHTMLDir 
               Height          =   285
               Left            =   3960
               TabIndex        =   112
               ToolTipText     =   "Enter the Name of the directory on the FTP Server for the HTML Template"
               Top             =   960
               Width           =   1395
            End
            Begin VB.TextBox txtFTPTemplate 
               Height          =   285
               Left            =   1380
               TabIndex        =   110
               ToolTipText     =   "Enter the name of the HTML Template"
               Top             =   960
               Width           =   1635
            End
            Begin VB.TextBox txtFTPPassword 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   3960
               PasswordChar    =   "*"
               TabIndex        =   108
               ToolTipText     =   "Enter the FTP site password"
               Top             =   600
               Width           =   1395
            End
            Begin VB.TextBox txtFTPUsername 
               Height          =   285
               Left            =   1080
               TabIndex        =   106
               ToolTipText     =   "Enter the FTP site username"
               Top             =   600
               Width           =   1515
            End
            Begin VB.TextBox txtFTPServer 
               Height          =   285
               Left            =   1080
               TabIndex        =   104
               ToolTipText     =   "Enter the URL of the FTP server"
               Top             =   240
               Width           =   4395
            End
            Begin VB.Label Label27 
               Caption         =   "Images Directory"
               Height          =   195
               Left            =   120
               TabIndex        =   113
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label26 
               Caption         =   "HTML Dir"
               Height          =   195
               Left            =   3180
               TabIndex        =   111
               Top             =   1020
               Width           =   795
            End
            Begin VB.Label Label25 
               Caption         =   "HTML Template"
               Height          =   195
               Left            =   120
               TabIndex        =   109
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label24 
               Caption         =   "Password"
               Height          =   255
               Left            =   3180
               TabIndex        =   107
               Top             =   660
               Width           =   795
            End
            Begin VB.Label Label23 
               Caption         =   "Username"
               Height          =   195
               Left            =   120
               TabIndex        =   105
               Top             =   660
               Width           =   795
            End
            Begin VB.Label Label22 
               Caption         =   "FTP Server"
               Height          =   255
               Left            =   120
               TabIndex        =   103
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   " FTP Options "
            Height          =   1095
            Left            =   2880
            TabIndex        =   90
            Top             =   240
            Width           =   2415
            Begin VB.CheckBox chkPasv 
               Alignment       =   1  'Right Justify
               Caption         =   "Use PASV Mode"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               ToolTipText     =   "Use PASV mode for FTP."
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txtTimeout 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   91
               Text            =   "30"
               ToolTipText     =   "Enter the timeout period for the FTP/HTTP server."
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label21 
               Caption         =   "Timeout:"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label20 
               Caption         =   "sec."
               Height          =   255
               Left            =   1800
               TabIndex        =   93
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Proxy Settings"
            Height          =   1815
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   2655
            Begin VB.CheckBox chkProxy 
               Alignment       =   1  'Right Justify
               Caption         =   "Use HTTP Proxy"
               Height          =   255
               Left            =   120
               TabIndex        =   87
               ToolTipText     =   "Enable this option if you use a proxy server"
               Top             =   360
               Width           =   2415
            End
            Begin VB.TextBox txtProxyIP 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1080
               TabIndex        =   86
               ToolTipText     =   "Enter the address of the proxy server"
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtProxyPort 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1080
               TabIndex        =   85
               ToolTipText     =   "Enter the proxy server port number"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.CheckBox chkFTPPRoxy 
               Alignment       =   1  'Right Justify
               Caption         =   "FTP Through HTTP Proxy"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               ToolTipText     =   "Enable FTP through HTTP proxy"
               Top             =   1440
               Width           =   2415
            End
            Begin VB.Label Label19 
               Caption         =   "Adress"
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label18 
               Caption         =   "Port"
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   1080
               Width           =   975
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   4155
         Left            =   -99969
         TabIndex        =   58
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":01D3
         Begin VB.Frame Frame15 
            Caption         =   " Second Observer "
            Height          =   1875
            Left            =   120
            TabIndex        =   68
            Top             =   2040
            Width           =   5505
            Begin VB.ComboBox cmbSecondObs 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   74
               ToolTipText     =   "Select the second observer from the list"
               Top             =   540
               Width           =   5055
            End
            Begin VB.TextBox txtSecondLat 
               Height          =   285
               Left            =   900
               TabIndex        =   73
               ToolTipText     =   "Enter the second observer Latitude in the format HH.MM.SS (N/S)"
               Top             =   960
               Width           =   885
            End
            Begin VB.TextBox txtSecondLong 
               Height          =   285
               Left            =   2700
               TabIndex        =   72
               ToolTipText     =   "Enter the second observer Longitude in the format HHH.MM.SS (E/W)"
               Top             =   960
               Width           =   1035
            End
            Begin VB.TextBox txtSecondName 
               Height          =   285
               Left            =   900
               TabIndex        =   71
               ToolTipText     =   "Enter the name of the second observer toplot on the map"
               Top             =   1365
               Width           =   4185
            End
            Begin VB.TextBox txtSecondHeight 
               Height          =   285
               Left            =   4500
               TabIndex        =   70
               ToolTipText     =   "Enter the second observer height in meters"
               Top             =   960
               Width           =   555
            End
            Begin VB.CheckBox chkEnableSecond 
               Caption         =   "Enable Mutual Calculations"
               Height          =   225
               Left            =   120
               TabIndex        =   69
               ToolTipText     =   "Enables mutual observer calculations"
               Top             =   240
               Width           =   2355
            End
            Begin VB.Label Label16 
               Caption         =   "Latitude"
               Height          =   195
               Left            =   180
               TabIndex        =   17
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "Longitude"
               Height          =   195
               Left            =   1860
               TabIndex        =   77
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label Label14 
               Caption         =   "Location"
               Height          =   195
               Left            =   180
               TabIndex        =   76
               Top             =   1395
               Width           =   675
            End
            Begin VB.Label Label13 
               Caption         =   "Height"
               Height          =   195
               Left            =   3900
               TabIndex        =   75
               Top             =   1020
               Width           =   555
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " Default Observer "
            Height          =   1875
            Left            =   120
            TabIndex        =   59
            Top             =   90
            Width           =   5505
            Begin VB.TextBox txtLatitude 
               Height          =   285
               Left            =   960
               TabIndex        =   63
               Text            =   "Text1"
               ToolTipText     =   "Enter your Latitude in the format HH.MM.SS"
               Top             =   240
               Width           =   1185
            End
            Begin VB.TextBox txtLongitude 
               Height          =   285
               Left            =   960
               TabIndex        =   62
               Text            =   "Text1"
               ToolTipText     =   "Enter your longitude in the format HHH.MM.SS"
               Top             =   640
               Width           =   1185
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Left            =   960
               TabIndex        =   61
               Text            =   "Text2"
               ToolTipText     =   "Enter the observer name to plot on the map"
               Top             =   1440
               Width           =   2205
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Left            =   960
               TabIndex        =   60
               Text            =   "Text1"
               ToolTipText     =   "Enter the default observer height in meters"
               Top             =   1040
               Width           =   555
            End
            Begin VB.Label Label2 
               Caption         =   "Latitude"
               Height          =   195
               Left            =   180
               TabIndex        =   67
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Longitude"
               Height          =   195
               Left            =   180
               TabIndex        =   66
               Top             =   700
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Location"
               Height          =   195
               Left            =   180
               TabIndex        =   65
               Top             =   1500
               Width           =   675
            End
            Begin VB.Label Label5 
               Caption         =   "Height"
               Height          =   195
               Left            =   180
               TabIndex        =   64
               Top             =   1100
               Width           =   675
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   4155
         Index           =   2
         Left            =   -99969
         TabIndex        =   20
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":01FB
         Begin VB.Frame Frame6 
            Caption         =   " Orthographic "
            Height          =   2175
            Left            =   270
            TabIndex        =   22
            Top             =   1860
            Width           =   5145
            Begin VB.Frame Frame9 
               Caption         =   " Size (Pixels) "
               Height          =   915
               Left            =   2280
               TabIndex        =   30
               Top             =   210
               Width           =   2685
               Begin VB.TextBox txtOrthY 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   420
                  TabIndex        =   33
                  Text            =   "Text1"
                  Top             =   570
                  Width           =   585
               End
               Begin VB.TextBox txtOrthX 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   420
                  TabIndex        =   32
                  Text            =   "Text1"
                  Top             =   240
                  Width           =   585
               End
               Begin VB.Label Label6 
                  Caption         =   "Y"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   34
                  Top             =   570
                  Width           =   195
               End
               Begin VB.Label Label1 
                  Caption         =   "X"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   31
                  Top             =   240
                  Width           =   195
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   " Locations "
               Height          =   945
               Left            =   2280
               TabIndex        =   27
               Top             =   1110
               Width           =   2685
               Begin VB.CommandButton Command1 
                  Caption         =   "Edit"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   1890
                  TabIndex        =   29
                  Top             =   210
                  Width           =   675
               End
               Begin VB.CheckBox Check2 
                  Caption         =   "Show Locations"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   210
                  Width           =   1455
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   " View From "
               Height          =   1005
               Left            =   120
               TabIndex        =   24
               Top             =   1080
               Width           =   2055
               Begin VB.OptionButton Option2 
                  Caption         =   "Observer"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   150
                  TabIndex        =   26
                  Top             =   570
                  Width           =   1365
               End
               Begin VB.OptionButton Option1 
                  Caption         =   " Satellite"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   150
                  TabIndex        =   25
                  Top             =   270
                  Width           =   1245
               End
            End
            Begin VB.CheckBox chkbOrthSahde 
               Caption         =   "Show day and night"
               Height          =   315
               Left            =   150
               TabIndex        =   23
               ToolTipText     =   "Enables night shading in orthographic views"
               Top             =   330
               Width           =   1815
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " General "
            Height          =   1455
            Left            =   270
            TabIndex        =   21
            Top             =   210
            Width           =   5175
            Begin MSComCtl2.UpDown UpDown4 
               Height          =   285
               Left            =   2176
               TabIndex        =   78
               Top             =   1080
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               Value           =   1
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtGTPS"
               BuddyDispid     =   196673
               OrigLeft        =   2490
               OrigTop         =   1050
               OrigRight       =   2730
               OrigBottom      =   1395
               Max             =   4
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtGTPS 
               Height          =   285
               Left            =   1920
               TabIndex        =   80
               ToolTipText     =   "Enter the size of the ground track dots."
               Top             =   1080
               Width           =   255
            End
            Begin VB.CheckBox chkIcons 
               Caption         =   "Display Satellites/Sun/Moon as pictures"
               Height          =   285
               Left            =   120
               TabIndex        =   57
               ToolTipText     =   "If enabled objects are displayed as pictures rather then dots"
               Top             =   600
               Width           =   3225
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   285
               Left            =   1680
               TabIndex        =   43
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtUpdate"
               BuddyDispid     =   196675
               OrigLeft        =   1800
               OrigTop         =   240
               OrigRight       =   2040
               OrigBottom      =   555
               Max             =   120
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtUpdate 
               Height          =   285
               Left            =   1230
               TabIndex        =   41
               ToolTipText     =   "Enter the number of seconds between screen updates"
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label17 
               Caption         =   "Ground track point size"
               Height          =   225
               Left            =   120
               TabIndex        =   79
               Top             =   1110
               Width           =   1755
            End
            Begin VB.Label Label8 
               Caption         =   "seconds"
               Height          =   225
               Left            =   2040
               TabIndex        =   42
               Top             =   270
               Width           =   705
            End
            Begin VB.Label Label7 
               Caption         =   "Update every"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   270
               Width           =   1005
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   4155
         Left            =   -99969
         TabIndex        =   14
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":0223
         Begin VB.PictureBox Picture1 
            Height          =   2715
            Left            =   120
            ScaleHeight     =   2655
            ScaleWidth      =   5475
            TabIndex        =   114
            Top             =   1260
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.Frame Frame3 
            Caption         =   " Timezone"
            Height          =   975
            Left            =   150
            TabIndex        =   15
            Top             =   180
            Width           =   5445
            Begin VB.CheckBox chkAutoAdjust 
               Caption         =   "Automatically adjust for daylight saving"
               Height          =   285
               Left            =   1830
               TabIndex        =   19
               ToolTipText     =   "If checked the daylight saving will be automatically enabled."
               Top             =   600
               Width           =   3195
            End
            Begin VB.CheckBox chkDaylightSaving 
               Caption         =   "Daylight Saving"
               Height          =   255
               Left            =   210
               TabIndex        =   18
               ToolTipText     =   "If checked daylight saving will be enabled."
               Top             =   600
               Width           =   1575
            End
            Begin VB.ComboBox cmbTimezones 
               Height          =   315
               Left            =   180
               Style           =   2  'Dropdown List
               TabIndex        =   16
               ToolTipText     =   "Select the observers timezone"
               Top             =   240
               Width           =   5115
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4155
         Left            =   -99969
         TabIndex        =   5
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":024B
         Begin VB.Frame Frame10 
            Caption         =   " Orthographic locations database "
            Height          =   645
            Left            =   90
            TabIndex        =   35
            Top             =   1860
            Width           =   5415
            Begin VB.CommandButton Command2 
               Caption         =   "..."
               Height          =   315
               Left            =   4860
               TabIndex        =   37
               Top             =   240
               Width           =   435
            End
            Begin VB.TextBox txtOrthLocations 
               Height          =   285
               Left            =   60
               TabIndex        =   36
               Top             =   240
               Width           =   4695
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " Default Raw Keplarian Elements Path"
            Height          =   645
            Left            =   90
            TabIndex        =   9
            Top             =   1110
            Width           =   5415
            Begin VB.TextBox txtDefKepsPath 
               Height          =   285
               Left            =   60
               TabIndex        =   11
               Text            =   "Text1"
               Top             =   240
               Width           =   4695
            End
            Begin VB.CommandButton cmdRawKepsBrowse 
               Caption         =   "..."
               Height          =   315
               Left            =   4860
               TabIndex        =   10
               Top             =   240
               Width           =   435
            End
         End
         Begin VB.Frame fraSample2 
            Caption         =   " Default Database Path"
            Height          =   645
            Left            =   90
            TabIndex        =   6
            Top             =   360
            Width           =   5415
            Begin VB.TextBox txtDefDatabasePath 
               Height          =   285
               Left            =   60
               TabIndex        =   8
               Text            =   "Text1"
               Top             =   240
               Width           =   4695
            End
            Begin VB.CommandButton cmdDatabaseBrowse 
               Caption         =   "..."
               Height          =   315
               Left            =   4860
               TabIndex        =   7
               Top             =   240
               Width           =   435
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4155
         Left            =   -99969
         TabIndex        =   4
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7329
         _Version        =   131082
         TabGuid         =   "frmProgramOptions.frx":0273
         Begin VB.Frame Frame14 
            Caption         =   " Speech "
            Height          =   1095
            Left            =   2340
            TabIndex        =   51
            Top             =   1950
            Width           =   3405
            Begin MSComCtl2.UpDown UpDown3 
               Height          =   315
               Left            =   1756
               TabIndex        =   55
               Top             =   600
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Value           =   5
               BuddyControl    =   "txtSpeech"
               BuddyDispid     =   196694
               OrigLeft        =   2100
               OrigTop         =   600
               OrigRight       =   2340
               OrigBottom      =   915
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtSpeech 
               Height          =   315
               Left            =   1410
               TabIndex        =   54
               Text            =   "Text1"
               ToolTipText     =   "How frequently the speech is audible"
               Top             =   600
               Width           =   345
            End
            Begin VB.CheckBox chkSpeech 
               Caption         =   "Announce time until AOS"
               Height          =   285
               Left            =   150
               TabIndex        =   52
               ToolTipText     =   "If enabled the time until current AOS will be spoken."
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label12 
               Caption         =   "updates"
               Height          =   255
               Left            =   2100
               TabIndex        =   56
               Top             =   630
               Width           =   825
            End
            Begin VB.Label Label11 
               Caption         =   "Announce every"
               Height          =   255
               Left            =   150
               TabIndex        =   53
               Top             =   660
               Width           =   1245
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   " General "
            Height          =   1095
            Left            =   60
            TabIndex        =   49
            Top             =   1950
            Width           =   2235
            Begin VB.CheckBox chkSysTray 
               Caption         =   "Minimise to System Tray"
               Height          =   315
               Left            =   90
               TabIndex        =   50
               ToolTipText     =   "If enabled the program will be minimised to the system tray."
               Top             =   270
               Width           =   2085
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   " Element Sets "
            Height          =   765
            Left            =   90
            TabIndex        =   44
            Top             =   1050
            Width           =   5655
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   285
               Left            =   3106
               TabIndex        =   47
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               Value           =   7
               BuddyControl    =   "txtKepsAge"
               BuddyDispid     =   196701
               OrigLeft        =   3450
               OrigTop         =   240
               OrigRight       =   3690
               OrigBottom      =   555
               Max             =   60
               Min             =   1
               SyncBuddy       =   -1  'True
               Wrap            =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtKepsAge 
               Height          =   285
               Left            =   2760
               TabIndex        =   46
               ToolTipText     =   "Elements older than this age will trigger a warning to update the elements"
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label10 
               Caption         =   "Days"
               Height          =   255
               Left            =   3420
               TabIndex        =   48
               Top             =   270
               Width           =   465
            End
            Begin VB.Label Label9 
               Caption         =   "Warn if average age is greater than"
               Height          =   255
               Left            =   150
               TabIndex        =   45
               Top             =   270
               Width           =   2535
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   " Startup "
            Height          =   825
            Left            =   2640
            TabIndex        =   38
            Top             =   120
            Width           =   2205
            Begin VB.CheckBox chkSaveOpen 
               Caption         =   "Save/Load Last View"
               Height          =   255
               Left            =   60
               TabIndex        =   39
               ToolTipText     =   "If enabled the current view will be saved"
               Top             =   270
               Width           =   1935
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   " Predictions View "
            Height          =   825
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   2235
            Begin VB.CheckBox chkSetupVis 
               Caption         =   "Indicate Visible Satellites"
               Height          =   285
               Left            =   60
               TabIndex        =   13
               ToolTipText     =   "If enabled visible satellites will be displayed in the predictions window"
               Top             =   300
               Width           =   2055
            End
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4845
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3690
      TabIndex        =   1
      Top             =   4845
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2460
      TabIndex        =   0
      Top             =   4845
      Width           =   1095
   End
End
Attribute VB_Name = "frmProgramOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTempDefDatabasePath As String
Dim strTempKepsPath As String
Dim strOrthLocs As String
Dim sCurZone As String
Dim bError As Boolean


Private Sub chkAlwaysTrack_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkAutoAdjust_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkbOrthSahde_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkDaylightSaving_Click()
  Me.cmdApply.Enabled = True
End Sub


Private Sub chkEnableSecond_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkFTPPRoxy_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkIcons_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkPasv_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkProxy_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkSaveOpen_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkSetupVis_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkSpeech_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub chkSysTray_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub cmbSecondObs_Click()
  Dim nPos As Integer

  nPos = Me.cmbSecondObs.ListIndex
  
  Me.txtSecondLat = sDxDetails(nPos).strLat
  Me.txtSecondLong = -sDxDetails(nPos).strLon
  Me.txtSecondName = sDxDetails(nPos).strName
  
  Me.cmdApply.Enabled = True

End Sub

Private Sub cmbTimezones_Click()
  CurrentTZI = LocTZI(Me.cmbTimezones.ListIndex)
  UpdateTZInfo
  Me.cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
  Dim fForm As Form
  Dim i As Integer
  
  bError = False
  If strValidateLatLon(Me, Me.txtLongitude, Me.txtLatitude, True) = "Ok" Then
    If strValidateLatLon(Me, Me.txtLatitude, Me.txtLongitude, False) = "Ok" Then

  If strValidateLatLon(Me, Me.txtSecondLong, Me.txtSecondLat, True) = "Ok" Then
    If strValidateLatLon(Me, Me.txtSecondLat, Me.txtSecondLong, False) = "Ok" Then

      sProgramOptions.strDefDatabasePath = strTempDefDatabasePath
      sProgramOptions.strDefKepsPath = strTempKepsPath
      sProgramOptions.strOrthLocations = strOrthLocs
  
      sProgramOptions.bIcons = IIf(Me.chkIcons.Value = 1, True, False)
  
      sProgramOptions.bSaveOpenlastView = IIf(Me.chkSaveOpen.Value = 1, True, False)

      sProgramOptions.bIndicateVis = IIf(Me.chkSetupVis.Value = 1, True, False)

      sProgramOptions.strTimeZone = LocTZI(Me.cmbTimezones.ListIndex).StandardName
      sProgramOptions.nTimezoneAdjust = LocTZI(Me.cmbTimezones.ListIndex).Bias / 60

      sProgramOptions.bAutoadjust = IIf(Me.chkAutoAdjust.Value = 1, True, False)
      sProgramOptions.bDaylightSaving = IIf(Me.chkDaylightSaving.Value = 1, True, False)
  
'      sProgramOptions.nLatitude = Me.txtLatitude
'      sProgramOptions.nLongitude = Me.txtLongitude
      sProgramOptions.nLatitude = sConvertPos(Me.txtLatitude) / 60
      sProgramOptions.nLongitude = sConvertPos(Me.txtLongitude) / 60
      sProgramOptions.nHeight = Me.txtHeight
      sProgramOptions.strLocation = Me.txtLocation
  
'      sProgramOptions.nSecondLatitude = Me.txtSecondLat
'      sProgramOptions.nSecondLongitude = Me.txtSecondLong
      sProgramOptions.nSecondLatitude = sConvertPos(Me.txtSecondLat) / 60
      sProgramOptions.nSecondLongitude = sConvertPos(Me.txtSecondLong) / 60
      sProgramOptions.nSecondHeight = Me.txtSecondHeight
      sProgramOptions.nSecondName = Me.txtSecondName
      sProgramOptions.bSecondUsed = IIf(Me.chkEnableSecond = 1, True, False)
  
      sProgramOptions.nKepsAge = Me.txtKepsAge

      sProgramOptions.bSysTray = IIf(Me.chkSysTray.Value = 1, True, False)
  
      sProgramOptions.nSpeechInterval = Me.txtSpeech.Text
      sProgramOptions.bSpeech = IIf(Me.chkSpeech.Value = 1, True, False)
  
'tab 2
      sProgramOptions.nOrthX = Me.txtOrthX
      sProgramOptions.nOrthY = Me.txtOrthY
      sProgramOptions.bOrthShade = IIf(Me.chkbOrthSahde.Value = 1, True, False)
      sProgramOptions.nUpdateInterval = Val(Me.txtUpdate)
      sProgramOptions.nGroundTrackPointSize = Me.txtGTPS.Text
      
' Rotator
  sProgramOptions.bRotatorAlwaysTrack = IIf(Me.chkAlwaysTrack.Value = 1, True, False)
  For i = 0 To Me.optRotator.UBound
    If Me.optRotator(i).Value Then
      sProgramOptions.nRotatorType = i
      Exit For
    End If
  Next i
  sProgramOptions.bRotatorEnabled = IIf(sProgramOptions.nRotatorType = 0, False, True)
  
  ' ftp
  sProgramOptions.strFTPHTMLDir = Me.txtFTPHTMLDir
  sProgramOptions.strFTPImagesDir = Me.txtFTPImagesDir
  sProgramOptions.strFTPPassword = Me.txtFTPPassword
  sProgramOptions.strFTPServer = Me.txtFTPServer
  sProgramOptions.strFTPHTMLTemplate = Me.txtFTPTemplate
  sProgramOptions.strFTPUserName = Me.txtFTPUsername
      
      For Each fForm In Forms
        With fForm
          If Left(.Caption, 7) = "SatView" Then
            .ocxSat.GroundTrackPointSize = sProgramOptions.nGroundTrackPointSize
          End If
        End With
      Next
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

Private Sub cmdDatabaseBrowse_Click()
  strTempDefDatabasePath = BrowseForFolder(Me.hwnd, "Select Path for Databases")
  Me.txtDefDatabasePath.Text = LongDirFix(strTempDefDatabasePath, 30)
End Sub

Private Sub cmdOk_Click()
  bError = False
  If Me.cmdApply.Enabled Then
    cmdApply_Click
  End If
  If Not bError Then
    Unload Me
  End If
End Sub

Private Sub cmdRawKepsBrowse_Click()
  strTempKepsPath = BrowseForFolder(Me.hwnd, "Select Path for Keps")
  Me.txtDefKepsPath.Text = LongDirFix(strTempKepsPath, 30)
End Sub


Private Sub Command2_Click()
  With Me.CommonDialog1
    .DialogTitle = "Select database file"
    .ShowOpen
    strOrthLocs = .FileName
    Me.txtOrthLocations = LongDirFix(strOrthLocs, 45)
    Me.cmdApply.Enabled = True
  End With
End Sub

Private Sub Form_Load()
  Dim n As Integer
  Dim i As Integer
  Dim strTemp As String
  Dim strLat As String
  Dim strLon As String
  
    'center the form
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
' Tab1 General options
  Me.chkSetupVis.Value = IIf(sProgramOptions.bIndicateVis, 1, 0)
  Me.chkSaveOpen.Value = IIf(sProgramOptions.bSaveOpenlastView, 1, 0)
  
  FormatLatLon sProgramOptions.nLatitude * 60, sProgramOptions.nLongitude * 60, strLat, strLon, True
  Me.txtLatitude = strLat
  Me.txtLongitude = strLon
'  Me.txtLatitude = sProgramOptions.nLatitude
'  Me.txtLongitude = sProgramOptions.nLongitude
  Me.txtHeight = sProgramOptions.nHeight
  Me.txtLocation = sProgramOptions.strLocation
  
  FormatLatLon sProgramOptions.nSecondLatitude * 60, sProgramOptions.nSecondLongitude * 60, strLat, strLon, True
'  Me.txtSecondLat = sProgramOptions.nSecondLatitude
'  Me.txtSecondLong = sProgramOptions.nSecondLongitude
  Me.txtSecondLat = strLat
  Me.txtSecondLong = strLon
  Me.txtSecondHeight = sProgramOptions.nSecondHeight
  Me.txtSecondName = sProgramOptions.nSecondName
  Me.chkEnableSecond.Value = IIf(sProgramOptions.bSecondUsed, 1, 0)
  
  Me.txtKepsAge = sProgramOptions.nKepsAge
  Me.chkSysTray.Value = IIf(sProgramOptions.bSysTray, 1, 0)
  Me.txtSpeech.Text = sProgramOptions.nSpeechInterval
  Me.chkSpeech.Value = IIf(sProgramOptions.bSpeech, 1, 0)
  Me.chkIcons.Value = IIf(sProgramOptions.bIcons, 1, 0)
  
' Tab 2 - views
  Me.txtOrthX = sProgramOptions.nOrthX
  Me.txtOrthY = sProgramOptions.nOrthY
  Me.chkbOrthSahde.Value = IIf(sProgramOptions.bOrthShade = True, 1, 0)
  Me.txtUpdate = sProgramOptions.nUpdateInterval
  Me.txtGTPS.Text = sProgramOptions.nGroundTrackPointSize
  
' Tab 3 Database Paths
  strTempDefDatabasePath = sProgramOptions.strDefDatabasePath
  strTempKepsPath = sProgramOptions.strDefKepsPath
  strOrthLocs = sProgramOptions.strOrthLocations
  Me.txtDefDatabasePath.Text = LongDirFix(strTempDefDatabasePath, 45)
  Me.txtDefKepsPath.Text = LongDirFix(strTempKepsPath, 45)
  Me.txtOrthLocations = LongDirFix(strOrthLocs, 45)
' Tab 4 - Timezone
  sCurZone = GetRegValueStr("System\CurrentControlSet\Control\TimeZoneInformation", "StandardName")
  For i = 0 To UBound(LocTZI)
    Me.cmbTimezones.AddItem LocTZI(i).DisplayName
    If LocTZI(i).StandardName = sCurZone Then n = i
    If LocTZI(i).StandardName = sProgramOptions.strTimeZone Then n = i
  Next
  Me.cmbTimezones.ListIndex = n
  DoEvents
  cmbTimezones_Click
  If sProgramOptions.bAutoadjust Then
    Me.chkAutoAdjust.Value = 1
   Else
    Me.chkAutoAdjust.Value = 0
  End If
  
  ' Rotator
  Me.chkAlwaysTrack = IIf(sProgramOptions.bRotatorAlwaysTrack, 1, 0)
  Me.optRotator(sProgramOptions.nRotatorType).Value = True
  
  'FTP
  Me.txtFTPHTMLDir = sProgramOptions.strFTPHTMLDir
  Me.txtFTPImagesDir = sProgramOptions.strFTPImagesDir
  Me.txtFTPPassword = sProgramOptions.strFTPPassword
  Me.txtFTPServer = sProgramOptions.strFTPServer
  Me.txtFTPTemplate = sProgramOptions.strFTPHTMLTemplate
  Me.txtFTPUsername = sProgramOptions.strFTPUserName
  
  For i = 0 To 2000
    If sDxDetails(i).strName = "" And sDxDetails(i).Callsign = "" Then Exit For
    strTemp = sDxDetails(i).Callsign & "  " & sDxDetails(i).strName
    Me.cmbSecondObs.AddItem strTemp
  Next i

  Me.cmdApply.Enabled = False
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

Private Sub optRotator_Click(Index As Integer)
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtDefDatabasePath_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtDefKepsPath_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPHTMLDir_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPImagesDir_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPPassword_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPServer_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPTemplate_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtFTPUsername_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtGTPS_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtGTPS_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtHeight_Click()
  Me.cmdApply.Enabled = True
End Sub


Private Sub txtHeight_KeyPress(KeyAscii As Integer)
  'If Not bCheckNumNeg(KeyAscii, False) And KeyAscii <> c0166 Then KeyAscii = 0
End Sub

Private Sub txtKepsAge_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtKepsAge_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) And KeyAscii <> cSPACE And KeyAscii <> cPoint Then KeyAscii = 0
End Sub

Private Sub txtLatitude_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtLatitude_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) And KeyAscii <> cSPACE And KeyAscii <> cPoint And KeyAscii <> Asc("N") And KeyAscii <> Asc("S") Then KeyAscii = 0
End Sub

Private Sub txtLocation_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtLongitude_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtLongitude_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) And KeyAscii <> cSPACE And KeyAscii <> cPoint And KeyAscii <> Asc("E") And KeyAscii <> Asc("W") Then KeyAscii = 0
End Sub

Private Sub txtOrthLocations_Change()
  Me.cmdApply.Enabled = True
  If Me.txtOrthLocations.Text = "" Then
    strOrthLocs = ""
  End If
End Sub

Private Sub txtOrthLocations_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtOrthX_Click()
  Me.cmdApply.Enabled = True

End Sub

Private Sub txtOrthY_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtProxyIP_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtProxyPort_Change()
  Me.cmdApply.Enabled = True
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

Private Sub txtSpeech_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtTimeout_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtUpdate_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtUpdate_Click()
  Me.cmdApply.Enabled = True
End Sub

Private Sub txtUpdate_KeyPress(KeyAscii As Integer)
  KeyAscii = nToUpper(KeyAscii)
  If Not bcheckNum(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub UpDown1_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub UpDown2_Change()
  Me.cmdApply.Enabled = True
End Sub

Private Sub UpDown4_Change()
  Me.cmdApply.Enabled = True
End Sub
Public Function strValidateLatLon(frmSource As Form, ctrControl1 As Control, ctrControl2 As Control, bCheckLat As Boolean) As String
Dim strData As String
Dim l018C As Integer
Dim l018E As Integer
Dim l0190 As Integer
Dim l0192 As Integer
Dim l0194 As Integer
Dim l0196 As Integer
Dim l0198 As String
Dim l019A As String
Dim l019C As String
Dim l019E As String
Dim l01A0 As Integer
Dim l01A2 As String
Dim l01A4 As Double
Dim l01A6 As Double
Dim sLat As Double
Dim sLon As Double
Dim strLon As String
Dim strLat As String
Dim l01B0 As Integer
On Error GoTo ErrorHandler

strData = Trim(ctrControl1.Text)
If strData = "" Then
    strValidateLatLon = ""
    Exit Function
End If
If fn2060(strData, l01A4, l01A6) Then
    sub2028 l01A6, l01A4, sLat, sLon
    FormatLatLon CSng(sLat * 60), CSng(sLon * 60), strLon, strLat, True
    If bCheckLat Then
        ctrControl1 = strLat
        ctrControl2 = strLon
    Else
        ctrControl1 = strLon
        ctrControl2 = strLat
    End If
    strValidateLatLon = True
    Exit Function
End If
If bCheckLat Then
    l019A = "E"
    l019C = "W"
    l019E = "000"
    l01A0 = 180
    l01A2 = "longitude"
Else
    l019A = "N"
    l019C = "S"
    l019E = "00"
    l01A0 = 90
    l01A2 = "latitude"
End If
l0198 = Right(strData, 1)
If l0198 <> l019A And l0198 <> l019C Then GoTo ErrorHandler
strData = Left$(strData, Len(strData) - 1)
If strData = "" Then GoTo ErrorHandler
l018E = InStr(strData, ".")
If l018E <> 0 And l018E <> Len(strData) Then l0190 = InStr(l018E + 1, strData, ".")
l018C = InStr(strData, " ")
If l018C <> 0 Then
    If l018E = 0 Or l018C < l018E Then GoTo ErrorHandler
End If
If l018E = 0 Then
    l0192 = Val(strData)
Else
    l0192 = Val(Left$(strData, l018E - 1))
    If l0190 = 0 Then
        l0194 = Val(Mid$(strData, l018E + 1))
    Else
        l0194 = Val(Mid$(strData, l018E + 1, l0190 - l018E - 1))
        l0196 = Val(Mid$(strData, l0190 + 1))
    End If
End If
If l0192 > l01A0 Or l0192 < 0 Then GoTo ErrorHandler
If l0192 = l01A0 And (l0194 <> 0 Or l0196 <> 0) Then GoTo ErrorHandler
If l0194 < 0 Or l0194 > 59 Then GoTo ErrorHandler
If l0196 < 0 Or l0196 > 99 Then GoTo ErrorHandler
If bCheckLat Then
    l018C = 4
    l01B0 = 9
Else
    l018C = 3
    l01B0 = 8
End If
If Mid$(strData, l018C, 1) <> "." Or Mid$(strData, l018C + 3, 1) <> "." Or Len(strData) <> l01B0 Then strData = Format$(l0192, l019E) & "." & Format$(l0194, "00") & "." & Format$(l0196, "00")
strData = strData & l0198
ctrControl1.Text = strData
strValidateLatLon = "Ok"
Exit Function

ErrorHandler:
strValidateLatLon = l01A2
Exit Function
End Function
Public Function fn2060(p04CC As String, p04CE As Double, p04D0 As Double) As Integer
Dim l04D2 As Integer
Dim l04D4 As Integer
Dim l04D6 As Integer
Dim l04D8 As Integer
Dim l04DA As Integer
Dim l04DC As String
Dim l04DE As Integer
Dim l04E0 As Integer
Dim l04E2 As Integer
Dim l04E4 As Integer
Dim l04E6 As String
Dim l04E8 As String
For l04D2 = 1 To Len(p04CC)
    l04E8 = Mid(p04CC, l04D2, 1)
    If l04E8 <> " " Then l04E6 = l04E6 & l04E8
Next l04D2
If Len(l04E6) = 6 Then
    l04E6 = Left(l04E6, 4) & "5" & Right(l04E6, 2) & "5"
ElseIf Len(l04E6) <> 8 Then
    Exit Function
End If
l04D2 = Asc(UCase$(l04E6)) - Asc("A")
l04D4 = Asc(UCase$(Mid$(l04E6, 2))) - Asc("A")
If l04D2 < 0 Or l04D2 > 25 Or l04D2 = 8 Or l04D4 < 0 Or l04D4 > 25 Or l04D4 = 8 Then Exit Function
l04D2 = l04D2 + (l04D2 > 8)
l04D4 = l04D4 + (l04D4 > 8)
l04DA = 0
l04DC = ""
For l04D6 = 3 To Len(l04E6)
        l04D8 = Asc(Mid$(l04E6, l04D6))
        If l04D8 >= Asc("0") And l04D8 <= Asc("9") Then
                l04DA = l04DA + 1
                l04DC = l04DC + Mid$(l04E6, l04D6, 1)
        End If
Next l04D6
If l04DA <> 6 Then Exit Function
p04CE = Val(Right$(l04DC, 3))
p04D0 = Val(Left$(l04DC, 3))
l04DE = l04D2 \ 5: l04E2 = l04D2 - l04DE * 5
l04E0 = l04D4 \ 5: l04E4 = l04D4 - l04E0 * 5
p04CE = 100 * (p04CE + 1000 * (5 * (3 - l04DE) + (4 - l04E0)))
p04D0 = 100 * (p04D0 + 1000 * (5 * (l04E2 - 2) + l04E4))
p04CC = l04E6
fn2060 = True
End Function

Public Sub sub2028(p046E As Double, p0470 As Double, p0472 As Double, p0474 As Double)
Dim l0476 As Double
Dim l0478 As Double
Dim l047A As Double
Dim l047C As Double
Dim l047E As Double
Dim l0480 As Double
Dim l0482 As Double
Dim l0484 As Double
Dim l0486 As Double
Dim l0488 As Double
Dim l048A As Double
Dim l048C As Double
Dim l048E As Double
Dim l0490 As Double
Dim l0492 As Double
Dim l0494 As Double
Dim l0496 As Double
Dim l0498 As Double

l0476 = 49 * c04C2 / 180
l0478 = -2 * c04C2 / 180
l047A = c04AA * c04AA
l047C = l047A * c04AA
l047E = l0476
l0480 = 0
Do
        l047E = l047E + (p0470 - c04BE - l0480) / c049A
        l0482 = l047E - l0476
        l0484 = l047E + l0476
        l0480 = (1 + c04AA + 1.25 * l047A + 1.25 * l047C) * l0482
        l0480 = l0480 - (3 * c04AA + 3 * l047A + 2.625 * l047C) * sIn(l0482) * Cos(l0484)
        l0480 = l0480 + (1.875 * l047A + 1.875 * l047C) * sIn(2 * l0482) * Cos(2 * l0484)
        l0480 = l0480 - (35 / 24 * l047C) * sIn(3 * l0482) * Cos(3 * l0484)
        l0480 = l0480 * c04A2
Loop While (Abs(p0470 - c04BE - l0480) >= 0.001)
l0488 = sIn(l047E)
l048A = Cos(l047E)
l048C = l0488 / l048A
l048E = l048C * l048C
l0490 = c049A / Sqr(1 - c04B2 * l0488 * l0488)
l0492 = l0490 * (1 - c04B2) / (1 - c04B2 * l0488 * l0488)
l0494 = l0490 / l0492 - 1
l0496 = (p046E - c04BA) / l0490
l0498 = l0496 * l0496
l0482 = l048C / 2
l0484 = l048C / 24 * (5 + l048E * (3 - 9 * l0494) + l0494)
l0486 = l048C / 720 * (61 + l048E * (90 + l048E * 45))
p0472 = (l047E - l0490 / l0492 * (l0498 * (l0482 - l0498 * (l0484 - l0498 * l0486))))
l0482 = (l0490 / l0492 + l048E * 2) / 6
l0484 = (5 + l048E * (28 + l048E * 24)) / 120
l0486 = (61 + l048E * (662 + l048E * (1320 + l048E * 720))) / 5040
p0474 = (l0478 + l0496 / l048A * (1 - l0498 * (l0482 - l0498 * (l0484 - l0498 * l0486))))
p0472 = p0472 * 180 / c04C2: p0474 = p0474 * 180 / c04C2
End Sub

Public Sub FormatLatLon(sLat As Single, sLon As Single, strLat As String, strLon As String, bIncludePeriod As Boolean)
Dim sAbsLon As Single
Dim sAbsLat As Single
Dim strLonType As String
Dim strLatType As String
Dim sLonHours As Integer
Dim sLatHours As Integer
Dim strSep As String

If bIncludePeriod Then strSep = "."
If sLon < 0 Then strLonType = "W" Else strLonType = "E"
If sLat < 0 Then strLatType = "S" Else strLatType = "N"
sAbsLon = Abs(sLon)
sAbsLat = Abs(sLat)
sLonHours = Int(sAbsLon) \ 60
sAbsLon = sAbsLon - sLonHours * 60
sLatHours = Int(sAbsLat) \ 60
sAbsLat = sAbsLat - sLatHours * 60
strLon = Format$(sLonHours, "000") & strSep & Format$(sAbsLon, "00.00") & strLonType
strLat = Format$(sLatHours, "00") & strSep & Format$(sAbsLat, "00.00") & strLatType
Mid$(strLon, Len(strLon) - 3, 1) = "."
Mid$(strLat, Len(strLat) - 3, 1) = "."
End Sub
Public Function sConvertPos(strSource As String) As Single
Dim nDotPos As Integer
Dim strTempSource As String
Dim strType As String
Dim sPos As Single
strType = Right(strSource, 1)
If InStr("NSEW", strType) = 0 Then Exit Function
strTempSource = Left(strSource, Len(strSource) - 1)
nDotPos = InStr(strTempSource, ".")
If nDotPos = 0 Then
    sPos = Val(strTempSource) * 60
Else
    sPos = Val(Left(strTempSource, nDotPos - 1) * 60) + Val(Mid(strTempSource, nDotPos + 1))
End If
If strType = "S" Or strType = "W" Then sPos = -sPos
sConvertPos = sPos
End Function
