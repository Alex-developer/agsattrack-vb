VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Tag             =   "Apply"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   10
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   9
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   210
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame3 
         Caption         =   " Position "
         Height          =   2055
         Left            =   60
         TabIndex        =   22
         Top             =   60
         Width           =   4395
         Begin VB.TextBox txtHeight 
            Height          =   285
            Left            =   1080
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   1040
            Width           =   555
         End
         Begin VB.TextBox txtLocation 
            Height          =   285
            Left            =   1080
            TabIndex        =   28
            Text            =   "Text2"
            Top             =   1440
            Width           =   1875
         End
         Begin VB.TextBox txtLongitude 
            Height          =   285
            Left            =   1080
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   640
            Width           =   555
         End
         Begin VB.TextBox txtLatitude 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "Height"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label4 
            Caption         =   "Location"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   1500
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Longitude"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Latitude"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   300
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame2 
         Caption         =   "Display Type"
         Height          =   1185
         Left            =   3300
         TabIndex        =   18
         Top             =   1140
         Width           =   1755
         Begin VB.OptionButton optDisplayType 
            Caption         =   "Horizon"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   780
            Width           =   1395
         End
         Begin VB.OptionButton optDisplayType 
            Caption         =   "Map"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   510
            Width           =   1395
         End
         Begin VB.OptionButton optDisplayType 
            Caption         =   "Table"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Center"
         Height          =   975
         Left            =   3300
         TabIndex        =   15
         Top             =   60
         Width           =   1755
         Begin VB.OptionButton optMapCenter 
            Caption         =   "180 Degrees"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optMapCenter 
            Caption         =   "0 Degrees"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame fraSample1 
         Caption         =   "Images"
         Height          =   2265
         Left            =   60
         TabIndex        =   4
         Tag             =   "Sample 1"
         Top             =   60
         Width           =   3120
         Begin VB.CommandButton cmdLoadNew 
            Caption         =   "Horizon"
            Height          =   435
            Index           =   2
            Left            =   180
            TabIndex        =   31
            Top             =   1620
            Width           =   795
         End
         Begin VB.CommandButton cmdLoadNew 
            Caption         =   "Load 180"
            Height          =   435
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   960
            Width           =   795
         End
         Begin VB.CommandButton cmdLoadNew 
            Caption         =   "Reset"
            Height          =   435
            Index           =   3
            Left            =   1260
            TabIndex        =   12
            Top             =   1620
            Width           =   795
         End
         Begin VB.CommandButton cmdLoadNew 
            Caption         =   "Load 0"
            Height          =   435
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Select a button to load a new bitmap image for the maps. The Reset button will return the maps to their defaults"
            Height          =   1215
            Left            =   1260
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Map"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Observer"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Group 3"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Group 4"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMapFile As String
Dim nMapType As Integer

Private Sub cmdApply_Click()
  'Dim Form As Form
  
  If optMapCenter(0).Value = True Then
    fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0
  Else
    fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 180
  End If

  If strMapFile <> "" Then
    fMainForm.ActiveForm.ocxSat.SetMap nMapType, strMapFile
  End If

  If optDisplayType(0).Value = True Then
    fMainForm.ActiveForm.ocxSat.OutputStyle = 0
  Else
    If optDisplayType(1).Value = True Then
      fMainForm.ActiveForm.ocxSat.OutputStyle = 1
    Else
      fMainForm.ActiveForm.ocxSat.OutputStyle = 2
    End If
  End If
' Tab 2 LOcation

  'sProgramOptions.nLatitude = Me.txtLatitude
  'sProgramOptions.nLongitude = Me.txtLongitude
  'sProgramOptions.strLocation = Me.txtLocation
  'sProgramOptions.nHeight = Me.txtHeight
  
  fMainForm.ActiveForm.ocxSat.ObserverLatitude = Me.txtLatitude
  fMainForm.ActiveForm.ocxSat.ObserverLongitude = Me.txtLongitude
  fMainForm.ActiveForm.ocxSat.ObserverHeight = Me.txtHeight
  fMainForm.ActiveForm.ocxSat.ObserverLocation = Me.txtLocation
  fMainForm.ActiveForm.ocxSat.PlotObserver
  
  cmdApply.Enabled = False

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdLoadNew_Click(Index As Integer)
  Dim sFile As String

  If Index <> 3 Then
    With CommonDialog1
      .Filter = "Bitmaps *.bmp|*.bmp|All Files (*.*)|*.*"
      .ShowOpen
      If Len(.FileName) = 0 Then
        Exit Sub
      End If
      sFile = .FileName
      strMapFile = .FileName
      nMapType = Index
      cmdApply.Enabled = True
    End With
  Else
    strMapFile = "Reset"
    nMapType = 0
    cmdApply.Enabled = True
  End If

End Sub

Private Sub cmdOk_Click()
  If cmdApply.Enabled = True Then
    cmdApply_Click
  End If
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  i = tbsOptions.SelectedItem.Index
  'handle ctrl+tab to move to the next tab
  If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
    If i = tbsOptions.Tabs.Count Then
      'last tab so we need to wrap to tab 1
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
    Else
      'increment the tab
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
    End If
  ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
    If i = 1 Then
      'last tab so we need to wrap to tab 1
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
    Else
      'increment the tab
      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
    End If
  End If
End Sub

Private Sub Form_Load()

  If fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0 Then
    optMapCenter(0).Value = True
  Else
    optMapCenter(1).Value = True
  End If

  optDisplayType(fMainForm.ActiveForm.ocxSat.OutputStyle).Value = True

  strMapFile = ""
' Tab 2 Location
  Me.txtLatitude = fMainForm.ActiveForm.ocxSat.ObserverLatitude
  Me.txtLongitude = fMainForm.ActiveForm.ocxSat.ObserverLongitude
  Me.txtHeight = fMainForm.ActiveForm.ocxSat.ObserverHeight
  Me.txtLocation = fMainForm.ActiveForm.ocxSat.ObserverLocation

  cmdApply.Enabled = False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub optDisplayType_Click(Index As Integer)
  cmdApply.Enabled = True
End Sub

Private Sub optMapCenter_Click(Index As Integer)
  cmdApply.Enabled = True
End Sub

Private Sub tbsOptions_Click()

  Dim i As Integer
  'show and enable the selected tab's controls
  'and hide and disable all others
  For i = 0 To tbsOptions.Tabs.Count - 1
    If i = tbsOptions.SelectedItem.Index - 1 Then
      picOptions(i).Left = 210
      picOptions(i).Enabled = True
    Else
      picOptions(i).Left = -20000
      picOptions(i).Enabled = False
    End If
  Next

End Sub

Private Sub txtHeight_Click()
  cmdApply.Enabled = True

End Sub

Private Sub txtLatitude_Click()
  cmdApply.Enabled = True
End Sub

Private Sub txtLocation_Click()
  cmdApply.Enabled = True

End Sub

Private Sub txtLongitude_Click()
  cmdApply.Enabled = True

End Sub
