VERSION 5.00
Begin VB.Form frmElements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keplarian Elements"
   ClientHeight    =   5535
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFields 
      DataField       =   "Epoch"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   35
      Top             =   720
      Width           =   1635
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Drag"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   34
      Top             =   1020
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Revolutionnumber"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   33
      Top             =   1340
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Inclination"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   32
      Top             =   1660
      Width           =   1335
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   8040
      TabIndex        =   26
      Top             =   4890
      Width           =   8040
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4505
         TabIndex        =   31
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   3409
         TabIndex        =   30
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   2313
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1217
         TabIndex        =   28
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   121
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\DEVELOP\Visual Basic\Radio\Satellite Tracking\Elements\Amateur.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [Elements] Order by [Designator]"
      Top             =   5190
      Width           =   8040
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ModeABeacon"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   25
      Top             =   4540
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ModeADownlink"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   23
      Top             =   4220
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ModeAUplink"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   21
      Top             =   3900
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Checksum"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   19
      Top             =   3580
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MeanMotion"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   17
      Top             =   3260
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MeanAnomoly"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   15
      Top             =   2940
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "AOP"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   13
      Top             =   2620
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Eccentricity"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   11
      Top             =   2300
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RAAN"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   9
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Name"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   380
      Width           =   1635
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Designator"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Caption         =   "ModeABeacon:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   24
      Top             =   4540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ModeADownlink:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   22
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ModeAUplink:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   20
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Checksum:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   18
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MeanMotion:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MeanAnomoly:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   14
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "AOP:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   12
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Eccentricity:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "RAAN:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Inclination:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Revolutionnumber:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Drag:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Epoch:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Designator:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
  datPrimaryRS.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  datPrimaryRS.Refresh
End Sub

Private Sub cmdUpdate_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'Throw away the error
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)

End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
datPrimaryRS.DatabaseName = sProgramOptions.strKepsDatabase

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

