VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRotatorSetup 
   HelpContextID   =  1380
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Rotator"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmRotatorSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3930
      TabIndex        =   2
      Top             =   630
      Width           =   825
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3930
      TabIndex        =   1
      Top             =   180
      Width           =   825
   End
   Begin MSComctlLib.ListView lstSats 
      Height          =   2565
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4524
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Satellite Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Norad Id"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keps Age"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRotatorSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  bMoveRotator = False
  Set frmRotatorForm = Nothing
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim i As Integer
  Dim bchecked As Boolean
  Dim strDesig As String
  
  bchecked = False
  For i = 1 To Me.lstSats.ListItems.count
    If Me.lstSats.ListItems(i).Checked And bchecked Then
      MsgBox "Please select ONE satellite to track antennas with", vbCritical + vbOKOnly, "Rotator Setup Error"
      bchecked = False
      strDesig = ""
      Exit For
    End If
    If Me.lstSats.ListItems(i).Checked Then
      bchecked = True
      strDesig = Me.lstSats.ListItems(i).SubItems(1)
    End If
  Next i
  
  If bchecked Then
    frmRotatorForm.Tag = strDesig
    frmRotatorForm.OpenRotatorLink
    Unload Me
  End If
  
End Sub

Private Sub Form_Load()
  Dim sItem As ListItem
  Dim i As Integer
  Dim nCentury As Integer
  Dim vEpochTime As Variant
  Dim vEpochYear As Variant
  Dim vEpochDate As Variant
  Dim strEpochDate As String
  
  With Me.lstSats
    For i = 1 To frmRotatorForm.ocxSat.SatelliteCount
      Set sItem = .ListItems.Add
      frmRotatorForm.ocxSat.SatelliteIndex = i
      sItem.Text = frmRotatorForm.ocxSat.SatelliteName
      sItem.SubItems(1) = frmRotatorForm.ocxSat.SatelliteDesignator
      strEpochDate = frmRotatorForm.ocxSat.KepsEpochTime
      If Val(Left$(strEpochDate, 2)) < 50 Then
        nCentury = 20
      Else
        nCentury = 19
      End If
      vEpochTime = Val(strEpochDate) - 1000 * Int(Val(strEpochDate) / 1000)
      vEpochYear = 100 * nCentury + Int(Val(strEpochDate) / 1000)
      vEpochDate = Format(vEpochTime, "dd/mm/") & vEpochYear
      vEpochDate = CDate(vEpochDate)
      sItem.SubItems(2) = vEpochDate
      If sItem.SubItems(1) = frmRotatorForm.Tag Then
        sItem.Checked = True
      End If
      Set sItem = Nothing
    Next i
  End With
  
End Sub
