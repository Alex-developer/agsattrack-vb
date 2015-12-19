VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Available satellites from"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   HelpContextID   =   35
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Caption         =   " Element Database "
      Height          =   1275
      Left            =   4260
      TabIndex        =   16
      Top             =   4230
      Width           =   1935
      Begin VB.ComboBox cmbFiles 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         WhatsThisHelpID =   5060
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Elements Age Details "
      Height          =   1275
      Left            =   60
      TabIndex        =   8
      Top             =   4230
      WhatsThisHelpID =   5080
      Width           =   4125
      Begin VB.Label lblYoungest 
         Caption         =   "Label2"
         Height          =   225
         Left            =   1500
         TabIndex        =   15
         Top             =   840
         WhatsThisHelpID =   5080
         Width           =   405
      End
      Begin VB.Label lblOldest 
         Caption         =   "Label2"
         Height          =   225
         Left            =   1500
         TabIndex        =   14
         Top             =   555
         WhatsThisHelpID =   5080
         Width           =   405
      End
      Begin VB.Label lblwarn 
         Caption         =   "The elements should be updated as they exceed the maximum age you have specified in the program options."
         ForeColor       =   &H000000FF&
         Height          =   1035
         Left            =   2010
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         WhatsThisHelpID =   5080
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Youngest age"
         Height          =   225
         Left            =   210
         TabIndex        =   12
         Top             =   840
         WhatsThisHelpID =   5080
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Oldest age"
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   570
         WhatsThisHelpID =   5080
         Width           =   975
      End
      Begin VB.Label lblAverage 
         Caption         =   "Label2"
         Height          =   225
         Left            =   1500
         TabIndex        =   10
         Top             =   270
         WhatsThisHelpID =   5080
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Average age"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   270
         WhatsThisHelpID =   5080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdInetUpdate 
      Caption         =   "Update"
      Height          =   345
      Left            =   4260
      TabIndex        =   7
      Top             =   2400
      WhatsThisHelpID =   5050
      Width           =   795
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details"
      Height          =   345
      Left            =   4260
      TabIndex        =   6
      Top             =   1995
      WhatsThisHelpID =   5040
      Width           =   795
   End
   Begin VB.FileListBox lstFiles 
      Height          =   480
      Left            =   4680
      TabIndex        =   5
      Top             =   2430
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView lstSatellites 
      Height          =   4065
      Left            =   60
      TabIndex        =   4
      Top             =   120
      WhatsThisHelpID =   5070
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7170
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NORAD ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keps Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5340
      Top             =   2940
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   345
      Left            =   4260
      TabIndex        =   3
      Top             =   1590
      WhatsThisHelpID =   5030
      Width           =   795
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   345
      Left            =   4260
      TabIndex        =   2
      Top             =   525
      WhatsThisHelpID =   5010
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   930
      WhatsThisHelpID =   5020
      Width           =   795
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   345
      Left            =   4260
      TabIndex        =   0
      Top             =   120
      WhatsThisHelpID =   5000
      Width           =   795
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCancelled As Boolean
Dim strTempDatabase As String


Private Sub cmbFiles_Click()
  
  On Error GoTo ERROR_cmbFiles_Click

 '   fMainForm.ActiveForm.ocxSat.DatabasePath = ""
    strTempDatabase = App.Path & "\Elements\" & Me.cmbFiles.Text
    ReadKeps strTempDatabase
    UpdateList False
    cmdApply.Enabled = False

EXIT_cmbFiles_Click:
  Exit Sub

ERROR_cmbFiles_Click:
  MsgBox "Error in ERROR_cmbFiles_Click : " & Error
  Resume EXIT_cmbFiles_Click

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdApply_Click()

  On Error GoTo ERROR_cmdApply_Click

  Dim i As Integer
  Dim SatNum As Integer
  Dim lvCount As Long
  Dim lvIndex As Long
  Dim strSatDesignator As String
  Dim lNode As ListItem
  Dim lCounter As Integer
  Dim l1 As String
  Dim l2 As String
  Dim l3 As String
  Dim bResult As Boolean
  Dim FormattedDateTime As String
  
  
  fMainForm.ActiveForm.ocxSat.EraseSatellites
  
  fMainForm.ActiveForm.ocxSat.DatabasePath = strTempDatabase
  
  lvCount = lstSatellites.ListItems.Count
  lvIndex = 1
  lCounter = 1
  
  Do
    If lstSatellites.ListItems(lvIndex).Checked Then
      strSatDesignator = lstSatellites.ListItems(lvIndex).Text
      For i = 0 To 500
        If sKeps(i).lDesignator = Val(strSatDesignator) Then
          fMainForm.ActiveForm.ocxSat.AddSatellite
          UpdateOcxKeps lCounter, i, fMainForm.ActiveForm
          lCounter = lCounter + 1
        End If
      Next i
    End If
    lvIndex = lvIndex + 1
  Loop Until lvIndex > lvCount Or lCounter > 20

  For lCounter = -1 To 0
    fMainForm.ActiveForm.ocxSat.SatelliteIndex = lCounter
    FormattedDateTime$ = Format$(Now, "yyyymmddhhmmss")
    fMainForm.ActiveForm.ocxSat.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
    fMainForm.ActiveForm.ocxSat.DisplayYear = Val(Left$(FormattedDateTime$, 4))
    fMainForm.ActiveForm.ocxSat.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
    fMainForm.ActiveForm.ocxSat.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
    fMainForm.ActiveForm.ocxSat.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
    fMainForm.ActiveForm.ocxSat.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
    fMainForm.ActiveForm.ocxSat.DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
  Next lCounter

  fMainForm.ActiveForm.ocxSat.CalculateALLPositions
  fMainForm.ActiveForm.ocxSat.DrawFootprints

  fMainForm.ActiveForm.ocxSat.SetSelectedSatellite = IIf(lCounter > 1, 1, 0)
  
  cmdApply.Enabled = False

  If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
    fMainForm.ActiveForm.UpdateSatelliteData
  End If
  
EXIT_cmdApply_Click:
  Exit Sub

ERROR_cmdApply_Click:
  MsgBox "Error in ERROR_cmdApply_Click : " & Error
  Resume EXIT_cmdApply_Click

End Sub

Sub UpdateOcxKeps(lCounter As Integer, i As Integer, fForm As Form)
  Dim FormattedDateTime As String

  With fForm
'    .ocxSat.SatelliteIndex = lCounter
'    .ocxSat.SatelliteDesignator = sKeps(i).lDesignator
'    .ocxSat.SatelliteName = Trim(sKeps(i).strName)
'    .ocxSat.KepsEpochTime = sKeps(i).strEpoch
'    .ocxSat.KepsDecayRate = sKeps(i).dDrag
'    .ocxSat.KepsOrbitNumber = sKeps(i).lRevolutionnumber
'    .ocxSat.KepsInclination = sKeps(i).dInclination
'    .ocxSat.KepsRAAN = sKeps(i).dRAAN
'    .ocxSat.KepsEccentricity = sKeps(i).dEccentricity
'    .ocxSat.KepsAOP = sKeps(i).dAOP
'    .ocxSat.KepsMeanAnomoly = sKeps(i).dMeanAnomoly
'    .ocxSat.KepsMeanMotion = sKeps(i).dMeanMotion
    .ocxSat.UpdateKeps sKeps(i).strLine1, sKeps(i).strLine2, sKeps(i).strLine3, lCounter, agplan13
    '          .ocxSat.SatelliteTXFrequency = sKeps(i).dModeABeacon
    .ocxSat.SatelliteIndex = lCounter
    
    FormattedDateTime$ = Format$(Now, "yyyymmddhhmmss")
    .ocxSat.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
    .ocxSat.DisplayYear = Val(Left$(FormattedDateTime$, 4))
    .ocxSat.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
    .ocxSat.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
    .ocxSat.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
    .ocxSat.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
    .ocxSat.DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))

    .ocxSat.DisplayDataFields = "1,2,-1,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,31,32,33"
  End With
End Sub
Private Sub cmdClear_Click()

  On Error GoTo ERROR_cmdClear_Click

  Dim i As Integer
  
  With Me.lstSatellites
    For i = 1 To .ListItems.Count
      .ListItems(i).Checked = False
    Next i
  End With
  cmdApply.Enabled = True

EXIT_cmdClear_Click:
  Exit Sub

ERROR_cmdClear_Click:
  MsgBox "Error in ERROR_cmdClear_Click : " & Error
  Resume EXIT_cmdClear_Click

End Sub

Private Sub cmdDetails_Click()

  On Error GoTo ERROR_cmdDetails_Click

  If Not (Me.lstSatellites.SelectedItem Is Nothing) Then
  
    frmSatDetails.Tag = Me.lstSatellites.SelectedItem.Tag
    frmSatDetails.Show vbModal
  End If

EXIT_cmdDetails_Click:
  Exit Sub

ERROR_cmdDetails_Click:
  MsgBox "Error in ERROR_cmdDetails_Click : " & Error
  Resume EXIT_cmdDetails_Click

End Sub

Private Sub cmdInetUpdate_Click()
  frmFTPMain.Show vbModal
  ReadKeps strTempDatabase
  UpdateList True
  cmdApply.Enabled = True
End Sub

Private Sub cmdOk_Click()

  On Error GoTo ERROR_cmdOk_Click

  If cmdApply.Enabled Then
    cmdApply_Click
  End If
  Unload Me

EXIT_cmdOk_Click:
  Exit Sub

ERROR_cmdOk_Click:
  MsgBox "Error in ERROR_cmdOk_Click : " & Error
  Resume EXIT_cmdOk_Click

End Sub

Private Sub Form_Load()

  'On Error GoTo ERROR_Form_Load

  CenterForm Me
  strTempDatabase = fMainForm.ActiveForm.ocxSat.DatabasePath
  If strTempDatabase = "" Then
    If FindFile(sProgramOptions.strKepsDatabase) Then
      strTempDatabase = sProgramOptions.strKepsDatabase
    Else
      Call MsgBox("Unable to locate Keplarian elements database file. Resetting default.", vbInformation + vbOKOnly + vbDefaultButton1, "Keps Database error")
      sProgramOptions.strKepsDatabase = App.Path & "\Elements\Amateur.txt"
      strTempDatabase = sProgramOptions.strKepsDatabase
    End If
  End If
  lstFiles.Path = App.Path & "\Elements"
  lstFiles.Pattern = "*.txt;*.tle"
  ReadKeps strTempDatabase
  UpdateList True
  cmdApply.Enabled = False

EXIT_Form_Load:
  Exit Sub

ERROR_Form_Load:
  MsgBox "Error in ERROR_Form_Load : " & Error
  Resume EXIT_Form_Load

End Sub

Private Sub UpdateList(bSetCombo As Boolean)

'  On Error GoTo ERROR_UpdateList

  Dim lNode As ListItem
  Dim i As Integer
  Dim nCentury As Integer
  Dim vEpochTime As Variant
  Dim vEpochYear As Variant
  Dim vEpochDate As Variant
  Dim lAverage As Long
  Dim vDate As Variant
  Dim vOldest As Variant
  Dim vYoungest As Variant

'  If fMainForm.ActiveForm.ocxSat.DatabasePath <> "" And fMainForm.ActiveForm.ocxSat.DatabasePath <> 0 Then
'    strTempDatabase = fMainForm.ActiveForm.ocxSat.DatabasePath
'  Else
'    strTempDatabase = sProgramOptions.strKepsDatabase
'  End If
  
  Me.Caption = "Satellites availble from " & GetFilenameNoExt(GetFilename(strTempDatabase))
  If bSetCombo Then
    Me.cmbFiles.Clear
    For i = 0 To lstFiles.ListCount
      Me.cmbFiles.AddItem Me.lstFiles.List(i)
    Next i
    Me.cmbFiles.Text = GetFilename(strTempDatabase)
  End If
  Me.lstSatellites.ListItems.Clear
    
  vDate = Date
  vOldest = Format(Date, "dd-mm-yyyy")
 ' vYoungest = Format(Date, "dd-mm-yyyy")
  For i = 0 To 500
  
    If sKeps(i).strName = "" Then Exit For
        
      Set lNode = lstSatellites.ListItems.Add(, "ID" & Str$(sKeps(i).lDesignator) & Str(i), sKeps(i).lDesignator)

    lNode.SubItems(1) = Trim(sKeps(i).strName)
    lNode.Tag = i
    
    If Val(Left$(sKeps(i).strEpoch, 2)) < 50 Then
      nCentury = 20
    Else
      nCentury = 19
    End If
      vEpochTime = Val(sKeps(i).strEpoch) - 1000 * Int(Val(sKeps(i).strEpoch) / 1000)
      vEpochYear = 100 * nCentury + Int(Val(sKeps(i).strEpoch) / 1000)
      vEpochDate = Format(vEpochTime, "dd/mm/") & vEpochYear
      vEpochDate = CDate(vEpochDate)
      lNode.SubItems(2) = vEpochDate
      lAverage = lAverage + DateDiff("d", vEpochDate, vDate)
      If fMainForm.ActiveForm.ocxSat.IsSatLoaded(sKeps(i).lDesignator) Then
        lNode.Checked = True
      End If
      If vEpochDate < vOldest Then
        vOldest = vEpochDate
      End If
      If vEpochDate > vYoungest Then
        vYoungest = vEpochDate
      End If
  Next i
  lAverage = lAverage / i

  If lAverage > sProgramOptions.nKepsAge Then
    Me.lblAverage.ForeColor = RGB(255, 0, 0)
    Me.lblwarn.Visible = True
  Else
    Me.lblAverage.ForeColor = RGB(0, 0, 0)
    Me.lblwarn.Visible = False
  End If
  
  Me.lblAverage.Caption = lAverage
  Me.lblOldest = DateDiff("d", vOldest, vDate)
  Me.lblYoungest = DateDiff("d", vYoungest, vDate)
EXIT_UpdateList:
  Exit Sub

ERROR_UpdateList:
  MsgBox "Error in ERROR_UpdateList : " & Error
  Resume EXIT_UpdateList

End Sub


Private Sub lstSatellites_Click()
  cmdApply.Enabled = True
End Sub

Private Sub lstSatellites_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  lstSatellites.SortKey = ColumnHeader.Index - 1
  If Me.lstSatellites.SortOrder = lvwAscending Then
    Me.lstSatellites.SortOrder = lvwDescending
  Else
    Me.lstSatellites.SortOrder = lvwAscending
  End If

End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub
