VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEditElementGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Internet Update Groups"
   ClientHeight    =   5850
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8955
   Icon            =   "frmEditElementGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " Groups "
      Height          =   1305
      Left            =   120
      TabIndex        =   10
      Top             =   3930
      Width           =   2625
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   345
         Left            =   1770
         TabIndex        =   17
         Top             =   780
         Width           =   765
      End
      Begin VB.CommandButton cmdDelGroup 
         Caption         =   "Delete"
         Height          =   345
         Left            =   930
         TabIndex        =   14
         Top             =   780
         Width           =   765
      End
      Begin VB.CommandButton cmdNewGroup 
         Caption         =   "Add"
         Height          =   345
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   765
      End
      Begin VB.TextBox txtGroupName 
         Height          =   315
         Left            =   720
         TabIndex        =   12
         ToolTipText     =   "Enter the name of a new group and click add"
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Details "
      Height          =   1305
      Left            =   2790
      TabIndex        =   6
      Top             =   3930
      Width           =   6015
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "select the element set to add to the current group"
         Top             =   240
         Width           =   2685
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Left            =   4890
         TabIndex        =   9
         ToolTipText     =   "Removes the selected element set from the current group"
         Top             =   720
         Width           =   885
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   345
         Left            =   4890
         TabIndex        =   8
         ToolTipText     =   "Adds the selected element set to the current group"
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   1170
         TabIndex        =   7
         Top             =   270
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lstDetails 
      Height          =   3570
      Left            =   2790
      TabIndex        =   5
      ToolTipText     =   "The elemen sets in the current selected group"
      Top             =   300
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   6297
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Server"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox lstGroups 
      Height          =   3570
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "The available element groups"
      Top             =   300
      Width           =   2595
   End
   Begin VB.FileListBox lstFiles 
      Height          =   675
      Left            =   6480
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7050
      TabIndex        =   1
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   5340
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Group Details"
      Height          =   255
      Left            =   2790
      TabIndex        =   16
      Top             =   30
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Groups"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   30
      Width           =   975
   End
End
Attribute VB_Name = "frmEditElementGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim lNode As ListItem
Dim strData(100) As String
Dim bUpdated As Boolean
Dim nSelected As Integer

Private Sub CancelButton_Click()
  Unload Me
End Sub


Private Sub cmdAdd_Click()

  On Error GoTo ERROR_cmdAdd_Click

  Dim nPos As Integer
  Dim vData() As Variant
  Dim strLine As String
  Dim i As Integer
  Dim bGotIt As Boolean
  
  nPos = Me.cmbNames.ItemData(Me.cmbNames.ListIndex)
  For i = 1 To Me.lstDetails.ListItems.Count
    If Me.lstDetails.ListItems.Item(i).Text = Me.cmbNames.List(Me.cmbNames.ListIndex) Then
      bGotIt = True
      Exit For
    End If
  Next i
  
  If Not bGotIt Then
    strLine = strData(nPos)
    vData = StrParse(strLine, ",")
    Set lNode = Me.lstDetails.ListItems.Add
    lNode.Text = vData(0)
    lNode.SubItems(1) = vData(1)
    lNode.SubItems(2) = vData(2)
    lNode.SubItems(3) = vData(3)
    bUpdated = True
  Else
    Call MsgBox("The element set you have selected is already in this group.", vbInformation + vbOKOnly + vbDefaultButton1, "Error")
  End If

EXIT_cmdAdd_Click:
  Exit Sub

ERROR_cmdAdd_Click:
  MsgBox "Error in ERROR_cmdAdd_Click : " & Error
  Resume EXIT_cmdAdd_Click

End Sub

Private Sub cmdDelete_Click()

  On Error GoTo ERROR_cmdDelete_Click

  If Not (Me.lstDetails.SelectedItem Is Nothing) Then
    Me.lstDetails.ListItems.Remove (Me.lstDetails.SelectedItem.Index)
    bUpdated = True
  End If

EXIT_cmdDelete_Click:
  Exit Sub

ERROR_cmdDelete_Click:
  MsgBox "Error in ERROR_cmdDelete_Click : " & Error
  Resume EXIT_cmdDelete_Click

End Sub

Private Sub cmdDelGroup_Click()

  On Error GoTo ERROR_cmdDelGroup_Click

  Dim strFilename As String
  
  If MsgBox("Are you sure that you wish to delete this group. If you select yes then the group will be permanently deleted.", vbQuestion + vbYesNo + vbDefaultButton1, "Delete Group") = vbYes Then
    strFilename = App.Path & "\Internet Updates\" & Me.lstGroups.List(Me.lstGroups.ListIndex)
    Kill strFilename
    Form_Load
  End If

EXIT_cmdDelGroup_Click:
  Exit Sub

ERROR_cmdDelGroup_Click:
  MsgBox "Error in ERROR_cmdDelGroup_Click : " & Error
  Resume EXIT_cmdDelGroup_Click

End Sub

Private Sub cmdNewGroup_Click()
  On Error GoTo ERROR_cmdNewGroup_Click
  Dim nFile As Integer
  Dim strName As String
  
  If Me.txtGroupName.Text = "" Then
    Call MsgBox("Please enter a valid group name.", vbExclamation + vbOKOnly + vbDefaultButton1, "Invalid Group Name")
    Me.txtGroupName.SetFocus
  Else
    strName = Me.txtGroupName.Text
    If UCase(Right(strName, 4)) <> ".DAT" Then
      strName = strName & ".dat"
      Me.txtGroupName.Text = strName
    End If
    nFile = FreeFile
    Open App.Path & "\Internet updates\" & Me.txtGroupName.Text For Output As #nFile
    Print #nFile, ""
    Close #nFile
    UpdateGroups
  End If

EXIT_cmdNewGroup_Click:
  Exit Sub

ERROR_cmdNewGroup_Click:
  Select Case Err
    Case 75
      Call MsgBox("Please enter a valid group name.", vbExclamation + vbOKOnly + vbDefaultButton1, "Invalid Group Name")
      Me.txtGroupName.Text = ""
      Me.txtGroupName.SetFocus
    Case Else
    MsgBox "Error in ERROR_cmdNewGroup_Click : " & Error
  End Select
  Resume EXIT_cmdNewGroup_Click
  
End Sub

Private Sub cmdRename_Click()

  On Error GoTo ERROR_cmdRename_Click

  Dim strSource As String
  Dim strDest As String
  Dim strFilename As String
  
  If Me.txtGroupName <> "" Then
    strFilename = Me.txtGroupName.Text
    If UCase(Right(strFilename, 4)) <> ".DAT" Then
      strFilename = strFilename & ".dat"
    End If
    strSource = App.Path & "\Internet Updates\" & Me.lstGroups.List(Me.lstGroups.ListIndex)
    strDest = App.Path & "\Internet Updates\" & strFilename
    If FileExists(strDest) Then
      Call MsgBox("The file you are trying to rename to already exists. Please select another name.", vbExclamation + vbOKOnly + vbDefaultButton1, "Cannot rename file")
    Else
      FileCopy strSource, strDest
      Kill strSource
      UpdateGroups
    End If
  End If

EXIT_cmdRename_Click:
  Exit Sub

ERROR_cmdRename_Click:
  Select Case Err
    Case 75
      Call MsgBox("the filename you have selected is invalid.", vbExclamation + vbOKOnly + vbDefaultButton1, "Invalid file operation")
    Case Else
      MsgBox "Error in ERROR_cmdRename_Click : " & Error
      Resume EXIT_cmdRename_Click
  End Select
End Sub

Private Sub Form_Load()

  On Error GoTo ERROR_Form_Load

  Dim i As Integer
  Dim nFile As Integer
  Dim strLine As String
  Dim vData() As Variant
  
  nSelected = -1
  UpdateGroups
  
  nFile = FreeFile
  i = 0
  Open App.Path & "\Internet Updates\List.src" For Input As #nFile
  While Not (EOF(nFile))
    Line Input #nFile, strLine
    strData(i) = strLine
    vData = StrParse(strLine, ",")
    Me.cmbNames.AddItem vData(0)
    Me.cmbNames.ItemData(Me.cmbNames.NewIndex) = i
    i = i + 1
  Wend
  Close #nFile

EXIT_Form_Load:
  Exit Sub

ERROR_Form_Load:
  MsgBox "Error in ERROR_Form_Load : " & Error
  Resume EXIT_Form_Load

End Sub

Private Sub UpdateGroups()

  On Error GoTo ERROR_UpdateGroups

  Dim i As Integer
  
  lstFiles.Path = App.Path & "\Internet Updates"
  lstFiles.Pattern = "*.daa"
  lstFiles.Pattern = "*.dat"

  Me.lstGroups.Clear
  For i = 0 To lstFiles.ListCount
    If Me.lstFiles.List(i) <> "" Then
      Me.lstGroups.AddItem Me.lstFiles.List(i)
    End If
  Next i
  Me.lstGroups.ListIndex = 0

EXIT_UpdateGroups:
  Exit Sub

ERROR_UpdateGroups:
  MsgBox "Error in ERROR_UpdateGroups : " & Error
  Resume EXIT_UpdateGroups

End Sub
Private Sub UpdateList(nPos As Integer)

  On Error GoTo ERROR_UpdateList

  Dim strFilename As String
  Dim nFile As Integer
  Dim strLine As String
  Dim vData() As Variant
  
  strFilename = lstFiles.Path & "\" & lstFiles.List(nPos)
  Me.lstDetails.ListItems.Clear
  
  nFile = FreeFile
  Open strFilename For Input As #nFile
  While Not EOF(nFile)
    Line Input #nFile, strLine
    If strLine <> "" Then
    vData = StrParse(strLine, ",")
      Set lNode = Me.lstDetails.ListItems.Add
      lNode.Text = vData(0)
      lNode.SubItems(1) = vData(1)
      lNode.SubItems(2) = vData(2)
      lNode.SubItems(3) = vData(3)
    End If
  Wend
  Close #nFile

EXIT_UpdateList:
  Exit Sub

ERROR_UpdateList:
  MsgBox "Error in ERROR_UpdateList : " & Error
  Resume EXIT_UpdateList

End Sub

Private Sub lstGroups_Click()
  SelectionChanged False
End Sub

Private Sub SelectionChanged(bForce As Boolean)

  On Error GoTo ERROR_SelectionChanged

  Dim nNewSelected As Integer
  Dim bOk As Boolean
  Dim bSave As Boolean
  Dim nResult As Integer
  
  nNewSelected = Me.lstGroups.ListIndex
  If nNewSelected <> nSelected Or bForce Then
    If bUpdated Then
      nResult = MsgBox("Do you want to save the changes to group " & Me.lstGroups.List(nSelected), vbInformation + vbYesNoCancel + vbDefaultButton1, "Save Changes")
      Select Case nResult
        Case vbYes
          bSave = True
          bOk = True
        Case vbNo
          bSave = False
          bOk = True
        Case vbCancel
          bOk = False
      End Select
    Else
      bOk = True
    End If
    If bOk Then
      If bSave Then
        SaveGroup
      End If
      UpdateList nNewSelected
      nSelected = Me.lstGroups.ListIndex
    End If
  End If

EXIT_SelectionChanged:
  Exit Sub

ERROR_SelectionChanged:
  MsgBox "Error in ERROR_SelectionChanged : " & Error
  Resume EXIT_SelectionChanged

End Sub

Private Sub SaveGroup()

  On Error GoTo ERROR_SaveGroup

  Dim strFilename As String
  Dim i As Integer
  Dim j As Integer
  Dim strLine As String
  Dim nFile As Integer
  
  strFilename = App.Path & "\Internet Updates\" & Me.lstGroups.List(nSelected)
  Kill strFilename

  nFile = FreeFile
  Open strFilename For Output As #nFile
  For i = 1 To Me.lstDetails.ListItems.Count
    strLine = Me.lstDetails.ListItems.Item(i).Text
    strLine = strLine & "," & Me.lstDetails.ListItems.Item(i).SubItems(1)
    strLine = strLine & "," & Me.lstDetails.ListItems.Item(i).SubItems(2)
    strLine = strLine & "," & Me.lstDetails.ListItems.Item(i).SubItems(3)
    Print #nFile, strLine
  Next i
  Close #nFile
  bUpdated = False

EXIT_SaveGroup:
  Exit Sub

ERROR_SaveGroup:
  MsgBox "Error in ERROR_SaveGroup : " & Error
  Resume EXIT_SaveGroup
  
End Sub

Private Sub OKButton_Click()
  If bUpdated Then
    SelectionChanged True
  End If
  Unload Me
End Sub
