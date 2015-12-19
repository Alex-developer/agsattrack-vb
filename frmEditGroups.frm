VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEditGroups 
   HelpContextID   =  1360
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Element Sets"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "frmEditGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   750
      Top             =   4440
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   ">"
      Height          =   375
      Left            =   4590
      TabIndex        =   14
      Top             =   1380
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2430
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Height          =   375
      Left            =   4590
      TabIndex        =   13
      Top             =   2310
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "<"
      Height          =   375
      Left            =   4590
      TabIndex        =   12
      Top             =   900
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   8040
      TabIndex        =   11
      Top             =   4530
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   6900
      TabIndex        =   10
      Top             =   4530
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Caption         =   " Source Group "
      Height          =   4215
      Left            =   5220
      TabIndex        =   5
      Top             =   90
      Width           =   4275
      Begin VB.Frame Frame3 
         Caption         =   " Find "
         Height          =   3435
         Left            =   2400
         TabIndex        =   16
         Top             =   630
         Width           =   1695
         Begin VB.ListBox lstResults 
            Height          =   2010
            Left            =   150
            TabIndex        =   19
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   150
            TabIndex        =   18
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Find"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   690
            Width           =   945
         End
      End
      Begin VB.TextBox txtSource 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   210
         Width           =   2415
      End
      Begin VB.CommandButton cmdSource 
         Caption         =   "..."
         Height          =   285
         Left            =   3420
         TabIndex        =   7
         Top             =   240
         Width           =   465
      End
      Begin VB.ListBox lstSource 
         Height          =   3375
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   2145
      End
      Begin VB.Label Label2 
         Caption         =   "Filename"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Edit Group "
      Height          =   4215
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4275
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   750
         Width           =   855
      End
      Begin VB.ListBox lstEdit 
         Height          =   3375
         Left            =   150
         TabIndex        =   4
         Top             =   720
         Width           =   2745
      End
      Begin VB.CommandButton cmdEditGroup 
         Caption         =   "..."
         Height          =   285
         Left            =   3420
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Filename"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEditGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sSourceKeps(500) As Keps
Private sEditKeps(500) As Keps
Private bChanged As Boolean

Private Sub cmdAdd_Click()
   Dim i As Integer
   Dim nSourcePos As Integer
   Dim nEditPos As Integer
   Dim strName As String

   strName = Me.lstSource.Text
   If strName <> "" Then
      For i = 0 To 500
         If sEditKeps(i).strLine1 = "" Then Exit For
         If sEditKeps(i).strLine1 = strName Then
            MsgBox "This satellite already exists", vbInformation + vbOKOnly, "Error"
            Exit Sub
         End If
      Next i
      nEditPos = i

      For i = 0 To 500
         If sSourceKeps(i).strName = strName Then
            nSourcePos = i
            Exit For
         End If
      Next i
      Me.lstEdit.AddItem strName
      sEditKeps(nEditPos).strLine1 = sSourceKeps(nSourcePos).strLine1
      sEditKeps(nEditPos).strLine2 = sSourceKeps(nSourcePos).strLine2
      sEditKeps(nEditPos).strLine3 = sSourceKeps(nSourcePos).strLine3
      bChanged = True
   End If
End Sub

Private Sub cmdDel_Click()
  Dim strName As String
  Dim i As Integer
  Dim j As Integer
  
  strName = Me.lstEdit.Text
  If strName <> "" Then
    For i = 0 To 500
      If sEditKeps(i).strLine1 = strName Then
        For j = i To 499
          sEditKeps(j) = sEditKeps(j + 1)
        Next j
        Exit For
      End If
    Next i
   Me.lstEdit.Clear
   For i = 0 To 500
      If sEditKeps(i).strLine1 = "" Then Exit For
      Me.lstEdit.AddItem sEditKeps(i).strName
   Next i
    bChanged = True
  End If
  
End Sub

Private Sub cmdOk_Click()
  CheckSave vbYesNo
  Unload Me
End Sub

Private Function CheckSave(nButtons As Integer) As Integer
  Dim nResult As Integer
  
  
  If bChanged And txtEdit.Text <> "" Then
    nResult = MsgBox("You have changed the group " & GetFilename(txtEdit) & "but not saved it. Do you wish to save the group now", vbQuestion + nButtons, "Save Changes")
    If nResult = vbYes Then
      cmdSave_Click
    End If
  End If
  CheckSave = nResult
  
End Function
Private Sub cmdRemove_Click()
   Dim i As Integer
   Dim nSourcePos As Integer
   Dim nEditPos As Integer
   Dim strName As String

   strName = Me.lstEdit.Text
   If strName <> "" Then
      For i = 0 To 500
         If sSourceKeps(i).strLine1 = "" Then Exit For
         If sSourceKeps(i).strLine1 = strName Then
            MsgBox "This satellite already exists", vbInformation + vbOKOnly, "Error"
            Exit Sub
         End If
      Next i
      nSourcePos = i

      For i = 0 To 500
         If sEditKeps(i).strName = strName Then
            nEditPos = i
            Exit For
         End If
      Next i
      Me.lstSource.AddItem strName
      sSourceKeps(nSourcePos).strLine1 = sEditKeps(nEditPos).strLine1
      sSourceKeps(nSourcePos).strLine2 = sEditKeps(nEditPos).strLine2
      sSourceKeps(nSourcePos).strLine3 = sEditKeps(nEditPos).strLine3
      bChanged = True
   End If

End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdEditGroup_Click()
  On Error GoTo ErrorHandler
  If CheckSave(vbYesNoCancel) <> vbCancel Then
   With Me.CommonDialog1
      .ShowOpen
      Me.txtEdit.Text = GetFilename(.FileName)
      UpdateEdit
   End With
   End If
    Exit Sub
    
ErrorHandler:
  Select Case Err
    Case cdlCancel
  End Select

End Sub
Private Sub UpdateEdit()
   Dim i As Integer

   Erase sEditKeps
   ReadKeps App.Path & "\elements\" & Me.txtEdit.Text, False
   Me.lstEdit.Clear
   For i = 0 To 500
      If sEditKeps(i).strLine1 = "" Then Exit For
      Me.lstEdit.AddItem sEditKeps(i).strName
   Next i
End Sub

Private Sub cmdFind_Click()
   Dim strPath As String
   Dim strFile As String
   Dim strFullPath As String
   Dim nFile As Integer
   Dim strLine As String

   strPath = App.Path & "\elements\"
   If txtFind.Text <> "" Then
      Me.lstResults.Clear
      strFile = Dir(strPath & "*.*")
      If strFile <> "" Then
         While strFile <> ""
            strFullPath = strPath & strFile
            nFile = FreeFile
            Open strFullPath For Input As #nFile
            Do
               Line Input #nFile, strLine
               If InStr(UCase(strLine), UCase(Me.txtFind.Text)) Then
                  Me.lstResults.AddItem strFile
                  Exit Do
               End If
            Loop While Not EOF(nFile)
            Close #nFile
            strFile = Dir
         Wend
      End If
   End If
End Sub


Private Sub cmdSave_Click()
   Dim i As Integer
   Dim strFile As String
   Dim nFile As Integer

   strFile = App.Path & "\elements\" & Me.txtEdit.Text
   Kill strFile

   nFile = FreeFile
   Open strFile For Output As #nFile
   For i = 0 To 500
      If sEditKeps(i).strLine1 = "" Then Exit For
      Print #nFile, sEditKeps(i).strLine1
      Print #nFile, sEditKeps(i).strLine2
      Print #nFile, sEditKeps(i).strLine3
   Next i
   Close #nFile
  bChanged = False
End Sub

Private Sub cmdSource_Click()
   With Me.CommonDialog1
      .ShowOpen
      Me.txtSource.Text = GetFilename(.FileName)
      UpdateSource
   End With

End Sub
Private Sub UpdateSource()
   Dim i As Integer

   Erase sSourceKeps
   ReadKeps App.Path & "\elements\" & Me.txtSource.Text, True
   Me.lstSource.Clear
   For i = 0 To 500
      If sSourceKeps(i).strLine1 = "" Then Exit For
      Me.lstSource.AddItem sSourceKeps(i).strName
   Next i
End Sub

Private Sub Form_Load()
  CenterForm Me
  
   With Me.CommonDialog1
      .DialogTitle = "Select Element Set"
      .flags = cdlOFNNoChangeDir + cdlOFNPathMustExist
      .FileName = App.Path & "\elements\*.*"
   End With

End Sub

Private Function ReadKeps(strFilePath As String, bSource As Boolean) As Integer
   Dim TempInc As Integer
   Dim StrTempSatName As String
   Dim CrLf As String
   Dim strDummy As String
   Dim KepLine1 As String
   Dim KepLine2 As String
   Dim KepLine3 As String
   Dim nLines As Integer
   Dim nCount As Integer
   Dim NumberOfSatellites As Integer
   Dim strFilename As String
   Dim strFilenameNoExt As String
   Dim strDatabaseFile As String
   Dim nFile As Integer
   Dim nCounter As Integer

   strFilename = GetFilename(strFilePath)
   strFilenameNoExt = GetFilenameNoExt(strFilename)
   If strFilename <> "" And strFilenameNoExt <> "" Then

      Erase sKeps

      CrLf$ = Chr$(13) + Chr$(10)

      nFile = FreeFile
      Open strFilePath For Input As #nFile
      nLines = 0
      Do
         Input #nFile, strDummy$
         nLines = nLines + 1
      Loop Until EOF(nFile)
      Close #nFile

      Open strFilePath For Input As #nFile

      NumberOfSatellites = 0
      nCount = 0

      Do
         If Not EOF(nFile) Then
            Input #nFile, KepLine1$
            If KepLine1$ = "" Then
               Do
                  If Not EOF(nFile) Then
                     Input #nFile, KepLine1$
                  End If
               Loop Until KepLine1$ <> "" Or EOF(nFile)
            End If
            If EOF(nFile) Then Exit Do
            Input #nFile, KepLine2$
            If KepLine2$ = "" Then
               Do
                  If Not EOF(nFile) Then
                     Input #nFile, KepLine2$
                  End If
               Loop Until KepLine2$ <> "" Or EOF(nFile)
            End If
            If EOF(nFile) Then Exit Do
            Input #nFile, KepLine3$
            If KepLine3$ = "" Then
               Do
                  If Not EOF(nFile) Then
                     Input #nFile, KepLine3$
                  End If
               Loop Until KepLine3$ <> "" Or EOF(nFile)
            End If
            If EOF(nFile) And KepLine3$ = "" Then Exit Do
            If Mid$(KepLine2$, 24, 1) <> "." And Mid$(KepLine2$, 35, 1) <> "." Then
               Do
                  KepLine1$ = KepLine2$
                  KepLine2$ = KepLine3$
                  Input #nFile, KepLine3$
                  If KepLine3$ = "" Then
                     Do
                        If Not EOF(nFile) Then
                           Input #nFile, KepLine3$
                        End If
                     Loop Until KepLine3$ <> "" Or EOF(nFile)
                  End If
               Loop Until Mid$(KepLine2$, 24, 1) = "." And Mid$(KepLine3$, 12, 1) = "." Or EOF(nFile)
            End If
            NumberOfSatellites = NumberOfSatellites + 1

            GetKeps nCounter, KepLine1$, KepLine2$, KepLine3$, bSource
            nCounter = nCounter + 1

         End If
      Loop Until EOF(nFile)

      Close #nFile

      ReadKeps = NumberOfSatellites

   End If

ExitSub:
   If NumberOfSatellites = 0 Then
      Call MsgBox("The file (" & strFilename & ") you have selected does not appear to contain any Keplarian elements. Please select another file.", vbExclamation + vbOKOnly + vbDefaultButton1, "Keplairan Element Load Error")
   End If
   Close #nFile
   Exit Function

ErrorHandler:
   Close #nFile
   Resume ExitSub
End Function
Private Function GetKeps(nPos As Integer, strLine1 As String, strLine2 As String, strLine3 As String, bSource As Boolean) As Boolean

   If bSource Then
      sSourceKeps(nPos).strLine1 = strLine1
      sSourceKeps(nPos).strLine2 = strLine2
      sSourceKeps(nPos).strLine3 = strLine3
      sSourceKeps(nPos).strName = Trim(strLine1)
      sSourceKeps(nPos).lDesignator = Val(Mid$(strLine2, 3, 5))
      sSourceKeps(nPos).strEpoch = Mid$(strLine2, 19, 14)
      sSourceKeps(nPos).dDrag = Val(Mid$(strLine2, 35, 9))
      sSourceKeps(nPos).lRevolutionnumber = Val(Mid$(strLine3, 64, 5))
      sSourceKeps(nPos).dInclination = Val(Mid$(strLine3, 9, 8))
      sSourceKeps(nPos).dRAAN = Val(Mid$(strLine3, 18, 8))
      sSourceKeps(nPos).dEccentricity = Val("0." + Mid$(strLine3, 27, 7))
      sSourceKeps(nPos).dAOP = Val(Mid$(strLine3, 35, 8))
      sSourceKeps(nPos).dMeanAnomoly = Val(Mid$(strLine3, 44, 8))
      sSourceKeps(nPos).dMeanMotion = Val(Mid$(strLine3, 53, 11))
      sSourceKeps(nPos).nElementSet = Val(Mid$(strLine2, 66, 3))
      sSourceKeps(nPos).lOrbitNUmber = Val(Mid$(strLine3, 64, 5))
   Else
      sEditKeps(nPos).strLine1 = strLine1
      sEditKeps(nPos).strLine2 = strLine2
      sEditKeps(nPos).strLine3 = strLine3
      sEditKeps(nPos).strName = Trim(strLine1)
      sEditKeps(nPos).lDesignator = Val(Mid$(strLine2, 3, 5))
      sEditKeps(nPos).strEpoch = Mid$(strLine2, 19, 14)
      sEditKeps(nPos).dDrag = Val(Mid$(strLine2, 35, 9))
      sEditKeps(nPos).lRevolutionnumber = Val(Mid$(strLine3, 64, 5))
      sEditKeps(nPos).dInclination = Val(Mid$(strLine3, 9, 8))
      sEditKeps(nPos).dRAAN = Val(Mid$(strLine3, 18, 8))
      sEditKeps(nPos).dEccentricity = Val("0." + Mid$(strLine3, 27, 7))
      sEditKeps(nPos).dAOP = Val(Mid$(strLine3, 35, 8))
      sEditKeps(nPos).dMeanAnomoly = Val(Mid$(strLine3, 44, 8))
      sEditKeps(nPos).dMeanMotion = Val(Mid$(strLine3, 53, 11))
      sEditKeps(nPos).nElementSet = Val(Mid$(strLine2, 66, 3))
      sEditKeps(nPos).lOrbitNUmber = Val(Mid$(strLine3, 64, 5))
   End If
End Function

Private Sub lstResults_Click()

   Me.txtSource.Text = Me.lstResults.Text
   UpdateSource

End Sub

Private Sub Timer1_Timer()
cmdEditGroup_Click
Timer1.Enabled = False
End Sub
