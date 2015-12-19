VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReadme 
   HelpContextID   =  1340
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Important Information"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmReadme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   6045
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10663
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReadme.frx":0442
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3750
      TabIndex        =   0
      Top             =   6270
      Width           =   1005
   End
End
Attribute VB_Name = "frmReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  On Error GoTo ERROR_Form_Load

  Dim nFile As Integer
  Dim strText As String
  Dim strLine As String
  
  CenterForm Me
  
  Me.rtf1.FileName = App.Path & "\agsattrack.rtf"
  
'  nFile = FreeFile
'  Open App.Path & "\readme.txt" For Input As #nFile
'  While Not (EOF(nFile))
'    Input #nFile, strLine
'    strText = strText & strLine & vbCrLf
'  Wend
'  Close #nFile
'  Me.Text1.Text = strText

EXIT_Form_Load:
  Exit Sub

ERROR_Form_Load:
  MsgBox "Error in ERROR_Form_Load : " & Error
  Resume EXIT_Form_Load
  
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Me.cmdOk.Top = Me.ScaleHeight - (Me.cmdOk.Height * 1.5)
  Me.rtf1.Top = 0
  Me.rtf1.Left = 0
  Me.rtf1.Width = Me.ScaleWidth
  Me.rtf1.Height = Me.ScaleHeight - (Me.cmdOk.Height * 2)
End Sub
