VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRegInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Information"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmRegInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRegInfo.frx":030A
   ScaleHeight     =   5625
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   2100
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   4395
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7752
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmRegInfo.frx":6A72
   End
End
Attribute VB_Name = "frmRegInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  On Error GoTo ERROR_Form_Load
  
  CenterForm Me
  
  Me.rtf1.FileName = App.Path & "\Register.rtf"

EXIT_Form_Load:
  Exit Sub

ERROR_Form_Load:
  MsgBox "Error in ERROR_Form_Load : " & Error
  Resume EXIT_Form_Load

End Sub

