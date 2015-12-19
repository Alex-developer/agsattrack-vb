VERSION 5.00
Begin VB.Form frmRegister 
   HelpContextID   =  85
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register AGSatTrack"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3630
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   180
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Code"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdRegister_Click()
  If KeyGen(Me.txtName, "6B8A-5063-205E-2A78", 3) = Me.txtCode Then
    bRegistered = True
    sProgramOptions.strUserName = Me.txtName
    sProgramOptions.strCode = Me.txtCode
    MsgBox "Thank you for registering this software", vbInformation + vbOKOnly, "Register"
    fMainForm.Caption = "AGSatTrak - Registered to " & Me.txtName
    Unload Me
  Else
    MsgBox "Invalid Registration code", vbCritical + vbOKOnly, "Registration code error"
    nRegCount = nRegCount + 1
    If nRegCount = 3 Then
      fMainForm.SSActiveToolBars1.Tools("id_Register").Enabled = False
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me
End Sub
