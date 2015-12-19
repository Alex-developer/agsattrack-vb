VERSION 5.00
Begin VB.Form frmModeSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Satellite Mode Selection"
   ClientHeight    =   1500
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmModeSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstModes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   150
      TabIndex        =   2
      Top             =   450
      Width           =   4395
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4740
      TabIndex        =   1
      Top             =   690
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4740
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Satellite    Downlink   Uplink"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   3945
   End
End
Attribute VB_Name = "frmModeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  bOpt = -1
  Unload Me
End Sub


Private Sub lstModes_DblClick()
  bOpt = Me.lstModes.ListIndex
  Unload Me
End Sub

Private Sub OKButton_Click()
  bOpt = Me.lstModes.ListIndex
  Unload Me
End Sub
