VERSION 5.00
Begin VB.Form frmGanttDetails 
   HelpContextID   =  1390
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmGanttDetails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   900
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblCaption 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   6765
   End
End
Attribute VB_Name = "frmGanttDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
  On Error Resume Next
  
  Me.lblCaption.Left = 0
  Me.lblCaption.Top = 0
  Me.lblCaption.Width = Me.ScaleWidth
  
  Me.lstData.Left = 0
  Me.lstData.Top = Me.lblCaption.Height
  Me.lstData.Width = Me.ScaleWidth
  Me.lstData.Height = Me.ScaleHeight - Me.lblCaption.Height
End Sub
