VERSION 5.00
Begin VB.Form frmReport 
   HelpContextID   =  1320
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   8250
   Begin VB.Timer tmrUpdate 
      Interval        =   65000
      Left            =   360
      Top             =   3660
   End
   Begin VB.TextBox txtReport 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmReport.frx":0442
      Top             =   60
      Width           =   7965
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nType As Integer
Public fForm As Form
Public strCaption As String
Public nSatIndex As Integer
Dim nOldIndex As Integer

Public Sub UpdateReport(Optional bInternal As Boolean)
  Me.Caption = "Updating Report. Please wait..."
  Select Case nType
    Case 2 And Not bInternal
      Me.txtReport = ""
      Me.txtReport = fForm.ocxSat.DisplayAOSReport(2)
    Case 3 And Not bInternal
      Me.txtReport = ""
      Me.txtReport = fForm.ocxSat.DisplayAOSReport(1)
    Case 4
      Me.txtReport = ""
      nOldIndex = fForm.ocxSat.SatelliteIndex
      fForm.ocxSat.SatelliteIndex = nSatIndex
      Me.txtReport = fForm.ocxSat.displayDX
      fForm.ocxSat.SatelliteIndex = nOldIndex
    Case 5 And Not bInternal
      Me.txtReport = ""
      Me.txtReport = fForm.ocxSat.DisplayAOSReport(3)
  End Select
  Me.Caption = Me.strCaption
End Sub

Private Sub Form_Resize()
  Me.txtReport.Left = 0
  Me.txtReport.Top = 0
  Me.txtReport.Width = Me.ScaleWidth
  Me.txtReport.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  lDocumentCount = lDocumentCount - 1
  fMainForm.UpdateToolbar
End Sub

Private Sub tmrUpdate_Timer()
  UpdateReport True
End Sub

Private Sub txtReport_GotFocus()
  fMainForm.UpdateToolbar
End Sub

Private Sub txtReport_LostFocus()
  fMainForm.UpdateToolbar
End Sub
