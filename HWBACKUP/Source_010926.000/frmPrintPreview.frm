VERSION 5.00
Begin VB.Form frmPrintPreview 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      Height          =   7515
      Left            =   60
      ScaleHeight     =   7455
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   120
      Width           =   9435
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim opreview As Preview
  Dim strCaption As String
  Dim nYpos As Single
  Const TopMargin = 0.25
  Const LeftMargin = 0.25
  Const RightMargin = 0.25
  Const BottomMargin = 0.25
  
Private Sub Form_Load()

  Set opreview = New Preview
  opreview.Container = picPreview.hwnd
  GeneratePreview
End Sub

Private Sub Form_Resize()
  On Local Error Resume Next
  
  Me.picPreview.Left = 0
  Me.picPreview.Top = 0
  Me.picPreview.Width = Me.ScaleWidth
  Me.picPreview.Height = Me.ScaleHeight
  opreview.Container = picPreview.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set opreview = Nothing
End Sub

Private Sub GeneratePreview()
  
  strCaption = frmForm.Caption
  
  If Left(strCaption, 7) = "SatView" Then
    PreviewSatWindow
  End If
  If Left(strCaption, 5) = "Gantt" Then
    PreviewGanttWindow
  End If
  If strCaption = "Predictions" Then
    PreviewPrediction
  End If
End Sub

Private Sub PreviewSatWindow()
  Dim nFontIndex As Integer

  Me.Caption = "AGSatTrack map preview"
  With opreview
    .Cls
    .OutputFile = "d:\temp\alex.htm"
    With .Pages
      .ScaleMode = vbInches
      .Width = 11
      .Height = 8.5
      .Add
      nFontIndex = .ActivePage.SetFont("Courier New", 10, False, False, False, False)
      With .ActivePage
        nYpos = 0
        PrintHeading "AGSatTrack map view", 1, nFontIndex
        .DrawPicture 1.2, 0.1, 8.5, 8, frmForm.ocxSat.Picture, True
        PrintFooter nFontIndex
      End With
    End With
    .Show
  End With
End Sub
Private Sub PreviewGanttWindow()
  Dim nFontIndex As Integer

  Me.Caption = "AGSatTrack Gantt preview"
  With opreview
    .Cls
    .OutputFile = "d:\temp\alex.htm"
    With .Pages
      .ScaleMode = vbInches
      .Width = 11
      .Height = 8.5
      .Add
      nFontIndex = .ActivePage.SetFont("Courier New", 10, False, False, False, False)
      With .ActivePage
        nYpos = 0
        PrintHeading "AGSatTrack Gantt view", 1, nFontIndex
        picPic.Picture = frmForm.picGantt.Image
        picPic.Refresh
        .DrawPicture 1.2, 0.1, 8.5, 8, picPic.Picture, True
        PrintFooter nFontIndex
      End With
    End With
    .Show
  End With
  
End Sub

Private Sub PreviewPrediction()
  Dim i As Single
  Dim nFontHeight As Single
  Dim nPage As Integer
  Dim nFontIndex As Integer
  
  Me.Caption = "AGSatTrack predictions preview"
  With opreview
    .Cls
    .OutputFile = "d:\temp\alex.htm"
    With .Pages
      .ScaleMode = vbInches
      .PaperSize = vbPRPSA4
      .Width = 8.5
      .Height = 11
      .Add
      nYpos = 0
      nPage = 1
      nFontIndex = .ActivePage.SetFont("Courier New", 10, False, False, False, False)
      nFontHeight = (10 / 72)
      PrintHeading "AGSatTrack prediction report", nPage, nFontIndex
      For i = 0 To frmForm.lstData.ListCount - 1
        nYpos = nYpos + nFontHeight
        If nYpos > .Height - BottomMargin - 0.5 Then
          PrintFooter nFontIndex
          .Add
          nFontIndex = .ActivePage.SetFont("Courier New", 10, False, False, False, False)
          nYpos = 0
          nPage = nPage + 1
          PrintHeading "AGSatTrack prediction report", nPage, nFontIndex
          nYpos = nYpos + nFontHeight
        End If
        .ActivePage.DrawText frmForm.lstData.List(i), 0, nYpos, 10, 1, vbBlack, vbWhite, vbLeftJustify, nFontIndex
      Next i
      PrintFooter nFontIndex
    End With
    .Show
  End With

End Sub

Private Sub PrintHeading(strText As String, nPage As Integer, nFontIndex As Integer)

  With opreview.Pages.ActivePage
    .DrawText strText, LeftMargin, TopMargin, 2, 1, , , vbLeftJustify, nFontIndex
    .DrawText "Page " & nPage, .Width - RightMargin - 2, TopMargin, 2, 1, , , vbLeftJustify, nFontIndex
    .DrawLine LeftMargin, TopMargin + 0.2, .Width - RightMargin, TopMargin + 0.2, RGB(0, 0, 0)
  End With
  nYpos = TopMargin + 0.2
End Sub

Private Sub PrintFooter(nFontIndex As Integer)
  
  With opreview.Pages.ActivePage
    .DrawLine LeftMargin, .Height - BottomMargin - 0.2, .Width - RightMargin, .Height - BottomMargin - 0.2, RGB(0, 0, 0)
    .DrawText "Printed on " & Format(Now, "General Date"), LeftMargin, .Height - BottomMargin - 0.2, 3, 1, , , vbLeftJustify, nFontIndex
    '.DrawText "Page " & nPage, .Width - RightMargin - 2, TopMargin, 2, 1
  End With
End Sub

