VERSION 5.00
Begin VB.UserControl ScrollingViewPort 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2565
   Begin VB.PictureBox Corner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   255
      Left            =   3000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   255
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1815
      Left            =   3000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "ScrollingViewPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Public Enum enumBorderStyle
    FixedSingle = 1
    None = 0
End Enum
Public Enum enumAppearance
   Flat = 0
   a3D = 1
End Enum
Public Enum enumBackStyle
    Transparent = 0
    Opaque = 1
End Enum
Public Enum enumScrollBar
    Always = 0
    Automatic = 1
    Never = 2
End Enum
'Default Property Values:
Const m_def_Appearance = 1
Const m_def_BackStyle = 1
Const m_def_BorderStyle = 1
Const m_def_ScrollBarVerticle = 1
Const m_def_ScrollBarHorizontal = 1
Const m_def_BackColor = &HE0E0E0
'Property Variables:
Dim m_ScrollBarVerticle As enumScrollBar
Dim m_ScrollBarHorizontal As enumScrollBar
Dim m_BackColor As OLE_COLOR
Dim m_ControledPictureBox As PictureBox
Dim m_PWidth As Integer
Dim m_PHeight As Integer
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get ControledPictureBox() As PictureBox
    Set ControledPictureBox = m_ControledPictureBox
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub HScroll_Change()
    Scroll
End Sub

Private Sub HScroll_Scroll()
    Scroll
End Sub

Private Sub UserControl_Initialize()
'    Info.Caption = "Place a picturebox on this control to have an instant scrollable viewport.  Yeah, it's a simple control, but it's useful and it's my first submission to PSC.  Vote if you like it." & vbNewLine & vbNewLine & "GKenny@Sprintmail.com"
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_ControledPictureBox = FindPictureBox
    m_BackColor = m_def_BackColor
    m_ScrollBarVerticle = m_def_ScrollBarVerticle
    m_ScrollBarHorizontal = m_def_ScrollBarHorizontal
    UserControl.Appearance = m_def_Appearance
    UserControl.BackColor = m_def_BackColor
    UserControl.BackStyle = m_def_BackStyle
    UserControl.BorderStyle = m_def_BorderStyle
End Sub

Private Sub UserControl_Paint()
    If Ambient.UserMode = True Then Exit Sub
    Dim pb As PictureBox
    Dim pw, ph As Integer
    Set pb = FindPictureBox
    If Not pb Is Nothing Then
        pw = pb.Width
        ph = pb.Height
    End If
    If Not pw = m_PWidth Or Not ph = m_PHeight Then
        FormatControl
    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = m_BackColor
    m_ScrollBarVerticle = PropBag.ReadProperty("ScrollBarVerticle", m_def_ScrollBarVerticle)
    m_ScrollBarHorizontal = PropBag.ReadProperty("ScrollBarHorizontal", m_def_ScrollBarHorizontal)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Height < 300 Then
        UserControl.Height = 300
        Exit Sub
    End If
    If UserControl.Width < 300 Then
        UserControl.Width = 300
        Exit Sub
    End If
    FormatControl
End Sub

Private Sub UserControl_Show()
    Dim pb As PictureBox
    Set pb = FindPictureBox
    FormatControl
    If Not pb Is Nothing Then
        pb.Top = -VScroll.Value
        pb.Left = -HScroll.Value
    End If
    'Info.Visible = Not Ambient.UserMode
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("ScrollBarVerticle", m_ScrollBarVerticle, m_def_ScrollBarVerticle)
    Call PropBag.WriteProperty("ScrollBarHorizontal", m_ScrollBarHorizontal, m_def_ScrollBarHorizontal)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,1
Public Property Get Appearance() As enumAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As enumAppearance)
    UserControl.Appearance = New_Appearance
    UserControl.BackColor = m_BackColor
    PropertyChanged "Appearance"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&hffffff
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,1
Public Property Get BackStyle() As enumBackStyle
Attribute BackStyle.VB_Description = "Returns/sets whether the background is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As enumBackStyle)
    UserControl.BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,1
Public Property Get BorderStyle() As enumBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As enumBorderStyle)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get ScrollBarHorizontal() As enumScrollBar
    ScrollBarHorizontal = m_ScrollBarHorizontal
End Property

Public Property Let ScrollBarHorizontal(ByVal New_ScrollBarHorizontal As enumScrollBar)
    m_ScrollBarHorizontal = New_ScrollBarHorizontal
    PropertyChanged "ScrollBarHorizontal"
    FormatControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get ScrollBarVerticle() As enumScrollBar
    ScrollBarVerticle = m_ScrollBarVerticle
End Property

Public Property Let ScrollBarVerticle(ByVal New_ScrollBarVerticle As enumScrollBar)
    m_ScrollBarVerticle = New_ScrollBarVerticle
    PropertyChanged "ScrollBarVerticle"
    FormatControl
End Property
Private Sub FormatControl()
    Dim pb As PictureBox
    Dim cw, ch As Integer
    Set pb = FindPictureBox
    If Not pb Is Nothing Then
        m_PWidth = pb.Width
        m_PHeight = pb.Height
        'Info.Visible = False
    Else
        m_PWidth = 0
        m_PHeight = 0
    End If
    cw = UserControl.Width - UserControl.BorderStyle * (2 + UserControl.Appearance) * Screen.TwipsPerPixelX
    ch = UserControl.Height - UserControl.BorderStyle * (2 + UserControl.Appearance) * Screen.TwipsPerPixelY
    If m_ScrollBarVerticle = Always Then
        VScroll.Left = cw - VScroll.Width
        VScroll.Visible = True
    Else
        VScroll.Left = cw
        VScroll.Visible = False
    End If
    If m_ScrollBarHorizontal = Always Then
        HScroll.Top = ch - HScroll.Height
        HScroll.Visible = True
    Else
        HScroll.Top = ch
        HScroll.Visible = False
    End If
    If m_ScrollBarVerticle = Automatic Then
        If m_PHeight > ch - Screen.TwipsPerPixelY Then
            VScroll.Left = cw - VScroll.Width
            VScroll.Visible = True
            If m_ScrollBarHorizontal = Automatic Then
                If m_PWidth > cw - VScroll.Width - Screen.TwipsPerPixelX Then
                    HScroll.Top = ch - HScroll.Height
                    HScroll.Visible = True
                End If
            End If
        End If
    End If
    If m_ScrollBarHorizontal = Automatic Then
        If m_PWidth > cw - Screen.TwipsPerPixelX Then
            HScroll.Top = ch - HScroll.Height
            HScroll.Visible = True
            If m_ScrollBarVerticle = Automatic Then
                If m_PHeight > ch - HScroll.Height - Screen.TwipsPerPixelY Then
                    VScroll.Left = cw - VScroll.Width
                    VScroll.Visible = True
                End If
            End If
        End If
    End If
    VScroll.Height = HScroll.Top
    HScroll.Width = VScroll.Left
    Corner.Left = cw - Corner.Width
    Corner.Top = ch - Corner.Height
    Corner.Visible = HScroll.Visible And VScroll.Visible
    Corner.ZOrder 0
    HScroll.ZOrder 0
    VScroll.ZOrder 0
    If m_PWidth > VScroll.Left + Screen.TwipsPerPixelX Then
        HScroll.Max = m_PWidth - VScroll.Left + Screen.TwipsPerPixelX
    Else
        HScroll.Max = 0
    End If
    If m_PHeight > HScroll.Top + Screen.TwipsPerPixelY Then
        VScroll.Max = m_PHeight - HScroll.Top + Screen.TwipsPerPixelY
    Else
        VScroll.Max = 0
    End If
    HScroll.LargeChange = m_PWidth * 0.75 + 1
    HScroll.SmallChange = m_PWidth * 0.05 + 1
    VScroll.LargeChange = m_PHeight * 0.75 + 1
    VScroll.SmallChange = m_PHeight * 0.05 + 1
End Sub
Private Function FindPictureBox() As PictureBox
    Dim c As Control
    On Error GoTo FindIt
    If m_ControledPictureBox.Name <> "" Then
        Set FindPictureBox = m_ControledPictureBox
        Exit Function
    End If
FindIt:
    Set m_ControledPictureBox = Nothing
    For Each c In UserControl.ContainedControls
        If TypeOf c Is PictureBox Then
            Set m_ControledPictureBox = c
            Exit For
        End If
    Next c
    Set FindPictureBox = m_ControledPictureBox
    PropertyChanged "ControledPictureBox"
End Function
Private Sub Scroll()
    Dim pb As PictureBox
    Set pb = FindPictureBox
    If Not pb Is Nothing Then
        pb.Top = -VScroll.Value
        pb.Left = -HScroll.Value
    End If
End Sub

Private Sub VScroll_Change()
    Scroll
End Sub

Private Sub VScroll_Scroll()
    Scroll
End Sub
