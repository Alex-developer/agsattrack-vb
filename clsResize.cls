VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ======================================================================================== '
' Component : clsResize                                               Created : 01/05/1999 '
' File Name : clsResize.cls                                           Author  : A D Moss   '
'                                                                                          '
' Purpose   : Resize and Reposition all controls on a Form.                                '
' ======================================================================================== '
' ======================================================================================== '
'                                                                                          '
Option Explicit
' ======================================================================================== '
' API Declarations.                                                                        '
' ======================================================================================== '
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

' ======================================================================================== '
' Constant Declarations.                                                                   '
' ======================================================================================== '
Private Const SSTAB_DIALOG_OFFSET As Long = 75000
Private Const SSTAB_TYPE_NAME As String = "SSTAB"

Private Const ResizeRepositionCommand As String = "@"
Private Const ResizeRepositionLeft As String = "L"
Private Const ResizeRepositionTop As String = "T"
Private Const ResizeRepositionWidth As String = "W"
Private Const ResizeRepositionHeight As String = "H"

' ======================================================================================== '
' Type Declarations.                                                                       '
' ======================================================================================== '
Private Type ControlPosition
  ControlInstance As Control                         'Reference to the control instance.
  OriginalLeft As Long                               'Original Left position of the Control.
  OriginalTop As Long                                'Original Top position of the Control.
  OriginalWidth As Long                              'Original Width of the Control.
  OriginalHeight As Long                             'Original Height of the Control.
End Type

' ======================================================================================== '
' Enumeration Declarations.                                                                '
' ======================================================================================== '

' ======================================================================================== '
' Event Declarations.                                                                      '
' ======================================================================================== '

' ======================================================================================== '
' Public Variable Declarations.                                                            '
' ======================================================================================== '

' ======================================================================================== '
' Private Variable Declarations.                                                           '
' ======================================================================================== '
Private m_SourceForm As Form                         'The form to be resized.
Private m_FormWidth As Long                          'Original form width.
Private m_FormHeight As Long                         'Original form height.
Private m_Controls() As ControlPosition              'Array for storing control information.
Private m_IsFirstResize As Boolean                   'Flag indicating first resize.

' ======================================================================================== '
' Routine     : Class_Initialize                                      Created : 01/05/1999 '
' Scope       : Private                                               Author  : A D Moss   '
' Description : Constructor used when an instance of this class is created.                '
' ======================================================================================== '
Private Sub Class_Initialize()
  Set m_SourceForm = Nothing

  m_IsFirstResize = True
End Sub

' ======================================================================================== '
' Routine     : Class_Terminate                                       Created : 01/05/1999 '
' Scope       : Private                                               Author  : A D Moss   '
' Description : Destructor used when an instance of this class is destroyed.               '
' ======================================================================================== '
Private Sub Class_Terminate()
  Erase m_Controls

  Set m_SourceForm = Nothing
End Sub

' ======================================================================================== '
' Routine     : SourceForm (Get)                                      Created : 01/05/1999 '
' Scope       : Public                                                Author  : A D Moss   '
' Description : Return the Form currently being used by an instance of this class.         '
' ======================================================================================== '
Public Property Get SourceForm() As Form
  Set SourceForm = m_SourceForm
End Property

' ======================================================================================== '
' Routine     : SourceForm (Let)                                      Created : 01/05/1999 '
' Scope       : Public                                                Author  : A D Moss   '
' Description : Set the Form to be used by an instance of this class.                      '
' ======================================================================================== '
Public Property Let SourceForm(New_SourceForm As Form)
  Set m_SourceForm = New_SourceForm
End Property

' ======================================================================================== '
' Routine     : SizeFormToScreen                                      Created : 01/05/1999 '
' Scope       : Public                                                Author  : A D Moss   '
' Description : Size the Form to n Percent of the Screen.                                  '
' ======================================================================================== '
Public Sub SizeFormToScreen(Optional Percent As Byte = 100)
  'Variable Declarations.
  Dim FormHeight As Long, FormWidth As Long

  'Check if the form has not been resized before.
  If m_IsFirstResize Then
    Call SaveInitialStates
  End If

  'Calculate the new height and width the form needs to be resized to, based on the current
  'screen resolution.
  FormHeight = Int(Screen.Height * (Percent / 100))
  FormWidth = Int(Screen.Width * (Percent / 100))

  'Use the Form that is to be resized.
  With m_SourceForm
    'Change the demensions and position of the form.
    .Top = (Screen.Height - FormHeight) / 2
    .Left = (Screen.Width - FormWidth) / 2
    .Height = FormHeight
    .Width = FormWidth
  End With

  'Resize all of the controls on the form.
  Call ResizeControls
End Sub

' ======================================================================================== '
' Routine     : SaveInitialStates                                     Created : 01/05/1999 '
' Scope       : Private                                               Author  : A D Moss   '
' Description : Store the dimensions and positions of all controls on the Form.            '
' ======================================================================================== '
Private Sub SaveInitialStates()
  'Variable Declarations.
  Dim i As Integer

  'Check if the form has any controls on it.
  If m_SourceForm.Controls.Count = 0 Then Exit Sub

  'Use the form that is being resized.
  With m_SourceForm
    'Check if the form being resized is an MDI form.
    If (TypeOf m_SourceForm Is MDIForm) Then
      'Set the FormWidth and FormHeight variables to the Form's Width and Height.
      m_FormWidth = .Width
      m_FormHeight = .Height
    Else
      'Set the FormWidth and FormHeight variables to the Form's Scale Width and Height.
      m_FormWidth = .ScaleWidth
      m_FormHeight = .ScaleHeight
    End If
  End With

  'Frees the memory used by the dynamic array storing all control information.
  Erase m_Controls

  'Redimension the array storing all control information to the number of controls on the
  'form that is being resized.
  ReDim m_Controls(0 To (m_SourceForm.Controls.Count - 1))

  'Iterate through each of the Controls on the form to be resized.
  For i = 0 To (m_SourceForm.Controls.Count - 1)
    'Use the control that is currently being referenced.
    With m_SourceForm.Controls(i)
      'Add the controls reference to the array containing all control information.
      Set m_Controls(i).ControlInstance = m_SourceForm.Controls(i)

      'Check if the control currently being referenced is a line control.
      If (TypeOf m_SourceForm.Controls(i) Is Line) Then
        'Add a reference to the control and it's position into the Controls array.
        m_Controls(i).OriginalLeft = .X1
        m_Controls(i).OriginalTop = .Y1
        m_Controls(i).OriginalWidth = .X2
        m_Controls(i).OriginalHeight = .Y2
      Else
        On Error Resume Next
          'Add a reference to the control and it's position into the Controls array.
          m_Controls(i).OriginalLeft = .Left
          m_Controls(i).OriginalTop = .Top
          m_Controls(i).OriginalWidth = .Width
          m_Controls(i).OriginalHeight = .Height
        On Error GoTo 0
      End If
    End With
  Next

  'Notify the class that the form has been resized.
  m_IsFirstResize = False
End Sub

' ======================================================================================== '
' Routine     : ReInitialise                                          Created : 01/05/1999 '
' Scope       : Public                                                Author  : A D Moss   '
' Description : Store the dimensions and positions of all controls on a Form.              '
' ======================================================================================== '
Public Sub ReInitialise()
  Call SaveInitialStates
End Sub

' ======================================================================================== '
' Routine     : ResizeControls                                        Created : 01/05/1999 '
' Scope       : Public                                                Author  : A D Moss   '
' Description : Resize all of the Controls on a Form.                                      '
' ======================================================================================== '
Public Sub ResizeControls()
  'Variable Declarations.
  Dim HeightChange As Double, WidthChange As Double
  Dim ControlInstance As Control
  Dim ResizeRepositionFlags As String
  Dim i As Integer

  'Check if the form has any controls on it.
  If m_SourceForm.Controls.Count = 0 Then Exit Sub

  'Check if the form is being minimized.
  If m_SourceForm.WindowState = vbMinimized Then Exit Sub

  'Check if it's the first time the form is to be resized.
  If m_IsFirstResize Then
    Call SaveInitialStates
    Exit Sub
  End If

  'Prevent the Window from redrawing until the user releases the mouse.
  Call LockWindowUpdate(m_SourceForm.hwnd)

  'Calculate the change in the form's size.
  If (TypeOf m_SourceForm Is MDIForm) Then
    HeightChange = m_SourceForm.Height / m_FormHeight
    WidthChange = m_SourceForm.Width / m_FormWidth
  Else
    HeightChange = m_SourceForm.ScaleHeight / m_FormHeight
    WidthChange = m_SourceForm.ScaleWidth / m_FormWidth
  End If

  'Iterate through the array containing the form's controls.
  For i = 0 To UBound(m_Controls)
    Set ControlInstance = m_Controls(i).ControlInstance

    'Use the Control that is currently being referenced.
    With ControlInstance
      On Error Resume Next
        'Check if the Tag Property contains valid resize or reposition instructions.
        If (Left(UCase(.Tag), 1) = ResizeRepositionCommand) Then
          'Retrieve any specific resize and reposition instructions.
          ResizeRepositionFlags = UCase(.Tag)
        Else
          'Default all resize and resposition instructions.
          ResizeRepositionFlags = ResizeRepositionCommand & ResizeRepositionLeft & ResizeRepositionTop & ResizeRepositionWidth & ResizeRepositionHeight
        End If

        'Resize and Reposition the control.
        If (TypeOf ControlInstance Is Line) Then
          'Line control.
           If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .X1 = (m_Controls(i).OriginalLeft * WidthChange)
           If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .Y1 = (m_Controls(i).OriginalTop * HeightChange)
           If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .X2 = (m_Controls(i).OriginalWidth * WidthChange)
           If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Y2 = (m_Controls(i).OriginalHeight * HeightChange)
        ElseIf (UCase(TypeName(ControlInstance.Container)) = SSTAB_TYPE_NAME) Then
          'Control is placed on a SSTab Dialog Control.
          If (.Left < 0) Then
            If (m_Controls(i).OriginalLeft > 0) Then
              m_Controls(i).OriginalLeft = m_Controls(i).OriginalLeft - SSTAB_DIALOG_OFFSET
            End If

            If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .Left = ((m_Controls(i).OriginalLeft + SSTAB_DIALOG_OFFSET) * WidthChange) - SSTAB_DIALOG_OFFSET
          Else
            If (m_Controls(i).OriginalLeft < 0) Then
              m_Controls(i).OriginalLeft = m_Controls(i).OriginalLeft + SSTAB_DIALOG_OFFSET
            End If

            If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .Left = m_Controls(i).OriginalLeft * WidthChange
          End If

          If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .Top = (m_Controls(i).OriginalTop * HeightChange)
          If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .Width = (m_Controls(i).OriginalWidth * WidthChange)
          If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Height = (m_Controls(i).OriginalHeight * HeightChange)
        Else
          'All other controls.
          If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .Left = (m_Controls(i).OriginalLeft * WidthChange)
          If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .Top = (m_Controls(i).OriginalTop * HeightChange)
          If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .Width = (m_Controls(i).OriginalWidth * WidthChange)
          If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Height = (m_Controls(i).OriginalHeight * HeightChange)
        End If
      On Error GoTo 0
    End With
  Next i

  'Release the memory being used by the ControlInstance variable.
  Set ControlInstance = Nothing

  'Allow Windows to refresh the form's display.
  Call LockWindowUpdate(0)
End Sub
