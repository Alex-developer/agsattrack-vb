Attribute VB_Name = "ContextIDs"
Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <AGSatTrack>
'=====================================================================
Global Const Hlp_frmSplash_ = 10    'Main Help Window
Global Const Hlp_Available_satellites = 35    'Main Help Window
Global Const Hlp_frmDocument_ = 45    'Main Help Window
Global Const Hlp_Program_Options = 55    'Main Help Window
Global Const Hlp_Window_Options = 65    'Main Help Window
Global Const Hlp_About_MyApp = 75    'Main Help Window
Global Const Hlp_Register_AGSatTrack = 85    'Main Help Window
Global Const Hlp_Registration_Information = 95    'Main Help Window
Global Const Hlp_Predictions_ = 105    'Main Help Window
Global Const Hlp_Form1_ = 115    'Main Help Window
Global Const Hlp_Form1_1 = 125    'Main Help Window
Global Const Hlp_Form1_2 = 135    'Main Help Window
Global Const Hlp_Gantt_Menu = 1310    'Main Help Window
Global Const Hlp_Form1_3 = 1320    'Main Help Window
Global Const Hlp_Edit_Internet = 1330    'Main Help Window
Global Const Hlp_Important_Information = 1340    'Main Help Window
Global Const Hlp_AnimatedGifCtl_ = 1350    'Main Help Window
Global Const Hlp_Edit_Element = 1360    'Main Help Window
Global Const Hlp_Form1_4 = 1370    'Main Help Window
Global Const Hlp_Setup_Rotator = 1380    'Main Help Window
Global Const Hlp_Form1_5 = 1390    'Main Help Window
Global Const Hlp_Internet_Keps = 1400    'Main Help Window
Global Const Hlp_Satellite_Details = 1410    'Main Help Window
Global Const Hlp_Chooser_Menu = 2310    'Main Help Window
Global Const Hlp_AGSatTrack_ = 2320    'Main Help Window
Global Const Hlp_SystrayMenu_ = 2410    'Main Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

    Type HELPWININFO
      wStructSize As Long
      X As Long
      Y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow as Long ' hWnd to tell WINHELP the helpfile owner

Dim MainWindowInfo as HELPWININFO
Public Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path),1) = "\" then
        App.HelpFile = App.Path + "AGSatTrack.HLP"
    else
        App.HelpFile = App.Path + "\AGSatTrack.HLP"
    end if
    MainWindowInfo.wStructSize = 26
    MainWindowInfo.X=256
    MainWindowInfo.Y=256
    MainWindowInfo.dX=512
    MainWindowInfo.dY=512
    MainWindowInfo.rgChMember=Chr$(0)+Chr$(0)
End Sub
Public Sub QuitHelp ()
    Dim Result as Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Public Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each 
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd02", HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd03", HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowGlossary ()
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(64000))

End Sub
Public Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, Clng(ContextID))

End Sub
Public Sub DoHelpMacro (ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal(Macrostring))

End Sub
Public Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, Clng(0))

End Sub
Public Sub ShowContentsTab ()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_FINDER, Clng(0))

End Sub
Public Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, Clng(0))

End Sub

Public Sub SearchHelp ()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal "" )

End Sub

Public Sub SearchHelpKeyWord (Argument as String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Public Sub HelpWindowSize (x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result as Variant
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dx = wx
    MainWindowInfo.dy = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub
