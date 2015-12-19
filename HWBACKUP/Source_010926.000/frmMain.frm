VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8E256E76-10D8-4532-BB33-230282428A57}#1.0#0"; "ftpOCX.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "AGSatTrack"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8520
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "SatData"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrGeneral 
      Interval        =   1000
      Left            =   540
      Top             =   840
   End
   Begin MSComctlLib.StatusBar stbStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrFTP 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4620
      Top             =   2340
   End
   Begin ftpOCX.FTP FTP1 
      Left            =   3120
      Top             =   2340
      _ExtentX        =   1058
      _ExtentY        =   1058
      TransferType    =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5820
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "newicon"
            Object.Tag             =   "newicon"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrReg 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4080
      Top             =   180
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   900
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   7
      ToolsCount      =   86
      PersonalizedMenus=   0
      Style           =   0
      Tools           =   "frmMain.frx":0626
      ToolBars        =   "frmMain.frx":22720
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2340
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuSysMenu 
      Caption         =   "SystrayMenu"
      Begin VB.Menu SysTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu systraysep 
         Caption         =   "-"
      End
      Begin VB.Menu SysTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bUpdating As Boolean
Dim nUploadTimer As Integer
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim strTempPath As String
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1

Private Sub FTP1_GetError(Error As String, Func As String, ErrorNum As Long)
  MsgBox Error, vbOKOnly, "FTP Upload Error"
End Sub

Private Sub MDIForm_Load()
  Dim sRegKeyTZI As String
  Dim lLen As Long

  Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
  Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
  Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
  Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
  
  sProgramOptions.nLatitude = GetSetting(App.Title, "Location", "Latitude", 52)
  sProgramOptions.nLongitude = GetSetting(App.Title, "Location", "Longitude", 0)
  sProgramOptions.strLocation = GetSetting(App.Title, "Location", "Location", "Home QTH")
  sProgramOptions.nHeight = GetSetting(App.Title, "Location", "Height", 80)
  
  sProgramOptions.nSecondLatitude = GetSetting(App.Title, "Location", "SecondLatitude", 52)
  sProgramOptions.nSecondLongitude = GetSetting(App.Title, "Location", "SecondLongitude", 0)
  sProgramOptions.nSecondName = GetSetting(App.Title, "Location", "SecondLocation", "Second QTH")
  sProgramOptions.nSecondHeight = GetSetting(App.Title, "Location", "SecondHeight", 80)
  sProgramOptions.bSecondUsed = GetSetting(App.Title, "Location", "SecondUsed", False)
  
  
  sProgramOptions.strDefFilename = GetSetting(App.Title, "Files", "DefaultFileName", "_Default.agv")
  sProgramOptions.bSaveOpenlastView = GetSetting(App.Title, "Files", "SaveOpenLast", True)
  sProgramOptions.strDefDatabasePath = GetSetting(App.Title, "Files", "DatabasePath", "")
  sProgramOptions.strDefKepsPath = GetSetting(App.Title, "Files", "KepsPath", "")
  sProgramOptions.strKepsDatabase = GetSetting(App.Title, "Files", "Default Database", App.Path & "\Elements\Amateur.txt")
  sProgramOptions.bDaylightSaving = GetSetting(App.Title, "TimeZone", "DayLightSaving", False)
  sProgramOptions.bIndicateVis = GetSetting(App.Title, "Predictions", "Indicate Visible", True)
  
  sProgramOptions.strUserName = GetSetting(App.Title, "User", "UserName", "Not Registered")
  sProgramOptions.strCode = GetSetting(App.Title, "User", "UserCode", "Not Registered")
      
  sProgramOptions.strTimeZone = GetSetting(App.Title, "Time", "Timezone", "Not Set")
  sProgramOptions.bAutoadjust = GetSetting(App.Title, "Time", "WinAdjust", True)
  sProgramOptions.nTimezoneAdjust = GetSetting(App.Title, "Time", "Adjust", 0)
      
  sProgramOptions.nOrthX = GetSetting(App.Title, "Orth", "xSize", 500)
  sProgramOptions.nOrthY = GetSetting(App.Title, "Orth", "ySize", 500)
  sProgramOptions.bOrthShade = GetSetting(App.Title, "Orth", "Shade", True)
  sProgramOptions.strOrthLocations = GetSetting(App.Title, "Orth", "Locations", "")
  
  sProgramOptions.bSetDesktop = GetSetting(App.Title, "Desktop", "SetWallpaper", False)
  sProgramOptions.nUpdateInterval = GetSetting(App.Title, "Views", "Update", 5)
  sProgramOptions.bDisplayRangeCircle = GetSetting(App.Title, "Views", "RangeCircle", False)
  sProgramOptions.nGroundTrackPointSize = GetSetting(App.Title, "Views", "GroundTrackPointSize", 1)
  
  sProgramOptions.bFirst = GetSetting(App.Title, "User", "First", True)
  sProgramOptions.nKepsAge = GetSetting(App.Title, "Keps", "ExpireAverage", 14)
  
  sProgramOptions.bForceReadme = GetSetting(App.Title, "General", "Version", "")
  sProgramOptions.bSysTray = GetSetting(App.Title, "General", "SysTray", False)
  sProgramOptions.bSpeech = GetSetting(App.Title, "General", "Speech", False)
  sProgramOptions.nSpeechInterval = GetSetting(App.Title, "General", "SpeechInterval", 5)
  sProgramOptions.bIcons = GetSetting(App.Title, "General", "Icons", False)
  sProgramOptions.bForceReset = GetSetting(App.Title, "General", "ResetToolbars", False)
  sProgramOptions.bShowListbar = GetSetting(App.Title, "General", "ShowListbar", True)
  
  sProgramOptions.nFTPTimeout = GetSetting(App.Title, "FTP", "Timeout", 30)
  sProgramOptions.bFTPbPasvMode = GetSetting(App.Title, "FTP", "PasvMode", True)
  sProgramOptions.bFTPProxy = GetSetting(App.Title, "FTP", "FTPProxy", False)
  sProgramOptions.strFTPProxyURL = GetSetting(App.Title, "FTP", "Proxy", "")
  sProgramOptions.nFTPPPort = GetSetting(App.Title, "FTP", "Port", 80)
  sProgramOptions.bFTPThruProxy = GetSetting(App.Title, "FTP", "FTPThruProxy", False)
  sProgramOptions.bFTPRollback = GetSetting(App.Title, "FTP", "FTPRollback", False)
  sProgramOptions.nFTPRollback = GetSetting(App.Title, "FTP", "FTPRollbackSize", 1024)
  
  sProgramOptions.strFTPHTMLDir = GetSetting(App.Title, "FTP", "HTMLDir", "")
  sProgramOptions.strFTPImagesDir = GetSetting(App.Title, "FTP", "ImagesDir", "")
  sProgramOptions.strFTPPassword = GetSetting(App.Title, "FTP", "Password", "")
  sProgramOptions.strFTPServer = GetSetting(App.Title, "FTP", "Server", "")
  sProgramOptions.strFTPHTMLTemplate = GetSetting(App.Title, "FTP", "Template", "")
  sProgramOptions.strFTPUserName = GetSetting(App.Title, "FTP", "UserName", "")
    
  sProgramOptions.bRotatorEnabled = GetSetting(App.Title, "ROTATOR", "Enabled", False)
  sProgramOptions.nRotatorType = GetSetting(App.Title, "ROTATOR", "Type", 0)
  sProgramOptions.bRotatorAlwaysTrack = GetSetting(App.Title, "ROTATOR", "AlwaysTrack", False)
    
  If IsNT Then
     sRegKeyTZI = "Software\Microsoft\Windows NT\CurrentVersion\Time Zones"
  Else
     sRegKeyTZI = "Software\Microsoft\Windows\CurrentVersion\Time Zones"
  End If
  If Not GetTZICollection(sRegKeyTZI) Then
     MsgBox "Unable to locate Time Zones information in Registry under the key: " & vbCrLf & sRegKeyTZI
     Unload Me
  End If
      
  strTempPath = String(255, Chr(0))
  lLen = GetTempPath(Len(strTempPath), strTempPath)
  strTempPath = Left(strTempPath, lLen)
      
  App.HelpFile = App.Path & "\AgSatTrack.hlp"
  
  If App.Major & "." & App.Minor & "." & App.Revision <> sProgramOptions.bForceReadme Then
    frmReadme.Show vbModal
    sProgramOptions.bForceReadme = App.Major & "." & App.Minor & "." & App.Revision
  End If
  
  If KeyGen(sProgramOptions.strUserName, "6B8A-5063-205E-2A78", 3) = sProgramOptions.strCode Then
    Me.tmrReg.Enabled = False
    bRegistered = True
    Me.Caption = "AGSatTrak - Registered to " & sProgramOptions.strUserName
    Me.SSActiveToolBars1.Tools("ID_Register").Enabled = False
  Else
'    Me.Caption = "AGSatTrak - <Not Registered> " & Str(30 - nRegTimer) & " mins until closedown"
    Me.Caption = "AGSatTrak - <Not Registered> "
    frmRegInfo.Show vbModal
    Me.tmrReg.Enabled = False
  End If
      
  If sProgramOptions.bFirst Then
    frmProgramOptions.Show vbModal
    sProgramOptions.bFirst = False
  End If
  If sProgramOptions.bForceReset Then
      fMainForm.SSActiveToolBars1.LoadConfiguration App.Path & "\DefTBar.atb"
      sProgramOptions.bForceReset = False
  Else
    If FileExists(App.Path & "\toolbars.atb") Then
      fMainForm.SSActiveToolBars1.LoadLayout App.Path & "\toolbars.atb"
    End If
  End If
  
  With fMainForm.SSActiveToolBars1
    .Tools("ID_DisplayFootprints").State = ssChecked
    .Tools("ID_ToggleTrack").State = ssChecked
    .Tools("ID_0DegreeCentre").State = ssChecked
    If sProgramOptions.bShowListbar Then
      .Tools("ID_ShowShortcutBar").State = ssChecked
    End If
    .Tools("ID_PrintPreview").Enabled = False
    .Tools("ID_EnableSatelliteMode").Enabled = False
    .Tools("ID_EnableFT847").Enabled = False
   ' .Tools("ID_ForceUpload").Enabled = False
   ' .Tools("ID_FTPUpload").Enabled = False
  End With
  
  If sProgramOptions.bSaveOpenlastView Then
    If FindFile(App.Path & "\Views\_Default.agv") Then
      OpenView "_Default.agv"
    Else
      LoadNewDoc 1, True, False
    End If
  Else
      LoadNewDoc 1, True, False
  End If
  
  Set gSysTray = New clsSysTray
  Set gSysTray.SourceWindow = Me
  gSysTray.ChangeIcon ImageList1.ListImages("newicon").Picture
  ReadDX
  bFTPConnected = False
  nUploadTimer = 0
End Sub

 Sub LoadNewDoc(nType As Integer, bFlag As Boolean, bResize As Boolean)
  Dim frmD As frmDocument
  Dim frmP As frmPasses
  Dim frmG As frmGantt
  Dim strReport As String

  lDocumentCount = lDocumentCount + 1

  Select Case nType
    Case 1
      Set frmD = New frmDocument
      If wPos.lLeft + wPos.lWidth <> 0 Then
        frmD.Left = wPos.lLeft
        frmD.Top = wPos.lTop
        frmD.Width = wPos.lWidth
        frmD.Height = wPos.lHeight
      End If
      If bFlag Then
        frmD.Tag = "Select"
      End If
      frmD.Caption = "SatView " & lDocumentCount
      frmD.Show

    Case 4
      If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
        Set frmR = New frmReport
        frmR.nType = 4
        frmR.nSatIndex = fMainForm.ActiveForm.ocxSat.SetSelectedSatellite
        frmR.strCaption = "DX Visible from " & fMainForm.ActiveForm.ocxSat.SelectedSatelliteName
        Set frmR.fForm = fMainForm.ActiveForm
        frmR.UpdateReport

        frmR.Show
      End If
    Case 6
        Set frmP = New frmPasses
     '   Set frmR.fForm = fMainForm.ActiveForm
        frmP.Show
    Case 7
      Set frmG = New frmGantt
      frmG.Caption = "Gantt View"
      frmG.Show
  End Select
  fMainForm.UpdateToolbar
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bRegistered Then
    If UnloadMode <> vbAppWindows Then
      If MsgBox("Are you sure you wish to exit.", vbQuestion + vbYesNoCancel + vbDefaultButton1, "AG SatTrack") = vbYes Then
      Else
        Cancel = 1
      End If
    End If
  End If
  If Not Cancel Then
    If sProgramOptions.bSaveOpenlastView Then
      SaveView "_Default.agv"
    End If
    gSysTray.RemoveFromSysTray
  End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized And sProgramOptions.bSysTray Then
        gSysTray.MinToSysTray
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
  End If
  
  SaveSetting App.Title, "Location", "Latitude", sProgramOptions.nLatitude
  SaveSetting App.Title, "Location", "Longitude", sProgramOptions.nLongitude
  SaveSetting App.Title, "Location", "Location", sProgramOptions.strLocation
  SaveSetting App.Title, "Location", "Height", sProgramOptions.nHeight
  
  SaveSetting App.Title, "Location", "SecondLatitude", sProgramOptions.nSecondLatitude
  SaveSetting App.Title, "Location", "SecondLongitude", sProgramOptions.nSecondLongitude
  SaveSetting App.Title, "Location", "SecondLocation", sProgramOptions.nSecondName
  SaveSetting App.Title, "Location", "SecondHeight", sProgramOptions.nSecondHeight
  SaveSetting App.Title, "Location", "SecondUsed", sProgramOptions.bSecondUsed
    
  SaveSetting App.Title, "Files", "DefaultFileName", sProgramOptions.strDefFilename
  SaveSetting App.Title, "Files", "SaveOpenLast", sProgramOptions.bSaveOpenlastView
  SaveSetting App.Title, "Files", "DatabasePath", sProgramOptions.strDefDatabasePath
  SaveSetting App.Title, "Files", "KepsPath", sProgramOptions.strDefKepsPath
  SaveSetting App.Title, "Files", "Default Database", sProgramOptions.strKepsDatabase
  SaveSetting App.Title, "TimeZone", "DayLightSaving", sProgramOptions.bDaylightSaving
  SaveSetting App.Title, "Predictions", "Indicate Visible", sProgramOptions.bIndicateVis
  SaveSetting App.Title, "Time", "Timezone", sProgramOptions.strTimeZone
  SaveSetting App.Title, "Time", "WinAdjust", sProgramOptions.bAutoadjust
  SaveSetting App.Title, "Time", "Adjust", sProgramOptions.nTimezoneAdjust
  SaveSetting App.Title, "Orth", "xSize", sProgramOptions.nOrthX
  SaveSetting App.Title, "Orth", "ySize", sProgramOptions.nOrthY
  SaveSetting App.Title, "Orth", "Shade", sProgramOptions.bOrthShade
  SaveSetting App.Title, "Orth", "Locations", sProgramOptions.strOrthLocations
  SaveSetting App.Title, "Desktop", "SetWallpaper", sProgramOptions.bSetDesktop
  SaveSetting App.Title, "Views", "Update", sProgramOptions.nUpdateInterval
  SaveSetting App.Title, "Views", "RangeCircle", sProgramOptions.bDisplayRangeCircle
  SaveSetting App.Title, "Views", "GroundTrackPointSize", sProgramOptions.nGroundTrackPointSize

  SaveSetting App.Title, "User", "First", sProgramOptions.bFirst
  SaveSetting App.Title, "Keps", "ExpireAverage", sProgramOptions.nKepsAge
  SaveSetting App.Title, "General", "Version", sProgramOptions.bForceReadme
  SaveSetting App.Title, "General", "SysTray", sProgramOptions.bSysTray
  SaveSetting App.Title, "General", "Speech", sProgramOptions.bSpeech
  SaveSetting App.Title, "General", "SpeechInterval", sProgramOptions.nSpeechInterval
  SaveSetting App.Title, "General", "Icons", sProgramOptions.bIcons
  SaveSetting App.Title, "General", "ResetToolbars", sProgramOptions.bForceReset
  SaveSetting App.Title, "General", "ShowListbar", sProgramOptions.bShowListbar
    
  SaveSetting App.Title, "FTP", "Timeout", sProgramOptions.nFTPTimeout
  SaveSetting App.Title, "FTP", "PasvMode", sProgramOptions.bFTPbPasvMode
  SaveSetting App.Title, "FTP", "FTPProxy", sProgramOptions.bFTPProxy
  SaveSetting App.Title, "FTP", "Proxy", sProgramOptions.strFTPProxyURL
  SaveSetting App.Title, "FTP", "Port", sProgramOptions.nFTPPPort
  SaveSetting App.Title, "FTP", "FTPThruProxy", sProgramOptions.bFTPThruProxy
  SaveSetting App.Title, "FTP", "FTPRollback", sProgramOptions.bFTPRollback
  SaveSetting App.Title, "FTP", "FTPRollbackSize", sProgramOptions.nFTPRollback
  
  SaveSetting App.Title, "FTP", "HTMLDir", sProgramOptions.strFTPHTMLDir
  SaveSetting App.Title, "FTP", "ImagesDir", sProgramOptions.strFTPImagesDir
  SaveSetting App.Title, "FTP", "Password", sProgramOptions.strFTPPassword
  SaveSetting App.Title, "FTP", "Server", sProgramOptions.strFTPServer
  SaveSetting App.Title, "FTP", "Template", sProgramOptions.strFTPHTMLTemplate
  SaveSetting App.Title, "FTP", "UserName", sProgramOptions.strFTPUserName
  
  SaveSetting App.Title, "ROTATOR", "Enabled", sProgramOptions.bRotatorEnabled
  SaveSetting App.Title, "ROTATOR", "Type", sProgramOptions.nRotatorType
  SaveSetting App.Title, "ROTATOR", "AlwaysTrack", sProgramOptions.bRotatorAlwaysTrack
    
  If bRegistered Then
    SaveSetting App.Title, "User", "UserName", sProgramOptions.strUserName
    SaveSetting App.Title, "User", "UserCode", sProgramOptions.strCode
  End If
  
  If FindFile(strTempPath & "\AGGlobe.bmp") Then
    Kill strTempPath & "\AGGlobe.bmp"
  End If
  If FindFile(strTempPath & "\AGMerc.bmp") Then
    Kill strTempPath & "\AGMerc.bmp"
  End If
  
  fMainForm.SSActiveToolBars1.SaveLayout App.Path & "\toolbars.atb"
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
  frmOptions.Show vbModal, Me
End Sub


Private Sub mnuHelpContents_Click()

  Dim nRet As Integer

  'if there is no helpfile for this project display a message to the user
  'you can set the HelpFile for your application in the
  'Project Properties dialog
  If Len(App.HelpFile) = 0 Then
    MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
  Else
    On Error Resume Next
    nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
    If Err Then
      MsgBox Err.Description
    End If
  End If
End Sub

Private Sub mnuHelpSearch_Click()

  Dim nRet As Integer

  'if there is no helpfile for this project display a message to the user
  'you can set the HelpFile for your application in the
  'Project Properties dialog
  If Len(App.HelpFile) = 0 Then
    MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
  Else
    On Error Resume Next
    nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
    If Err Then
      MsgBox Err.Description
    End If
  End If
End Sub

Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub

Private Sub mnuFileOpen_Click()
  On Error GoTo ERROR_mnuFileOpen_Click

  With Me.dlgCommonDialog
    .CancelError = True
    .DialogTitle = "Open Views"
    .FileName = App.Path & "\Views\*.agv"
    .ShowOpen
    OpenView GetFilename(.FileName)
  End With
  
EXIT_mnuFileOpen_Click:
  Exit Sub

ERROR_mnuFileOpen_Click:
  Select Case Err
    Case cdlCancel
    Case Else
      MsgBox "Error in ERROR_mnuFileOpen_Click : " & Error
  End Select
  Resume EXIT_mnuFileOpen_Click
End Sub

Private Sub mnuFileSave_Click()

  On Error GoTo ERROR_mnuFileSave_Click

  With Me.dlgCommonDialog
    .CancelError = True
    .DialogTitle = "Save Views as"
    .Filter = "View Files (*.agv)|*.agv|All Filed|*.*"
    .FileName = App.Path & "\Views\*.agv"
    .ShowSave
    SaveView GetFilename(.FileName)
  End With
  
EXIT_mnuFileSave_Click:
  Exit Sub

ERROR_mnuFileSave_Click:
  Select Case Err
    Case cdlCancel
    Case Else
      MsgBox "Error in ERROR_mnuFileSave_Click : " & Error
  End Select
  Resume EXIT_mnuFileSave_Click
End Sub
            
Private Sub mnuFileSavePredictions_Click()
  Dim nFile As Integer
  Dim i As Integer
  
  On Error GoTo ERROR_mnuFileSavePredictions_Click

  With Me.dlgCommonDialog
    .CancelError = True
    .DialogTitle = "Save Prediction as"
    .Filter = "Text Files (*.txt)|*.txt|All Filed|*.*"
    .FileName = App.Path & "\*.txt"
    .ShowSave
    nFile = FreeFile
    Open .FileName For Output As #nFile
      For i = 0 To fMainForm.ActiveForm.lstData.ListCount - 1
        Print #nFile, fMainForm.ActiveForm.lstData.List(i)
      Next i
    Close #nFile
  End With
  
EXIT_mnuFileSavePredictions_Click:
  Exit Sub

ERROR_mnuFileSavePredictions_Click:
  Select Case Err
    Case cdlCancel
    Case Else
      MsgBox "Error in ERROR_mnuFileSavePredictions_Click : " & Error
  End Select
  Resume EXIT_mnuFileSavePredictions_Click
End Sub

Private Sub mnuFilePageSetup_Click()
  dlgCommonDialog.ShowPrinter
End Sub

Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub

Private Sub mnuFileNew_Click()
  
  LoadNewDoc 1, True, False
End Sub


Private Sub SSActiveToolBars1_ComboCloseUp(ByVal Tool As ActiveToolBars.SSTool)
  Select Case Tool.ID
    Case "ID_Orbits"
      fMainForm.ActiveForm.ocxSat.SatelliteTrackOrbits = Tool.ComboBox.Text
      fMainForm.ActiveForm.UpdatePosTable
      fMainForm.ActiveForm.ocxSat.DrawFootprints
    Case "ID_GroundsTrackInterval"
      Select Case Tool.ComboBox.Text
        Case "15 Seconds"
          fMainForm.ActiveForm.ocxSat.GroundTrackInterval = ag15
        Case "30 Seconds"
          fMainForm.ActiveForm.ocxSat.GroundTrackInterval = ag30
        Case "60 Seconds"
          fMainForm.ActiveForm.ocxSat.GroundTrackInterval = ag60
        Case "Automatic"
          fMainForm.ActiveForm.ocxSat.GroundTrackInterval = agAuto
      End Select
      fMainForm.ActiveForm.UpdatePosTable
      fMainForm.ActiveForm.ocxSat.DrawFootprints
  End Select

End Sub

Private Sub SSActiveToolBars1_ToolChange(ByVal Tool As ActiveToolBars.SSTool)
  If Tool.ID = "ID_Orbits" Then
    Beep
  End If

End Sub

Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
  Dim nStart As Integer
  Dim nEnd As Integer
  Dim i As Integer
  Dim fForm As Form
  Dim strCaption As String
  
  ' NOTE: Place this code in the control's ToolClick event.
  '
  ' Determine which tool was clicked.
  If Not bUpdating Then
    Select Case Tool.ID
      Case "ID_ExperimentalSatellite"
          If Not (fMainForm.ActiveForm Is Nothing) Then
            If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
              frmSatDetails.Tag = "New"
              frmSatDetails.Show vbModal
            End If
          End If
      Case "ID_FTPUpload"
        Select Case Tool.State
          Case ssChecked
            With sProgramOptions
              If .strFTPImagesDir = "" Or .strFTPPassword = "" Or .strFTPServer = "" Or .strFTPUserName = "" Then
                MsgBox "Please ensure that all of the FTP upload options are set in the program options before attempting to upload images.", vbCritical + vbOKOnly, "FTP Upload error"
                Me.SSActiveToolBars1.Tools("ID_FTPUpload").State = ssUnchecked
              Else
                Me.tmrFTP.Enabled = True
              End If
            End With
          Case ssUnchecked
            Me.tmrFTP.Enabled = False
        End Select
      Case "ID_ForceUpload"
        With sProgramOptions
          If .strFTPImagesDir = "" Or .strFTPPassword = "" Or .strFTPServer = "" Or .strFTPUserName = "" Then
            MsgBox "Please ensure that all of the FTP upload options are set in the program options before attempting to upload images.", vbCritical + vbOKOnly, "FTP Upload error"
          Else
            nUploadTimer = 9999
            tmrFTP_Timer
          End If
        End With
      Case "ID_ShowShortcutBar"
        sProgramOptions.bShowListbar = Tool.State
        If Not (fMainForm.ActiveForm Is Nothing) Then
          For Each fForm In Forms
            With fForm
              If Left(.Caption, 7) = "SatView" Then
                fForm.Resize
              End If
            End With
          Next
        End If
      Case "ID_Rotator"
          If Not (fMainForm.ActiveForm Is Nothing) Then
            If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
              If Tool.State = ssChecked Then
                If frmRotatorForm Is Nothing Then
                  Set frmRotatorForm = fMainForm.ActiveForm
                  frmRotatorSetup.Show vbModal
                Else
                  MsgBox "The rotator interface is in use on another form", vbCritical + vbOKOnly, "Rotator"
                End If
              Else
                If Not (frmRotatorForm Is Nothing) Then
                  bMoveRotator = False
                  frmRotatorForm.CloseRotatorLink
                  Set frmRotatorForm = Nothing
                End If
              End If
            End If
          End If
          
        Case "ID_PrintPreview"
          If Not (fMainForm.ActiveForm Is Nothing) Then
            Set frmForm = fMainForm.ActiveForm
            frmPrintPreview.Show vbModal
            Set frmForm = Nothing
          End If
        Case "ID_MoveSatellites"
            frmEditGroups.Show vbModal
'    Case "id_SatelliteLabel"
'        If Not (fMainForm.ActiveForm Is Nothing) Then
'          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
'            With fMainForm.ActiveForm
'              .ocxSat.DisplaySatelliteLabel = IIf(Tool.State = ssChecked, True, False)
'              .bViewSatLabel = IIf(Tool.State = ssChecked, True, False)
'            End With
'            fMainForm.ActiveForm.ocxSat.DrawFootprints
'          End If
'        End If
'    Case "id_ViewStatusBar"
'        If Not (fMainForm.ActiveForm Is Nothing) Then
'          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
'            With fMainForm.ActiveForm
'              .ocxSat.DisplayStatusBar = IIf(Tool.State = ssChecked, True, False)
'              .bViewStatusBar = IIf(Tool.State = ssChecked, True, False)
'            End With
'            fMainForm.ActiveForm.ocxSat.DrawFootprints
'          End If
'        End If
    
      Case "ID_PointsOrLines"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
            If Tool.State = ssChecked Then
              fMainForm.ActiveForm.ocxSat.DisplayGroundTrackAsPoints = False
            Else
              fMainForm.ActiveForm.ocxSat.DisplayGroundTrackAsPoints = True
            End If
            fMainForm.ActiveForm.UpdatePosTable
            fMainForm.ActiveForm.ocxSat.DrawFootprints
          End If
        End If
        
      Case "ID_ResetToolbars"
        If MsgBox("Are you sure you wish to reset the toolbars. If you answer yes then you will have to restart AGSatTrack to access the new menus.", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm toolbar reset") = vbYes Then
          sProgramOptions.bForceReset = True
          Unload Me
        End If
      Case "ID_EnableSpeech"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
            If Tool.State = ssChecked Then
              fMainForm.ActiveForm.bSpeech = True
            Else
              fMainForm.ActiveForm.bSpeech = False
            End If
          End If
        End If
      Case "ID_ViewReadme"
        frmReadme.Show vbModal
      Case "ID_ViewElements"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
            ReadKeps fMainForm.ActiveForm.ocxSat.DatabasePath
            nOldSat = fMainForm.ActiveForm.ocxSat.SatelliteIndex
            fMainForm.ActiveForm.ocxSat.SatelliteIndex = fMainForm.ActiveForm.ocxSat.SetSelectedSatellite
            For i = 0 To UBound(sKeps)
              If fMainForm.ActiveForm.ocxSat.SatelliteDesignator = sKeps(i).lDesignator Then
                fMainForm.ActiveForm.ocxSat.SatelliteIndex = nOldSat
                frmSatDetails.Tag = i
                frmSatDetails.Show vbModal
                Exit For
              End If
            Next i
            fMainForm.ActiveForm.ocxSat.SatelliteIndex = nOldSat
          End If
        End If
      Case "ID_EditGroups"
        frmEditElementGroups.Show vbModal
      Case "ID_Desktop"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
            For Each fForm In Forms
              With fForm
                If Left(.Caption, 7) = "SatView" Then
                  .ocxSat.SetActiveWindowAsWallpaper = False
                  .bOnDesktop = False
                
                End If
              End With
            Next
            If Tool.State = ssChecked Then
              fMainForm.ActiveForm.bOnDesktop = True
              fMainForm.ActiveForm.ocxSat.SetActiveWindowAsWallpaper = True
            Else
              fMainForm.ActiveForm.bOnDesktop = False
              fMainForm.ActiveForm.ocxSat.SetActiveWindowAsWallpaper = False
            End If
          End If
        End If
      Case "ID_Register"
        frmRegister.Show vbModal
      Case "ID_RegistrationInformation"
        frmRegInfo.Show vbModal
      Case "id_FileOpen"    '(Button)
        mnuFileOpen_Click
      Case "id_FileSave"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          strCaption = fMainForm.ActiveForm.Caption
          If strCaption = "Predictions" Then
            mnuFileSavePredictions_Click
          Else
            mnuFileSave_Click
          End If
        End If
      Case "ID_UpdatefromtheInternet"
'        frmInetUpdate.Show vbModal
        frmFTPMain.Show vbModal
      Case "id_FileSaveAs"    '(Button)
        mnuFileSave_Click
      Case "id_FileProperties"    '(Button)
        frmProgramOptions.Show vbModal
      Case "id_FileExit"    '(Button)
        Unload Me
      Case "id_WindowNewWindow"    '(Button)
        LoadNewDoc 1, True, False
      Case "id_WindowCascade"    '(Button)
        mnuWindowCascade_Click
      Case "id_WindowTileHorizontal"    '(Button)
        mnuWindowTileHorizontal_Click
      Case "id_WindowTileVertical"    '(Button)
        mnuWindowTileVertical_Click
      Case "id_WindowArrangeIcons"    '(Button)
        mnuWindowArrangeIcons_Click
      Case "id_Help"    '(Menu)
      Case "id_HelpContents"    '(Button)
        mnuHelpContents_Click
      Case "id_HelpSearchForHelpOn"    '(Button)
        mnuHelpSearch_Click
      Case "id_HelpAboutAGSatTrack"    '(Button)
        mnuHelpAbout_Click
      Case "ID_TrackingView"    '(Button)
        LoadNewDoc 1, True, False
      Case "ID_DXReport"
        LoadNewDoc 4, True, False
      Case "ID_Predictions"
        LoadNewDoc 6, True, False
      Case "ID_GanttView"
        LoadNewDoc 7, True, False
      Case "ID_NextAOS"
        nStart = 0
        nEnd = 0
        If fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text = "All Satellites" Then
          nStart = 1
          nEnd = fMainForm.ActiveForm.ocxSat.SatelliteCount
        Else
          For i = 1 To fMainForm.ActiveForm.ocxSat.SatelliteCount
            fMainForm.ActiveForm.ocxSat.SatelliteIndex = i
            If fMainForm.ActiveForm.ocxSat.SatelliteName = fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text Then
              nStart = i
              nEnd = i
              Exit For
            End If
          Next i
        End If
        If nStart <> 0 Then
          For i = nStart To nEnd
            fMainForm.ActiveForm.ocxSat.DisplayAOS 2, i
            fMainForm.ActiveForm.UpdateToolBarTime
          Next i
          fMainForm.ActiveForm.ocxSat.CalculateALLPositions
          fMainForm.ActiveForm.ocxSat.DrawFootprints
        End If
      Case "ID_PreviousAOS"
        nStart = 0
        nEnd = 0
        If fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text = "All Satellites" Then
          nStart = 1
          nEnd = fMainForm.ActiveForm.ocxSat.SatelliteCount
        Else
          For i = 1 To fMainForm.ActiveForm.ocxSat.SatelliteCount
            fMainForm.ActiveForm.ocxSat.SatelliteIndex = i
            If fMainForm.ActiveForm.ocxSat.SatelliteName = fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text Then
              nStart = i
              nEnd = i
              Exit For
            End If
          Next i
        End If
        If nStart <> 0 Then
          For i = nStart To nEnd
            fMainForm.ActiveForm.ocxSat.DisplayAOS 1, i
            fMainForm.ActiveForm.UpdateToolBarTime
          Next i
          fMainForm.ActiveForm.ocxSat.CalculateALLPositions
          fMainForm.ActiveForm.ocxSat.DrawFootprints
        End If
      Case "ID_ResetTime"
        nStart = 1
        nEnd = fMainForm.ActiveForm.ocxSat.SatelliteCount
        For i = nStart To nEnd
          fMainForm.ActiveForm.ocxSat.SatelliteIndex = i
          fMainForm.ActiveForm.ocxSat.ResetSatellite
          fMainForm.ActiveForm.UpdateToolBarTime
        Next i
      Case "ID_Stop"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.nSatSpeed = 0
        End If
      Case "ID_BackFast"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.nSatSpeed = -10
        End If
      Case "ID_BackSlow"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.nSatSpeed = -1
        End If
      Case "ID_ForwardSlow"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.nSatSpeed = 1
        End If
      Case "ID_ForwardFast"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.nSatSpeed = 10
        End If

      Case "ID_DisplayMoon"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.DisplayMoon = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.DisplayMoon = False
            fMainForm.SSActiveToolBars1.Tools("ID_DisplayMoonFootprint").State = ssUnchecked
        End Select
        fMainForm.ActiveForm.ocxSat.DrawFootprints

      Case "ID_DisplayMoonFootprint"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.DisplayMoonFootprint = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.DisplayMoonFootprint = False
        End Select
        fMainForm.ActiveForm.ocxSat.DrawFootprints

      Case "ID_DisplaySun"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.DisplaySun = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.DisplaySun = False
            fMainForm.SSActiveToolBars1.Tools("ID_DisplaySunFootprint").State = ssUnchecked
        End Select
        fMainForm.ActiveForm.ocxSat.DrawFootprints

      Case "ID_DisplaySunFootprint"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.DisplaySunFootprint = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.DisplaySunFootprint = False
        End Select
        fMainForm.ActiveForm.ocxSat.DrawFootprints

      Case "ID_0DegreeCentre"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.ocxSat.OutputStyle = 1
          fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0
        End If
      
      Case "ID_180DegreeCentre"    '(Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          fMainForm.ActiveForm.ocxSat.OutputStyle = 1
          fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 180
        End If
      
      Case "ID_TableView"    '(Button)
        fMainForm.ActiveForm.ocxSat.OutputStyle = 0
      Case "ID_HorizonView"    '(Button)
        fMainForm.ActiveForm.ocxSat.OutputStyle = 2
      Case "ID_Globe"    '(Button)
        fMainForm.ActiveForm.ocxSat.OutputStyle = 3
      Case "ID_DisplayFootprints"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If fMainForm.SSActiveToolBars1.Tools("ID_DisplayFootprints").State = ssChecked Then
            fMainForm.ActiveForm.ocxSat.DisplayFootprints = True
          Else
            fMainForm.ActiveForm.ocxSat.DisplayFootprints = False
          End If
          fMainForm.ActiveForm.ocxSat.DrawFootprints
        End If
      Case "ID_SelectSatellites"
        If FileExists(sProgramOptions.strKepsDatabase) Then
          frmSelect.Show vbModal
          If Left$(fMainForm.ActiveForm.Caption, 5) = "Gantt" Then
            fMainForm.ActiveForm.UpdateGannt
          End If
        Else
          Call MsgBox("The selected database does not exist. Please import some keplarian elements.", vbExclamation + vbOKOnly + vbDefaultButton1, "Database Error")
        End If
      Case "ID_ToggleTrack"    '(State Button)
        If Not (fMainForm.ActiveForm Is Nothing) Then
          Select Case Tool.State
            Case ssChecked
              fMainForm.ActiveForm.ocxSat.DisplayTracks = True
            Case ssUnchecked
              fMainForm.ActiveForm.ocxSat.DisplayTracks = False
          End Select
          fMainForm.ActiveForm.ocxSat.CalculateALLPositions
          fMainForm.ActiveForm.ocxSat.DrawFootprints
        End If
      Case "ID_Refresh"
        If Not (fMainForm.ActiveForm Is Nothing) Then
          If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
            fMainForm.ActiveForm.UpdatePosTable
            fMainForm.ActiveForm.ocxSat.DrawFootprints
          Else
            fMainForm.ActiveForm.UpdateReport
          End If
        End If
      Case "ID_WindowProperties"
        frmWindowOptions.Show vbModal

      Case "ID_EnableFT847"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.Enable847 = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.Enable847 = False
        End Select

      Case "ID_EnableSatelliteMode"    '(State Button)
        Select Case Tool.State
          Case ssChecked
            fMainForm.ActiveForm.ocxSat.Enable847Sat = True
          Case ssUnchecked
            fMainForm.ActiveForm.ocxSat.Enable847Sat = False
        End Select

      Case Else

    End Select
  End If
End Sub

Private Sub SSActiveToolBars1_ToolGotFocus(ByVal Tool As ActiveToolBars.SSTool)
  ' frmDateSelect.Show
End Sub

Public Sub UpdateToolbar()

  If fMainForm.ActiveForm Is Nothing Then
    Exit Sub
  End If

  With fMainForm.SSActiveToolBars1
    bUpdating = True
    .Redraw = False
    If Left$(fMainForm.ActiveForm.Caption, 7) = "SatView" Then
      .Tools("ID_Rotator").Enabled = True
      .Tools("ID_PointsOrLines").Enabled = True
      .Tools("ID_NextAOS").Enabled = True
      .Tools("ID_PreviousAOS").Enabled = True
      .Tools("ID_ResetTime").Enabled = True
      .Tools("ID_Stop").Enabled = True
      .Tools("ID_BackFast").Enabled = True
      .Tools("ID_BackSlow").Enabled = True
      .Tools("ID_ForwardSlow").Enabled = True
      .Tools("ID_ForwardFast").Enabled = True
      .Tools("ID_DisplayMoon").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplayMoon Then
        .Tools("ID_DisplayMoon").State = ssChecked
      Else
        .Tools("ID_DisplayMoon").State = ssUnchecked
      End If
      .Tools("ID_DisplayMoonFootprint").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplayMoonFootprint Then
        .Tools("ID_DisplayMoonFootprint").State = ssChecked
      Else
        .Tools("ID_DisplayMoonFootprint").State = ssUnchecked
      End If
      .Tools("ID_DisplaySun").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplaySun Then
        .Tools("ID_DisplaySun").State = ssChecked
      Else
        .Tools("ID_DisplaySun").State = ssUnchecked
      End If
      .Tools("ID_DisplaySunFootprint").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplaySunFootprint Then
        .Tools("ID_DisplaySunFootprint").State = ssChecked
      Else
        .Tools("ID_DisplaySunFootprint").State = ssUnchecked
      End If
      .Tools("ID_0DegreeCentre").Enabled = True
      .Tools("ID_180DegreeCentre").Enabled = True
      .Tools("ID_TableView").Enabled = True
      .Tools("ID_HorizonView").Enabled = True
      .Tools("ID_Globe").Enabled = True
      .Tools("ID_DisplayFootprints").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplayFootprints Then
        .Tools("ID_DisplayFootprints").State = ssChecked
      Else
        .Tools("ID_DisplayFootprints").State = ssUnchecked
      End If
      .Tools("ID_SelectSatellites").Enabled = True
      .Tools("ID_ToggleTrack").Enabled = True
      If fMainForm.ActiveForm.ocxSat.DisplayTracks Then
        .Tools("ID_ToggleTrack").State = ssChecked
      Else
        .Tools("ID_ToggleTrack").State = ssUnchecked
      End If
      .Tools("ID_WindowProperties").Enabled = True
      Select Case fMainForm.ActiveForm.ocxSat.OutputStyle
        Case 0
          .Tools("ID_TableView").State = ssChecked
        Case 1
          If fMainForm.ActiveForm.ocxSat.ObserverMapCentre = 0 Then
            .Tools("ID_0DegreeCentre").State = ssChecked
          Else
            .Tools("ID_180DegreeCentre").State = ssChecked
          End If
        Case 2
          .Tools("ID_HorizonView").State = ssChecked
        Case 3
          .Tools("ID_Globe").State = ssChecked
        Case 4
      End Select
      Select Case fMainForm.ActiveForm.nSatSpeed
        Case -10
          .Tools("ID_BackFast").State = ssChecked
        Case -1
          .Tools("ID_BackSlow").State = ssChecked
        Case 0
          .Tools("ID_Stop").State = ssChecked
        Case 1
          .Tools("ID_ForwardSlow").State = ssChecked
        Case 10
          .Tools("ID_ForwardFast").State = ssChecked
      End Select
      If fMainForm.ActiveForm.bOnDesktop Then
        .Tools("ID_Desktop").State = ssChecked
      Else
        .Tools("ID_Desktop").State = ssUnchecked
      End If

      .Tools("ID_EnableSpeech").Enabled = True
      If fMainForm.ActiveForm.bSpeech Then
        .Tools("ID_EnableSpeech").State = ssChecked
      Else
        .Tools("ID_EnableSpeech").State = ssUnchecked
      End If
      .Tools("ID_Orbits").Enabled = True
      
      .Tools("ID_GroundsTrackInterval").Enabled = True
      .Tools("ID_Desktop").Enabled = True
     ' .Tools("ID_EnableFT847").Enabled = True
     ' .Tools("ID_EnableSatelliteMode").Enabled = True
    Else
      .Tools("ID_Rotator").Enabled = False
      .Tools("ID_PointsOrLines").Enabled = False
      .Tools("ID_EnableSatelliteMode").Enabled = False
      .Tools("ID_EnableFT847").Enabled = False
      .Tools("ID_Desktop").Enabled = False
      .Tools("ID_GroundsTrackInterval").Enabled = False
      .Tools("ID_EnableSpeech").Enabled = False
      .Tools("ID_Orbits").Enabled = False
      .Tools("ID_NextAOS").Enabled = False
      .Tools("ID_PreviousAOS").Enabled = False
      .Tools("ID_ResetTime").Enabled = False
      .Tools("ID_Stop").Enabled = False
      .Tools("ID_BackFast").Enabled = False
      .Tools("ID_BackSlow").Enabled = False
      .Tools("ID_ForwardSlow").Enabled = False
      .Tools("ID_ForwardFast").Enabled = False
      .Tools("ID_DisplayMoon").Enabled = False
      .Tools("ID_DisplayMoonFootprint").Enabled = False
      .Tools("ID_DisplaySun").Enabled = False
      .Tools("ID_DisplaySunFootprint").Enabled = False
      .Tools("ID_0DegreeCentre").Enabled = False
      .Tools("ID_180DegreeCentre").Enabled = False
      .Tools("ID_TableView").Enabled = False
      .Tools("ID_HorizonView").Enabled = False
      .Tools("ID_Globe").Enabled = False
      .Tools("ID_DisplayFootprints").Enabled = False
      .Tools("ID_SelectSatellites").Enabled = False
      .Tools("ID_ToggleTrack").Enabled = False
      .Tools("ID_WindowProperties").Enabled = False
    End If
    If Left$(fMainForm.ActiveForm.Caption, 5) = "Gantt" Then
      .Tools("ID_SelectSatellites").Enabled = True
    End If
    If lDocumentCount = 0 Then
      .Tools("ID_Refresh").Enabled = False
    Else
      .Tools("ID_Refresh").Enabled = True
    End If
    .Redraw = True
  bUpdating = False
  
  End With
End Sub

Private Sub SSActiveToolBars1_ToolLostFocus(ByVal Tool As ActiveToolBars.SSTool)
  Dim vTime As Variant
  Dim vDate As Variant
  Dim nStart As Integer
  Dim nEnd As Integer
  Dim FormattedDateTime As String
  
  Select Case Tool.ID
    Case "ID_SatelliteTime", "ID_SatelliteDate"
      vTime = fMainForm.SSActiveToolBars1.Tools("ID_SatelliteTime").Edit.Text
      vDate = fMainForm.SSActiveToolBars1.Tools("ID_satelliteDate").Edit.Text
      nStart = 0
      nEnd = 0
      If fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text = "All Satellites" Then
        nStart = 1
        nEnd = fMainForm.ActiveForm.ocxSat.SatelliteCount
      Else
        For i = 1 To fMainForm.ActiveForm.ocxSat.SatelliteCount
          fMainForm.ActiveForm.ocxSat.SatelliteIndex = i
          If fMainForm.ActiveForm.ocxSat.SatelliteName = fMainForm.SSActiveToolBars1.Tools("ID_Satellite").ComboBox.Text Then
            nStart = i
            nEnd = i
            Exit For
          End If
        Next i
      End If
      If nStart <> 0 Then
        For i = nStart To nEnd
          FormattedDateTime$ = Format$(vDate & " " & vTime, "yyyymmddhhmmss")
          fMainForm.ActiveForm.ocxSat.DisplayCentury = Val(Left$(FormattedDateTime$, 2))
          fMainForm.ActiveForm.ocxSat.DisplayYear = Val(Left$(FormattedDateTime$, 4))
          fMainForm.ActiveForm.ocxSat.DisplayMonth = Val(Mid$(FormattedDateTime$, 5, 2))
          fMainForm.ActiveForm.ocxSat.DisplayDay = Val(Mid$(FormattedDateTime$, 7, 2))
          fMainForm.ActiveForm.ocxSat.DisplayHour = Val(Mid$(FormattedDateTime$, 9, 2))
          fMainForm.ActiveForm.ocxSat.DisplayMinute = Val(Mid$(FormattedDateTime$, 11, 2))
          fMainForm.ActiveForm.ocxSat.DisplaySecond = Val(Mid$(FormattedDateTime$, 13, 2))
              
          fMainForm.ActiveForm.ocxSat.SatelliteBusy = True
        Next i
        fMainForm.ActiveForm.ocxSat.CalculateALLPositions
        fMainForm.ActiveForm.ocxSat.DrawFootprints
      End If
    Case "ID_Satellite"
  End Select
End Sub

Private Sub SysTrayExit_Click()
  Unload Me
End Sub

Private Sub SysTrayRestore_Click()
  Me.WindowState = vbNormal
  Me.Show
  App.TaskVisible = True
  gSysTray.RemoveFromSysTray
End Sub
Private Sub gSysTray_RButtonUP()
    PopupMenu Me.mnuSysMenu
End Sub

Private Sub tmrFTP_Timer()
Dim fForm As Form
Dim nView As Integer
Dim nCount As Integer
Dim strFilename As String
Dim strFilenameBMP As String
Dim strFilenameGIF As String
Dim ctrImage As New cRVTVBIMG
Dim lResult As Long
Dim strTempPath As String
Dim lLen As Long
Dim nTotal As Integer
Dim nTotalGlobes As Integer
Dim nTotalMaps As Integer
Dim nTotalHorizons As Integer
Dim strPage As String

  If Not bFTPConnected Then
  nUploadTimer = nUploadTimer + 1
  If nUploadTimer > 15 Then
    If Not fMainForm.ActiveForm Is Nothing Then
      bFTPConnected = True
      Me.stbStatusBar1.Panels(1).Text = "Connecting to Web Server"
      If Me.FTP1.Connect(App.Title, sProgramOptions.strFTPServer, "21", sProgramOptions.strFTPUserName, sProgramOptions.strFTPPassword) Then
        If sProgramOptions.strFTPHTMLDir <> "" Then
          Me.FTP1.SetDirectory sProgramOptions.strFTPHTMLDir
        End If
        If sProgramOptions.strFTPHTMLTemplate <> "" Then
          strPage = ProcessWebPage(App.Path & "\Web Pages\" & sProgramOptions.strFTPHTMLTemplate)
          Me.FTP1.UploadFile sProgramOptions.strFTPHTMLTemplate, strPage
        End If
        Me.FTP1.SetDirectory sProgramOptions.strFTPImagesDir
        Me.stbStatusBar1.Panels(2).Text = "Uploading Image(s) Web Server"
        strTempPath = String$(255, " ")
        lLen = GetTempPath(Len(strTempPath), strTempPath)
        strTempPath = Left(strTempPath, lLen)
        For Each fForm In Forms
          With fForm
            If Left(.Caption, 7) = "SatView" Then
              nView = fForm.ocxSat.OutputStyle
              Select Case nView
                Case 1
                  nTotal = nTotal + 1
                Case 2
                  nTotal = nTotal + 1
                Case 3
                  nTotal = nTotal + 1
                Case Else
                  strFilename = ""
              End Select
            End If
          End With
        Next
        nCount = 1
        nTotalGlobes = 1
        nTotalMaps = 1
        nTotalHorizons = 1
        For Each fForm In Forms
          Me.stbStatusBar1.Panels(2).Text = "Uploading Image" & Str(nCount) & " of " & Str(nTotal)
          With fForm
            If Left(.Caption, 7) = "SatView" Then
              nView = fForm.ocxSat.OutputStyle
              Select Case nView
                Case 1
                  strFilename = "sat" & Trim(Str(nTotalMaps))
                  nTotalMaps = nTotalMaps + 1
                Case 2
                  strFilename = "horizon" & Trim(Str(nTotalHorizons))
                  nTotalHorizons = nTotalHorizons + 1
                Case 3
                  strFilename = "Globe" & Trim(Str(nTotalGlobes))
                  nTotalGlobes = nTotalGlobes + 1
                Case Else
                  strFilename = ""
              End Select
              If strFilename <> "" Then
                strFilenameGIF = strFilename & ".gif"
                strFilenameBMP = strTempPath & "\" & strFilename & ".bmp"
                SavePicture fForm.ocxSat.Picture, strFilenameBMP
                lResult = ctrImage.SetPipeline(0, 0, 1, 5, 0)
                lResult = ctrImage.LoadAsBMP(strFilenameBMP)
                lResult = ctrImage.DoAPIOperations(PIC_FLIP_VERT, 0, 0)
                lResult = ctrImage.DoColorMapping
                lResult = ctrImage.SaveAsGIF(strTempPath & "\" & strFilenameGIF)
                Me.FTP1.UploadFile strFilenameGIF, strTempPath & "\" & strFilenameGIF
                Kill strTempPath & "\" & strFilenameGIF
                Kill strFilenameBMP
                nCount = nCount + 1
              End If
            End If
          End With
        Next
        Me.FTP1.Disconnect
      Else
        MsgBox "Connect failed", vbOKOnly, "FTP Upload Error"
      End If
     ' Me.tmrFTP.Enabled = False
      bFTPConnected = False
      Me.stbStatusBar1.Panels(1).Text = "Ready..."
      Me.stbStatusBar1.Panels(2).Text = ""
      nUploadTimer = 1
    End If
    End If
  End If
End Sub

Private Function ProcessWebPage(strFilename As String) As String
  Dim nFile As Integer
  Dim nTempFile As Integer
  Dim strTempPath As String
  Dim strExtension As String
  
  If InStr(strFilename, ".") <> 0 Then
    strExtension = Mid(strFilename, InStr(strFilename, "."))
  Else
    strExtension = ".htm"
  End If
  strTempPath = String$(255, " ")
  lLen = GetTempPath(Len(strTempPath), strTempPath)
  strTempPath = Left(strTempPath, lLen) & "web" & strExtension
  ProcessWebPage = strTempPath
  
  If FileExists(strTempPath) Then
    Kill strTempPath
  End If
  
  nFile = FreeFile
  Open strFilename For Input As #nFile
  nTempFile = FreeFile
  Open strTempPath For Output As #nTempFile
  Do While Not (EOF(nFile))
    Line Input #nFile, strLine
    strLine = Replace(strLine, "&lt;", "<")
    strLine = Replace(strLine, "&gt;", ">")
    strLine = Replace(strLine, "<TIMEDATE>", Now)
    Print #nTempFile, strLine
  Loop
  Close #nFile
  Close #nTempFile
End Function

Private Sub tmrGeneral_Timer()
Dim nTimeLeft As Integer
Dim nSeconds As Integer

  If Me.SSActiveToolBars1.Tools("ID_FTPUpload").State = ssChecked Then
    nTimeLeft = 15 - nUploadTimer
    Me.stbStatusBar1.Panels(2).Text = nTimeLeft & " Mins until upload"
  End If
  
End Sub

Private Sub tmrReg_Timer()
  
  nRegTimer = nRegTimer + 1
  Me.Caption = "AGSatTrak - <Not Registered> " & Str(30 - nRegTimer) & " mins until closedown"
  If nRegTimer = 30 Then
    MsgBox "This program is not registered and will now close down. Please re run the program", vbCritical + vbOKOnly, "Shutdown"
    Unload Me
  End If

End Sub
