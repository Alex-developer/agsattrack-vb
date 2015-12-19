Attribute VB_Name = "Module1"

Type tOPts
  nLatitude As Single
  nLongitude As Single
  nHeight As Integer
  strLocation As String
  nSecondLatitude As Single
  nSecondLongitude As Single
  nSecondHeight As Integer
  nSecondName As String
  bSecondUsed As Boolean
  strDefFilename As String
  bSaveOpenlastView As Boolean
  strDefDatabasePath As String
  strDefKepsPath As String
  strKepsDatabase As String
  bDaylightSaving As Boolean
  strUserName As String
  strCode As String
  bIndicateVis As Boolean
  strTimeZone As String
  bAutoadjust As Boolean
  nTimezoneAdjust As Integer
  nOrthX As Long
  nOrthY As Long
  bOrthShade As Boolean
  strOrthLocations As String
  bSetDesktop As Boolean
  nUpdateInterval As Integer
  bFirst As Boolean
  nKepsAge As Integer
  bForceReadme As String
  bSysTray As Boolean
  bSpeech As Boolean
  nSpeechInterval As Integer
  bIcons As Boolean
  bForceReset As Boolean
  bDisplayRangeCircle As Boolean
  nGroundTrackPointSize As Integer
  nFTPTimeout As Integer
  bFTPbPasvMode As Boolean
  bFTPProxy As Boolean
  strFTPProxyURL As String
  nFTPPPort As Integer
  bFTPThruProxy As Boolean
  bFTPRollback As Boolean
  nFTPRollback As Integer
  bShowListbar As Boolean
  bRotatorEnabled As Boolean
  bRotatorAlwaysTrack As Boolean
  nRotatorType As Integer
  strFTPServer As String
  strFTPUserName As String
  strFTPPassword As String
  strFTPHTMLDir As String
  strFTPHTMLTemplate As String
  strFTPImagesDir As String
End Type

Type Keps
  strName As String
  lDesignator As Long
  strEpoch As String
  dDrag As Double
  lRevolutionnumber As Long
  dInclination As Double
  dRAAN As Double
  dEccentricity As Double
  dAOP As Double
  dMeanAnomoly As Double
  dMeanMotion As Double
  nElementSet As Integer
  lOrbitNUmber As Long
  strLine1 As String
  strLine2 As String
  strLine3 As String
End Type


Public Type tPOS
  lLeft As Long
  lTop As Long
  lWidth As Long
  lHeight As Long
End Type
Global wPos As tPOS

Global sKeps(500) As Keps

Public sProgramOptions As tOPts
Public kepSatName(500) As String
Public kepDesignation(500) As String
Public kepEpoch(500) As Double
Public kepDrag(500) As Single
Public kepRevolutions(500) As Long
Public kepInclination(500) As Single
Public kepRAAN(500) As Single
Public kepEccentricity(500) As Single
Public kepArgOfPerigee(500) As Single
Public kepMeanAnomoly(500) As Single
Public kepMeanMotion(500) As Double

Public lDocumentCount As Long

Type DX
  Callsign As String
  strLat As String
  strLon As String
  strName As String
End Type

Global sDxDetails(2000) As DX



Public nCounter As Long
'Public cFunctions As New AllinOne
Public fMainForm As frmMain
Global bRegistered As Boolean
Global nRegCount As Integer
Global nRegTimer As Integer

Global frmForm As Form

Global frmRotatorForm As Form
Global bMoveRotator As Boolean

Global gbCancel As Boolean
Global Const gc0528 = 75369793000000#
Global Const cEarthRadius = 6367
Global Const cFileVersion = "AGSatTrack View V4.00"
Global Const cBack = 8
Global Const c0 = 48
Global Const c9 = 57
Global Const cA = 65
Global Const cZ = 90

Global Const c04AA = 0.00167322025
Global Const c04B2 = 0.006670539762
Global Const c04BA = 400000  ' &H61A80&
Global Const c04BE = -100000 ' &HFFFE7960&
Global Const c04C2 = 3.14159265358979
Global Const c049A = 6375020.481
Global Const c04A2 = 6353722.49


Global Const cSPACE = 32
Global Const cPoint = 46

'--------START GLOBAL STRINGS FOR THIS PROJECT-----
Public strSvrURL                        As String
Public strSvrPort                       As String
Public bProxy                           As Boolean
Public URL                              As String
Public RESUMEFILE                       As Boolean
Public FilePathName                     As String
Public FileName                         As String
Public FileLength                       As Single
Public Sec                              As Integer
Public Min                              As Integer
Public Hr                               As Integer

Global BytesAlreadySent                As Single
Global BytesRemaining                  As Single

Global bFTPConnected As Boolean


Sub Main()
  frmSplash.Show
  frmSplash.Refresh
  Set fMainForm = New frmMain
  Load fMainForm
  Unload frmSplash

  fMainForm.Show
End Sub

Public Sub CenterForm(frmTarget As Form)

  frmTarget.Move (fMainForm.Left + (fMainForm.Width - frmTarget.Width) / 2), (fMainForm.Top + (fMainForm.Height - frmTarget.Height) / 2)

End Sub
Public Function nToUpper(nChar As Integer) As Integer
If nChar >= Asc("a") And nChar <= Asc("z") Then nToUpper = nChar + Asc("A") - Asc("a") Else nToUpper = nChar
End Function
Public Function bcheckNum(nChar As Integer) As Boolean
If (nChar >= Asc("0") And nChar <= Asc("9")) Or nChar = cBack Then bcheckNum = True
End Function
Public Function bCheckNumNeg(nChar As Integer, bAllowneg As Boolean) As Boolean
If nChar = 8 Or nChar = 9 Or nChar = 13 Or nChar = 3 Or nChar = 22 Or (nChar >= Asc("0") And nChar <= Asc("9")) Or (nChar = Asc("-") And bAllowneg) Then bCheckNumNeg = True
End Function

