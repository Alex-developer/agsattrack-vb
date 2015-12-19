VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AllinOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private TotPhys As Long
Private TotVirt As Long
Private TotPag As Long
Private AvPhys As Long
Private AvVirt As Long
Private AvPag As Long
'System Name
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Disk Space
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
'Sreen Res.
#If Win16 Then
Private Declare Sub GetWindowRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT)
    Private Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
    Private Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
    Private Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
    Private Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Private Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
    Private Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Private Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Private Declare Function SetBkColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Private Declare Function Rectangle Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
    Private Declare Function SelectObject Lib "User32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Const BITSPIXEL = 12
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const PLANES = 14
Private hdesktopwnd
Private hdccaps
#If Win16 Then
    Private Type RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
    End Type
#Else
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
'*** Check System Version Function
Private Declare Function GetVersionX Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type
Dim OSInfo As OSVERSIONINFO
'*** Display Res. Functions
Private Declare Function GetDesktopWindow Lib "User32" () As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Memorystatus
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Dim lpInfoBuffer As MEMORYSTATUS
'PlaySound
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYN = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
'SET DISK LABEL
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
'ADD To DOCUMENTSMenu
Private Const SHARD_PIDL = &H1&
Private Const SHARD_PATH = &H2&
Private Declare Function SHAddToRecentDocs Lib "shell32.dll" (ByVal dwFlags As Long, ByVal dwData As String) As Long
'FORMAT DISK
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'ALWAYS ON TOP
Private Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
'GET IDE MODE
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'MAKE A SHORTCUT
Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
'MAKE ASSOCIATIONS

'CHANGE RESOLUTION
Private Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
'LOAD ANY DOCUMENT
Private Const SW_SHOWNORMAL = 1
#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpdirectory As String, ByVal nShowCmd As Long) As Long
#Else
Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, ByVal lpszDir$, ByVal fsShowCmd%) As Integer

'ASSOCIATIONS
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
'GRADIENT
Public Type Colour
    R As Single
    G As Single
    B As Single
End Type
Private S As Colour
Private E As Colour
'SHORT NAMES
Private Declare Function GetShortPathNameA Lib "kernel32" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'GET DRIVE TYPE
Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_UNKNOWN = 0
Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long
'GET *PATH
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'DELETE
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
End Type
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Function AddBackslash(S As String) As String
If Len(S) > 0 Then
If Right$(S, 1) <> "\" Then
AddBackslash = S + "\"
Else
AddBackslash = S
End If
Else
AddBackslash = "\"
End If
End Function
Public Function GetSystemDirectory() As String
   Dim S As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   S = Space(i)
   Call GetSystemDirectoryA(S, i)
   GetSystemDirectory = AddBackslash(Left$(S, i - 1))
End Function
Public Function GetWindowsDirectory() As String
   Dim S As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   S = Space(i)
   Call GetWindowsDirectoryA(S, i)
   GetWindowsDirectory = AddBackslash(Left$(S, i - 1))
End Function
Public Function GetTempPath() As String
   Dim S As String
   Dim i As Integer
   i = GetTempPathA(0, "")
   S = Space(i)
   Call GetTempPathA(i, S)
   GetTempPath = AddBackslash(Left$(S, i - 1))
End Function
#End If
Public Function GetIDEmode() As Boolean
         Dim strFileName As String
         Dim lngCount As Long
         strFileName = String(255, 0)
         lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
         strFileName = Left(strFileName, lngCount)
         If UCase(Right(strFileName, 7)) <> "VB5.EXE" Then
GetIDEmode = False
         Else
GetIDEmode = True
         End If
End Function

Public Function GetshortFilename(LongFilename As String) As String
   Dim S As String
   Dim i As Long
   i = Len(longpath) + 1
   S = String(i, 0)
   GetShortPathNameA longpath, S, i
   GetshortFilename = Left$(S, InStr(S, Chr$(0)) - 1)
End Function



Sub MSCAssociateProgram()
'COPY THE FOLLOWING CODE INTO FORM LOAD
            'If Command$ <> "%1" Then
            'MsgBox (Command$ & " is the file you need to open!"), vbInformation
            'Add to Recent file folder
            '    lReturn = fCreateShellLink("..\..\Recent", _
            Command$, Command$, "") UN-REMARK this line if you want the file to be added to the DOCUMENTS MENU
            'End If
'END COPY

'*****NOTES******
'Change the following code to suit the extension you wish to customize
'I Used a .FWL file, so change the parts with .FWL in it

'create an entry in the class key
Call savestring(HKEY_CLASSES_ROOT, "\.fwl", "", "fwlfile")
'content type
Call savestring(HKEY_CLASSES_ROOT, "\.fwl", "Content Type", "text/plain")
'name
Call savestring(HKEY_CLASSES_ROOT, "\fwlfile", "", "Fruit World Level")
'edit flags
Call SaveDword(HKEY_CLASSES_ROOT, "\fwlfile", "EditFlags", "0000")
'file's icon (can be an icon file, or an icon located within a dl
'     l file)
Call savestring(HKEY_CLASSES_ROOT, "\fwlfile\DefaultIcon", "", App.Path & "\ICON.ico")
'Shell
Call savestring(HKEY_CLASSES_ROOT, "\fwlfile\Shell", "", "")
'Shell Open
Call savestring(HKEY_CLASSES_ROOT, "\fwlfile\Shell\Open", "", "")
'Shell open command
Call savestring(HKEY_CLASSES_ROOT, "\fwlfile\Shell\Open\command", "", App.Path & "\Editor.exe %1")
End Sub

#End If 'dont know why these are here, but it dont work otherwise
#End If
Public Function GetHdiskSpace(DrivePath As String, RoundtoMB As Boolean) As Double
  Dim Drive As String
  Dim SectorsPerCluster As Long, BytesPerSector As Long
  Dim NumberOfFreeClusters As Long, TotalClusters As Long, Sts As Long
  Dim DS
  Drive = Left(Trim(DrivePath), 1) & ":\"     ' Ensure path is at the root.
  Sts = GetDiskFreeSpace(Drive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalClusters)
  If Sts <> 0 Then
    GetHdiskSpace = SectorsPerCluster * BytesPerSector * NumberOfFreeClusters
    DS = Format$(GetHdiskSpace, "###,###")
    If RoundtoMB = True Then
    GetHdiskSpace = DS / 1024 / 1024
    End If
  Else
    GetHdiskSpace = -1
  End If
End Function


Public Function GetWinComputerName() As String
  Dim sBuffer As String
  Dim lBufSize As Long
  Dim lStatus As Long
  lBufSize = 255
  sBuffer = String$(lBufSize, " ")
  lStatus = GetComputerName(sBuffer, lBufSize)
  GetWinComputerName = ""
  If lStatus <> 0 Then
     GetWinComputerName = Left(sBuffer, lBufSize)
  End If
  'frmSys.lblName = sGetComputerName
End Function
Public Function GetWinDisplayColors(RETURNx As String) As String
     RETURNx = UCase(RETURNx)
     Dim DisplayBits
    Dim DisplayPlanes
    Dim DisplayWidth
    Dim DisplayHeight
    Dim RetVal
    hdccaps = GetDC(hdesktopwnd)
    DisplayBits = GetDeviceCaps(hdccaps, 12)
    DisplayPlanes = GetDeviceCaps(hdccaps, 14)
    DisplayWidth = GetDeviceCaps(hdccaps, 8)
    DisplayHeight = GetDeviceCaps(hdccaps, 10)
    RetVal = ReleaseDC(hdesktopwnd, hdccaps)
Select Case RETURNx
        Case "CLS_BITS"
                    If DisplayBits = 1 Then
                    If DisplayPlanes = 1 Then
                        GetWinDisplayColors = "1 Bit"
                        ElseIf DisplayPlanes = 4 Then
                        GetWinDisplayColors = "4 Bit"
                    End If
                    ElseIf DisplayBits = 8 Then
                        GetWinDisplayColors = "8 Bit"
                    ElseIf DisplayBits = 16 Then
                    GetWinDisplayColors = "16 Bit"
                    Else
                        GetWinDisplayColors = "24 / 32 Bit"
                    End If
        Case "CLS_COL"
                    If DisplayBits = 1 Then
                    If DisplayPlanes = 1 Then
                        GetWinDisplayColors = "2"
                        ElseIf DisplayPlanes = 4 Then
                        GetWinDisplayColors = "16"
                    End If
                    ElseIf DisplayBits = 8 Then
                        GetWinDisplayColors = "256"
                    ElseIf DisplayBits = 16 Then
                    GetWinDisplayColors = "65000"
                    Else
                        GetWinDisplayColors = "16000000"
                    End If
        Case Else
End Select
End Function
Public Function GetWinVersion() As String
Dim RetVal As Long
RetVal = GetVersionX(OSInfo)
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space(128)
RetVal = GetVersionX(OSInfo)
GetWinVersion = OSInfo.dwMajorVersion & "." & OSInfo.dwMinorVersion
End Function
Public Function GetWinResX() As Integer
GetWinResX = Screen.Width / Screen.TwipsPerPixelX
End Function
Public Function GetWinResY() As Integer
GetWinResY = Screen.Height / Screen.TwipsPerPixelY
End Function
Public Function GetWinResXY() As String
GetWinResXY = Screen.Width / Screen.TwipsPerPixelX & " x " & Screen.Height / Screen.TwipsPerPixelY
End Function


Public Function GetPhysMemTotal() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetPhysMemTotal = lpInfoBuffer.dwTotalPhys '/ 1024 & " KByte"
TotPhys = lpInfoBuffer.dwTotalPhys / 1024
End Function
Public Function GetVirtMemTotal() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetVirtMemTotal = lpInfoBuffer.dwTotalVirtual '/ 1024 & " KByte"
TotVirt = lpInfoBuffer.dwTotalVirtual / 1024
End Function
Public Function GetPageFileMemTotal() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetPageFileMemTotal = lpInfoBuffer.dwTotalPageFile ' / 1024 & " KByte"
TotPag = lpInfoBuffer.dwTotalPageFile / 1024
End Function

Public Function GetPhysMemAvailable() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetPhysMemAvailable = lpInfoBuffer.dwAvailPhys ' / 1024 & " KByte"
AvPhys = lpInfoBuffer.dwAvailPhys '/ 1024
End Function
Public Function GetVirtMemAvailable() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetVirtMemAvailable = lpInfoBuffer.dwAvailVirtual ' / 1024 & " KByte"
AvVirt = lpInfoBuffer.dwAvailVirtual / 1024
End Function
Public Function GetPageFileMemAvailable() As Long
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetPageFileMemAvailable = lpInfoBuffer.dwAvailPageFile '/ 1024 & " KByte"
AvPag = lpInfoBuffer.dwAvailPageFile / 1024
End Function





Public Function GetMemLoad() As Integer
lpInfoBuffer.dwLength = Len(lpInfoBuffer)
GlobalMemoryStatus lpInfoBuffer
GetMemLoad = lpInfoBuffer.dwMemoryLoad '& "%"
End Function



Public Function MSCdeleteToRecycle(ParamArray vntFileName() As Variant) As Boolean
Dim i As Integer
Dim sFileNames As String
Dim SHFileOp As SHFILEOPSTRUCT
For i = LBound(vntFileName) To UBound(vntFileName)
sFileNames = sFileNames & vntFileName(i) & vbNullChar
Next
sFileNames = sFileNames & vbNullChar
With SHFileOp
.wFunc = FO_DELETE
.pFrom = sFileNames
.fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
End With
i = SHFileOperation(SHFileOp)
If i = 0 Then
ShellDelete = True
Else
ShellDelete = False
End If
End Function

Public Function MSCcheckFileExists(FileName As String) As Boolean
On Error GoTo FREAK
Dim Tester As String
Open FileName For Input As #254
Input #254, Tester
Close #254
MSCcheckFileExists = True
Exit Function
FREAK:
MSCcheckFileExists = False
End Function

Sub MSCgradient(sr As Integer, sg As Integer, sb As Integer, er As Integer, eg As Integer, eb As Integer, Direction As Integer, Where As Object)
On Error Resume Next
S.R = sr: S.G = sg: S.B = sb
E.R = er: E.G = eg: E.B = eb
If Direction < 0.5 Then Direction = 0
If Direction > 0.5 Then Direction = 1
If Direction = 0.5 Then Direction = 1
Select Case Direction
Case 0
Dim d As Colour
    d.R = (E.R - S.R) / Where.ScaleWidth
    d.G = (E.G - S.G) / Where.ScaleWidth
    d.B = (E.B - S.B) / Where.ScaleWidth
Dim c As Colour
c.R = S.R: c.G = S.G: c.B = S.B
For x = 0 To Where.ScaleWidth
    Where.Line (x, 0)-Step(0, Where.ScaleHeight), RGB(c.R, c.G, c.B)
    c.R = c.R + d.R
    c.G = c.G + d.G
    c.B = c.B + d.B
Next x
Case 1
Dim v As Colour
    v.R = (E.R - S.R) / Where.ScaleHeight
    v.G = (E.G - S.G) / Where.ScaleHeight
    v.B = (E.B - S.B) / Where.ScaleHeight
Dim f As Colour
f.R = S.R: f.G = S.G: f.B = S.B
For Y = 0 To Where.ScaleHeight
    Where.Line (0, Y)-Step(Where.ScaleWidth, 0), RGB(f.R, f.G, f.B)
    f.R = f.R + v.R
    f.G = f.G + v.G
    f.B = f.B + v.B
Next Y
End Select
End Sub


Public Function MSCloadDoc(Doc As String)
      Dim Scr_hDC As Long
      Scr_hDC = GetDesktopWindow()
      StartDoc = ShellExecute(Scr_hDC, "Open", Doc, "", "C:\", SW_SHOWNORMAL)
End Function

Sub MSCPlayWave(WaveFile As String, ASYNC As Boolean)
On Error Resume Next
Dim sru As String
sru = WaveFile
If ASYNC = True Then
R = sndPlaySound(ByVal CStr(sru), SND_ASYN And SND_NODEFAULT)
Else
R = sndPlaySound(ByVal CStr(sru), SND_SYNC And SND_NODEFAULT)
End If
End Sub

Sub SETfileToDocumentMenu(NewFile As String)
Dim R As Long
Dim shFlag As Long
Dim shData As String
shFlag = SHARD_PATH
shData = (NewFile)
R = SHAddToRecentDocs(shFlag, shData)
End Sub

Public Function SETdiskLabel(Disk As String, NewVal As String)
SETdiskLabel = SetVolumeLabel(Disk, NewVal)
End Function

Sub SETformatDisk(Disk As String, srcForm As Object)
Dim DriveLetter$, DriveNumber&, DriveType&
Dim RetVal&, RetFromMsg%
DriveLetter = UCase(Disk)
DriveNumber = (Asc(DriveLetter) - 65) ' Change letter to Number: A=0
DriveType = GetDriveType(DriveLetter)
If DriveType = 2 Then 'Floppies, etc
RetVal = SHFormatDrive(srcForm.hwnd, DriveNumber, 0&, 0&)
Else
RetFromMsg = MsgBox("This drive is NOT a removeable" & vbCrLf & _
"drive! Format this drive?", 276, "Format Drive Warning")
Select Case RetFromMsg
Case 6 'Yes
RetVal = SHFormatDrive(srcForm.hwnd, DriveNumber, 0&, 0&)
Case 7 'No
End Select
End If
End Sub


Sub SETmakeShortcut(FolderName As String, ShortcutName As String, Executable As String, Arguments As String)
iL = fCreateShellLink(FolderName, ShortcutName, Executable, Arguments)
End Sub

Sub SETtopmostwindow(FRM As Object, Status As Boolean)
If Status = True Then
        SetWindowPos FRM.hwnd, HWND_TOPMOST, FRM.Left / 15, _
                    FRM.Top / 15, FRM.Width / 15, _
                    FRM.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
        SetWindowPos FRM.hwnd, HWND_NOTOPMOST, FRM.Left / 15, _
                     FRM.Top / 15, FRM.Width / 15, _
                     FRM.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End Sub


Sub SETwinRes(x As Integer, Y As Integer)
Dim a As Boolean
Dim i&
i = 0
Do
a = EnumDisplaySettings(0&, i&, DevM)
i = i + 1
Loop Until (a = False)
Dim B&
DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
DevM.dmPelsWidth = x
DevM.dmPelsHeight = Y
B = ChangeDisplaySettings(DevM, 0)
End Sub


