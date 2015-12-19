Attribute VB_Name = "Timezone"
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type LOCALE_TIME_ZONE_INFORMATION
    Bias As Long
    StandardBias As Long
    DaylightBias As Long
    StandardDate As SYSTEMTIME
    DaylightDate As SYSTEMTIME
    DisplayName As String
    StandardName As String
    DaylightName As String
    MapID As String
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
    

Private Const MAX_SIZE = 2048
Private Const HKLM = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const KEY_ALL_ACCESS = &HF003F
Private Const REG_SZ = 1
Private Const REG_BINARY = 3

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long)
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

Public LocTZI() As LOCALE_TIME_ZONE_INFORMATION
Public CurrentTZI As LOCALE_TIME_ZONE_INFORMATION
Public Function GetTZICollection(sRegKey As String) As Boolean
   GetTZICollection = EnumSubKeys(HKLM, sRegKey)
End Function

Private Function EnumSubKeys(TopKey As Long, SubKey As String, Optional colTZ As Collection) As Boolean
   Dim hKey As Long, curidx As Long, KeyName As String, KeyValue As String
   Dim s As String
   RegOpenKeyEx TopKey, SubKey, 0&, KEY_ALL_ACCESS, hKey
   EnumSubKeys = hKey
   Do
     KeyName = Space$(MAX_SIZE)
     KeyValue = Space$(MAX_SIZE)
     If RegEnumKey(hKey, curidx, KeyName, MAX_SIZE) <> ERROR_SUCCESS Then Exit Do
     ReDim Preserve LocTZI(curidx)
     KeyName = StripNulls(KeyName)
     LocTZI(curidx) = GetRegValueLTZI(SubKey & "\" & KeyName)
     curidx = curidx + 1
   Loop
   RegCloseKey hKey
End Function

Private Function GetRegValueLTZI(sKey As String) As LOCALE_TIME_ZONE_INFORMATION
  Dim hKey As Long, ltzi As LOCALE_TIME_ZONE_INFORMATION, sTemp As String
  If RegOpenKeyEx(HKLM, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
  Call RegQueryValueEx(hKey, "TZI", 0&, REG_BINARY, ltzi, 44&)
  sTemp = Space$(MAX_SIZE)
  Call RegQueryValueEx(hKey, "Display", 0&, REG_SZ, ByVal sTemp, MAX_SIZE)
  ltzi.DisplayName = StripNulls(sTemp)
  sTemp = Space$(MAX_SIZE)
  Call RegQueryValueEx(hKey, "Std", 0&, REG_SZ, ByVal sTemp, MAX_SIZE)
  ltzi.StandardName = StripNulls(sTemp)
  sTemp = Space$(MAX_SIZE)
  Call RegQueryValueEx(hKey, "Dlt", 0&, REG_SZ, ByVal sTemp, MAX_SIZE)
  ltzi.DaylightName = StripNulls(sTemp)
  sTemp = Space$(MAX_SIZE)
  Call RegQueryValueEx(hKey, "MapID", 0&, REG_SZ, ByVal sTemp, MAX_SIZE)
  ltzi.MapID = StripNulls(sTemp)
  RegCloseKey hKey
  GetRegValueLTZI = ltzi
End Function

Public Function GetRegValueStr(sKey As String, sSubKey As String) As String
  Dim hKey As Long, sTemp As String
  GetRegValueStr = ""
  If RegOpenKeyEx(HKLM, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
  sTemp = Space$(MAX_SIZE)
  If RegQueryValueEx(hKey, sSubKey, 0&, REG_SZ, ByVal sTemp, MAX_SIZE) = ERROR_SUCCESS Then
     GetRegValueStr = Trim$(StripNulls(sTemp))
  End If
  RegCloseKey hKey
End Function

Private Function StripNulls(ByVal sText As String) As String
  Dim nPosition&
  StripNulls = sText
  nPosition = InStr(sText, vbNullChar)
  If nPosition Then StripNulls = Left$(sText, nPosition - 1)
  If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = ""
End Function

Public Function UTCToLocalDate(dt As Date, tz As LOCALE_TIME_ZONE_INFORMATION) As Date
   Dim d As Date
   d = DateSerial(Year(dt), Month(dt), Day(dt)) + TimeSerial(Hour(dt), Minute(dt) - tz.Bias - tz.DaylightBias, Second(dt))
   If Not IsDayLight(d, tz) Then
      d = DateSerial(Year(dt), Month(dt), Day(dt)) + TimeSerial(Hour(dt), Minute(dt) - tz.Bias, Second(dt))
   End If
   UTCToLocalDate = d
End Function

Public Function LocalDateToUTC(dt As Date, tz As LOCALE_TIME_ZONE_INFORMATION) As Date
   If IsDayLight(dt, tz) Then
      LocalDateToUTC = DateSerial(Year(dt), Month(dt), Day(dt)) + TimeSerial(Hour(dt), Minute(dt) + tz.Bias + tz.DaylightBias, Second(dt))
   Else
      LocalDateToUTC = DateSerial(Year(dt), Month(dt), Day(dt)) + TimeSerial(Hour(dt), Minute(dt) + tz.Bias, Second(dt))
   End If
End Function

Public Function IsDayLight(dt As Date, tz As LOCALE_TIME_ZONE_INFORMATION) As Boolean
  Dim dlBegin As Date, dlEnd As Date
  With tz.DaylightDate
       If .wYear Then
          dlBegin = DateSerial(Year(dt), .wMonth, .wDay)
       Else
          dlBegin = WeekDayToDate(Year(dt), .wMonth, .wDayOfWeek, .wDay)
       End If
       dlBegin = dlBegin + TimeSerial(.wHour, .wMinute, .wSecond)
  End With
  With tz.StandardDate
       If .wYear Then
          dlEnd = DateSerial(Year(dt), .wMonth, .wDay)
       Else
          dlEnd = WeekDayToDate(Year(dt), .wMonth, .wDayOfWeek, .wDay)
       End If
       dlEnd = dlEnd + TimeSerial(.wHour, .wMinute, .wSecond)
  End With
  If dlBegin = dlEnd Then Exit Function
  If dlBegin < dlEnd Then
     If dt > dlBegin And dt < dlEnd Then IsDayLight = True
  Else
    'Australia!!!
     If dt < dlEnd Or dt > dlBegin Then IsDayLight = True
  End If
End Function

Public Function WeekDayToDate(Y As Integer, m As Integer, wDay As Integer, nDay As Integer) As Date
  Dim d As Date, n As Integer, s As String, count As Integer
  For i = 1 To 31
      d = DateSerial(Y, m, i)
      If Month(d) > m Then Exit For
      If wDay = Weekday(d) - 1 Then
         count = count + 1
         WeekDayToDate = d
         If count = nDay Then Exit For
      End If
  Next i
End Function

Public Function IsNT() As Boolean
  Dim verinfo As OSVERSIONINFO
  verinfo.dwOSVersionInfoSize = Len(verinfo)
  If (GetVersionEx(verinfo)) = 0 Then Exit Function
  If verinfo.dwPlatformId = 2 Then IsNT = True
End Function


