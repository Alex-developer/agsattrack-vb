Attribute VB_Name = "File_Code"
Private Type BrowseInfo
  hwndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
  (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" _
  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
  (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
  
  Dim iNull As Integer
  Dim lpIDList As Long
  Dim lResult As Long
  Dim sPath As String
  Dim udtBI As BrowseInfo
  With udtBI
    .hwndOwner = hwndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  
  lpIDList = SHBrowseForFolder(udtBI)
  
  If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    
    If iNull Then
      sPath = Left$(sPath, iNull - 1)
    End If
    
  End If
  
  BrowseForFolder = sPath
End Function

Public Function CheckDirExists(strFile) As Boolean
  Dim strDIR As String
  
  CheckDirExists = True
  
  strDIR = Dir(strFile, vbDirectory)
  If strDIR = "" Then
    CheckDirExists = False
  End If
  
End Function
Public Function FindFile(sPath As String) As Boolean
  If Dir(sPath) = TrimPath(sPath) Then
    FindFile = True
  Else
    FindFile = False
  End If
End Function
Public Function TrimPath(sPath As String) As String
  Dim i As Integer
  
  For i = Len(sPath) To 1 Step -1
    If InStr(i, sPath, "\", 1) = i Then Exit For
  Next i
  
  TrimPath = Right$(sPath, Len(sPath) - i)
End Function
Public Function FileExists(ByVal strPathName As String) As Boolean
  Dim intFileNum As Integer
  
  On Error Resume Next
  
  If Right$(strPathName, 1) = "\" Then
    strPathName = Left$(strPathName, Len(strPathName) - 1)
  End If
  intFileNum = FreeFile
  Open strPathName For Input As intFileNum
  FileExists = IIf(Err, False, True)
  Close intFileNum
  Err = 0
End Function
Function LongDirFix(TargetString As String, Max As Integer) As String

  Dim i, LblLen, StringLen As Integer
  Dim TempString As String
  TempString = TargetString
  LblLen = Max

  If Len(TempString) <= LblLen Then
    LongDirFix = TempString
    Exit Function
  End If

  LblLen = LblLen - 6

  For i = Len(TempString) - LblLen To Len(TempString)

    If Mid$(TempString, i, 1) = "\" Then Exit For
  Next

  LongDirFix = Left$(TempString, 3) & "..." & Right$(TempString, Len(TempString) - (i - 1))
End Function

Function GetFilename(strFilePath As String) As String
  Dim i As Integer
  
  GetFilename = ""
  If Right$(strFilePath, 1) = "\" Then
    strFilePath = Left$(strFilePath, Len(strFilePath) - 1)
  End If
  
  For i = Len(strFilePath) To 1 Step -1
    If Mid$(strFilePath, i, 1) = "\" Then
      GetFilename = Mid$(strFilePath, i + 1)
      Exit For
    End If
  Next i
  
End Function

Function GetFilenameNoExt(strFilename As String) As String
  Dim i As Integer

  GetFilenameNoExt = strFilename
  For i = 1 To Len(strFilename)
    If Mid$(strFilename, i, 1) = "." Then
      GetFilenameNoExt = Left$(strFilename, i - 1)
    End If
  Next i
End Function

Public Function StrParse(Source As String, PChar As String) As Variant
  Dim TmpArray() As Variant

  Dim Pos As Integer, i As Integer
  Source = Left(Source, Len(Source)) + PChar
  Do
    Pos% = InStr(Source, PChar)

    If Pos% Then
      ReDim Preserve TmpArray(i%)
      TmpArray(i%) = Trim(Left$(Source, Pos% - 1))
      Source = Mid$(Source, Pos% + 1, Len(Source))
      i% = i% + 1
    End If

  Loop Until Source = "" Or Source = Chr(0)

  StrParse = TmpArray
End Function

Function SaveView(strFilename As String)

  On Error GoTo ERROR_SaveView

  Dim strFile As String
  Dim nFile As Integer
  Dim fForm As Form
  Dim i As Integer
  Dim nSats As Integer
  Dim nTotalForms As Integer

  For Each fForm In Forms
    If Left(fForm.Caption, 7) = "SatView" Then
      nTotalForms = nTotalForms + 1
    End If
  Next
  If nTotalForms > 0 Then
    strFile = App.Path & "\Views\" & strFilename

    nFile = FreeFile

    Open strFile For Output As #nFile
    Write #nFile, cFileVersion
    Write #nFile, nTotalForms
    For Each fForm In Forms
      With fForm
        If Left(.Caption, 7) = "SatView" Then
          Write #nFile, .Left
          Write #nFile, .Top
          Write #nFile, .Width
          Write #nFile, .Height
          nSats = .ocxSat.SatelliteCount
          Write #nFile, .nUpdate
          Write #nFile, nSats
          Write #nFile, .ocxSat.DatabasePath
          For i = 1 To nSats
            .ocxSat.SatelliteIndex = i
            Write #nFile, .ocxSat.SatelliteDesignator
            Write #nFile, .ocxSat.SatelliteTrackOrbits
' new in "AGSatTrack View V4.00"
            Write #nFile, .ocxSat.OrbitModel
          Next i
          Write #nFile, .ocxSat.ObserverMapCentre
          Write #nFile, .ocxSat.OutputStyle
          Write #nFile, .ocxSat.DisplayMoon
          Write #nFile, .ocxSat.DisplayMoonFootprint
          Write #nFile, .ocxSat.DisplaySun
          Write #nFile, .ocxSat.DisplaySunFootprint
' new in "AGSatTrack View V2.00"
          Write #nFile, .ocxSat.DisplayFootprints
          Write #nFile, .ocxSat.DisplayTracks
          Write #nFile, .ocxSat.DisplayIcons
          Write #nFile, .bSpeech
          Write #nFile, .nSpeechInterval
          Write #nFile, .nUpdate
          Write #nFile, .bOnDesktop
          Write #nFile, .ocxSat.SetSelectedSatellite
          Write #nFile, .strMap0
          Write #nFile, .strMap180
          Write #nFile, .strMapHorizon
          Write #nFile, .ocxSat.DisplayAOSCircle
' new in "AGSatTrack View V2.00"
          Write #nFile, .ocxSat.SecondObserverEnabled
          Write #nFile, .ocxSat.SecondObserverHeight
          Write #nFile, .ocxSat.SecondObserverLatitude
          Write #nFile, .ocxSat.SecondObserverLongitude
          Write #nFile, .ocxSat.SecondObserverLocation
' new in "AGSatTrack View V4.00"
          Write #nFile, .ocxSat.TimeZoneName
          Write #nFile, .ocxSat.Timezone
        End If
      End With
    Next
    Close #nFile
  End If

EXIT_SaveView:
  Exit Function

ERROR_SaveView:
  MsgBox "Error in ERROR_SaveView : " & Error
  Resume EXIT_SaveView

End Function
Function OpenView(strFilename As String)

  On Error GoTo ERROR_OpenView

  Dim strFile As String
  Dim nFile As Integer
  Dim fForm As Form
  Dim i As Integer
  Dim j As Integer
  Dim nViews As Integer
  Dim nSats As Integer
  Dim strTemp As String
  Dim nTotalForms As Integer
  Dim strDatabase As String
  Dim strDesignator As String
  Dim nOrbits As Integer
  Dim bTemp As Boolean
  Dim nTemp As Integer
  Dim lTop As Long
  Dim lLeft As Long
  Dim lWidth As Long
  Dim lHeight As Long
  Dim nUpdate As Integer
  Dim sTemp As Single
  Dim nPos As Integer
  
  strFile = App.Path & "\Views\" & strFilename

  nFile = FreeFile

  Open strFile For Input As #nFile
  Input #nFile, strTemp
  Select Case strTemp
    Case "AGSatTrack View V1.00"
      Input #nFile, nTotalForms
      If nTotalForms > 0 Then
        For Each fForm In Forms
          If Left(fForm.Caption, 7) = "SatView" Then
            Unload fForm
          End If
        Next
        For nViews = 1 To nTotalForms
          Input #nFile, lLeft
          Input #nFile, lTop
          Input #nFile, lWidth
          Input #nFile, lHeight
          wPos.lLeft = lLeft
          wPos.lTop = lTop
          wPos.lWidth = lWidth
          wPos.lHeight = lHeight
          Input #nFile, nUpdate
          fMainForm.LoadNewDoc 1, False, True
          With fMainForm.ActiveForm.ocxSat
            'With frmDocument.ocxSat
            fMainForm.ActiveForm.nUpdate = nUpdate
            Input #nFile, nSats
            Input #nFile, strDatabase
            .DatabasePath = strDatabase
            ReadKeps strDatabase
            For i = 1 To nSats
              Input #nFile, strDesignator
              Input #nFile, nOrbits
              For j = 0 To UBound(sKeps())
                If CLng(strDesignator) = sKeps(j).lDesignator Then
                  .AddSatellite
                  frmSelect.UpdateOcxKeps i, j, fMainForm.ActiveForm
                  .SatelliteTrackOrbits = nOrbits
                  Exit For
                End If
              Next j
            Next i
            .SetSelectedSatellite = 1
            Input #nFile, nTemp
            .ObserverMapCentre = nTemp
            Input #nFile, nTemp
            .OutputStyle = nTemp
            Input #nFile, nTemp
            .DisplayMoon = nTemp
            Input #nFile, nTemp
            .DisplayMoonFootprint = nTemp
            Input #nFile, nTemp
            .DisplaySun = nTemp
            Input #nFile, nTemp
            .DisplaySunFootprint = nTemp
          End With
          fMainForm.ActiveForm.UpdatePosTable
          fMainForm.UpdateToolbar
        Next nViews
      End If
    Case "AGSatTrack View V2.00"
      Input #nFile, nTotalForms
      If nTotalForms > 0 Then
        For Each fForm In Forms
          If Left(fForm.Caption, 7) = "SatView" Then
            Unload fForm
          End If
        Next
        For nViews = 1 To nTotalForms
          Input #nFile, lLeft
          Input #nFile, lTop
          Input #nFile, lWidth
          Input #nFile, lHeight
          wPos.lLeft = lLeft
          wPos.lTop = lTop
          wPos.lWidth = lWidth
          wPos.lHeight = lHeight
          Input #nFile, nUpdate
          fMainForm.LoadNewDoc 1, False, True
          With fMainForm.ActiveForm.ocxSat
            'With frmDocument.ocxSat
            fMainForm.ActiveForm.nUpdate = nUpdate
            Input #nFile, nSats
            Input #nFile, strDatabase
            .DatabasePath = strDatabase
            ReadKeps strDatabase
            For i = 1 To nSats
              Input #nFile, strDesignator
              Input #nFile, nOrbits
              For j = 0 To UBound(sKeps())
                If CLng(strDesignator) = sKeps(j).lDesignator Then
                  .AddSatellite
                  frmSelect.UpdateOcxKeps i, j, fMainForm.ActiveForm
                  .SatelliteTrackOrbits = nOrbits
                  Exit For
                End If
              Next j
            Next i
            .SetSelectedSatellite = 1
            Input #nFile, nTemp
            .ObserverMapCentre = nTemp
            Input #nFile, nTemp
            .OutputStyle = nTemp
            Input #nFile, bTemp
            .DisplayMoon = bTemp
            Input #nFile, bTemp
            .DisplayMoonFootprint = bTemp
            Input #nFile, bTemp
            .DisplaySun = bTemp
            Input #nFile, bTemp
            .DisplaySunFootprint = bTemp
            Input #nFile, bTemp
            .DisplayFootprints = bTemp
            Input #nFile, bTemp
            .DisplayTracks = bTemp
            Input #nFile, bTemp
            .DisplayIcons = bTemp
            Input #nFile, bTemp
            .EnableSpeech = bTemp
            fMainForm.ActiveForm.bSpeech = bTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nSpeechInterval = nTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nUpdate = nTemp
            Input #nFile, bTemp
            fMainForm.ActiveForm.bOnDesktop = bTemp
            .SetActiveWindowAsWallpaper = bTemp
            Input #nFile, nTemp
            .SetSelectedSatellite = nTemp
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap0 = strTemp
            If strTemp <> "" Then
              .SetMap 0, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap180 = strTemp
            If strTemp <> "" Then
              .SetMap 1, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMapHorizon = strTemp
            If strTemp <> "" Then
              .SetMap 2, strTemp
            End If
            Input #nFile, bTemp
            .DisplayAOSCircle = bTemp
            .Refresh
          End With
          fMainForm.ActiveForm.UpdatePosTable
          fMainForm.UpdateToolbar
        Next nViews
      End If
    Case "AGSatTrack View V3.00"
      Input #nFile, nTotalForms
      If nTotalForms > 0 Then
        For Each fForm In Forms
          If Left(fForm.Caption, 7) = "SatView" Then
            Unload fForm
          End If
        Next
        For nViews = 1 To nTotalForms
          Input #nFile, lLeft
          Input #nFile, lTop
          Input #nFile, lWidth
          Input #nFile, lHeight
          wPos.lLeft = lLeft
          wPos.lTop = lTop
          wPos.lWidth = lWidth
          wPos.lHeight = lHeight
          Input #nFile, nUpdate
          fMainForm.LoadNewDoc 1, False, True
          With fMainForm.ActiveForm.ocxSat
            'With frmDocument.ocxSat
            fMainForm.ActiveForm.nUpdate = nUpdate
            Input #nFile, nSats
            Input #nFile, strDatabase
            .DatabasePath = strDatabase
            ReadKeps strDatabase
            For i = 1 To nSats
              Input #nFile, strDesignator
              Input #nFile, nOrbits
              For j = 0 To UBound(sKeps())
                If CLng(strDesignator) = sKeps(j).lDesignator Then
                  .AddSatellite
                  frmSelect.UpdateOcxKeps i, j, fMainForm.ActiveForm
                  .SatelliteTrackOrbits = nOrbits
                  Exit For
                End If
              Next j
            Next i
            .SetSelectedSatellite = 1
            Input #nFile, nTemp
            .ObserverMapCentre = nTemp
            Input #nFile, nTemp
            .OutputStyle = nTemp
            Input #nFile, bTemp
            .DisplayMoon = bTemp
            Input #nFile, bTemp
            .DisplayMoonFootprint = bTemp
            Input #nFile, bTemp
            .DisplaySun = bTemp
            Input #nFile, bTemp
            .DisplaySunFootprint = bTemp
            Input #nFile, bTemp
            .DisplayFootprints = bTemp
            Input #nFile, bTemp
            .DisplayTracks = bTemp
            Input #nFile, bTemp
            .DisplayIcons = bTemp
            Input #nFile, bTemp
            .EnableSpeech = bTemp
            fMainForm.ActiveForm.bSpeech = bTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nSpeechInterval = nTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nUpdate = nTemp
            Input #nFile, bTemp
            fMainForm.ActiveForm.bOnDesktop = bTemp
            .SetActiveWindowAsWallpaper = bTemp
            Input #nFile, nTemp
            .SetSelectedSatellite = nTemp
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap0 = strTemp
            If strTemp <> "" Then
              .SetMap 0, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap180 = strTemp
            If strTemp <> "" Then
              .SetMap 1, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMapHorizon = strTemp
            If strTemp <> "" Then
              .SetMap 2, strTemp
            End If
            Input #nFile, bTemp
            .DisplayAOSCircle = bTemp
            .Refresh
            Input #nFile, bTemp
            .SecondObserverEnabled = bTemp
            Input #nFile, sTemp
            .SecondObserverHeight = sTemp
            Input #nFile, sTemp
            .SecondObserverLatitude = sTemp
            Input #nFile, sTemp
            .SecondObserverLongitude = sTemp
            Input #nFile, strTemp
            .SecondObserverLocation = strTemp
          End With
          fMainForm.ActiveForm.UpdatePosTable
          fMainForm.UpdateToolbar
        Next nViews
      End If
    Case "AGSatTrack View V4.00"
      Input #nFile, nTotalForms
      If nTotalForms > 0 Then
        For Each fForm In Forms
          If Left(fForm.Caption, 7) = "SatView" Then
            Unload fForm
          End If
        Next
        For nViews = 1 To nTotalForms
          Input #nFile, lLeft
          Input #nFile, lTop
          Input #nFile, lWidth
          Input #nFile, lHeight
          wPos.lLeft = lLeft
          wPos.lTop = lTop
          wPos.lWidth = lWidth
          wPos.lHeight = lHeight
          Input #nFile, nUpdate
          fMainForm.LoadNewDoc 1, False, True
          With fMainForm.ActiveForm.ocxSat
            'With frmDocument.ocxSat
            fMainForm.ActiveForm.nUpdate = nUpdate
            Input #nFile, nSats
            Input #nFile, strDatabase
            .DatabasePath = strDatabase
            ReadKeps strDatabase
            For i = 1 To nSats
              Input #nFile, strDesignator
              Input #nFile, nOrbits
              Input #nFile, nTemp
              For j = 0 To UBound(sKeps())
                If CLng(strDesignator) = sKeps(j).lDesignator Then
                  nPos = .AddSatellite
                  frmSelect.UpdateOcxKeps i, j, fMainForm.ActiveForm
                  .SatelliteTrackOrbits = nOrbits
                  .SatelliteIndex = nPos
                  .OrbitModel = nTemp
                  Exit For
                End If
              Next j
            Next i
            .SetSelectedSatellite = 1
            Input #nFile, nTemp
            .ObserverMapCentre = nTemp
            Input #nFile, nTemp
            .OutputStyle = nTemp
            Input #nFile, bTemp
            .DisplayMoon = bTemp
            Input #nFile, bTemp
            .DisplayMoonFootprint = bTemp
            Input #nFile, bTemp
            .DisplaySun = bTemp
            Input #nFile, bTemp
            .DisplaySunFootprint = bTemp
            Input #nFile, bTemp
            .DisplayFootprints = bTemp
            Input #nFile, bTemp
            .DisplayTracks = bTemp
            Input #nFile, bTemp
            .DisplayIcons = bTemp
            Input #nFile, bTemp
            .EnableSpeech = bTemp
            fMainForm.ActiveForm.bSpeech = bTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nSpeechInterval = nTemp
            Input #nFile, nTemp
            fMainForm.ActiveForm.nUpdate = nTemp
            Input #nFile, bTemp
            fMainForm.ActiveForm.bOnDesktop = bTemp
            .SetActiveWindowAsWallpaper = bTemp
            Input #nFile, nTemp
            .SetSelectedSatellite = nTemp
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap0 = strTemp
            If strTemp <> "" Then
              .SetMap 0, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMap180 = strTemp
            If strTemp <> "" Then
              .SetMap 1, strTemp
            End If
            Input #nFile, strTemp
            fMainForm.ActiveForm.strMapHorizon = strTemp
            If strTemp <> "" Then
              .SetMap 2, strTemp
            End If
            Input #nFile, bTemp
            .DisplayAOSCircle = bTemp
            .Refresh
            Input #nFile, bTemp
            .SecondObserverEnabled = bTemp
            Input #nFile, sTemp
            .SecondObserverHeight = sTemp
            Input #nFile, sTemp
            .SecondObserverLatitude = sTemp
            Input #nFile, sTemp
            .SecondObserverLongitude = sTemp
            Input #nFile, strTemp
            .SecondObserverLocation = strTemp
            Input #nFile, strTemp
            .TimeZoneName = strTemp
            Input #nFile, nTemp
            .Timezone = nTemp
          End With
          fMainForm.ActiveForm.UpdatePosTable
          fMainForm.UpdateToolbar
        Next nViews
      End If
  End Select
  Close #nFile

EXIT_OpenView:
  Exit Function

ERROR_OpenView:
  MsgBox "Error in ERROR_OpenView : " & Error
  Resume EXIT_OpenView

End Function
 Sub ReadDX()
    On Error GoTo ERROR_HANDLER
    
    Dim nFile As Integer
    Dim strPath As String
    Dim strData As String
    Dim strTemp As String
    Dim i As Integer
    
    strPath = App.Path & "\Observer\ObserverLocations.txt"

    nFile = FreeFile
      
    Open strPath For Input As #nFile
    
    While Not EOF(1)
      Line Input #nFile, strTemp
      strData = strTemp
      sDxDetails(i).Callsign = csvParser(strData, 2)
      strData = strTemp
      sDxDetails(i).strName = csvParser(strData, 3)
      strData = strTemp
      sDxDetails(i).strLat = csvParser(strData, 4)
      strData = strTemp
      sDxDetails(i).strLon = csvParser(strData, 5)
      i = i + 1
    Wend
    Close #nFile
    Exit Sub
    
ERROR_HANDLER:
  MsgBox "Unable to read the DX locations database", vbCritical, "Open file error"
End Sub

Public Function csvParser(dataline As String, FieldNum As Integer) As String
Dim i As Integer
Dim startPos As Integer
Dim stopPos As Integer
Dim quote As Integer
Dim out As Boolean

On Error Resume Next
i = 1
out = False
startPos = 1
While (i < FieldNum) And (Not out)
 quote = 1
 
 If Mid(dataline, 1, 1) = """" Then
  startPos = 2
  While (Mid(dataline, startPos, 1) <> """") And (startPos <= Len(dataline))
   startPos = startPos + 1
  Wend
 Else
  startPos = InStr(1, dataline, ",", vbBinaryCompare)
  quote = 0
 End If
 i = i + 1
 If quote = 1 Then
  dataline = Right(dataline, Len(dataline) - startPos - 1)
  If Err.Number <> 0 Then dataline = Right(dataline, Len(dataline) - startPos)
 Else
  dataline = Right(dataline, Len(dataline) - startPos)
 End If
Wend

If Mid(dataline, 1, 1) = """" Then
 dataline = Right(dataline, Len(dataline) - 1)
 stopPos = InStr(dataline, """") - 1
 If stopPos <> 0 Then
  csvParser = Mid(dataline, 1, stopPos)
 Else
  csvParser = Mid(dataline, 2, Len(stopPos) - 2)
 End If
Else
 stopPos = InStr(dataline, ",") - 1
End If

If stopPos = -1 Then
 csvParser = dataline
Else
 csvParser = ""
End If
If stopPos > 0 Then
 csvParser = Mid(dataline, 1, stopPos)
End If


End Function

Public Function FileCheck(Path$) As Boolean
    Dim Disregard                       As Long
    'USAGE: If FileCheck("C:\windows\kewl.exe") then msgbox "it was found"
    FileCheck = True 'Assume Success
    On Error Resume Next
    Disregard = FileLen(Path)
    If Err <> 0 Then
        FileCheck = False
    End If
End Function

Public Function UpdateProgress(PB As Control, ByVal Percent As Integer, Optional ByVal ShowPercent = False)
    'Replacement for progress bar..looks nicer also
    Dim sNum                            As String    'use percent
    If Not PB.AutoRedraw Then 'picture in memory ?
        PB.AutoRedraw = -1 'no, make one
    End If
    PB.Cls 'clear picture in memory
    PB.ScaleWidth = 100 'new sclaemodus
    PB.DrawMode = 10 'not XOR Pen Modus
    If ShowPercent = True Then
    Num$ = Format$(Percent, "###0") + "%"
    PB.CurrentX = 50 - PB.TextWidth(Num$) / 2
    PB.CurrentY = (PB.ScaleHeight - PB.TextHeight(Num$)) / 2
    PB.Print Num$ 'print percent
    End If
    PB.Line (0, 0)-(Percent, PB.ScaleHeight), , BF
    PB.Refresh 'show differents
End Function

Public Function FormsOnTop(frmForm As Form, fOnTop As Boolean)
    'USAGE: ONTOP ME,TRUE   -ONTOP MOST
    '       ONTOP ME,FALSE  -NOT TOP MOST
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Dim lState                          As Long
    Dim iLeft                           As Integer
    Dim iTop                            As Integer
    Dim iWidth                          As Integer
    Dim iHeight                         As Integer
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
  '  Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Function

Public Function File_ByteConversion(NumberOfBytes As Single) As String
    On Error Resume Next
    If NumberOfBytes < 1024 Then 'checks to see if its so small that it cant be converted into larger grouping
        File_ByteConversion = NumberOfBytes & " Bytes"
    End If
    If NumberOfBytes > 1024 Then  'Checks to see if file is big enough to convert into KB
        File_ByteConversion = Format(NumberOfBytes / 1024, "0.00") & " KB"
    End If
    If NumberOfBytes > 1048576 Then  'Checks to see if its big enough to convert into MB
        File_ByteConversion = Format(NumberOfBytes / 1048576, "###,###,##0.00") & " MB"
    End If
End Function

