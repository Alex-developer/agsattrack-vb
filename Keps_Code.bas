Attribute VB_Name = "Keps_Code"
Option Explicit

Function ReadKeps(strFilePath As String) As Integer
  Dim TempInc As Integer
  Dim StrTempSatName As String
  Dim CrLf As String
  Dim strDummy As String
  Dim KepLine1 As String
  Dim KepLine2 As String
  Dim KepLine3 As String
  Dim nLines As Integer
  Dim nCount As Integer
  Dim NumberOfSatellites As Integer
  Dim strFilename As String
  Dim strFilenameNoExt As String
  Dim strDatabaseFile As String
  Dim nfile As Integer
  Dim nCounter As Integer
  
  strFilename = GetFilename(strFilePath)
  strFilenameNoExt = GetFilenameNoExt(strFilename)
  If strFilename <> "" And strFilenameNoExt <> "" Then
  
  Erase sKeps

  CrLf$ = Chr$(13) + Chr$(10)

  nfile = FreeFile
  Open strFilePath For Input As #nfile
  nLines = 0
  Do
    Input #nfile, strDummy$
    nLines = nLines + 1
  Loop Until EOF(nfile)
  Close #nfile
  
  Open strFilePath For Input As #nfile

  NumberOfSatellites = 0
  nCount = 0

  Do
    If Not EOF(nfile) Then
      Input #nfile, KepLine1$
      If KepLine1$ = "" Then
        Do
          If Not EOF(nfile) Then
            Input #nfile, KepLine1$
          End If
        Loop Until KepLine1$ <> "" Or EOF(nfile)
      End If
      If EOF(nfile) Then Exit Do
      Input #nfile, KepLine2$
      If KepLine2$ = "" Then
        Do
          If Not EOF(nfile) Then
            Input #nfile, KepLine2$
          End If
        Loop Until KepLine2$ <> "" Or EOF(nfile)
      End If
      If EOF(nfile) Then Exit Do
      Input #nfile, KepLine3$
      If KepLine3$ = "" Then
        Do
          If Not EOF(nfile) Then
            Input #nfile, KepLine3$
          End If
        Loop Until KepLine3$ <> "" Or EOF(nfile)
      End If
      If EOF(nfile) And KepLine3$ = "" Then Exit Do
      If Mid$(KepLine2$, 24, 1) <> "." And Mid$(KepLine2$, 35, 1) <> "." Then
        Do
          KepLine1$ = KepLine2$
          KepLine2$ = KepLine3$
          Input #nfile, KepLine3$
          If KepLine3$ = "" Then
            Do
              If Not EOF(nfile) Then
                Input #nfile, KepLine3$
              End If
            Loop Until KepLine3$ <> "" Or EOF(nfile)
          End If
        Loop Until Mid$(KepLine2$, 24, 1) = "." And Mid$(KepLine3$, 12, 1) = "." Or EOF(nfile)
      End If
      NumberOfSatellites = NumberOfSatellites + 1

      GetKeps nCounter, KepLine1$, KepLine2$, KepLine3$
      nCounter = nCounter + 1
      
    End If
  Loop Until EOF(nfile)

  Close #nfile
  
  ReadKeps = NumberOfSatellites
  
End If

ExitSub:
  If NumberOfSatellites = 0 Then
    Call MsgBox("The file (" & strFilename & ") you have selected does not appear to contain any Keplarian elements. Please select another file.", vbExclamation + vbOKOnly + vbDefaultButton1, "Keplairan Element Load Error")
  End If
  Close #nfile
  Exit Function

ErrorHandler:
  Close #nfile
  Resume ExitSub
End Function
Public Function GetKeps(nPos As Integer, strLine1 As String, strLine2 As String, strLine3 As String) As Boolean

  sKeps(nPos).strLine1 = strLine1
  sKeps(nPos).strLine2 = strLine2
  sKeps(nPos).strLine3 = strLine3
  sKeps(nPos).strName = Trim(strLine1)
  sKeps(nPos).lDesignator = Val(Mid$(strLine2, 3, 5))
  sKeps(nPos).strEpoch = Mid$(strLine2, 19, 14)
  sKeps(nPos).dDrag = Val(Mid$(strLine2, 35, 9))
  sKeps(nPos).lRevolutionnumber = Val(Mid$(strLine3, 64, 5))
  sKeps(nPos).dInclination = Val(Mid$(strLine3, 9, 8))
  sKeps(nPos).dRAAN = Val(Mid$(strLine3, 18, 8))
  sKeps(nPos).dEccentricity = Val("0." + Mid$(strLine3, 27, 7))
  sKeps(nPos).dAOP = Val(Mid$(strLine3, 35, 8))
  sKeps(nPos).dMeanAnomoly = Val(Mid$(strLine3, 44, 8))
  sKeps(nPos).dMeanMotion = Val(Mid$(strLine3, 53, 11))
  sKeps(nPos).nElementSet = Val(Mid$(strLine2, 66, 3))
  sKeps(nPos).lOrbitNUmber = Val(Mid$(strLine3, 64, 5))
End Function
Function PercentDone(nTotal As Integer, nCount As Integer) As Integer
  nCount = nCount + 1
  PercentDone = Int(nCount / nTotal * 100)
End Function

