Attribute VB_Name = "Module1"
'ALWAYS ON TOP
Type T0344
  M3FD3 As Single
  M3FD9 As Single
  M3FE6 As Single
  M3FF2 As Integer
  M3FFA As Single
  M4009 As Integer
  M4014 As Integer
End Type

Global nFields(100) As Variant
Global bSatDetailsVisible As Boolean
Global nSatDetailsTag As Integer

Private Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Global bOpt As Integer



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

Public Function StrParse(Source As String, PChar As String) As Variant

    Dim pos As Integer
    Dim i As Integer
    Dim TmpArray() As Variant
    
    Source = Left(Source, Len(Source)) + PChar 'Make sure we get the last element in string.

    Do
        pos% = InStr(Source, PChar)
         

        If pos% Then
            ReDim Preserve TmpArray(i%)
            TmpArray(i%) = Trim(Left$(Source, pos% - 1))
            Source = Mid$(Source, pos% + 1, Len(Source))
            i% = i% + 1
        End If

        Loop Until Source = "" Or Source = Chr(0)

            StrParse = TmpArray 'Return a New Populated Array.
        End Function


Public Function BuildParseStr(vArray As Variant) As String

  Dim i As Integer, BldStr As String

  If Not IsArray(vArray) Then 'If not an array then return zero length string.
    BuildParseStr = ""
    Exit Function
  End If

  For i = LBound(vArray) To UBound(vArray) 'Go thru each element in the array

    If VarType(vArray(i)) = vbString Then ' Make sure all element are string type
      vArray(i) = CStr(vArray(i)) ' If Not Convert them to strings.
    End If

    If i = UBound(vArray) Then 'Keep from Appending last "," at the end of the final returned string
      BldStr = BldStr & vArray(i)
    Else
      BldStr = BldStr & vArray(i) & "," 'Build the String on the Fly.
    End If

  Next i

  BuildParseStr = BldStr ' Return Parseable String.
End Function

Public Sub sRemove(String1 As String, String2 As String)

  Dim i As Integer
  i = 1

  Do

    If (i > Len(String1)) Then Exit Do
    i = InStr(i, String1, String2)

    If i Then
      String1 = Left$(String1, i - 1) + Mid$(String1, i + Len(String2) + 1)
      i = i + 2

      DoEvents
    End If

  Loop While i

End Sub

Function sReplace(SearchLine As String, SearchFor As String, ReplaceWith As String)

    Dim vSearchLine As String, found As Integer
    found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine

    If found <> 0 Then
        vSearchLine = ""

        If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
            vSearchLine = vSearchLine + ReplaceWith

            If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
                vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
            End If

                sReplace = vSearchLine
            End Function

Function sReplaceCharacters(strMainString As String, strOld As String, strNew As String) As String

    sReplaceCharacters = ""
    Dim strNewString As String
    Dim i As Integer

    For i = 1 To Len(strMainString)

        If Mid(strMainString, i, Len(strOld)) = strOld Then
            strNewString = strNewString & strNew
            i = i + Len(strOld) - 1
        Else
            strNewString = strNewString & Mid(strMainString, i, 1)
        End If

        Next i

            sReplaceCharacters = strNewString
        End Function



