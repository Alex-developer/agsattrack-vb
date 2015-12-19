Attribute VB_Name = "Module2"
Option Explicit


 Function StrParse(Source As String, PChar As String) As Variant

    Dim pos As Integer
    Dim i As Integer
    Dim TmpArray() As Variant
    
'    Source = Left(Source, Len(Source) - 1) + PChar 'Make sure we get the last element in string.
    Source = Left(Source, Len(Source)) + PChar 'Make sure we get the last element in string.
    '     'begin a loop

    Do
        '     'find the first separating PChar
        pos% = InStr(Source, PChar)
         
        '     'if there's one, then...

        If pos% Then
            ReDim Preserve TmpArray(i%)
            '     'extract the string up to the PChar
            TmpArray(i%) = Trim(Left$(Source, pos% - 1))
            '     'and remove that from the Source string,
            '     'so it won't be checked again
            Source = Mid$(Source, pos% + 1, Len(Source))
            i% = i% + 1
        End If

        Loop Until Source = "" Or Source = Chr(0)

            StrParse = TmpArray 'Return a New Populated Array.

        End Function


 Function BuildParseStr(vArray As Variant) As String

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


'****************************************************************
' Name: sRemove
'     ' Description:Remove a string within a string
' By: David J Berube
'
' Inputs:None
' Returns:None
' Assumes:None
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************

 Sub sRemove(String1 As String, String2 As String)


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
'****************************************************************



'****************************************************************
' Name: sReplace
' Description:search for a specific string and rep
'     lace it with another. http://137.56.41.168:2080/Vi
'     sualBasicSource/vbsearch&replace.txt
' By: Found on the World Wide Web
'
' Inputs:SearchLine is input, SearchFor is what to search for, ReplaceWith is the replacement
' Returns:None
' Assumes:None
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************


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
        '     ' Check if the strOld is in the strMainString

        If Mid(strMainString, i, Len(strOld)) = strOld Then
            ' It is in the string so concatinate the strNew to
            '      the strNewString
            strNewString = strNewString & strNew
            ' increase i to skip past the strOld in the strMai
            '     nString -1 since i will be incremented.
            i = i + Len(strOld) - 1
        Else
            ' Just concatinate the character to the strNewStri
            '     ng
            strNewString = strNewString & Mid(strMainString, i, 1)
        End If

        Next i

            sReplaceCharacters = strNewString

        End Function




