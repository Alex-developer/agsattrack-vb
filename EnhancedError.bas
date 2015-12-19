Attribute VB_Name = "EnhancedErrors_Module"
' #VBIDEUtils#************************************************************
' * Programmer Name  : Waty Thierry
' * Web Site         : www.geocities.com/ResearchTriangle/6311/
' * E-Mail           : waty.thierry@usa.net
' * Date             : 14/01/99
' * Time             : 16:48
' * Module Name      : Errors
' * Module Filename  : Errors.bas
' **********************************************************************
' * Comments         : Error files for general error handling
' *
' *
' **********************************************************************

Option Explicit

' *** Error Collection
Global gcErrors               As New Collection

' *** Exceptions Collection
Global gcExceptions           As New Collection

' *** View message box or not
Global gbViewMessage          As Boolean

' *** Const ***
Global Const MESSAGE_YES = True

Public Function TreatErrorHandler() As Integer
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 16:49
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : TreatErrorHandler
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *  This function :
   ' *  - Gets the error
   ' *  - Verifies the error collections
   ' *  - Shows a message box to the user
   ' *  - Tell to the faulty calling routine what to do
   ' ********************************************************
   ' *  Returned value  Identification of what to do.
   ' *                  0 = Resume
   ' *                  1 = Resume next
   ' *                  2 = Exit from the procedure
   ' *                  3 = Cancel the application
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR# 'Diseable error handler
   
   ' *** Some variables for the message box
   Dim sMsgTitle        As String         ' Message Box title
   Dim sMsgTxt          As String         ' Text to show in the Message Box
   Dim nMsgAnswer       As Integer        ' Button pressed by the user
   Dim nButtons         As Long
   Dim sSeparator       As String

   ' *** Some variables for the error to treat
   Dim nJ               As Integer
   Dim nI               As Integer
   Dim nPos             As Integer
   Dim nCurrentError    As Integer

   nCurrentError = Err.Number
   sSeparator = Chr$(10)      ' vbCrLf

   ' *** Formatting the error message
   sMsgTitle = "Error N° " & Str$(nCurrentError)

   ' *** Text message
   sMsgTxt = "< " & Err.Description & " >"

   If (gcErrors(gcErrors.Count).Details = True) Then
      sMsgTxt = sMsgTxt & sSeparator & sSeparator
      ' *** Add the comment of the user
      sMsgTxt = sMsgTxt & "Comment :" & sSeparator
      sMsgTxt = sMsgTxt & Chr$(9) & gcErrors(gcErrors.Count).Comment & sSeparator & sSeparator

      ' *** Localisation of the error
      sMsgTxt = sMsgTxt & "Error Localisation :" & sSeparator
      If (Trim(gcErrors(gcErrors.Count).FormCaption) <> "") Then sMsgTxt = sMsgTxt & Chr$(9) & "Form caption : " & gcErrors(gcErrors.Count).FormCaption & sSeparator

      ' *** Identification of the procedure
      If (Trim(gcErrors(gcErrors.Count).ProcName) <> "") Then sMsgTxt = sMsgTxt & Chr$(9) & "Procedure : " & gcErrors(gcErrors.Count).ProcName & sSeparator

      ' *** Add the line number if available
      If Erl > 0 Then sMsgTxt = sMsgTxt & Chr$(9) & "Line :" & Chr$(9) & Chr$(9) & Erl & sSeparator

      ' *** Show the call stack
      sMsgTxt = sMsgTxt & sSeparator
      sMsgTxt = sMsgTxt & "Calling sequence  :" & sSeparator

      nI = 1
      If (gcErrors.Count > 4) Then
         sMsgTxt = sMsgTxt & Chr$(9) & ". . . " & sSeparator
         nI = gcErrors.Count - 4
      End If

      For nJ = nI To gcErrors.Count
         sMsgTxt = sMsgTxt & Chr$(9) & gcErrors(nJ).LevelCascade & " : " & gcErrors(nJ).ProcName & sSeparator
      Next
   End If

   ' *** Store in the logfile
   LogFile CStr(Now) & " " & sMsgTitle & " " & sMsgTxt

   ' *** Show the messagebox if needed
   If (gcErrors(gcErrors.Count).NeedMessageBox = True) Then
      nButtons = gcErrors(gcErrors.Count).Parametres
      If (nButtons = 0) Then nButtons = vbCritical
      nMsgAnswer = MsgBox(sMsgTxt, nButtons, sMsgTitle)
   End If

   ' *** Return the value according the selected button
   If (gcErrors(gcErrors.Count).NeedMessageBox = True) Then
      TreatErrorHandler = nMsgAnswer
   Else
      TreatErrorHandler = 1     ' resume next
   End If

End Function

Private Sub LogFile(sMessage As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 16:59
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : LogFile
   ' * Parameters       :
   ' *                    sMessage As String
   ' **********************************************************************
   ' * Comments         : Store in the log file
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR#
   On Error GoTo ERROR_LogFile

   On Error Resume Next

   Dim nLogFile      As Integer
   Dim sFileName     As String

   ' *** Name of the logfile
   sFileName = App.Path + "\LogError.Log"

   nLogFile = FreeFile

   If (FileLen(sFileName) > 1024000) Then Kill sFileName
   Open sFileName For Append As #nLogFile
   Print #nLogFile, sMessage
   Close #nLogFile

EXIT_LogFile:
   Exit Sub

   ' #VBIDEUtilsERROR#
ERROR_LogFile:
   Resume EXIT_LogFile

End Sub

Sub ErrorHandlerEnd(ByVal sProcedureName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 17:01
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : ErrorHandlerEnd
   ' * Parameters       :
   ' *                    ByVal sProcedureName As String
   ' **********************************************************************
   ' * Comments         : Called at the end of each procedure
   ' * Used to remove the procedure from the call stack
   ' *
   ' **********************************************************************

   If (gcErrors.Count > 0) Then gcErrors.Remove gcErrors.Count

End Sub

Sub ErrorHandlerBegin(ByVal sProcedureName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 17:02
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : ErrorHandlerBegin
   ' * Parameters       :
   ' *                    ByVal sProcedureName As String
   ' **********************************************************************
   ' * Comments         : Called at the beginning of each procedure
   ' * Used to create the call stack
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR# 'Diseable error handler
   
   Dim clsError            As New class_Error

   ' *** Store the name of the active procedure
   clsError.ProcName = sProcedureName

   ' *** Going down from one level
   clsError.LevelCascade = gcErrors.Count + 1

   ' *** By default, there is no error
   clsError.ErrorNumber = 0
   clsError.ErrorName = ""

   ' *** Add in the collection
   gcErrors.Add Item:=clsError, Key:=CStr(clsError.LevelCascade)

End Sub

Sub ErrorHandlerParameter(ByVal sComment As String, ByVal nParametres As Long, ByVal bNeedMessageBox As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 17:05
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : ErrorHandlerParameter
   ' * Parameters       :
   ' *                    ByVal sComment As String
   ' *                    ByVal nParametres As Long
   ' *                    ByVal bNeedMessageBox As Boolean
   ' **********************************************************************
   ' * Comments         : Send parameters
   ' *  - Add more comments
   ' *  - Configure to remove the messagebox
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR# 'Diseable error handler
   
   ' *** Store the comment
   gcErrors(gcErrors.Count).Comment = sComment

   ' *** Supress or not the error message
   gcErrors(gcErrors.Count).NeedMessageBox = bNeedMessageBox

   ' *** Do you want all the details
   gcErrors(gcErrors.Count).Details = True

   ' *** Add in the collection
   gcErrors(gcErrors.Count).Parametres = nParametres

End Sub

Sub ErrorHandlerStartProcedure(ByVal sFormCaption As String, ByVal sProcedureName As String)
   ' #VBIDEUtils#************************************************************
   ' * Programmer Name  : Waty Thierry
   ' * Web Site         : www.geocities.com/ResearchTriangle/6311/
   ' * E-Mail           : waty.thierry@usa.net
   ' * Date             : 14/01/99
   ' * Time             : 17:06
   ' * Module Name      : Errors
   ' * Module Filename  : Errors.bas
   ' * Procedure Name   : ErrorHandlerStartProcedure
   ' * Parameters       :
   ' *                    ByVal sFormCaption As String
   ' *                    ByVal sProcedureName As String
   ' **********************************************************************
   ' * Comments         : Init for a new event
   ' *
   ' *
   ' **********************************************************************

   ' #VBIDEUtilsERROR# 'Diseable error handler
   
   ' *** Init a new event
   ErrorHandlerBegin sProcedureName

   gcErrors(gcErrors.Count).FormCaption = sFormCaption
   gcErrors(gcErrors.Count).ProcName = sProcedureName
   gcErrors(gcErrors.Count).Comment = ""

End Sub
