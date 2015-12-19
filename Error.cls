VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** Number of the error
Dim nErrorNumber              As Long

' *** Error name
Dim nErrorName                As String

' *** Function or procedure name
Dim sProcName                 As String

' *** A comment
Dim sComment                  As String

' *** Indicates if we want a message box or not
Dim bMessageBox               As Boolean

' *** Level in the error cascade
Dim nLevelCascade             As Integer

' *** Title of the windows
Dim sFormCaption              As String

' *** Indicate if we want the details or not
Dim bDetails                  As Boolean

' *** Parameters for the message box
Dim nParametres               As Long

Public Property Get ErrorNumber()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We send the error number

   ErrorNumber = nErrorNumber

End Property

Public Property Let ErrorNumber(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the error number

   nErrorNumber = vNewValue

End Property

Public Property Get ErrorName()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We send the error name

   ErrorName = nErrorName

End Property

Public Property Let ErrorName(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the error name

   nErrorName = vNewValue

End Property

Public Property Get ProcName()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We send the function/procedure name

   ProcName = sProcName

End Property

Public Property Let ProcName(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the fucntion/procedure name

   sProcName = vNewValue

End Property

Public Property Get Comment()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We get the comment

   Comment = sComment

End Property

Public Property Let Comment(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the comment

   sComment = vNewValue

End Property

Public Property Get NeedMessageBox()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** Do we need a message box

   NeedMessageBox = bMessageBox

End Property

Public Property Let NeedMessageBox(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler

   ' *** We set the need of a message box

   bMessageBox = vNewValue

End Property

Public Property Get LevelCascade()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We get the error level cascade

   LevelCascade = nLevelCascade

End Property

Public Property Let LevelCascade(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the error level cascade

   nLevelCascade = vNewValue

End Property

Public Property Get FormCaption()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We get the caption of the window

   FormCaption = sFormCaption

End Property

Public Property Let FormCaption(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the caption of the window

   sFormCaption = vNewValue

End Property

Public Property Get Details()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** Do we need all details

   Details = bDetails

End Property

Public Property Let Details(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the need of all details

   bDetails = vNewValue

End Property

Public Property Get Parametres()
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We get the parameters

   Parametres = nParametres

End Property

Public Property Let Parametres(vNewValue)
   ' #VBIDEUtilsERROR# 'Deseable error handler
   ' *** We set the parameters

   nParametres = vNewValue

End Property
