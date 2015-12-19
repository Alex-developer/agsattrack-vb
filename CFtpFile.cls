VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFileName                    As String
Private mvarLastWriteTime               As Date
Private mvarFileSize                    As Long
Private mvarIsDirectory                 As Boolean
Private mvarFilePath                    As String
'
Private mvarPermissions                 As String
Private mvarOwner                       As String
Private mvarGroup                       As String

Public Property Let IsDirectory(ByVal vData As Boolean)
    mvarIsDirectory = vData
End Property

Public Property Get IsDirectory() As Boolean
    IsDirectory = mvarIsDirectory
End Property

Public Property Let FileSize(ByVal vData As Long)
    mvarFileSize = vData
End Property

Public Property Get FileSize() As Long
    FileSize = mvarFileSize
End Property

Public Property Let LastWriteTime(ByVal vData As Date)
    mvarLastWriteTime = vData
End Property

Public Property Get LastWriteTime() As Date
    LastWriteTime = mvarLastWriteTime
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let FilePath(ByVal vData As String)
    mvarFilePath = vData
End Property

Public Property Get FilePath() As String
    FilePath = mvarFilePath
End Property

Public Property Let Permissions(ByVal vData As String)
    mvarPermissions = vData
End Property

Public Property Get Permissions() As String
    Permissions = mvarPermissions
End Property

Public Property Let Owner(ByVal vData As String)
    mvarOwner = vData
End Property

Public Property Get Owner() As String
    Owner = mvarOwner
End Property

Public Property Let Group(ByVal vData As String)
    mvarGroup = vData
End Property

Public Property Get Group() As String
    Group = mvarGroup
End Property
