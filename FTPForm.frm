VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FTPForm 
   Caption         =   "Update Keplarian Elements via the Internet"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   780
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "FTPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FTPHostname As String
Dim Response As String

Public Sub writefile(pathname As String, filename As String, IPaddress As String)

  FTPLogin
  FTPHostname = IPaddress
  Inet1.Execute FTPHostname, "PUT " & pathname & filename & " /" & filename

  Do While Inet1.StillExecuting
    DoEvents
  Loop
  Exit Sub
End Sub

Public Sub getfile(pathname As String, filename As String, IPaddress As String)

  FTPLogin
  FTPHostname = IPaddress
  Inet1.Execute FTPHostname, "GET " & filename & " " & pathname & filename

  Do While Inet1.StillExecuting
    DoEvents
  Loop
  Exit Sub
End Sub

Private Sub FTPLogin()

  With Inet1
    .Password = "ag"
    .UserName = "sanderson"
    .AccessType = icNamedProxy
    .Proxy = "192.168.108.176:80"
  
  End With

End Sub

Private Sub Form_Load()
  getfile "c:\temp\", "amateur.txt", "ftp://ftp.celestrak.com/pub/elements/"
End Sub
