VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNetConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Karl E. Peterson, 1998
' All Rights Reserved.
' http://www.mvps.org/vb
' ********************************************
' May be freely used in your applications.
' Redistribution of source prohibited.
' ********************************************
Option Explicit
'
' Required WinInet.dll declarations
'
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
'
' // Flags for InternetAutodial
'
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1         ' Forces an online Internet connection.
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2     ' Forces an unattended Internet dial-up.
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4  ' Essentially undocumented???
' Indicates to use config info from registry
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
' WinInet flags
Private Const INTERNET_FLAG_RELOAD = &H80000000 'read from wire even if locally cached
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000    '// use keep-alive semantics
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000    '// don't write this item to the cache
'
' Enumerated dial modes
'
Public Enum AutodialFlags
   InternetAutodialForceOnline = INTERNET_AUTODIAL_FORCE_ONLINE
   InternetAutodialUnattended = INTERNET_AUTODIAL_FORCE_UNATTENDED
   InternetAutodialFailIfSecurityCheck = INTERNET_AUTODIAL_FAILIFSECURITYCHECK
End Enum
'
' // Flags for InternetGetConnectedState
'
Private Const INTERNET_CONNECTION_MODEM = 1
Private Const INTERNET_CONNECTION_LAN = 2
Private Const INTERNET_CONNECTION_PROXY = 4
Private Const INTERNET_CONNECTION_MODEM_BUSY = 8
'
' Enumerated connection states
'
Public Enum InetConnectionStates
   InternetConnectionModem = INTERNET_CONNECTION_MODEM
   InternetConnectionLan = INTERNET_CONNECTION_LAN
   InternetConnectionProxy = INTERNET_CONNECTION_PROXY
   InternetConnectionModemBusy = INTERNET_CONNECTION_MODEM_BUSY
End Enum
'
' Set aside storage for private member variables.
'
Private m_Connected As Boolean
Private m_ConnectMode As InetConnectionStates
Private m_ConnectModeDesc As String
Private m_KnownSites As Collection

' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
   Set m_KnownSites = New Collection
   m_KnownSites.Add "http://www.yahoo.com", "http://www.yahoo.com"
   m_KnownSites.Add "http://www.cnn.com", "http://www.cnn.com"
End Sub

Private Sub Class_Terminate()
   Set m_KnownSites = Nothing
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Get KnownTestSites() As Collection
   Set KnownTestSites = m_KnownSites
End Property

Public Property Let KnownTestSites(ByVal NewVal As Collection)
   Set m_KnownSites = NewVal
End Property

' ********************************************
'  Public Properties // Read-Only
' ********************************************
Public Property Get Connected() As Boolean
   Call Me.Refresh
   Connected = m_Connected
End Property

Public Property Get ConnectMode() As InetConnectionStates
   Call Me.Refresh
   ConnectMode = m_ConnectMode
End Property

Public Property Get ConnectModeDesc() As String
   Call Me.Refresh
   ConnectModeDesc = m_ConnectModeDesc
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Function Dial(Optional Flags As AutodialFlags = InternetAutodialUnattended) As Boolean
   Dial = False
   If Not Connected Then
      If CBool(InternetAutodial(Flags, 0&)) Then
         Dial = True
      Else
         Debug.Print WinInetErrorTextEx(Err.LastDllError)
      End If
   End If
End Function

Public Function HangUp() As Boolean
   If Connected Then
      HangUp = CBool(InternetAutodialHangup(0&))
   Else
      HangUp = False  ' maybe should be True???
   End If
End Function

Public Sub Refresh()
   Dim Flags As Long
   
   m_Connected = InternetGetConnectedState(Flags, 0&)
   m_ConnectMode = Flags
   
   If Flags And INTERNET_CONNECTION_MODEM Then
      m_ConnectModeDesc = "Modem"
      If Flags And INTERNET_CONNECTION_MODEM_BUSY Then
         m_ConnectModeDesc = "Modem (Busy)"
         m_Connected = TryTheWire()
      End If
   ElseIf Flags And INTERNET_CONNECTION_LAN Then
      m_ConnectModeDesc = "LAN"
      m_Connected = TryTheWire()
   ElseIf Flags And INTERNET_CONNECTION_PROXY Then
      m_ConnectModeDesc = "Proxy Server"
      m_Connected = TryTheWire()
   End If
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Function TryTheWire() As Boolean
   Dim hInet As Long
   Dim hUrl As Long
   Dim Flags As Long
   Dim url As Variant
   '
   ' Try opening each URL in the collection, until one succeeds.
   '
   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
   If hInet Then
      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
      For Each url In m_KnownSites
         hUrl = InternetOpenUrl(hInet, CStr(url), vbNullString, 0, Flags, 0)
         If hUrl Then
            Call InternetCloseHandle(hUrl)
            TryTheWire = True
            Exit For
         End If
      Next url
   End If
   Call InternetCloseHandle(hInet)
End Function

