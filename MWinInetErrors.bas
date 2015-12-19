Attribute VB_Name = "MWinInetErrors"
Option Explicit

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long

'//
'// Used to retrieve error text from system DLL errors
'//
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
   (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
   ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
   Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Const INTERNET_ERROR_BASE = 12000
Private Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Private Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Private Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Private Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Private Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Private Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Private Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Private Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Private Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Private Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Private Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Private Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Private Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Private Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Private Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Private Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Private Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Private Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Private Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Private Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Private Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Private Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Private Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Private Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Private Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Private Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Private Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Private Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Private Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Private Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Private Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Private Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)
Private Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Private Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Private Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Private Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Private Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Private Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Private Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Private Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Private Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Private Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Private Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Private Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Private Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Private Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Private Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Private Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Private Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)
'//
'// FTP API errors
'//
Private Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Private Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Private Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)
'//
'// gopher API errors
'//
Private Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Private Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Private Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Private Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Private Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Private Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Private Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Private Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Private Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)
'//
'// HTTP API errors
'//
Private Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Private Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Private Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Private Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Private Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Private Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Private Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Private Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Private Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Private Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Private Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)
'//
'// additional Internet API error codes
'//
Private Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Private Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Private Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Private Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Private Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Private Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)
Private Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Private Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Private Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Private Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)
'//
'// InternetAutodial specific errors
'//
Private Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)
Private Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK
'//
'// Other general API errors that can result from WinInet calls
'//
Private Const ERROR_INVALID_HANDLE = 6&  '   The handle is invalid.

Public Function WinInetErrorText(ByVal ErrNum As Long) As String
   Select Case ErrNum
      Case ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP
         WinInetErrorText = "Client authorization is not set up on this computer."
      Case ERROR_INTERNET_OUT_OF_HANDLES
         WinInetErrorText = "No more handles could be generated at this time."
      Case ERROR_INTERNET_TIMEOUT
         WinInetErrorText = "The request has timed out."
      Case ERROR_INTERNET_EXTENDED_ERROR
         ' An extended error was returned from the server. This is typically
         ' a string or buffer containing a verbose error message. Call
         ' InternetGetLastResponseInfo to retrieve the error text.
         WinInetErrorText = WinInetErrorTextEx(ErrNum)
      Case ERROR_INTERNET_INTERNAL_ERROR
         WinInetErrorText = "An internal error has occurred."
      Case ERROR_INTERNET_INVALID_URL
         WinInetErrorText = "The URL is invalid."
      Case ERROR_INTERNET_UNRECOGNIZED_SCHEME
         WinInetErrorText = "The URL scheme could not be recognized, or is not supported."
      Case ERROR_INTERNET_NAME_NOT_RESOLVED
         WinInetErrorText = "The server name could not be resolved."
      Case ERROR_INTERNET_PROTOCOL_NOT_FOUND
         WinInetErrorText = "The requested protocol could not be located."
      Case ERROR_INTERNET_INVALID_OPTION
         WinInetErrorText = "A request to InternetQueryOption or InternetSetOption specified an invalid option value."
      Case ERROR_INTERNET_BAD_OPTION_LENGTH
         WinInetErrorText = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified."
      Case ERROR_INTERNET_OPTION_NOT_SETTABLE
         WinInetErrorText = "The request option cannot be set, only queried."
      Case ERROR_INTERNET_SHUTDOWN
         WinInetErrorText = "The Win32 Internet function support is being shut down or unloaded."
      Case ERROR_INTERNET_INCORRECT_USER_NAME
         WinInetErrorText = "The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect."
      Case ERROR_INTERNET_INCORRECT_PASSWORD
         WinInetErrorText = "The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect."
      Case ERROR_INTERNET_LOGIN_FAILURE
         WinInetErrorText = "The request to connect and log on to an FTP server failed."
      Case ERROR_INTERNET_INVALID_OPERATION
         WinInetErrorText = "The requested operation is invalid."
      Case ERROR_INTERNET_OPERATION_CANCELLED
         WinInetErrorText = "The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed."
      Case ERROR_INTERNET_INCORRECT_HANDLE_TYPE
         WinInetErrorText = "The type of handle supplied is incorrect for this operation."
      Case ERROR_INTERNET_INCORRECT_HANDLE_STATE
         WinInetErrorText = "The requested operation cannot be carried out because the handle supplied is not in the correct state."
      Case ERROR_INTERNET_NOT_PROXY_REQUEST
         WinInetErrorText = "The request cannot be made via a proxy."
      Case ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND
         WinInetErrorText = "A required registry value could not be located."
      Case ERROR_INTERNET_BAD_REGISTRY_PARAMETER
         WinInetErrorText = "A required registry value was located but is an incorrect type or has an invalid value."
      Case ERROR_INTERNET_NO_DIRECT_ACCESS
         WinInetErrorText = "Direct network access cannot be made at this time."
      Case ERROR_INTERNET_NO_CONTEXT
         WinInetErrorText = "An asynchronous request could not be made because a zero context value was supplied."
      Case ERROR_INTERNET_NO_CALLBACK
         WinInetErrorText = "An asynchronous request could not be made because a callback function has not been set."
      Case ERROR_INTERNET_REQUEST_PENDING
         WinInetErrorText = "The required operation could not be completed because one or more requests are pending."
      Case ERROR_INTERNET_INCORRECT_FORMAT
         WinInetErrorText = "The format of the request is invalid."
      Case ERROR_INTERNET_ITEM_NOT_FOUND
         WinInetErrorText = "The requested item could not be located."
      Case ERROR_INTERNET_CANNOT_CONNECT
         WinInetErrorText = "The attempt to connect to the server failed."
      Case ERROR_INTERNET_CONNECTION_ABORTED
         WinInetErrorText = "The connection with the server has been terminated."
      Case ERROR_INTERNET_CONNECTION_RESET
         WinInetErrorText = "The connection with the server has been reset."
      Case ERROR_INTERNET_FORCE_RETRY
         WinInetErrorText = "The Win32 Internet function needs to redo the request."
      'Case ERROR_INTERNET_ZONE_CROSSING
      '   WinInetErrorText = "Not used in this release."
      Case ERROR_INTERNET_MIXED_SECURITY
         WinInetErrorText = "The content is not entirely secure. Some of the content being viewed may have come from unsecured servers."
      'Case ERROR_INTERNET_SSL_CERT_CN_INVALID
      '   WinInetErrorText = "The certificate returned by an SSL/PCT server is invalid because of a mismatched server name. The server name that was given by the caller does not match the common name inside the certificate."
      Case ERROR_INTERNET_HANDLE_EXISTS
         WinInetErrorText = "The request failed because the handle already exists."
      Case ERROR_FTP_TRANSFER_IN_PROGRESS
         WinInetErrorText = "The requested operation cannot be made on the FTP session handle because an operation is already in progress."
      Case ERROR_FTP_DROPPED
         WinInetErrorText = "The FTP operation was not completed because the session was aborted."
      Case ERROR_GOPHER_PROTOCOL_ERROR
         WinInetErrorText = "An error was detected while parsing data returned from the gopher server."
      Case ERROR_GOPHER_NOT_FILE
         WinInetErrorText = "The request must be made for a file locator."
      Case ERROR_GOPHER_DATA_ERROR
         WinInetErrorText = "An error was detected while receiving data from the gopher server."
      Case ERROR_GOPHER_END_OF_DATA
         WinInetErrorText = "The end of the data has been reached."
      Case ERROR_GOPHER_INVALID_LOCATOR
         WinInetErrorText = "The supplied locator is not valid."
      Case ERROR_GOPHER_INCORRECT_LOCATOR_TYPE
         WinInetErrorText = "The type of the locator is not correct for this operation."
      Case ERROR_GOPHER_NOT_GOPHER_PLUS
         WinInetErrorText = "The requested operation can only be made against a Gopher+ server, or with a locator that specifies a Gopher+ operation."
      Case ERROR_GOPHER_ATTRIBUTE_NOT_FOUND
         WinInetErrorText = "The requested attribute could not be located."
      Case ERROR_GOPHER_UNKNOWN_LOCATOR
         WinInetErrorText = "The locator type is unknown."
      Case ERROR_HTTP_HEADER_NOT_FOUND
         WinInetErrorText = "The requested header could not be located."
      Case ERROR_HTTP_DOWNLEVEL_SERVER
         WinInetErrorText = "The server did not return any headers."
      Case ERROR_HTTP_INVALID_SERVER_RESPONSE
         WinInetErrorText = "The server response could not be parsed."
      Case ERROR_HTTP_INVALID_HEADER
         WinInetErrorText = "The supplied header is invalid."
      Case ERROR_HTTP_INVALID_QUERY_REQUEST
         WinInetErrorText = "The request made to HttpQueryInfo is invalid."
      Case ERROR_HTTP_HEADER_ALREADY_EXISTS
         WinInetErrorText = "The header could not be added because it already exists."
      Case ERROR_INVALID_HANDLE
         WinInetErrorText = "The handle that was passed to the API has been either invalidated or closed."
      Case Else
         WinInetErrorText = ApiErrorText(ErrNum)
   End Select
End Function

Public Function WinInetErrorTextEx(ByVal ErrNum As Long) As String
   Static Buffer As String
   Static nRet As Long
   Static BufLen As Long
   Const ERROR_INSUFFICIENT_BUFFER = 122
   
   'assume failure
   WinInetErrorTextEx = "Error Code " & ErrNum & " not defined."

   nRet = InternetGetLastResponseInfo(ErrNum, Buffer, BufLen)
   If nRet = False Then
      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
         Buffer = Space(BufLen)
         nRet = InternetGetLastResponseInfo(ErrNum, Buffer, BufLen)
      End If
   End If
   
   If nRet Then
      If BufLen Then
         WinInetErrorTextEx = Left(Buffer, BufLen)
      End If
   End If
End Function

Public Function ApiErrorText(ByVal ErrNum As Long) As String
   Dim msg As String
   Dim nRet As Long
   
   msg = Space(1024)
   nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, msg, Len(msg), ByVal 0&)
   If nRet Then
      msg = Left(msg, nRet)
      If Right(msg, 2) = vbCrLf Then
         msg = Left(msg, Len(msg) - 2)
      End If
      ApiErrorText = msg
   Else
      ApiErrorText = "Error (" & Format(ErrNum) & ") not defined."
   End If
End Function


