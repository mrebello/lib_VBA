Attribute VB_Name = "m_VBA_Lib"
Option Explicit
' Remove above line if not in ACCESS
Option Compare Database

Public Const NORMAL_PRIORITY_CLASS As Long = &H20&
Public Const DUPLICATE_CLOSE_SOURCE = &H1
Public Const DUPLICATE_SAME_ACCESS = &H2
Public Const STARTF_USESTDHANDLES As Long = &H100&
Public Const STARTF_USESHOWWINDOW As Long = &H1&
Public Const SW_HIDE As Long = 0&
Public Const ERROR_BROKEN_PIPE = 109
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

Public Const SCUSERAGENT = "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT 5.1)"
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_ASYNC = &H10000000
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000   ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
    
Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111

Public Const REG_SZ = 1
Public Const KEY_ALL_ACCESS = &H2003F
Public Const HKEY_CURRENT_USER = &H80000001
Public Const UNIQUE_NAME = &H0
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = &HFFFF
Public Const WAIT_OBJECT_0 = 0
Public Const WAIT_TIMEOUT = &H102

Public Const ODBC_ADD_DSN = 1
Public Const ODBC_CONFIG_DSN = 2
Public Const ODBC_REMOVE_DSN = 3

Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Const GHND = &H42
Const CF_TEXT = 1

Public Const BLOCK_SIZE = 16384

Public Type GUID
  Data(0 To 15) As Byte
End Type

Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Type STARTUPINFO
  Cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Public Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Public Enum InfoLevelEnum
  http_QUERY_CONTENT_TYPE = 1
  http_QUERY_CONTENT_LENGTH = 5
  http_QUERY_EXPIRES = 10
  http_QUERY_LAST_MODIFIED = 11
  http_QUERY_PRAGMA = 17
  http_QUERY_VERSION = 18
  http_QUERY_STATUS_CODE = 19
  http_QUERY_STATUS_TEXT = 20
  http_QUERY_RAW_HEADERS = 21
  http_QUERY_RAW_HEADERS_CRLF = 22
  http_QUERY_FORWARDED = 30
  http_QUERY_SERVER = 37
  http_QUERY_USER_AGENT = 39
  http_QUERY_SET_COOKIE = 43
  http_QUERY_REQUEST_METHOD = 45
  http_STATUS_DENIED = 401
  http_STATUS_PROXY_AUTH_REQ = 407
End Enum

Private Type IP_ADDR_STRING
  Next As Long
  IpAddress As String * 16
  IpMask As String * 16
  Context As Long
End Type

Private Type IP_ADAPTER_INFO
  Next As Long
  ComboIndex As Long
  AdapterName As String * MAX_ADAPTER_NAME_LENGTH
  Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
  AddressLength As Long
  Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
  Index As Long
  Type As Long
  DhcpEnabled As Long
  CurrentIpAddress As Long
  IpAddressList As IP_ADDR_STRING
  GatewayList As IP_ADDR_STRING
  DhcpServer As IP_ADDR_STRING
  HaveWins As Byte
  PrimaryWinsServer As IP_ADDR_STRING
  SecondaryWinsServer As IP_ADDR_STRING
  LeaseObtained As Long
  LeaseExpires As Long
End Type

Private Type FIXED_INFO
  Hostname As String * MAX_HOSTNAME_LEN
  domainname As String * MAX_DOMAIN_NAME_LEN
  CurrentDnsServer As Long
  DnsServerList As IP_ADDR_STRING
  NodeType As Long
  ScopeId  As String * MAX_SCOPE_ID_LEN
  EnableRouting As Long
  EnableProxy As Long
  EnableDns As Long
End Type

Private Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * 257
    szSystemStatus As String * 129
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Public Type sockaddr_in
  sin_family       As Integer
  sin_port         As Integer
  sin_addr         As Long
  sin_zero(1 To 8) As Byte
End Type

Public Type HOSTENT
  hName     As Long
  hAliases  As Long
  hAddrType As Integer
  hLength   As Integer
  hAddrList As Long
End Type


#If VBA7 Then
Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare PtrSafe Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare PtrSafe Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Declare PtrSafe Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Declare PtrSafe Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare PtrSafe Function httpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hhttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Declare PtrSafe Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Declare PtrSafe Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, ByVal lpdwBufferLength As Long) As Integer
Declare PtrSafe Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Declare PtrSafe Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Declare PtrSafe Function GetAdaptersInfo Lib "IPHlpApi.dll" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Boolean
Declare PtrSafe Function Lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function CharToOem& Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String)
Declare PtrSafe Function OemToChar& Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String)
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function aht_apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare PtrSafe Function aht_apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare PtrSafe Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare PtrSafe Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Public Declare PtrSafe Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare PtrSafe Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal CP As String) As Long
Public Declare PtrSafe Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
#Else
Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function httpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hhttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, ByVal lpdwBufferLength As Long) As Integer
Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Declare Function GetAdaptersInfo Lib "IPHlpApi.dll" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Boolean
Declare Function Lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function CharToOem& Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String)
Declare Function OemToChar& Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function aht_apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function aht_apiGetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal CP As String) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
#End If

'============= Functions from APIs ==============

Public Function Get_Temp_Dir() As String
  Get_Temp_Dir = Environ$("TEMP")
End Function


Public Function Get_Temp_File(Optional sPrefix As String = "VBA", Optional sExtensao As String = "") As String
  Dim sTmpPath As String * 512
  Dim sTmpName As String * 576
  Dim nRet As Long
  Dim f As String
  nRet = GetTempPath(512, sTmpPath)
  If (nRet > 0 And nRet < 512) Then
    nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then f = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    If sExtensao > "" Then
      Kill f
      If Right(f, 4) = ".tmp" Then f = Left(f, Len(f) - 4)
      f = f & sExtensao
    End If
    Get_Temp_File = f
  End If
End Function


Public Function New_GUID()
  Dim udtGUID As GUID
  If (CoCreateGuid(udtGUID) = 0) Then New_GUID = udtGUID.Data
End Function


Public Function ShellEx(CommandLine As String, Optional WinStyle As VbAppWinStyle = vbNormalFocus, Optional Wait As Boolean = False, Optional ByExtension As Boolean = False) As Long
  ' If ByExtension, then use ShellExecute (exec a program associate to the extension of file in CommandLine)
  ' If Wait, then wait program to finalize until return
  Dim lPid As Long
  Dim lHnd As Long
  Dim lRet As Long
  If ByExtension Then
    If Wait Then
      MsgBox "Not implemented."
    Else
      ShellEx = ShellExecute(0, "", CommandLine, "", "", WinStyle)
    End If
  Else
    lPid = Shell(CommandLine, WinStyle)
    ShellEx = lPid
    If Wait And (lPid <> 0) Then
      lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
      If lHnd <> 0 Then
        lRet = WaitForSingleObject(lHnd, INFINITE)
        CloseHandle (lHnd)
      End If
    End If
  End If
End Function


Public Function Shell_InOut(ByVal CommandLine As String, Optional StrIn As String = "") As String
  ' tratamento especial para arquivo %out%
  Dim T_BAT As String
  Dim T_OUT As String
  Dim T_IN As String
  Dim r As String
  T_BAT = Get_Temp_File(, ".bat")
  T_OUT = Get_Temp_File
  If StrIn > "" Then
    T_IN = Get_Temp_File
    File_Save T_IN, StrIn
    CommandLine = CommandLine & " < " & T_IN
  End If
  If InStr(CommandLine, "%out%") > 0 Then
    CommandLine = Replace(CommandLine, "%out%", T_OUT)
  Else
    CommandLine = CommandLine & " > " & T_OUT & " 2>&1"
  End If
  File_Save T_BAT, CommandLine
  ShellEx T_BAT, vbHide, True
  Shell_InOut = File_Load(T_OUT)
  If T_IN > "" Then Kill T_IN
  Kill T_BAT
  Kill T_OUT
End Function


Public Function Shell_Escape(ByVal Text As String) As String
  ' return ecaped text to shell
  If InStr(Text, " ") > 0 And Left(Text, 1) <> """" Then Text = """" & Text & """"
  Shell_Escape = Text
End Function


Public Sub DSN(ByVal sDSN As String, ByVal sDriver As String, ByVal sServer As String, ByVal sBD As String, ByVal lAction As Long, Optional sUserOrParameters As String = "", Optional sPass As String = "")
  ' sUserOrParameters - if have "=" and sPass="", = Parameters, and replace ";" by vbNull
  Dim sAttributes As String
  Dim sDBQ As String
  Dim lngRet As Long
  
  Dim hKey As Long
  Dim regValue As String
  Dim valueType As Long

  ' consulta o registro para verificar se o DSN já esta instalado
  If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\" & sDSN, 0, KEY_ALL_ACCESS, hKey) = 0 Then   ' zero significa sem errosr => Retorna o valor de chave "DBQ"
    regValue = String$(1024, 0)        ' Aloca espaço para a variável
    If RegQueryValueEx(hKey, "DBQ", 0, valueType, regValue, Len(regValue)) = 0 Then       ' zero signifia sem erros, podemo retornar o valor
      If valueType = REG_SZ Then sDBQ = Left$(regValue, InStr(regValue, vbNullChar) - 1)
    End If
    RegCloseKey hKey
  End If
  
  ' Realiza a ação somente se você esta incluindo um DSN que nao existe ou remove um existente
  If (sDBQ = "" And lAction = ODBC_ADD_DSN) Or (sDBQ <> "" And lAction = ODBC_REMOVE_DSN) Then
    sAttributes = "DSN=" & sDSN & vbNullChar & "Server=" & sServer & vbNullChar & "Database=" & sBD & vbNullChar
    If sUserOrParameters = "" Then
      sAttributes = sAttributes & "Trusted_Connection=Yes"
    Else
      If sPass = "" And InStr(sUserOrParameters, "=") > 0 Then
        sAttributes = sAttributes & Replace(sUserOrParameters, ";", vbNullChar)
      Else
        sAttributes = sAttributes & "User Id=" & sUserOrParameters & vbNullChar & "Password=" & sPass
      End If
    End If
    lngRet = SQLConfigDataSource(0&, lAction, sDriver, sAttributes)
  End If
End Sub


Public Function Open_Url(ByVal sURL As String, Optional LFtoCRLF As Boolean = False, Optional Flags As Long = 0) As String
'::  ::orig. author R. Woodward::Ryan_Woodward@yahoo.com::
'::DESC:
'::  Retrieve the page specified by "url"
'::  Returns string of page source
'::  On error, returns "error #"
'::     e.g. page not found returns "error 404"
'::
  Dim s As String
  Dim sReadBuf As String * 2048   'a data buffer for InternetOpen fcns
  Dim bytesRead As Long
  Dim hInet As Long       'wininet handle
  Dim hUrl As Long        'url request handle
  Dim flagMoreData As Boolean
  Dim ret As String
  ' used for callling httpQueryInfo
  Dim sErrBuf As String * 255
  Dim sErrBufLen As Long
  Dim dwIndex As Long
  ' return codes and err code saves
  Dim lastErr As Long
  Dim bRet As Boolean
  Dim wRet As Integer
  ' http status code
  Dim httpCode As Integer
  ' grab a handle for using wininet
  hInet = InternetOpen(SCUSERAGENT, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If Err.LastDllError <> 0 Then
    lastErr = Err.LastDllError
    ret = "error (wininet.dll," & lastErr & ")"
    GoTo exitfunc
  End If
  ' retrieve the requested URL
  hUrl = InternetOpenUrl(hInet, sURL, vbNullString, 0, Flags, 0)
  If Err.LastDllError <> 0 Then
    lastErr = Err.LastDllError
    ret = "error (wininet.dll," & lastErr & ")"
    GoTo exitfunc
  End If
  ' get query info, this should give us a status code among other things
  sErrBufLen = 255
  bRet = httpQueryInfo(hUrl, 19, ByVal sErrBuf, sErrBufLen, dwIndex)
  If Err.LastDllError <> 0 Then
    lastErr = Err.LastDllError
    ret = "error (wininet.dll," & lastErr & ")"
    GoTo exitfunc
  End If
  ' sErrBuf should now hopefully contain http status code stuff
  ' if the call failed, no status info was returned (i.e. sErrBuf is empty)
  '   then throw error
  If sErrBufLen = 0 Or Not bRet Then
    ret = "error"
    GoTo exitfunc
  Else
    ' retrieve the http status code
    httpCode = CInt(Left(sErrBuf, sErrBufLen))
    If httpCode >= 300 Then
      ret = "error " & httpCode
      GoTo exitfunc
    End If
  End If
  ' if we made it this far, then we can begin retrieving data
  flagMoreData = True
  Do While flagMoreData
    sReadBuf = vbNullString
    wRet = InternetReadFile(hUrl, sReadBuf, Len(sReadBuf), bytesRead)
    If Err.LastDllError <> 0 Then
      lastErr = Err.LastDllError
      ret = "error (wininet.dll," & lastErr & ")"
      GoTo exitfunc
    End If
    If wRet <> 1 Then
      ret = "error"
      GoTo exitfunc
    End If
    s = s & Left$(sReadBuf, bytesRead)
    If Not CBool(bytesRead) Then flagMoreData = False
  Loop
  ret = s
exitfunc:
  If hUrl <> 0 Then InternetCloseHandle (hUrl)
  If hInet <> 0 Then InternetCloseHandle (hInet)
  If LFtoCRLF Then
    Open_Url = Replace(ret, vbLf, vbCrLf)
  Else
    Open_Url = ret
  End If
End Function

Function TCPConnection(strHostName As String, RemotePort As Long, strSend As String) As String
  Dim udtWinsockData  As WSAData
  Dim lngSocket As Long
  Dim lngAddress As Long
  Dim lngPtrToHOSTENT As Long
  Dim udtHostent      As HOSTENT
  Dim lngPtrToIP      As Long
  Dim udtSocketAddress As sockaddr_in
  Dim arrBuffer()     As Byte
  Dim lngResult As Long
  Const MAX_BUFFER_LENGTH As Long = 8192
  Dim arrBuffer2(1 To MAX_BUFFER_LENGTH)   As Byte
  Dim lngBytesReceived                    As Long
  Dim r As String
   
  WSAStartup &H202, udtWinsockData
  lngSocket = socket(2, 1, 6)
  If lngSocket < 0 Then
    MsgBox "Erro."
    Exit Function
  End If
  ' GetAddressLong
  lngAddress = inet_addr(strHostName)
  If lngAddress = &HFFFF Then   ' INADDR_NONE
    lngPtrToHOSTENT = gethostbyname(strHostName)
    If lngPtrToHOSTENT <> 0 Then
      RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
      RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
      RtlMoveMemory lngAddress, lngPtrToIP, udtHostent.hLength
    Else
      lngAddress = &HFFFF
    End If
  End If
  ' Connect
  With udtSocketAddress
    .sin_addr = lngAddress
    .sin_port = htons(UnsignedToInteger(RemotePort))
    .sin_family = 2
  End With
  lngResult = Connect(lngSocket, udtSocketAddress, LenB(udtSocketAddress))
  If lngResult = 0 Then ' Conected
    ' Send
    arrBuffer() = StrConv(strSend, vbFromUnicode)
    lngResult = Send(lngSocket, arrBuffer(0), Len(strSend), 0&)
    ' Loop receiving data
    lngResult = 1
    While lngResult > 0    ' erro quando a conexão for fechada
      lngResult = recv(lngSocket, arrBuffer2(1), MAX_BUFFER_LENGTH, 0&)
      If lngResult > 0 Then
        r = r + Left$(StrConv(arrBuffer2, vbUnicode), lngResult)
      End If
    Wend
    ' close connection
    closesocket lngSocket
  Else
    r = "*ERR*"
  End If
  WSACleanup
  TCPConnection = r
End Function


Public Function UnsignedToInteger(value As Long) As Integer
 ' This function takes a Long containing a value in the range of an unsigned Integer and returns an Integer that you can pass to an API that requires an unsigned Integer
 If value < 0 Or value >= 65536 Then Error 6 ' Overflow
 If value <= 32767 Then
   UnsignedToInteger = value
 Else
   UnsignedToInteger = value - 65536
 End If
End Function

Public Function IP_Gateway() As String
  Dim error As Long, i As Integer
  Dim AdapterInfoBuffer() As Byte, AdapterInfo As IP_ADAPTER_INFO, AdapterInfoSize As Long
  Dim Buffer As IP_ADDR_STRING
  Dim pAdapt As Long
  AdapterInfoSize = 0
  error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
  If error <> 0 Then If error <> ERROR_BUFFER_OVERFLOW Then IP_Gateway = "": Exit Function
  ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
  error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
  If error <> 0 Then IP_Gateway = "": Exit Function
  CopyMemory AdapterInfo, AdapterInfoBuffer(0), Len(AdapterInfo)  '  AdapterInfoSize
  pAdapt = AdapterInfo.Next
  Do
    If Asc(AdapterInfo.GatewayList.IpAddress) > 0 Then
      IP_Gateway = Before(AdapterInfo.GatewayList.IpAddress, Chr$(0))
      Exit Function
    End If
    pAdapt = AdapterInfo.Next
    If pAdapt <> 0 Then CopyMemory AdapterInfo, ByVal pAdapt, Len(AdapterInfo)  ' AdapterInfoSize
  Loop Until pAdapt = 0
End Function
'Sub Main()
'    Dim error As Long
'    Dim FixedInfoSize As Long
'    Dim AdapterInfoSize As Long
'    Dim i As Integer
'    Dim PhysicalAddress  As String
'    Dim NewTime As Date
'    Dim AdapterInfo As IP_ADAPTER_INFO
'    Dim AddrStr As IP_ADDR_STRING
'    Dim FixedInfo As FIXED_INFO
'    Dim Buffer As IP_ADDR_STRING
'    Dim pAddrStr As Long
'    Dim pAdapt As Long
'    Dim Buffer2 As IP_ADAPTER_INFO
'    Dim FixedInfoBuffer() As Byte
'    Dim AdapterInfoBuffer() As Byte
'
'    ' Get the main IP configuration information for this machine
'    ' using a FIXED_INFO structure.
'    FixedInfoSize = 0
'    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
'    If error <> 0 Then
'        If error <> ERROR_BUFFER_OVERFLOW Then
'           MsgBox "GetNetworkParams sizing failed with error " & error
'           Exit Sub
'        End If
'    End If
'    ReDim FixedInfoBuffer(FixedInfoSize - 1)
'
'    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
'    If error = 0 Then
'            CopyMemory FixedInfo, FixedInfoBuffer(0), FixedInfoSize
'            MsgBox "Host Name:  " & FixedInfo.Hostname
'            MsgBox "DNS Servers:  " & FixedInfo.DnsServerList.IpAddress
'            pAddrStr = FixedInfo.DnsServerList.Next
'            Do While pAddrStr <> 0
'                  CopyMemory Buffer, ByVal pAddrStr, LenB(Buffer)
'                  MsgBox "DNS Servers:  " & Buffer.IpAddress
'                  pAddrStr = Buffer.Next
'            Loop
'
'            Select Case FixedInfo.NodeType
'                       Case 1
'                                  MsgBox "Node type: Broadcast"
'                       Case 2
'                                  MsgBox "Node type: Peer to peer"
'                       Case 4
'                                  MsgBox "Node type: Mixed"
'                       Case 8
'                                  MsgBox "Node type: Hybrid"
'                       Case Else
'                                  MsgBox "Unknown node type"
'            End Select
'
'            MsgBox "NetBIOS Scope ID:  " & FixedInfo.ScopeId
'            If FixedInfo.EnableRouting Then
'                       MsgBox "IP Routing Enabled "
'            Else
'                       MsgBox "IP Routing not enabled"
'            End If
'            If FixedInfo.EnableProxy Then
'                       MsgBox "WINS Proxy Enabled "
'            Else
'                       MsgBox "WINS Proxy not Enabled "
'            End If
'            If FixedInfo.EnableDns Then
'                      MsgBox "NetBIOS Resolution Uses DNS "
'            Else
'                      MsgBox "NetBIOS Resolution Does not use DNS  "
'            End If
'    Else
'            MsgBox "GetNetworkParams failed with error " & error
'            Exit Sub
'    End If
'
'    ' Enumerate all of the adapter specific information using the
'    ' IP_ADAPTER_INFO structure.
'    ' Note:  IP_ADAPTER_INFO contains a linked list of adapter entries.
'
'    AdapterInfoSize = 0
'    error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
'    If error <> 0 Then
'        If error <> ERROR_BUFFER_OVERFLOW Then
'           MsgBox "GetAdaptersInfo sizing failed with error " & error
'           Exit Sub
'        End If
'    End If
'    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)
'
'    ' Get actual adapter information
'    error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
'    If error <> 0 Then
'       MsgBox "GetAdaptersInfo failed with error " & error
'       Exit Sub
'    End If
'
'    ' Allocate memory
'     CopyMemory AdapterInfo, AdapterInfoBuffer(0), AdapterInfoSize
'    pAdapt = AdapterInfo.Next
'
'    Do
'     CopyMemory Buffer2, AdapterInfo, AdapterInfoSize
'       Select Case Buffer2.Type
'              Case MIB_IF_TYPE_ETHERNET
'                   MsgBox "Adapter name: Ethernet adapter "
'              Case MIB_IF_TYPE_TOKENRING
'                   MsgBox "Adapter name: Token Ring adapter "
'              Case MIB_IF_TYPE_FDDI
'                   MsgBox "Adapter name: FDDI adapter "
'              Case MIB_IF_TYPE_PPP
'                   MsgBox "Adapter name: PPP adapter"
'              Case MIB_IF_TYPE_LOOPBACK
'                   MsgBox "Adapter name: Loopback adapter "
'              Case MIB_IF_TYPE_SLIP
'                   MsgBox "Adapter name: Slip adapter "
'              Case Else
'                   MsgBox "Adapter name: Other adapter "
'       End Select
'       MsgBox "AdapterDescription: " & Buffer2.Description
'
'       PhysicalAddress = ""
'       For i = 0 To Buffer2.AddressLength - 1
'           PhysicalAddress = PhysicalAddress & Hex(Buffer2.Address(i))
'           If i < Buffer2.AddressLength - 1 Then
'              PhysicalAddress = PhysicalAddress & "-"
'           End If
'       Next
'       MsgBox "Physical Address: " & PhysicalAddress
'
'       If Buffer2.DhcpEnabled Then
'          MsgBox "DHCP Enabled "
'       Else
'          MsgBox "DHCP disabled"
'       End If
'
'       MsgBox "IP Address: " & Buffer2.IpAddressList.IpAddress
'       MsgBox "Subnet Mask: " & Buffer2.IpAddressList.IpMask
'       pAddrStr = Buffer2.IpAddressList.Next
'       Do While pAddrStr <> 0
'          CopyMemory Buffer, Buffer2.IpAddressList, LenB(Buffer)
'          MsgBox "IP Address: " & Buffer.IpAddress
'          MsgBox "Subnet Mask: " & Buffer.IpMask
'          pAddrStr = Buffer.Next
'          If pAddrStr <> 0 Then
'             CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, _
'                        LenB(Buffer2.IpAddressList)
'          End If
'       Loop
'
'       MsgBox "Default Gateway: " & Buffer2.GatewayList.IpAddress
'       pAddrStr = Buffer2.GatewayList.Next
'       Do While pAddrStr <> 0
'          CopyMemory Buffer, Buffer2.GatewayList, LenB(Buffer)
'          MsgBox "IP Address: " & Buffer.IpAddress
'          pAddrStr = Buffer.Next
'          If pAddrStr <> 0 Then
'             CopyMemory Buffer2.GatewayList, ByVal pAddrStr, _
'                        LenB(Buffer2.GatewayList)
'          End If
'       Loop
'
'       MsgBox "DHCP Server: " & Buffer2.DhcpServer.IpAddress
'       MsgBox "Primary WINS Server: " & _
'              Buffer2.PrimaryWinsServer.IpAddress
'       MsgBox "Secondary WINS Server: " & _
'              Buffer2.SecondaryWinsServer.IpAddress
'
'       ' Display time.
'       NewTime = DateAdd("s", Buffer2.LeaseObtained, #1/1/1970#)
'       MsgBox "Lease Obtained: " & _
'              CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))
'
'       NewTime = DateAdd("s", Buffer2.LeaseExpires, #1/1/1970#)
'       MsgBox "Lease Expires :  " & _
'              CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))
'       pAdapt = Buffer2.Next
'       If pAdapt <> 0 Then
'           CopyMemory AdapterInfo, ByVal pAdapt, AdapterInfoSize
'        End If
'      Loop Until pAdapt = 0
'End Sub


Sub SetClipboard(szText As String)
    Dim wLen As Integer, hMemory As Long, lpMemory As Long
    Dim RetVal As Variant, wFreeMemory As Boolean
    ' Get the length, including one extra for a CHR$(0) at the end.
    wLen = Len(szText) + 1
    szText = szText & Chr$(0)
    hMemory = GlobalAlloc(GHND, wLen + 1)
    If hMemory = 0 Then MsgBox "Unable to allocate memory.": Exit Sub
    wFreeMemory = True
    lpMemory = GlobalLock(hMemory)
    If lpMemory = 0 Then MsgBox "Unable to lock memory.": GoTo T2CB_Free
    ' Copy our string into the locked memory.
    RetVal = Lstrcpy(lpMemory, szText)
    ' Don't send clipboard locked memory.
    RetVal = GlobalUnlock(hMemory)
    If OpenClipboard(0&) = 0 Then MsgBox "Unable to open Clipboard.  Perhaps some other application is using it.": GoTo T2CB_Free
    If EmptyClipboard() = 0 Then MsgBox "Unable to empty the clipboard.": GoTo T2CB_Close
    If SetClipboardData(CF_TEXT, hMemory) = 0 Then MsgBox "Unable to set the clipboard data.": GoTo T2CB_Close
    wFreeMemory = False
T2CB_Close:
    If CloseClipboard() = 0 Then MsgBox "Unable to close the Clipboard."
    If wFreeMemory Then GoTo T2CB_Free
    Exit Sub
T2CB_Free:
    If GlobalFree(hMemory) <> 0 Then MsgBox "Unable to free global memory."
End Sub


Function GetClipboard()
    Dim wLen As Integer, hMemory As Long, hMyMemory As Long
    Dim lpMemory As Long, lpMyMemory As Long, szText As String, wSize As Long
    Dim RetVal As Variant, wFreeMemory As Boolean, wClipAvail As Integer
    If IsClipboardFormatAvailable(CF_TEXT) = 0 Then GetClipboard = Null: Exit Function
    If OpenClipboard(0&) = 0 Then MsgBox "Unable to open Clipboard.  Perhaps some other application is using it.": GoTo CB2T_Free
    hMemory = GetClipboardData(CF_TEXT)
    If hMemory = 0 Then MsgBox "Unable to retrieve text from the Clipboard.": Exit Function
    wSize = GlobalSize(hMemory)
    szText = space(wSize)
    wFreeMemory = True
    lpMemory = GlobalLock(hMemory)
    If lpMemory = 0 Then MsgBox "Unable to lock clipboard memory.": GoTo CB2T_Free
    ' Copy our string into the locked memory.
    RetVal = Lstrcpy(szText, lpMemory)
    ' Get rid of trailing stuff.
    szText = Trim(szText)
    ' Get rid of trailing 0.
    GetClipboard = Left(szText, Len(szText) - 1)
    wFreeMemory = False
CB2T_Close:
    If abCloseClipboard() = 0 Then MsgBox "Unable to close the Clipboard."
    If wFreeMemory Then GoTo CB2T_Free
    Exit Function
CB2T_Free:
    If abGlobalFree(hMemory) <> 0 Then MsgBox "Unable to free global clipboard memory."
End Function

Function Convert2ansi(in_string) As String
  Dim Out_String As String
  Out_String = space(Len(in_string))
  OemToChar in_string, Out_String
  Convert2ansi = Out_String
End Function

Function ahtCommonFileOpenSave(Optional ByRef Flags As Variant, Optional ByVal InitialDir As Variant, _
            Optional ByVal Filter As Variant, Optional ByVal FilterIndex As Variant, Optional ByVal DefaultExt As Variant, _
            Optional ByVal FileName As Variant, Optional ByVal DialogTitle As Variant, _
            Optional ByVal hwnd As Variant, Optional ByVal OpenFile As Variant) As Variant
    ' In:
    ' Flags: one or more of the ahtOFN_* constants, OR'd together.
    ' InitialDir: the directory in which to first look
    ' Filter: a set of file filters, with pipe "Databases|*.mdb;*.mda|Text|*.txt"
    ' FilterIndex: 1-based integer indicating which filter
    ' set to use, by default (1 if unspecified)
    ' DefaultExt: Extension to use if the user doesn't enter one.
    ' Only useful on file saves.
    ' FileName: Default value for the file name text box.
    ' DialogTitle: Title for the dialog.
    ' hWnd: parent window handle
    ' OpenFile: Boolean(True=Open File/False=Save As)
    ' Out:
    ' Return Value: Either Null or the selected filename
    Dim OFN As tagOPENFILENAME
    Dim strFileName As String
    Dim strFileTitle As String
    Dim fResult As Boolean
    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFileName = Left(FileName & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = Replace(Filter & "|", "|", vbNullChar)
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If
    If fResult Then
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        If Flags And ahtOFN_ALLOWMULTISELECT Then
            ' Return the full array.
            Dim items As Variant
            Dim value As String
            value = OFN.strFile
            ' Get rid of empty items:
            Dim i As Integer
            For i = Len(value) To 1 Step -1
              If Mid$(value, i, 1) <> Chr$(0) Then
                Exit For
              End If
            Next i
            value = Mid(value, 1, i)
            ' Break the list up at null characters:
            items = Split(value, Chr(0))
            ' Loop through the items in the "array",
            ' and build full file names:
            Dim numItems As Integer
            Dim result() As String
            numItems = UBound(items) + 1
            If numItems > 1 Then
                ReDim result(0 To numItems - 2)
                For i = 1 To numItems - 1
                    result(i - 1) = RTrimEx((items(0)), "\") & "\" & items(i)
                Next i
                ahtCommonFileOpenSave = result
            Else
                ahtCommonFileOpenSave = items(0)
            End If
        Else
          Dim intPos As Integer
          intPos = InStr(OFN.strFile, vbNullChar)
          If intPos > 0 Then
            ahtCommonFileOpenSave = Left(OFN.strFile, intPos - 1)
          Else
            ahtCommonFileOpenSave = OFN.strFile
          End If
        End If
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function


'============= WMI =================

Public Function Ping(ByVal ComputerName As String)
  Dim oPingResult As Variant
  Ping = False
  For Each oPingResult In GetObject("winmgmts://./root/cimv2").ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & ComputerName & "'")
    If IsObject(oPingResult) Then
      If oPingResult.StatusCode = 0 Then
        Ping = True
        Exit Function
      End If
    End If
  Next
End Function


'============= String Functions ==============

Public Function SQL_Value(x As Variant, Optional Server As String = "ACCESS") As String
  ' Return the value in a string escaped format to SQL Sintax
  ' Server = ACCESS / POSTGRE / PG / SQL / SQP_SP
  ' Server = SQL_SP  -> convert to value passed as a parameter in 'exec'
  Dim tipo As String
  If IsNull(x) Then
    SQL_Value = "NULL"
    Exit Function
  End If
  Server = UCase(Server)
  tipo = TypeName(x)
  If tipo = "AccessField" Or tipo = "Field" Then
    tipo = TypeName(x.value)
  ElseIf tipo = "Field2" Then
    tipo = x.Type
  End If
  If tipo = "String" Then
    SQL_Value = "'" & Replace(x, "'", "''") & "'"
  ElseIf tipo = "Double" Or tipo = "Single" Then
    SQL_Value = Replace(x, ",", ".")
  ElseIf tipo = "Date" Then
    If Server = "ACCESS" Then
      ' use of dateserial avoids date format mistakes
      Dim r As String
      If Year(x) <> 1899 Or Month(x) <> 12 Or Day(x) <> 30 Then
        r = "dateserial(" & Year(x) & "," & Month(x) & "," & Day(x) & ")"
      End If
      If Hour(x) <> 0 Or Minute(x) <> 0 Or Second(x) <> 0 Then
        If r > "" Then r = r & "+"
        r = r & "timeserial(" & Hour(x) & "," & Minute(x) & "," & Second(x) & ")"
      End If
      SQL_Value = r
    ElseIf Server = "POSTGRES" Or Server = "POSTGRESQL" Or Server = "PG" Then
      SQL_Value = " date '" & Year(x) & "-" & Month(x) & "-" & Day(x) & "'"
    ElseIf Server = "SQL_SP" Then
      SQL_Value = "'" & Format(x, "yyyy-mm-dd hh:mm:ss") & "'"
'    ElseIf Server = "SQL" Then
'      SQL_Value = "datefromparts(" & Year(x) & "," & Month(x) & "," & Day(x) & "," & Hour(x) & "," & Minute(x) & "," & Second(x) & ")"
    Else
      ' use of dateadd avoids date format mistakes
      r = 0
      If Year(x) <> 1899 Or Month(x) <> 12 Or Day(x) <> 30 Then
        r = "dateadd(d," & (Day(x) - 1) & ",dateadd(m," & ((Year(x) - 1900) * 12 + Month(x) - 1) & ",0))"
      End If
      If Hour(x) <> 0 Or Minute(x) <> 0 Or Second(x) <> 0 Then
        r = "dateadd(s," & (Hour(x) * 3600 + Minute(x) * 60 + Second(x)) & ", " & r & ")"
      End If
      SQL_Value = r
    End If
  ElseIf tipo = "15" Then ' GUID para Field2
    SQL_Value = GUID_Clean(x)
  Else
    SQL_Value = "" & x
  End If
End Function


Public Function SQL_Format(Server As String, command As String, ParamArray var() As Variant) As String
  Dim i As Integer, r As String, p As Integer, pa As Integer
  p = InStr(command, "?")
  pa = 1
  Do While (p > 0)
    r = r & Mid(command, pa, p - pa) & SQL_Value(var(i), Server)
    pa = p + 1
    i = i + 1
    If i > UBound(var) Then Exit Do
    p = InStr(pa, command, "?")
  Loop
  SQL_Format = r & Mid(command, pa)
End Function


Public Function Remove_Accents(ByVal a As String) As String
  Dim x&, l&, c$, p&, cs$, s1$, S2$
  s1 = "áàãâäéèêëíìîïóòõôöúùûüçÁÀÃÂÄÉÈÊËÍÍÎÏÓÒÕÔÖÚÙÛÜÇ"
  S2 = "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC"
  l = Len(s1)
  For x = 1 To l
    c = Mid(s1, x, 1)
    cs = Mid(S2, x, 1)
    p = InStr(1, a, c, vbBinaryCompare)
    While p > 0
      Mid(a, p, 1) = cs
      p = InStr(p + 1, a, c, vbBinaryCompare)
    Wend
  Next x
  Remove_Accents = a
End Function


Public Function Item(Text As String, Index As Integer, Optional Delimiter As String = ";", Optional NotEmpty As Boolean = False) As String
  Dim r
  If NotEmpty Then
    Dim x%, i
    x = 0
    For Each i In Split(Text, Delimiter)
      If i > "" Then x = x + 1
      If x = Index Then
        Item = i
        Exit Function
      End If
    Next
  Else
    r = Split(Text, Delimiter)
    If Index < 0 Then Index = UBound(r) + Index + 2
    If Index >= 1 And Index - 1 <= UBound(r) Then Item = r(Index - 1)
  End If
End Function


Public Function Subst_Car(ByVal Text As String, CharSet1 As String, CharSet2 As String) As String
' Substitui caracteres de s1 por caracteres de s2
' substituicao carater a carater: Subst_Car("1234","13","AB")="A2B4"
  Dim x&, l&, c$, p&, cs$
  l = Len(CharSet1)
  For x = 1 To l
    c = Mid(CharSet1, x, 1)
    cs = Mid(CharSet2, x, 1)
    p = InStr(Text, c)
    While p > 0
      Mid(Text, p, 1) = cs
      p = InStr(p + 1, Text, c)
    Wend
  Next x
  Subst_Car = Text
End Function


Public Function Remove_Car(ByVal Text As String, CharSet1 As String) As String
' Remove todos os caracteres CharSet1 de Text
' Ex: Remove_Car("1.23-4/3",".-/")="12343"
  Dim x&, l&, c$, p&, cs$
  l = Len(CharSet1)
  For x = 1 To l
    c = Mid(CharSet1, x, 1)
    Text = Replace(Text, c, "")
  Next x
  Remove_Car = Text
End Function


Public Function File_Legal_Name(FileName As String)
  ' Remove illegal characters from FileName
  If Right(FileName, 1) = "." Then FileName = Left(FileName, Len(FileName) - 1)
  File_Legal_Name = Replace(Subst_Car(Remove_Accents(FileName), "º,:!@#$%¨&*/\ .()", "o_____________+++"), "+", "")
End Function


Public Function HTMLs2HTML(HTMLs As String) As String
  ' HTMLs = HTML simple - only <B> and <I>
  Dim r$
  r = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><HTML><HEAD>" & _
       "<META http-equiv=Content-Type content=""text/html; charset=unicode"">" & _
       "<META content=""MSHTML 6.00.6000.16414"" name=GENERATOR></HEAD>" & _
       "<BODY><P class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><SPAN style=""FONT-FAMILY: Arial"">"
  r = r & HTML2HTMLs(HTMLs)
  r = r & "</SPAN></P></BODY></HTML>"
  HTMLs2HTML = r
End Function


Public Function HTML2HTMLs(HTML As String) As String
  Dim r$, b$
  Dim p As Long, pa As Long, PF As Long
  p = InStr(HTML, "<")
  While p > 0
    r = r & Mid(HTML, pa + 1, p - pa - 1)
    PF = InStr(p + 1, HTML, ">")
    b = UCase(Mid(HTML, p + 1, PF - p - 1))
    If b = "STRONG" Then b = "B"
    If b = "/STRONG" Then b = "/B"
    If b = "EM" Then b = "I"
    If b = "/EM" Then b = "/I"
    If InStr("B,I,/B,/I,", b + ",") > 0 Then r = r & "<" & b & ">"
    pa = PF
    p = InStr(PF + 1, HTML, "<")
  Wend
  r = r & Mid(HTML, pa + 1)
  HTML2HTMLs = Replace(r, vbCrLf, "")
End Function


Public Function RTrimEx(Text As String, Optional Delimiters As String = vbTab + vbCrLf + " ") As String
  Dim p As Long
  p = Len(Text) + 1
  If p = 1 Then Exit Function
  Do While InStr(Delimiters, Mid(Text, p - 1, 1)) > 0
    p = p - 1
  Loop
  RTrimEx = Left(Text, p - 1)
End Function


Public Function LTrimEx(Text As String, Optional Delimiters As String = vbTab + vbCrLf + " ") As String
  Dim p As Long
  p = 1
  If Len(Text) = 0 Then Exit Function
  Do While InStr(Delimiters, Mid(Text, p, 1)) > 0
    p = p + 1
  Loop
  LTrimEx = Mid(Text, p)
End Function


Public Function LPad(Text As String, tam As Long, pad As String) As String
  LPad = Right(String(tam, pad) & Text, tam)
End Function


Public Function RPad(Text As String, tam As Long, pad As String) As String
  RPad = Left(Text & String(tam, pad), tam)
End Function


Public Function After(Text As String, Delimiter As String) As String
  Dim p As Long
  If Len(Text) = 0 Then Exit Function
  p = InStr(Text, Delimiter)
  If p > 0 Then
    After = Mid(Text, p + Len(Delimiter))
  End If
End Function


Public Function Before(Text As String, Delimiter As String) As String
  Dim p As Long
  p = InStr(Text, Delimiter)
  If p = 0 Then p = Len(Text) + 1
  Before = Left(Text, p - 1)
End Function


Public Function End_With(Text As String, Ends As String) As Boolean
  End_With = (Right(Text, Len(Ends)) = Ends)
End Function


Public Function XML_Escape(ByVal Text As String)
  Text = Replace(Text, "<", "&lt;")
  Text = Replace(Text, ">", "&gt;")
  Text = Replace(Text, "&", "&amp;")
  Text = Replace(Text, """", "&quot;")
  XML_Escape = Replace(Text, "'", "&#39;")
End Function


Public Function Module11(ByVal value As String, Optional MaxMult As Integer = 9) As Integer
  Dim Mult As Integer, Sum As Long, p As Integer, dv As Integer
  value = RTrim(LTrim(value))
  p = Len(value)
  Sum = 0
  Mult = 2
  While p > 0
    Sum = Sum + Val(Mid(value, p, 1)) * Mult
    Mult = Mult + 1
    If Mult > MaxMult Then Mult = 2
    p = p - 1
  Wend
  dv = 11 - (Sum Mod 11)
  If dv >= 10 Then dv = 0
  Module11 = dv
End Function


Public Function Checa_CPF(CPF As String) As Boolean
  ' Validade Brazilian CPF digits-check
  Dim dc$, c$, D1, D2, x%
  c = Replace(Replace(Replace(CPF, "/", ""), ".", ""), "-", "")
  If Len(c) <> 11 Then
    Checa_CPF = False
  Else
    dc = Right(c, 2)
    For x = 1 To 9
      D1 = D1 + Val(Mid(c, x, 1)) * (11 - x)
      D2 = D2 + Val(Mid(c, x, 1)) * (12 - x)
    Next x
    D1 = D1 Mod 11
    D1 = IIf(D1 < 2, 0, 11 - D1)
    D2 = (D2 + D1 * 2) Mod 11
    D2 = IIf(D2 < 2, 0, 11 - D2)
    Checa_CPF = (D1 & D2 = dc)
  End If
End Function


Public Function IP_Start(IP As String, ByVal BitsMask As Integer, Optional Format_ As Boolean = False) As String
  ' Return the initial IP (network) in the range of IP/BitsMask
  Dim i, m As Integer, x As Integer
  If IP = "" Then Exit Function
  i = Split(IP, ".", 4)
  If UBound(i) < 3 Then Exit Function
  For x = 0 To 3
    If BitsMask >= 8 Then
      m = 255
      BitsMask = BitsMask - 8
    Else
      m = 255 Xor (2 ^ (8 - BitsMask) - 1)
      BitsMask = 0
    End If
    i(x) = Val(i(x)) And m
    If Format_ Then i(x) = Format(i(x), "000")
  Next x
  IP_Start = Join(i, ".")
End Function


Public Function IP_End(IP As String, ByVal BitsMask As Integer, Optional Format_ As Boolean = False) As String
  ' Return the final IP (broadcast) in the range of IP/BitsMask
  Dim i, m As Integer, x As Integer
  If IP = "" Then Exit Function
  i = Split(IP, ".", 4)
  If UBound(i) < 3 Then Exit Function
  For x = 0 To 3
    If BitsMask >= 8 Then
      m = 255
      BitsMask = BitsMask - 8
    Else
      m = 255 Xor (2 ^ (8 - BitsMask) - 1)
      BitsMask = 0
    End If
    i(x) = (Val(i(x)) And m) Or (255 Xor m)
    If Format_ Then i(x) = Format(i(x), "000")
  Next x
  IP_End = Join(i, ".")
End Function


Public Function IP_Format(IP As String) As String
  Dim i, m As Integer, x As Integer
  If IP = "" Then
    IP_Format = "000.000.000.000"
    Exit Function
  End If
  i = Split(IP, ".", 4)
  If UBound(i) < 3 Then Exit Function
  For x = 0 To 3
    i(x) = Val(i(x))
    i(x) = Format(i(x), "000")
  Next x
  IP_Format = Join(i, ".")
End Function


Public Function IP_Mask(IP As String, ByVal BitsMask As Integer) As String
  Dim i, m As Integer, x As Integer
  If IP = "" Then Exit Function
  i = Split(IP, ".", 4)
  If UBound(i) < 3 Then Exit Function
  For x = 0 To 3
    If BitsMask >= 8 Then
      m = 255
      BitsMask = BitsMask - 8
    Else
      m = 255 Xor (2 ^ (8 - BitsMask) - 1)
      BitsMask = 0
    End If
    i(x) = m
  Next x
  IP_Mask = Join(i, ".")
End Function


Public Function IP_single(IP As String) As Double
  ' Returns IP as a number (Double, because VBA don't have unsigned long)
  Dim i, r As Single, x As Integer
  If IP = "" Then Exit Function
  i = Split(IP, ".", 4)
  If UBound(i) < 3 Then Exit Function
  r = 0
  For x = 0 To 3
    r = r + Val(i(x)) * (256 ^ (3 - x))
  Next x
  IP_single = r
End Function


Public Sub MkDirEx(ByVal path As String, Optional ErrorMsg As Boolean = True)
  ' Create enteri path Path if necessary
  Dim p As String, x As Long
  If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
  If ErrorMsg Then On Error GoTo er
  MkDir path
  Exit Sub
er:
  If Err = 76 Then
    x = InStrRev(path, "\")
    If x > 2 Then
      MkDirEx Left(path, x - 1)
      Resume
    Else
      MsgBox "Error in MkDirEx " & path & "."
    End If
  ElseIf Err <> 75 Then
    MsgBox error
  End If
End Sub


Public Function File_Path(FileName As String)
  Dim r As String, p As Long
  p = InStrRev(FileName, "\")
  If p = 0 Then p = 1
  File_Path = Left(FileName, p - 1)
End Function

Public Sub Create_Recursive_Path(path As String)
  ' cria o caminha criando todos os níveis necessários
  path = RTrimEx(path, "\")
  If Dir(path, vbDirectory) > "" Then Exit Sub
  Create_Recursive_Path File_Path(path)
  MkDir path
End Sub

Public Function File_Name(FileName As String)
  Dim r As String, p As Long
  p = InStrRev(FileName, "\")
  If p = 0 Then p = InStrRev(FileName, ":")
  File_Name = Mid(FileName, p + 1)
End Function


Public Sub File_Save(FileName As String, Text As String)
  On Error GoTo er
  Dim f As Integer
  f = FreeFile
  Open FileName For Output As #f
  Print #f, Text;
  Close f
  Exit Sub
er:
  If Err = 76 Then
    MkDirEx File_Path(FileName)
    Resume
  Else
    MsgBox error
  End If
End Sub


Public Function File_Load(FileName As String) As String
   Dim nSourceFile As Integer, sText As String
   nSourceFile = FreeFile
   Open FileName For Binary As #nSourceFile
   sText = Input$(LOF(nSourceFile), nSourceFile)
   Close
   File_Load = sText
End Function


Public Sub File_Delete(FileName As String)
  On Error Resume Next
  Kill FileName
End Sub


Public Function File_Exist(FileName As String) As Boolean
  On Error GoTo ErrorHandler
  Call FileLen(FileName)
  File_Exist = True
  Exit Function
ErrorHandler:
  File_Exist = False
End Function


Public Function DirEx(path As String, Optional Atrib As VbFileAttribute = vbNormal) As String
  Dim r As String, D As String
  D = Dir(path, Atrib)
  While D > ""
    r = r & D & vbCrLf
    D = Dir()
  Wend
  DirEx = r
End Function




'============= Numéricas / Gerais =============

Public Function MinEx(a, b, Optional c = Null)
  Dim r
  r = a
  If IsNull(a) Then r = b
  If IsNull(b) Then r = c
  If Not (IsNull(a)) And Not (IsNull(b)) Then
    If a < b Then r = a Else r = b
  End If
  If Not IsNull(c) Then
    If c < r Then r = c
  End If
  MinEx = r
End Function


Public Function MaxEx(a, b, Optional c = Null)
  Dim r
  r = a
  If IsNull(a) Then r = b
  If IsNull(b) Then r = c
  If Not (IsNull(a)) And Not (IsNull(b)) Then
    If a > b Then r = a Else r = b
  End If
  If Not IsNull(c) Then
    If c > r Then r = c
  End If
  MaxEx = r
End Function


Public Function Alternate(Optional Max As Integer = 2, Optional Incr As Integer = 1) As Integer
  Static x As Integer
  x = x + Incr
  If x >= Max Then x = 0
  Alternate = x
End Function



Public Function CNAB(Valor, largura As Integer, tipo As String, Optional Dec As Integer = 0) As String
  ' Brazilian format for Banks files
  Dim r As String, v As Double
  If tipo = "9" Then
    Select Case TypeName(Valor)
      Case "String"
        v = Val(Valor)
      Case "Date"
        If largura = 8 Then  ' DATA
          r = Format(Valor, "ddmmyyyy")
        ElseIf largura = 6 Then  ' HORA
          r = Format(Valor, "HhNnSs")
        Else
          MsgBox "Formato de data não suportado."
        End If
      Case Else
        v = Nz(Valor, 0)
    End Select
    If r = "" Then  ' não entra se é data
      If Dec > 0 Then v = v * (10 ^ Dec)
      r = Left(Format(v, String(largura, "0")), largura)
    End If
  ElseIf tipo = "D" Then
    If largura = 8 Then  ' DATA
      r = Format(Valor, "ddmmyyyy")
    ElseIf largura = 6 Then  ' HORA
      r = Format(Valor, "HhNnSs")
    Else
      MsgBox "Formato de data não suportado."
    End If
  ElseIf tipo = "X" Then
    r = UCase(Left(Remove_Accents(Nz(Valor, "")) & space(largura), largura))
  Else
    MsgBox "Tipo Errado."
  End If
  CNAB = r
End Function


Public Sub ODBC_Control()
  If Dir(Environ("SYSTEMROOT") & "\SysWOW64\odbcad32.exe") > "" Then
    Shell Environ("SYSTEMROOT") + "\SysWOW64\odbcad32.exe"
  Else
    Shell Environ("SYSTEMROOT") + "\System32\odbcad32.exe"
  End If
End Sub


Public Function UTF8_Decode(ByVal sStr As String)
  Dim l As Long, sUtf8 As String, iChar As Long, iChar2 As Integer
  If Left(sStr, 3) = "ï»¿" Then sStr = Mid(sStr, 4)
  For l = 1 To Len(sStr)
    iChar = Asc(Mid(sStr, l, 1))
    If iChar > 127 Then
      If Not iChar And 32 Then ' 2 chars
        iChar2 = Asc(Mid(sStr, l + 1, 1))
        sUtf8 = sUtf8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
        l = l + 1
      Else
        Dim iChar3 As Integer
        iChar2 = Asc(Mid(sStr, l + 1, 1))
        iChar3 = Asc(Mid(sStr, l + 2, 1))
        sUtf8 = sUtf8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
        l = l + 2
      End If
    Else
      sUtf8 = sUtf8 & Chr$(iChar)
    End If
  Next l
  UTF8_Decode = sUtf8
End Function


Public Function UTF8_Encode(ByVal sStr As String)
  Dim l As Long, lChar As Integer, sUtf8 As String
  For l = 1 To Len(sStr)
    lChar = AscW(Mid(sStr, l, 1))
    If lChar < 128 Then
      sUtf8 = sUtf8 + Mid(sStr, l, 1)
    ElseIf ((lChar > 127) And (lChar < 2048)) Then
      sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
      sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
    Else
      sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
      sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
      sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
    End If
  Next l
  UTF8_Encode = sUtf8
End Function


Public Function UTF8_Encode_Escaped(ByVal sStr As String)
  Dim l As Long, lChar As Integer, sUtf8 As String
  For l = 1 To Len(sStr)
    lChar = AscW(Mid(sStr, l, 1))
    If lChar < 128 Then
      sUtf8 = sUtf8 + Mid(sStr, l, 1)
    ElseIf ((lChar > 127) And (lChar < 2048)) Then
      sUtf8 = sUtf8 + "\x" + Hex(((lChar \ 64) Or 192))
      sUtf8 = sUtf8 + "\x" + Hex(((lChar And 63) Or 128))
    Else
      sUtf8 = sUtf8 + "\x" + Hex(((lChar \ 144) Or 234))
      sUtf8 = sUtf8 + "\x" + Hex((((lChar \ 64) And 63) Or 128))
      sUtf8 = sUtf8 + "\x" + Hex(((lChar And 63) Or 128))
    End If
  Next l
  UTF8_Encode_Escaped = sUtf8
End Function


Function Date_Between(Period_Ini As Date, Period_Fim As Date, Data_to_check_Ini As Date, Optional Data_to_check_Fim As Date = #1/1/1900#) As Boolean
  Dim r As Boolean
  r = False
  If Data_to_check_Fim < Data_to_check_Ini Then Data_to_check_Fim = Data_to_check_Ini
  If Period_Ini > Data_to_check_Fim Then
    r = False
  ElseIf Data_to_check_Ini < Period_Ini Then
    r = True
  ElseIf Data_to_check_Ini <= Period_Fim Then
    r = True
  End If
  Date_Between = r
End Function
