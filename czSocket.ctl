VERSION 5.00
Begin VB.UserControl czSocket 
   BackColor       =   &H80000018&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
   Begin VB.Timer tmrPoll 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   240
   End
   Begin VB.Label labLogo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cz"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   480
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "czSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' czSocket User Control v1.1 (c) 2026
' Single-file VB6 socket control with full networking support
' Inspired by VbAsyncSocket Project (wqweto@gmail.com)
'
' Features:
'   - TCP/UDP client and server
'   - TLS 1.0-1.3 via SSPI/Schannel (native Windows)
'   - HTTP/HTTPS requests with response parsing (chunked + Content-Length)
'   - WebSocket (WS/WSS) with RFC 6455 framing
'   - File upload (multipart/form-data) and download (save to disk)
'   - Chunked file sending with progress, speed, and ETA tracking
'   - Universal API: auto-detects protocol from URL scheme
'   - Zero external dependencies — uses only built-in Windows DLLs
'   - Minimum OS: Windows 7
'
' Required DLLs (all built into Windows):
'   ws2_32.dll   — WinSock TCP/UDP
'   secur32.dll  — SSPI/Schannel TLS
'   advapi32.dll — CryptoAPI (SHA-1, random bytes)
'   crypt32.dll  — Base64 encode/decode
'   kernel32.dll — Memory, string conversion, timing
'   user32.dll   — Window properties (server Accept pattern)
'
' This project is licensed under the terms of the MIT license
'
'=========================================================================
' PUBLIC API REFERENCE
'=========================================================================
'
' METHODS (Universal — auto-detect protocol):
'   Connect [RemoteHost], [RemotePort]
'       Connects to a remote host. If RemoteHost contains a URL scheme
'       (ws://, wss://, http://, https://), protocol is auto-detected.
'       Examples: Connect "example.com", 80
'                 Connect "wss://ws.postman-echo.com/raw"
'                 Connect "https://jsonplaceholder.typicode.com/posts/1"
'
'   SendData data
'       Sends string or byte array. If WebSocket is open, automatically
'       wraps data in WS text frame (string) or binary frame (byte array).
'
'   SendFile FilePath, [ChunkSize=8192]
'       Sends a file in chunks from disk. Auto-detects mode:
'       - WebSocket: wraps each chunk in WS binary frame
'       - Raw TCP/TLS: sends raw bytes
'       Fires Progress with speed and ETA.
'
'   Download Url, SavePath, [ExtraHeaders]
'       Downloads a file via HTTP/HTTPS and saves to disk.
'       Fires Progress during download, Response on complete.
'
'   Upload Url, FilePath, [FieldName="file"], [ExtraHeaders]
'       Uploads a file via HTTP POST multipart/form-data.
'       Fires Response on server reply.
'
'   Request Method, Url, [Body], [ContentType], [ExtraHeaders]
'       Sends an HTTP/HTTPS request. Auto-connects and auto-TLS.
'       Fires Response when complete.
'
'   Disconnect
'       Closes the connection. If WebSocket is open, sends a graceful
'       close frame before disconnecting.
'
'   Listen [CertFile], [Password], ...
'       Starts listening for incoming connections (server mode).
'
'   Accept requestID
'       Accepts an incoming connection from ConnectionRequest event.
'
'   Bind [LocalPort], [LocalIP]
'       Binds to a local port before Connect or Listen.
'
'   GetData data, [type], [maxLen]
'       Retrieves received data as string or byte array.
'
'   PeekData data, [type], [maxLen]
'       Peeks at received data without removing it from buffer.
'
'   Ping [Data]
'       Sends a WebSocket ping frame (only when connected via ws/wss).
'
' EVENTS:
'   Connect()
'       Fired when connection is established (TCP, TLS, or WS handshake).
'
'   Disconnected(Code, Reason)
'       Fired when connection is closed (TCP or WebSocket).
'       TCP: Code=0, Reason="". WS: Code=close code, Reason=close reason.
'
'   DataArrival(bytesTotal As Long)
'       Fired when raw data is available (TCP/UDP mode only).
'
'   Response(Status, ContentType, Body, Headers)
'       Fired when HTTP response is fully received.
'       In Download mode, Body contains the saved file path.
'
'   Progress(BytesSent, BytesTotal, BytesPerSec, SecondsRemaining)
'       Fired during file transfer with speed and ETA.
'       Works for Download, Upload, and SendFile operations.
'
'   Receive(Data, IsBinary)
'       Fired when a WebSocket message is received.
'
'   ConnectionRequest(requestID As Long)
'       Fired on server when a client connects (call Accept).
'
'   SendProgress(bytesSent, bytesRemaining)
'       Fired during buffer flush with send progress.
'
'   SendComplete()
'       Fired when all queued/chunked data has been sent.
'
'   Error(Number, Description, Scode, Source, ...)
'       Fired on socket or TLS error.
'
' PROPERTIES:
'   RemoteHost As String     — Remote host name or IP
'   RemotePort As Long       — Remote port number
'   RemoteHostIP As String   — (Read-only) Resolved remote IP
'   LocalPort As Long        — Local port for Bind/Listen
'   LocalHostName As String  — (Read-only) Local machine name
'   LocalIP As String        — (Read-only) Local IP address
'   Protocol As UcsProtocolConstants — sckTCPProtocol / sckUDPProtocol / sckTLSProtocol
'   State As UcsStateConstants       — (Read-only) Current socket state
'   SocketHandle As Long     — (Read-only) Raw WinSock handle
'   Timeout As Long          — Connection timeout in seconds
'   StatusCode As Long       — (Read-only) Last HTTP response status code
'   ReadyState As Long          — (Read-only) WebSocket state (0=None,1=Connecting,2=Open,3=Closing)
'   SockOpt(OptionName, [Level]) — Get/Set raw socket options
'
' ENUMS:
'   UcsProtocolConstants: sckTCPProtocol=0, sckUDPProtocol=1, sckTLSProtocol=2
'   UcsStateConstants:    sckClosed=0 .. sckConnected=7 .. sckClosing=8 .. sckError=9
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "czSocket"

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
    sckTLSProtocol = 2
End Enum

Public Enum UcsStateConstants
    sckClosed = 0
    sckOpen = 1
    sckListening = 2
    sckConnectionPending = 3
    sckResolvingHost = 4
    sckHostResolved = 5
    sckConnecting = 6
    sckConnected = 7
    sckClosing = 8
    sckError = 9
End Enum

Public Enum UcsErrorConstants
    sckInvalidPropertyValue = 380
    sckGetNotSupported = 394
    sckSetNotSupported = 383
    sckOutOfMemory = 7
    sckBadState = 40006
    sckInvalidArg = 40014
    sckSuccess = 40017
    sckUnsupported = 40018
    sckInvalidOp = 40020
    sckOutOfRange = 40021
    sckWrongProtocol = 40026
    sckOpCanceled = 10004
    sckInvalidArgument = 10014
    sckWouldBlock = 10035
    sckInProgress = 10036
    sckAlreadyComplete = 10037
    sckNotSocket = 10038
    sckMsgTooBig = 10040
    sckPortNotSupported = 10043
    sckAddressInUse = 10048
    sckAddressNotAvailable = 10049
    sckNetworkSubsystemFailed = 10050
    sckNetworkUnreachable = 10051
    sckNetReset = 10052
    sckConnectAborted = 10053
    sckConnectionReset = 10054
    sckNoBufferSpace = 10055
    sckAlreadyConnected = 10056
    sckNotConnected = 10057
    sckSocketShutdown = 10058
    sckTimedout = 10060
    sckConnectionRefused = 10061
    sckNotInitialized = 10093
    sckHostNotFound = 11001
    sckHostNotFoundTryAgain = 11002
    sckNonRecoverableError = 11003
    sckNoData = 11004
End Enum

Public Enum UcsSckLocalFeaturesEnum '--- bitmask
    ucsSckSupportTls10 = 2 ^ 0
    ucsSckSupportTls11 = 2 ^ 1
    ucsSckSupportTls12 = 2 ^ 2
    ucsSckSupportTls13 = 2 ^ 3
    ucsSckIgnoreServerCertificateErrors = 2 ^ 4
    ucsSckIgnoreServerCertificateRevocation = 2 ^ 5
    ucsSckSupportAll = ucsSckSupportTls10 Or ucsSckSupportTls11 Or ucsSckSupportTls12 Or ucsSckSupportTls13
End Enum

Public Enum UcsSckOptionLevelEnum
    ucsSckIP = 0
    ucsSckICMP = 1
    ucsSckIGMP = 2
    ucsSckTCP = 6
    ucsSckUDP = 17
    ucsSckSocket = &HFFFF&
End Enum

Public Enum UcsSckOptionNameEnum
    ucsSckDebug = &H1
    ucsSckAcceptConnection = &H2
    ucsSckReuseAddress = &H4
    ucsSckKeepAlive = &H8
    ucsSckDontRoute = &H10
    ucsSckBroadcast = &H20
    ucsSckUseLoopback = &H40
    ucsSckLinger = &H80
    ucsSckOutOfBandInline = &H100
    ucsSckDontLinger = Not ucsSckLinger
    ucsSckExclusiveAddressUse = Not ucsSckReuseAddress
    ucsSckSendBuffer = &H1001
    ucsSckReceiveBuffer = &H1002
    ucsSckSendLowWater = &H1003
    ucsSckReceiveLowWater = &H1004
    ucsSckSendTimeout = &H1005
    ucsSckReceiveTimeout = &H1006
    ucsSckError = &H1007
    ucsSckType = &H1008
    ucsSckMaxMsgSize = &H2003
    ucsSckProtocolInfo = &H2004
    ucsSckReuseUnicastPort = &H3007
    ucsSckMaxConnections = &H7FFFFFFF
    '-- IP
    ucsSckIPOptions = 1
    ucsSckHeaderIncluded = 2
    ucsSckTypeOfService = 3
    ucsSckIpTimeToLive = 4
    ucsSckMulticastInterface = 9
    ucsSckMulticastTimeToLive = 10
    ucsSckMulticastLoopback = 11
    ucsSckAddMembership = 12
    ucsSckDropMembership = 13
    ucsSckDontFragment = 14
    '-- TCP
    ucsSckNoDelay = 1
    ucsSckExpedited = 2
    '-- UDP
    ucsSckNoChecksum = 1
    ucsSckChecksumCoverage = 20
End Enum

'=========================================================================
' Events
'=========================================================================

Event Connect()
Event Disconnected(ByVal Code As Long, ByVal Reason As String)
Event ConnectionRequest(ByVal requestID As Long)
Event DataArrival(ByVal bytesTotal As Long)
Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Event SendComplete()
Event Error(ByVal Number As Long, Description As String, ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Event Response(ByVal Status As Long, ByVal ContentType As String, Body As String, Headers As String)
Event Receive(ByVal Data As String, ByVal IsBinary As Boolean)
Event Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)

'=========================================================================
' API declarations - WinSock
'=========================================================================

Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32" () As Long
Private Declare Function WSACreateEvent Lib "ws2_32" () As Long
Private Declare Function WSACloseEvent Lib "ws2_32" (ByVal hEvent As Long) As Long
Private Declare Function WSAEventSelect Lib "ws2_32" (ByVal s As Long, ByVal hEventObject As Long, ByVal lNetworkEvents As Long) As Long
Private Declare Function WSAWaitForMultipleEvents Lib "ws2_32" (ByVal cEvents As Long, lphEvents As Long, ByVal fWaitAll As Long, ByVal dwTimeout As Long, ByVal fAlertable As Long) As Long
Private Declare Function WSAEnumNetworkEvents Lib "ws2_32" (ByVal s As Long, ByVal hEventObject As Long, lpNetworkEvents As Any) As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, Name As Any, ByVal namelen As Long) As Long
Private Declare Function ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, Name As Any, ByVal namelen As Long) As Long
Private Declare Function ws_listen Lib "ws2_32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function ws_accept Lib "ws2_32" Alias "accept" (ByVal s As Long, addr As Any, addrlen As Long) As Long
Private Declare Function ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_sendto Lib "ws2_32" Alias "sendto" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long, toAddr As Any, ByVal tolen As Long) As Long
Private Declare Function ws_recvfrom Lib "ws2_32" Alias "recvfrom" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long, fromAddr As Any, fromlen As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_getpeername Lib "ws2_32" Alias "getpeername" (ByVal s As Long, Name As Any, namelen As Long) As Long
Private Declare Function ws_getsockname Lib "ws2_32" Alias "getsockname" (ByVal s As Long, Name As Any, namelen As Long) As Long
Private Declare Function ws_gethostbyname Lib "ws2_32" Alias "gethostbyname" (ByVal Name As String) As Long
Private Declare Function ws_gethostname Lib "ws2_32" Alias "gethostname" (ByVal Name As String, ByVal namelen As Long) As Long
Private Declare Function ws_inet_addr Lib "ws2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function ws_inet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function ws_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function ws_shutdown Lib "ws2_32" Alias "shutdown" (ByVal s As Long, ByVal How As Long) As Long

'=========================================================================
' API declarations - SSPI / Schannel
'=========================================================================

Private Declare Function AcquireCredentialsHandle Lib "secur32" Alias "AcquireCredentialsHandleA" ( _
    ByVal pszPrincipal As Long, ByVal pszPackage As String, ByVal fCredentialUse As Long, _
    ByVal pvLogonId As Long, pAuthData As Any, ByVal pGetKeyFn As Long, _
    ByVal pvGetKeyArgument As Long, phCredential As Any, ptsExpiry As Currency) As Long
Private Declare Function InitializeSecurityContext Lib "secur32" Alias "InitializeSecurityContextA" ( _
    phCredential As Any, phContext As Any, ByVal pszTargetName As String, _
    ByVal fContextReq As Long, ByVal Reserved1 As Long, ByVal TargetDataRep As Long, _
    pInput As Any, ByVal Reserved2 As Long, phNewContext As Any, _
    pOutput As Any, pfContextAttr As Long, ptsExpiry As Currency) As Long
Private Declare Function AcceptSecurityContext Lib "secur32" ( _
    phCredential As Any, phContext As Any, pInput As Any, _
    ByVal fContextReq As Long, ByVal TargetDataRep As Long, _
    phNewContext As Any, pOutput As Any, pfContextAttr As Long, _
    ptsExpiry As Currency) As Long
Private Declare Function EncryptMessage Lib "secur32" ( _
    phContext As Any, ByVal fQOP As Long, pMessage As Any, ByVal MessageSeqNo As Long) As Long
Private Declare Function DecryptMessage Lib "secur32" ( _
    phContext As Any, pMessage As Any, ByVal MessageSeqNo As Long, pfQOP As Long) As Long
Private Declare Function QueryContextAttributes Lib "secur32" Alias "QueryContextAttributesA" ( _
    phContext As Any, ByVal ulAttribute As Long, pBuffer As Any) As Long
Private Declare Function FreeContextBuffer Lib "secur32" (ByVal pvContextBuffer As Long) As Long
Private Declare Function DeleteSecurityContext Lib "secur32" (phContext As Any) As Long
Private Declare Function FreeCredentialsHandle Lib "secur32" (phCredential As Any) As Long
Private Declare Function ApplyControlToken Lib "secur32" (phContext As Any, pInput As Any) As Long

'=========================================================================
' API declarations - Kernel / User
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Args As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpStr As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetPropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetPropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemovePropA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'=========================================================================
' API declarations - CryptoAPI (SHA-1 for WebSocket, Base64)
'=========================================================================

Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal algId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Any) As Long
Private Declare Function CryptBinaryToStringA Lib "crypt32" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long

'=========================================================================
' Private Types
'=========================================================================

Private Type SOCKADDR_IN
    sin_family      As Integer
    sin_port        As Integer
    sin_addr        As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type SecHandle
    dwLower         As Long
    dwUpper         As Long
End Type

Private Type SecBuffer
    cbBuffer        As Long
    BufferType      As Long
    pvBuffer        As Long
End Type

Private Type SecBufferDesc
    ulVersion       As Long
    cBuffers        As Long
    pBuffers        As Long
End Type

Private Type SCHANNEL_CRED
    dwVersion       As Long
    cCreds          As Long
    paCred          As Long
    hRootStore      As Long
    cMappers        As Long
    aphMappers      As Long
    cSupportedAlgs  As Long
    palgSupportedAlgs As Long
    grbitEnabledProtocols As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan As Long
    dwFlags         As Long
    dwCredFormat    As Long
End Type

Private Type SecPkgContext_StreamSizes
    cbHeader        As Long
    cbTrailer       As Long
    cbMaximumMessage As Long
    cBuffers        As Long
    cbBlockSize     As Long
End Type

Private Type WSANETWORKEVENTS_TYPE
    lNetworkEvents  As Long
    iErrorCode(0 To 9) As Long
End Type

'=========================================================================
' Constants
'=========================================================================

'--- WinSock
Private Const INVALID_SOCKET            As Long = -1
Private Const SOCKET_ERROR              As Long = -1
Private Const AF_INET                   As Long = 2
Private Const SOCK_STREAM               As Long = 1
Private Const SOCK_DGRAM                As Long = 2
Private Const IPPROTO_TCP               As Long = 6
Private Const IPPROTO_UDP               As Long = 17
Private Const SD_SEND                   As Long = 1
Private Const FIONREAD                  As Long = &H4004667F
Private Const SOL_SOCKET                As Long = &HFFFF&
Private Const INADDR_NONE               As Long = -1
Private Const INADDR_ANY                As Long = 0

'--- WSA errors
Private Const WSAEWOULDBLOCK            As Long = 10035
Private Const WSAENOTCONN               As Long = 10057
Private Const WSAEINPROGRESS            As Long = 10036

'--- WSA event flags
Private Const FD_READ                   As Long = &H1&
Private Const FD_WRITE                  As Long = &H2&
Private Const FD_OOB                    As Long = &H4&
Private Const FD_ACCEPT                 As Long = &H8&
Private Const FD_CONNECT                As Long = &H10&
Private Const FD_CLOSE                  As Long = &H20&
Private Const FD_READ_BIT               As Long = 0
Private Const FD_WRITE_BIT              As Long = 1
Private Const FD_ACCEPT_BIT             As Long = 3
Private Const FD_CONNECT_BIT            As Long = 4
Private Const FD_CLOSE_BIT              As Long = 5

'--- WSA wait
Private Const WSA_WAIT_EVENT_0          As Long = 0

'--- SSPI / Schannel
Private Const UNISP_NAME               As String = "Microsoft Unified Security Protocol Provider"
Private Const SECPKG_CRED_OUTBOUND      As Long = 2
Private Const SECPKG_CRED_INBOUND       As Long = 1
Private Const SCHANNEL_CRED_VERSION     As Long = 4
Private Const SCH_CRED_NO_DEFAULT_CREDS As Long = &H10&
Private Const SCH_CRED_AUTO_CRED_VALIDATION As Long = &H20&
Private Const SCH_CRED_MANUAL_CRED_VALIDATION As Long = &H8&
Private Const SCH_CRED_IGNORE_NO_REVOCATION_CHECK As Long = &H800&
Private Const SCH_CRED_IGNORE_REVOCATION_OFFLINE As Long = &H1000&
Private Const SCH_CRED_NO_SERVERNAME_CHECK As Long = &H4&

Private Const ISC_REQ_SEQUENCE_DETECT   As Long = &H8&
Private Const ISC_REQ_REPLAY_DETECT     As Long = &H4&
Private Const ISC_REQ_CONFIDENTIALITY   As Long = &H10&
Private Const ISC_REQ_EXTENDED_ERROR    As Long = &H4000&
Private Const ISC_REQ_ALLOCATE_MEMORY   As Long = &H100&
Private Const ISC_REQ_STREAM            As Long = &H8000&
Private Const ISC_REQ_MANUAL_CRED_VALIDATION As Long = &H80000

Private Const ASC_REQ_SEQUENCE_DETECT   As Long = &H8&
Private Const ASC_REQ_REPLAY_DETECT     As Long = &H4&
Private Const ASC_REQ_CONFIDENTIALITY   As Long = &H10&
Private Const ASC_REQ_EXTENDED_ERROR    As Long = &H4000&
Private Const ASC_REQ_ALLOCATE_MEMORY   As Long = &H100&
Private Const ASC_REQ_STREAM            As Long = &H10000

Private Const SEC_E_OK                  As Long = 0
Private Const SEC_I_CONTINUE_NEEDED     As Long = &H90312
Private Const SEC_E_INCOMPLETE_MESSAGE  As Long = &H80090318
Private Const SEC_I_INCOMPLETE_CREDENTIALS As Long = &H90320
Private Const SEC_I_RENEGOTIATE         As Long = &H90321
Private Const SEC_I_CONTEXT_EXPIRED     As Long = &H90317

Private Const SECBUFFER_VERSION         As Long = 0
Private Const SECBUFFER_EMPTY           As Long = 0
Private Const SECBUFFER_DATA            As Long = 1
Private Const SECBUFFER_TOKEN           As Long = 2
Private Const SECBUFFER_EXTRA           As Long = 5
Private Const SECBUFFER_STREAM_TRAILER  As Long = 6
Private Const SECBUFFER_STREAM_HEADER   As Long = 7
Private Const SECBUFFER_ALERT           As Long = 17

Private Const SECPKG_ATTR_STREAM_SIZES  As Long = 4

Private Const SCHANNEL_SHUTDOWN_TOKEN   As Long = 1

'--- Schannel protocol flags
Private Const SP_PROT_TLS1_0            As Long = &HC0&
Private Const SP_PROT_TLS1_1            As Long = &H300&
Private Const SP_PROT_TLS1_2            As Long = &HC00&
Private Const SP_PROT_TLS1_3            As Long = &H3000&

'--- FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

'--- Defaults
Private Const DEF_LOCALPORT             As Long = 0
Private Const DEF_PROTOCOL              As Long = 0
Private Const DEF_REMOTEHOST            As String = vbNullString
Private Const DEF_REMOTEPORT            As Long = 0
Private Const DEF_TIMEOUT               As Long = 5000
Private Const STR_LOGO                  As String = "cz"

'--- TLS state
Private Const TLS_NONE                  As Long = 0
Private Const TLS_HANDSHAKE             As Long = 1
Private Const TLS_CONNECTED             As Long = 2
Private Const TLS_SHUTDOWN              As Long = 3

'--- Recv buffer size
Private Const RECV_BUFFER_SIZE          As Long = 16384

'--- Global property name for Accept pattern
Private Const PROP_REQUEST_SOCKET       As String = "czSocket_RequestSocket"
Private Const PROP_REQUEST_PROTOCOL     As String = "czSocket_RequestProtocol"

'--- WebSocket opcodes
Private Const WS_OPCODE_CONTINUATION    As Long = 0
Private Const WS_OPCODE_TEXT            As Long = 1
Private Const WS_OPCODE_BINARY          As Long = 2
Private Const WS_OPCODE_CLOSE           As Long = 8
Private Const WS_OPCODE_PING            As Long = 9
Private Const WS_OPCODE_PONG            As Long = 10

'--- WebSocket state
Private Const WS_NONE                   As Long = 0
Private Const WS_CONNECTING             As Long = 1
Private Const WS_OPEN                   As Long = 2
Private Const WS_CLOSING                As Long = 3

'--- HTTP parser state
Private Const HTTP_NONE                 As Long = 0
Private Const HTTP_WAITING              As Long = 1

'--- CryptoAPI
Private Const PROV_RSA_FULL             As Long = 1
Private Const CRYPT_VERIFYCONTEXT       As Long = &HF0000000
Private Const CALG_SHA1                 As Long = &H8004&
Private Const HP_HASHVAL                As Long = 2
Private Const CRYPT_STRING_BASE64       As Long = 1
Private Const CRYPT_STRING_NOCRLF       As Long = &H40000000

'--- WebSocket GUID
Private Const WS_MAGIC_GUID             As String = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"

'=========================================================================
' Member variables
'=========================================================================

'--- Socket state
Private m_hSocket               As Long
Private m_eState                As UcsStateConstants
Private m_eProtocol             As UcsProtocolConstants
Private m_sRemoteHost           As String
Private m_lRemotePort           As Long
Private m_lLocalPort            As Long
Private m_lTimeout              As Long
Private m_lLocalFeatures        As UcsSckLocalFeaturesEnum

'--- Async engine
Private m_hEvent                As Long
Private m_bListening            As Boolean
Private m_bConnecting           As Boolean

'--- TLS state
Private m_lTlsState             As Long
Private m_hCred(0 To 1)         As Long     '--- SecHandle (dwLower, dwUpper)
Private m_hCtx(0 To 1)          As Long     '--- SecHandle (dwLower, dwUpper)
Private m_bHasCred              As Boolean
Private m_bHasCtx               As Boolean
Private m_uStreamSizes          As SecPkgContext_StreamSizes
Private m_baTlsPending()        As Byte     '--- Pending encrypted data for decrypt/handshake

'--- Data buffers
Private m_baRecvBuffer()        As Byte     '--- Decrypted/received data available to user
Private m_baSendBuffer()        As Byte     '--- Pending data to send (already encrypted if TLS)
Private m_lSendPos              As Long

'--- Accepted socket tracking
Private m_hAcceptedSocket       As Long     '--- Socket handle from Accept
Private m_bWsaInitialized       As Boolean

'--- HTTP state
Private m_lHttpState            As Long
Private m_sHttpRawResponse      As String
Private m_lHttpStatus           As Long
Private m_sHttpHeaders          As String
Private m_sHttpContentType      As String
Private m_lHttpContentLength    As Long
Private m_bHttpChunked          As Boolean
Private m_sHttpBody             As String

'--- WebSocket state
Private m_lWsState              As Long
Private m_sWsKey                As String
Private m_baWsBuffer()          As Byte
Private m_sWsHost               As String
Private m_sWsPath               As String

'--- File transfer state
Private m_sDownloadPath         As String   '--- Path to save downloaded file
Private m_bDownloadMode         As Boolean  '--- True when HttpDownloadFile active
Private m_sSendFilePath         As String   '--- File being sent
Private m_lSendFileChunk        As Long     '--- Chunk size for file sending
Private m_lSendFileOffset       As Long     '--- Current offset in file
Private m_lSendFileSize         As Long     '--- Total file size
Private m_bSendingFile          As Boolean  '--- True when SendFile active
Private m_dTransferStart        As Double   '--- Timer value when transfer started

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties - Design-time
'=========================================================================

Property Get LocalPort() As Long
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        Dim uAddr As SOCKADDR_IN
        Dim lLen As Long
        lLen = LenB(uAddr)
        If ws_getsockname(m_hSocket, uAddr, lLen) = 0 Then
            LocalPort = ws_ntohs(uAddr.sin_port) And &HFFFF&
            Exit Property
        End If
    End If
    LocalPort = m_lLocalPort
End Property

Property Let LocalPort(ByVal lValue As Long)
    If m_lLocalPort <> lValue Then
        Disconnect
        m_lLocalPort = lValue
        PropertyChanged
    End If
End Property

Property Get Protocol() As UcsProtocolConstants
    Protocol = m_eProtocol
End Property

Property Let Protocol(ByVal eValue As UcsProtocolConstants)
    If m_eProtocol <> eValue Then
        Disconnect
        m_eProtocol = eValue
        PropertyChanged
    End If
End Property

Property Get RemoteHost() As String
    RemoteHost = m_sRemoteHost
End Property

Property Let RemoteHost(sValue As String)
    If m_sRemoteHost <> sValue Then
        m_sRemoteHost = sValue
        m_baSendBuffer = vbNullString
        PropertyChanged
    End If
End Property

Property Get RemotePort() As Long
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        Dim uAddr As SOCKADDR_IN
        Dim lLen As Long
        lLen = LenB(uAddr)
        If ws_getpeername(m_hSocket, uAddr, lLen) = 0 Then
            RemotePort = ws_ntohs(uAddr.sin_port) And &HFFFF&
            Exit Property
        End If
    End If
    RemotePort = m_lRemotePort
End Property

Property Let RemotePort(ByVal lValue As Long)
    If m_lRemotePort <> lValue Then
        m_lRemotePort = lValue
        m_baSendBuffer = vbNullString
        PropertyChanged
    End If
End Property

Property Get Timeout() As Long
    Timeout = m_lTimeout
End Property

Property Let Timeout(ByVal lValue As Long)
    If m_lTimeout <> lValue Then
        m_lTimeout = lValue
        PropertyChanged
    End If
End Property

'=========================================================================
' Properties - Run-time
'=========================================================================

Property Get SocketHandle() As Long
    SocketHandle = m_hSocket
End Property

Property Get State() As UcsStateConstants
    State = m_eState
End Property

Property Get LocalHostName() As String
    Dim sBuf As String
    sBuf = String$(256, 0)
    If ws_gethostname(sBuf, 256) = 0 Then
        LocalHostName = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    End If
End Property

Property Get LocalIP() As String
    Dim sHost As String
    sHost = LocalHostName
    If LenB(sHost) <> 0 Then
        Dim pHost As Long
        pHost = ws_gethostbyname(sHost)
        If pHost <> 0 Then
            LocalIP = pvGetHostIP(pHost)
        End If
    End If
End Property

Property Get RemoteHostIP() As String
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        Dim uAddr As SOCKADDR_IN
        Dim lLen As Long
        lLen = LenB(uAddr)
        If ws_getpeername(m_hSocket, uAddr, lLen) = 0 Then
            RemoteHostIP = pvInetNtoa(uAddr.sin_addr)
        End If
    End If
End Property

Property Get SockOpt(ByVal OptionName As UcsSckOptionNameEnum, Optional ByVal Level As UcsSckOptionLevelEnum = ucsSckSocket) As Long
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        Dim lValue As Long
        Dim lLen As Long
        lLen = 4
        ws_getsockopt m_hSocket, Level, OptionName, lValue, lLen
        SockOpt = lValue
    End If
End Property

Property Let SockOpt(ByVal OptionName As UcsSckOptionNameEnum, Optional ByVal Level As UcsSckOptionLevelEnum = ucsSckSocket, ByVal Value As Long)
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        ws_setsockopt m_hSocket, Level, OptionName, Value, 4
    End If
End Property

Property Get StatusCode() As Long
    StatusCode = m_lHttpStatus
End Property

Property Get ReadyState() As Long
    ReadyState = m_lWsState
End Property

'=========================================================================
' Public Methods
'=========================================================================

Public Sub Disconnect()
    On Error GoTo EH
    '--- WebSocket graceful close if open
    If m_lWsState = WS_OPEN Then
        m_lWsState = WS_CLOSING
        Dim baEmpty() As Byte
        baEmpty = vbNullString
        Dim baCode(0 To 1) As Byte
        baCode(0) = 3: baCode(1) = 232  '--- 1000 = normal close
        Dim baCloseFrame() As Byte
        pvWsBuildFrame WS_OPCODE_CLOSE, baCode, baCloseFrame
        If m_lTlsState = TLS_CONNECTED Then
            Dim baEnc() As Byte
            pvTlsEncryptData baCloseFrame, baEnc
            pvRawSend baEnc
        Else
            pvRawSend baCloseFrame
        End If
    End If
    m_lWsState = WS_NONE
    m_lHttpState = HTTP_NONE
    m_bDownloadMode = False
    m_bSendingFile = False
    '--- Clear HTTP/WS response state (prevents stale data contamination)
    m_sHttpRawResponse = vbNullString
    m_lHttpStatus = 0
    m_sHttpBody = vbNullString
    m_baWsBuffer = vbNullString
    If m_eState <> sckClosed Then
        pvState = sckClosing
        '--- TLS shutdown
        If m_lTlsState = TLS_CONNECTED Then
            pvTlsShutdown
        End If
        '--- Close socket
        If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
            ws_shutdown m_hSocket, SD_SEND
            ws_closesocket m_hSocket
            m_hSocket = INVALID_SOCKET
        End If
        '--- Close WSA event
        If m_hEvent <> 0 Then
            WSACloseEvent m_hEvent
            m_hEvent = 0
        End If
        '--- TLS cleanup
        pvTlsCleanup
        '--- Stop polling
        tmrPoll.Enabled = False
        '--- Reset state
        m_bListening = False
        m_bConnecting = False
        pvState = sckClosed
    End If
    Exit Sub
EH:
    PrintError "Disconnect"
End Sub

Public Sub Bind(Optional ByVal LocalPort As Long, Optional LocalIP As String)
    On Error GoTo EH
    Disconnect
    If LocalPort <> 0 Then
        m_lLocalPort = LocalPort
    End If
    '--- Create socket
    If Not pvCreateSocket() Then GoTo QH
    '--- Bind
    Dim uAddr As SOCKADDR_IN
    uAddr.sin_family = AF_INET
    uAddr.sin_port = ws_htons(m_lLocalPort)
    If LenB(LocalIP) <> 0 Then
        uAddr.sin_addr = ws_inet_addr(LocalIP)
    Else
        uAddr.sin_addr = INADDR_ANY
    End If
    If ws_bind(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
        pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
        GoTo QH
    End If
    pvState = sckOpen
QH:
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Connect(Optional RemoteHost As String, Optional ByVal RemotePort As Long, Optional ByVal LocalFeatures As UcsSckLocalFeaturesEnum)
    On Error GoTo EH
    '--- Auto-detect URL scheme
    If InStr(RemoteHost, "://") > 0 Then
        Dim sScheme As String, sHost As String, lPort As Long, sPath As String
        pvParseUrl RemoteHost, sScheme, sHost, lPort, sPath
        Select Case LCase$(sScheme)
        Case "ws", "wss"
            '--- WebSocket mode
            pvWsConnect RemoteHost
            Exit Sub
        Case "https"
            m_eProtocol = sckTLSProtocol
            RemoteHost = sHost
            If RemotePort = 0 Then RemotePort = lPort
        Case "http"
            m_eProtocol = sckTCPProtocol
            RemoteHost = sHost
            If RemotePort = 0 Then RemotePort = lPort
        End Select
    End If
    Disconnect
    If LenB(RemoteHost) <> 0 Then
        m_sRemoteHost = RemoteHost
    End If
    If RemotePort <> 0 Then
        m_lRemotePort = RemotePort
    End If
    m_lLocalFeatures = LocalFeatures
    '--- Resolve host
    pvState = sckResolvingHost
    Dim lAddr As Long
    lAddr = pvResolveHost(m_sRemoteHost)
    If lAddr = INADDR_NONE Then
        pvSetError LastDllError:=sckHostNotFound, RaiseError:=True
        GoTo QH
    End If
    pvState = sckHostResolved
    '--- Create socket
    If Not pvCreateSocket() Then GoTo QH
    '--- Bind if local port specified
    If m_lLocalPort <> 0 Then
        Dim uBind As SOCKADDR_IN
        uBind.sin_family = AF_INET
        uBind.sin_port = ws_htons(m_lLocalPort)
        uBind.sin_addr = INADDR_ANY
        ws_bind m_hSocket, uBind, LenB(uBind)
    End If
    '--- Setup async events
    If Not pvSetupAsyncEvents(FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE) Then GoTo QH
    '--- Connect
    pvState = sckConnecting
    m_bConnecting = True
    Dim uAddr As SOCKADDR_IN
    uAddr.sin_family = AF_INET
    uAddr.sin_addr = lAddr
    uAddr.sin_port = ws_htons(m_lRemotePort)
    Dim lResult As Long
    lResult = ws_connect(m_hSocket, uAddr, LenB(uAddr))
    If lResult = SOCKET_ERROR Then
        Dim lErr As Long
        lErr = WSAGetLastError()
        If lErr <> WSAEWOULDBLOCK Then
            pvSetError LastDllError:=lErr, RaiseError:=True
            GoTo QH
        End If
    End If
    '--- For UDP, consider connected immediately
    If m_eProtocol = sckUDPProtocol Then
        m_bConnecting = False
        pvState = sckConnected
        RaiseEvent Connect
    End If
QH:
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Listen( _
            Optional CertFile As String, _
            Optional Password As String, _
            Optional CertSubject As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Collection, _
            Optional AlpnProtocols As String, _
            Optional ByVal LocalFeatures As UcsSckLocalFeaturesEnum)
    On Error GoTo EH
    '--- Create socket if not already bound
    If m_hSocket = INVALID_SOCKET Or m_hSocket = 0 Then
        If Not pvCreateSocket() Then GoTo QH
        '--- Bind
        Dim uAddr As SOCKADDR_IN
        uAddr.sin_family = AF_INET
        uAddr.sin_port = ws_htons(m_lLocalPort)
        uAddr.sin_addr = INADDR_ANY
        If ws_bind(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
            pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
            GoTo QH
        End If
    End If
    '--- Listen
    If m_eProtocol <> sckUDPProtocol Then
        If ws_listen(m_hSocket, 5) = SOCKET_ERROR Then
            pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
            GoTo QH
        End If
    End If
    '--- Setup async events
    If Not pvSetupAsyncEvents(FD_ACCEPT Or FD_READ Or FD_CLOSE) Then GoTo QH
    m_bListening = True
    pvState = sckListening
QH:
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Accept(ByVal requestID As Long)
    On Error GoTo EH
    Disconnect
    '--- Check global request socket first
    Dim hGlobalSocket As Long
    hGlobalSocket = GetPropA(GetDesktopWindow(), PROP_REQUEST_SOCKET)
    If hGlobalSocket <> 0 And hGlobalSocket = requestID Then
        '--- Take ownership of the accepted socket
        m_hSocket = hGlobalSocket
        m_eProtocol = GetPropA(GetDesktopWindow(), PROP_REQUEST_PROTOCOL)
    Else
        '--- Use requestID directly as socket handle
        m_hSocket = requestID
    End If
    '--- Setup async events for the accepted socket
    If Not pvSetupAsyncEvents(FD_READ Or FD_WRITE Or FD_CLOSE) Then GoTo QH
    pvState = sckConnected
    '--- If TLS, server handshake would be needed here
    '--- (basic TCP accept is immediate)
QH:
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub SendData(data As Variant)
    On Error GoTo EH
    '--- Convert to byte array
    Dim baData() As Byte
    Select Case VarType(data)
    Case vbString
        '--- Auto-route: WebSocket text frame or raw send
        If m_lWsState = WS_OPEN Then
            pvWsSend CStr(data)
            Exit Sub
        End If
        baData = pvToAcpArray(CStr(data))
    Case vbByte + vbArray
        '--- Auto-route: WebSocket binary frame or raw send
        If m_lWsState = WS_OPEN Then
            baData = data
            pvWsSendBinary baData
            Exit Sub
        End If
        baData = data
    Case Else
        Err.Raise vbObjectError, , "Unsupported data type: " & TypeName(data)
    End Select
    If pvArraySize(baData) = 0 Then Exit Sub
    '--- If TLS, encrypt first
    If m_lTlsState = TLS_CONNECTED Then
        Dim baEncrypted() As Byte
        If Not pvTlsEncryptData(baData, baEncrypted) Then
            pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
            Exit Sub
        End If
        pvAppendBuffer m_baSendBuffer, baEncrypted
    Else
        pvAppendBuffer m_baSendBuffer, baData
    End If
    m_lSendPos = 0
    '--- Try to send immediately
    pvFlushSendBuffer
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub GetData(data As Variant, Optional ByVal type_ As Long, Optional ByVal maxLen As Long = -1)
    On Error GoTo EH
    If type_ = 0 Then type_ = VarType(data)
    Select Case type_
    Case vbString, vbByte + vbArray
    Case Else
        Err.Raise vbObjectError, , "Unsupported data type: " & type_
    End Select
    Dim baResult() As Byte
    baResult = vbNullString
    Dim lAvail As Long
    lAvail = pvArraySize(m_baRecvBuffer)
    If lAvail > 0 Then
        If maxLen < 0 Then
            '--- Return all
            baResult = m_baRecvBuffer
            m_baRecvBuffer = vbNullString
        ElseIf maxLen = 0 Then
            baResult = vbNullString
        Else
            If maxLen >= lAvail Then
                baResult = m_baRecvBuffer
                m_baRecvBuffer = vbNullString
            Else
                '--- Return partial
                ReDim baResult(0 To maxLen - 1)
                CopyMemory baResult(0), m_baRecvBuffer(0), maxLen
                '--- Shift remaining
                Dim lRemain As Long
                lRemain = lAvail - maxLen
                Dim baTemp() As Byte
                ReDim baTemp(0 To lRemain - 1)
                CopyMemory baTemp(0), m_baRecvBuffer(maxLen), lRemain
                m_baRecvBuffer = baTemp
            End If
        End If
    End If
    '--- Convert to requested type
    Select Case type_
    Case vbString
        data = pvFromAcpArray(baResult)
    Case vbByte + vbArray
        data = baResult
    End Select
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub PeekData(data As Variant, Optional ByVal type_ As Long, Optional ByVal maxLen As Long = -1)
    On Error GoTo EH
    If type_ = 0 Then type_ = VarType(data)
    Select Case type_
    Case vbString, vbByte + vbArray
    Case Else
        Err.Raise vbObjectError, , "Unsupported data type: " & type_
    End Select
    Dim baResult() As Byte
    baResult = vbNullString
    Dim lAvail As Long
    lAvail = pvArraySize(m_baRecvBuffer)
    If lAvail > 0 Then
        If maxLen < 0 Or maxLen >= lAvail Then
            baResult = m_baRecvBuffer
        ElseIf maxLen > 0 Then
            ReDim baResult(0 To maxLen - 1)
            CopyMemory baResult(0), m_baRecvBuffer(0), maxLen
        End If
    End If
    '--- Convert
    Select Case type_
    Case vbString
        data = pvFromAcpArray(baResult)
    Case vbByte + vbArray
        data = baResult
    End Select
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

'=========================================================================
' Public Methods - HTTP
'=========================================================================

Public Sub Request(Method As String, Url As String, Optional Body As String, Optional ContentType As String = "application/json", Optional ExtraHeaders As String)
    On Error GoTo EH
    '--- Parse URL
    Dim sScheme As String, sHost As String, lPort As Long, sPath As String
    pvParseUrl Url, sScheme, sHost, lPort, sPath
    '--- Set protocol based on scheme
    Select Case LCase$(sScheme)
    Case "https"
        m_eProtocol = sckTLSProtocol
    Case Else
        m_eProtocol = sckTCPProtocol
    End Select
    '--- Build request
    Dim sReq As String
    sReq = Method & " " & sPath & " HTTP/1.1" & vbCrLf & _
           "Host: " & sHost & vbCrLf
    If LenB(Body) <> 0 Then
        sReq = sReq & "Content-Type: " & ContentType & vbCrLf & _
               "Content-Length: " & Len(Body) & vbCrLf
    End If
    If LenB(ExtraHeaders) <> 0 Then
        sReq = sReq & ExtraHeaders & vbCrLf
    End If
    sReq = sReq & "Connection: close" & vbCrLf & vbCrLf
    If LenB(Body) <> 0 Then sReq = sReq & Body
    '--- Connect first
    m_sRemoteHost = sHost
    m_lRemotePort = lPort
    Connect m_sRemoteHost, m_lRemotePort
    '--- Setup HTTP state AFTER Connect (Connect calls Disconnect which resets state)
    m_lHttpState = HTTP_WAITING
    m_sHttpRawResponse = vbNullString
    m_lHttpStatus = 0
    m_sHttpHeaders = vbNullString
    m_sHttpContentType = vbNullString
    m_lHttpContentLength = -1
    m_bHttpChunked = False
    m_dTransferStart = Timer
    '--- Store request to send after connect
    m_sHttpBody = sReq
    Exit Sub
EH:
    m_lHttpState = HTTP_NONE
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Upload(Url As String, FilePath As String, _
    Optional FieldName As String = "file", _
    Optional ExtraHeaders As String)
    On Error GoTo EH
    '--- Read file into byte array
    Dim baFile() As Byte
    Dim lFileSize As Long
    pvReadFile FilePath, baFile, lFileSize
    If lFileSize = 0 Then Err.Raise vbObjectError, , "File is empty or not found: " & FilePath
    '--- Extract filename
    Dim sFileName As String
    Dim lSlash As Long
    lSlash = InStrRev(FilePath, "\")
    If lSlash = 0 Then lSlash = InStrRev(FilePath, "/")
    If lSlash > 0 Then
        sFileName = Mid$(FilePath, lSlash + 1)
    Else
        sFileName = FilePath
    End If
    '--- Build multipart body
    Dim sBoundary As String
    sBoundary = "----czSocket" & Hex$(GetTickCount()) & Hex$(CLng(Timer * 1000))
    Dim sPrefix As String
    sPrefix = "--" & sBoundary & vbCrLf & _
        "Content-Disposition: form-data; name=""" & FieldName & """; filename=""" & sFileName & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    Dim sSuffix As String
    sSuffix = vbCrLf & "--" & sBoundary & "--" & vbCrLf
    '--- Convert to byte arrays
    Dim baPrefix() As Byte
    baPrefix = pvToAcpArray(sPrefix)
    Dim baSuffix() As Byte
    baSuffix = pvToAcpArray(sSuffix)
    '--- Combine: prefix + file + suffix
    Dim lTotalSize As Long
    lTotalSize = pvArraySize(baPrefix) + lFileSize + pvArraySize(baSuffix)
    Dim baBody() As Byte
    ReDim baBody(0 To lTotalSize - 1)
    Dim lPos As Long
    CopyMemory baBody(0), baPrefix(0), pvArraySize(baPrefix)
    lPos = pvArraySize(baPrefix)
    CopyMemory baBody(lPos), baFile(0), lFileSize
    lPos = lPos + lFileSize
    CopyMemory baBody(lPos), baSuffix(0), pvArraySize(baSuffix)
    '--- Send via HttpRequest
    Dim sContentType As String
    sContentType = "multipart/form-data; boundary=" & sBoundary
    Dim sBodyStr As String
    sBodyStr = pvFromAcpArray(baBody)
    Request "POST", Url, sBodyStr, sContentType, ExtraHeaders
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Download(Url As String, SavePath As String, Optional ExtraHeaders As String)
    On Error GoTo EH
    m_sDownloadPath = SavePath
    m_bDownloadMode = True
    m_dTransferStart = Timer
    Request "GET", Url, , , ExtraHeaders
    Exit Sub
EH:
    m_bDownloadMode = False
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub SendFile(FilePath As String, Optional ByVal ChunkSize As Long = 8192)
    On Error GoTo EH
    '--- Validate: must be connected (raw or WebSocket)
    If m_eState <> sckConnected And m_lWsState <> WS_OPEN Then
        Err.Raise vbObjectError, , "Not connected"
    End If
    '--- Get file size
    Dim lSize As Long
    lSize = FileLen(FilePath)
    If lSize = 0 Then Err.Raise vbObjectError, , "File is empty: " & FilePath
    '--- Initialize chunked send state
    m_sSendFilePath = FilePath
    m_lSendFileChunk = ChunkSize
    m_lSendFileOffset = 0
    m_lSendFileSize = lSize
    m_bSendingFile = True
    m_dTransferStart = Timer
    '--- Send first chunk (auto-detects WS/raw mode)
    pvSendNextFileChunk
    Exit Sub
EH:
    m_bSendingFile = False
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

'=========================================================================
' Public Methods - WebSocket
'=========================================================================

Private Sub pvWsConnect(Url As String, Optional Protocols As String)
    On Error GoTo EH
    '--- Parse URL
    Dim sScheme As String, sHost As String, lPort As Long, sPath As String
    pvParseUrl Url, sScheme, sHost, lPort, sPath
    m_sWsHost = sHost
    m_sWsPath = sPath
    '--- Set protocol based on scheme
    Select Case LCase$(sScheme)
    Case "wss"
        m_eProtocol = sckTLSProtocol
    Case Else
        m_eProtocol = sckTCPProtocol
    End Select
    '--- Generate random key
    m_sWsKey = pvGenerateWsKey()
    '--- Connect first
    m_sRemoteHost = sHost
    m_lRemotePort = lPort
    Connect m_sRemoteHost, m_lRemotePort
    '--- Set WS state AFTER Connect (Connect calls Disconnect which resets state)
    m_lWsState = WS_CONNECTING
    m_baWsBuffer = vbNullString
    Exit Sub
EH:
    m_lWsState = WS_NONE
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Private Sub pvWsSend(Data As String)
    On Error GoTo EH
    If m_lWsState <> WS_OPEN Then Err.Raise vbObjectError, , "WebSocket not connected"
    Dim baPayload() As Byte
    baPayload = pvToUtf8Array(Data)
    Dim baFrame() As Byte
    pvWsBuildFrame WS_OPCODE_TEXT, baPayload, baFrame
    '--- Send raw frame through TLS or TCP
    If m_lTlsState = TLS_CONNECTED Then
        Dim baEnc() As Byte
        pvTlsEncryptData baFrame, baEnc
        pvAppendBuffer m_baSendBuffer, baEnc
    Else
        pvAppendBuffer m_baSendBuffer, baFrame
    End If
    m_lSendPos = 0
    pvFlushSendBuffer
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Private Sub pvWsSendBinary(Data() As Byte)
    On Error GoTo EH
    If m_lWsState <> WS_OPEN Then Err.Raise vbObjectError, , "WebSocket not connected"
    Dim baFrame() As Byte
    pvWsBuildFrame WS_OPCODE_BINARY, Data, baFrame
    If m_lTlsState = TLS_CONNECTED Then
        Dim baEnc() As Byte
        pvTlsEncryptData baFrame, baEnc
        pvAppendBuffer m_baSendBuffer, baEnc
    Else
        pvAppendBuffer m_baSendBuffer, baFrame
    End If
    m_lSendPos = 0
    pvFlushSendBuffer
    Exit Sub
EH:
    pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
End Sub

Public Sub Ping(Optional Data As String)
    On Error GoTo EH
    If m_lWsState <> WS_OPEN Then Exit Sub
    Dim baPayload() As Byte
    If LenB(Data) <> 0 Then
        baPayload = pvToUtf8Array(Data)
    Else
        baPayload = vbNullString
    End If
    Dim baFrame() As Byte
    pvWsBuildFrame WS_OPCODE_PING, baPayload, baFrame
    If m_lTlsState = TLS_CONNECTED Then
        Dim baEnc() As Byte
        pvTlsEncryptData baFrame, baEnc
        pvAppendBuffer m_baSendBuffer, baEnc
    Else
        pvAppendBuffer m_baSendBuffer, baFrame
    End If
    m_lSendPos = 0
    pvFlushSendBuffer
    Exit Sub
EH:
    PrintError "Ping"
End Sub

Private Sub pvWsClose(Optional ByVal Code As Long = 1000, Optional Reason As String)
    On Error GoTo EH
    If m_lWsState <> WS_OPEN Then Exit Sub
    m_lWsState = WS_CLOSING
    '--- Build close payload: 2-byte code + reason
    Dim baPayload() As Byte
    Dim baReason() As Byte
    baReason = vbNullString
    If LenB(Reason) <> 0 Then baReason = pvToUtf8Array(Reason)
    ReDim baPayload(0 To 1 + pvArraySize(baReason))
    baPayload(0) = CByte((Code \ 256) And &HFF)
    baPayload(1) = CByte(Code And &HFF)
    If pvArraySize(baReason) > 0 Then
        CopyMemory baPayload(2), baReason(0), pvArraySize(baReason)
    End If
    Dim baFrame() As Byte
    pvWsBuildFrame WS_OPCODE_CLOSE, baPayload, baFrame
    If m_lTlsState = TLS_CONNECTED Then
        Dim baEnc() As Byte
        pvTlsEncryptData baFrame, baEnc
        pvAppendBuffer m_baSendBuffer, baEnc
    Else
        pvAppendBuffer m_baSendBuffer, baFrame
    End If
    m_lSendPos = 0
    pvFlushSendBuffer
    Exit Sub
EH:
    PrintError "pvWsClose"
End Sub

'=========================================================================
' Private - Socket engine
'=========================================================================

Private Function pvCreateSocket() As Boolean
    Dim lType As Long
    Dim lProto As Long
    Select Case m_eProtocol
    Case sckUDPProtocol
        lType = SOCK_DGRAM
        lProto = IPPROTO_UDP
    Case Else '--- TCP and TLS
        lType = SOCK_STREAM
        lProto = IPPROTO_TCP
    End Select
    m_hSocket = ws_socket(AF_INET, lType, lProto)
    If m_hSocket = INVALID_SOCKET Then
        pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
        Exit Function
    End If
    pvCreateSocket = True
End Function

Private Function pvSetupAsyncEvents(ByVal lEvents As Long) As Boolean
    '--- Create WSA event if not exists
    If m_hEvent = 0 Then
        m_hEvent = WSACreateEvent()
        If m_hEvent = 0 Then
            pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
            Exit Function
        End If
    End If
    '--- Associate socket with event
    If WSAEventSelect(m_hSocket, m_hEvent, lEvents) = SOCKET_ERROR Then
        pvSetError LastDllError:=WSAGetLastError(), RaiseError:=True
        Exit Function
    End If
    '--- Enable polling timer
    tmrPoll.Enabled = True
    pvSetupAsyncEvents = True
End Function

Private Function pvResolveHost(sHost As String) As Long
    '--- Try as IP address first
    pvResolveHost = ws_inet_addr(sHost)
    If pvResolveHost <> INADDR_NONE Then Exit Function
    '--- DNS lookup (synchronous)
    Dim pHost As Long
    pHost = ws_gethostbyname(sHost)
    If pHost = 0 Then
        pvResolveHost = INADDR_NONE
        Exit Function
    End If
    '--- Extract first IP address from hostent
    Dim lAddrList As Long
    CopyMemory lAddrList, ByVal pHost + 12, 4   '--- h_addr_list offset
    If lAddrList = 0 Then
        pvResolveHost = INADDR_NONE
        Exit Function
    End If
    Dim lFirstAddr As Long
    CopyMemory lFirstAddr, ByVal lAddrList, 4
    If lFirstAddr = 0 Then
        pvResolveHost = INADDR_NONE
        Exit Function
    End If
    CopyMemory pvResolveHost, ByVal lFirstAddr, 4
End Function

Private Sub pvFlushSendBuffer()
    Dim lSent As Long
    Dim lRemaining As Long
    Do While pvArraySize(m_baSendBuffer) > 0 And m_lSendPos <= UBound(m_baSendBuffer)
        lRemaining = UBound(m_baSendBuffer) + 1 - m_lSendPos
        If m_eProtocol = sckUDPProtocol Then
            Dim uAddr As SOCKADDR_IN
            uAddr.sin_family = AF_INET
            uAddr.sin_addr = pvResolveHost(m_sRemoteHost)
            uAddr.sin_port = ws_htons(m_lRemotePort)
            lSent = ws_sendto(m_hSocket, m_baSendBuffer(m_lSendPos), lRemaining, 0, uAddr, LenB(uAddr))
        Else
            lSent = ws_send(m_hSocket, m_baSendBuffer(m_lSendPos), lRemaining, 0)
        End If
        If lSent = SOCKET_ERROR Then
            Dim lErr As Long
            lErr = WSAGetLastError()
            If lErr = WSAEWOULDBLOCK Then
                Exit Do     '--- Wait for FD_WRITE
            Else
                pvSetError LastDllError:=lErr
                Exit Sub
            End If
        ElseIf lSent = 0 Then
            Exit Do
        Else
            m_lSendPos = m_lSendPos + lSent
            RaiseEvent SendProgress(m_lSendPos, UBound(m_baSendBuffer) + 1 - m_lSendPos)
        End If
    Loop
    '--- Check if all sent
    If pvArraySize(m_baSendBuffer) > 0 And m_lSendPos > UBound(m_baSendBuffer) Then
        m_lSendPos = 0
        m_baSendBuffer = vbNullString
        RaiseEvent SendComplete
    End If
End Sub

'=========================================================================
' Private - TLS engine (SSPI/Schannel)
'=========================================================================

Private Function pvTlsAcquireCredentials(Optional ByVal IsServer As Boolean) As Boolean
    Dim uCred As SCHANNEL_CRED
    Dim tsExpiry As Currency
    Dim lResult As Long
    uCred.dwVersion = SCHANNEL_CRED_VERSION
    '--- Set protocol flags
    If m_lLocalFeatures = 0 Then
        '--- Let Schannel choose best protocol
        uCred.grbitEnabledProtocols = 0
    Else
        If (m_lLocalFeatures And ucsSckSupportTls10) <> 0 Then
            uCred.grbitEnabledProtocols = uCred.grbitEnabledProtocols Or SP_PROT_TLS1_0
        End If
        If (m_lLocalFeatures And ucsSckSupportTls11) <> 0 Then
            uCred.grbitEnabledProtocols = uCred.grbitEnabledProtocols Or SP_PROT_TLS1_1
        End If
        If (m_lLocalFeatures And ucsSckSupportTls12) <> 0 Then
            uCred.grbitEnabledProtocols = uCred.grbitEnabledProtocols Or SP_PROT_TLS1_2
        End If
        If (m_lLocalFeatures And ucsSckSupportTls13) <> 0 Then
            uCred.grbitEnabledProtocols = uCred.grbitEnabledProtocols Or SP_PROT_TLS1_3
        End If
    End If
    '--- Set credential flags
    uCred.dwFlags = SCH_CRED_NO_DEFAULT_CREDS
    If (m_lLocalFeatures And ucsSckIgnoreServerCertificateErrors) <> 0 Then
        uCred.dwFlags = uCred.dwFlags Or SCH_CRED_MANUAL_CRED_VALIDATION
    Else
        uCred.dwFlags = uCred.dwFlags Or SCH_CRED_AUTO_CRED_VALIDATION
    End If
    If (m_lLocalFeatures And ucsSckIgnoreServerCertificateRevocation) <> 0 Then
        uCred.dwFlags = uCred.dwFlags Or SCH_CRED_IGNORE_NO_REVOCATION_CHECK Or SCH_CRED_IGNORE_REVOCATION_OFFLINE
    End If
    Dim lUsage As Long
    If IsServer Then
        lUsage = SECPKG_CRED_INBOUND
    Else
        lUsage = SECPKG_CRED_OUTBOUND
    End If
    lResult = AcquireCredentialsHandle(0&, UNISP_NAME, lUsage, 0&, uCred, 0&, 0&, m_hCred(0), tsExpiry)
    If lResult <> SEC_E_OK Then
        pvSetError LastDllError:=lResult
        Exit Function
    End If
    m_bHasCred = True
    pvTlsAcquireCredentials = True
End Function

Private Function pvTlsBeginHandshake() As Boolean
    '--- Acquire credentials
    If Not pvTlsAcquireCredentials(False) Then Exit Function
    '--- Initialize buffers
    Dim baEmpty() As Byte
    baEmpty = vbNullString
    m_baTlsPending = baEmpty
    m_lTlsState = TLS_HANDSHAKE
    '--- First InitializeSecurityContext call (no input)
    pvTlsBeginHandshake = pvTlsHandshakeStep(baEmpty)
End Function

Private Function pvTlsHandshakeStep(baInput() As Byte) As Boolean
    Const FUNC_NAME As String = "pvTlsHandshakeStep"
    Dim lFlags As Long
    Dim lAttr As Long
    Dim tsExpiry As Currency
    Dim lResult As Long
    On Error GoTo EH
    lFlags = ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or _
             ISC_REQ_EXTENDED_ERROR Or ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM
    If (m_lLocalFeatures And ucsSckIgnoreServerCertificateErrors) <> 0 Then
        lFlags = lFlags Or ISC_REQ_MANUAL_CRED_VALIDATION
    End If
    '--- Setup output buffer descriptor
    Dim uOutBuf(0 To 0) As SecBuffer
    Dim uOutDesc As SecBufferDesc
    uOutBuf(0).BufferType = SECBUFFER_TOKEN
    uOutBuf(0).cbBuffer = 0
    uOutBuf(0).pvBuffer = 0
    uOutDesc.ulVersion = SECBUFFER_VERSION
    uOutDesc.cBuffers = 1
    uOutDesc.pBuffers = VarPtr(uOutBuf(0))
    Dim lInputSize As Long
    lInputSize = pvArraySize(baInput)
    If lInputSize = 0 And Not m_bHasCtx Then
        '--- First call: no context, no input
        lResult = InitializeSecurityContext( _
            m_hCred(0), ByVal 0&, m_sRemoteHost, lFlags, 0, 0, _
            ByVal 0&, 0, m_hCtx(0), uOutDesc, lAttr, tsExpiry)
    Else
        '--- Subsequent calls: have context and input
        Dim uInBuf(0 To 1) As SecBuffer
        Dim uInDesc As SecBufferDesc
        uInBuf(0).BufferType = SECBUFFER_TOKEN
        uInBuf(0).cbBuffer = lInputSize
        If lInputSize > 0 Then
            uInBuf(0).pvBuffer = VarPtr(baInput(0))
        End If
        uInBuf(1).BufferType = SECBUFFER_EMPTY
        uInBuf(1).cbBuffer = 0
        uInBuf(1).pvBuffer = 0
        uInDesc.ulVersion = SECBUFFER_VERSION
        uInDesc.cBuffers = 2
        uInDesc.pBuffers = VarPtr(uInBuf(0))
        lResult = InitializeSecurityContext( _
            m_hCred(0), m_hCtx(0), m_sRemoteHost, lFlags, 0, 0, _
            uInDesc, 0, m_hCtx(0), uOutDesc, lAttr, tsExpiry)
    End If
    m_bHasCtx = True
    '--- Send output token if any
    If uOutBuf(0).cbBuffer > 0 And uOutBuf(0).pvBuffer <> 0 Then
        Dim baSend() As Byte
        ReDim baSend(0 To uOutBuf(0).cbBuffer - 1)
        CopyMemory baSend(0), ByVal uOutBuf(0).pvBuffer, uOutBuf(0).cbBuffer
        FreeContextBuffer uOutBuf(0).pvBuffer
        '--- Direct send (during handshake)
        pvRawSend baSend
    End If
    '--- Handle extra data from input
    If lInputSize > 0 Then
        If uInBuf(1).BufferType = SECBUFFER_EXTRA And uInBuf(1).cbBuffer > 0 Then
            '--- Save extra data for next step
            Dim baExtra() As Byte
            ReDim baExtra(0 To uInBuf(1).cbBuffer - 1)
            CopyMemory baExtra(0), baInput(lInputSize - uInBuf(1).cbBuffer), uInBuf(1).cbBuffer
            m_baTlsPending = baExtra
        Else
            m_baTlsPending = vbNullString
        End If
    End If
    '--- Process result
    Select Case lResult
    Case SEC_E_OK
        '--- Handshake complete!
        QueryContextAttributes m_hCtx(0), SECPKG_ATTR_STREAM_SIZES, m_uStreamSizes
        m_lTlsState = TLS_CONNECTED
        m_bConnecting = False
        pvState = sckConnected
        '--- Auto-send for HTTPS/WSS modes
        pvPostConnectActions
        '--- Don't fire Connect for WebSocket (pvWsProcessHandshakeResponse will fire it)
        If m_lWsState <> WS_CONNECTING Then RaiseEvent Connect
        '--- Check if there's extra data that might be application data
        If pvArraySize(m_baTlsPending) > 0 Then
            pvTlsProcessReceivedData
        End If
    Case SEC_I_CONTINUE_NEEDED
        '--- Need more data from server - wait for FD_READ
    Case SEC_E_INCOMPLETE_MESSAGE
        '--- Need more data - restore input to pending
        If lInputSize > 0 Then
            m_baTlsPending = baInput
        End If
    Case SEC_I_INCOMPLETE_CREDENTIALS
        '--- Server wants client cert - continue without
        Dim baEmptyCred() As Byte
        baEmptyCred = vbNullString
        pvTlsHandshakeStep baEmptyCred
    Case Else
        '--- Error
        pvSetError LastDllError:=lResult
        m_lTlsState = TLS_NONE
        Exit Function
    End Select
    pvTlsHandshakeStep = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvTlsEncryptData(baPlain() As Byte, baEncrypted() As Byte) As Boolean
    Dim lPlainSize As Long
    lPlainSize = pvArraySize(baPlain)
    If lPlainSize = 0 Then Exit Function
    Dim lBufSize As Long
    Dim lOffset As Long
    Dim lChunkSize As Long
    baEncrypted = vbNullString
    lOffset = 0
    Do While lOffset < lPlainSize
        lChunkSize = lPlainSize - lOffset
        If lChunkSize > m_uStreamSizes.cbMaximumMessage Then
            lChunkSize = m_uStreamSizes.cbMaximumMessage
        End If
        '--- Build encrypt buffer: [Header][Data][Trailer]
        lBufSize = m_uStreamSizes.cbHeader + lChunkSize + m_uStreamSizes.cbTrailer
        Dim baBuf() As Byte
        ReDim baBuf(0 To lBufSize - 1)
        CopyMemory baBuf(m_uStreamSizes.cbHeader), baPlain(lOffset), lChunkSize
        '--- Setup SecBuffers
        Dim uBuf(0 To 3) As SecBuffer
        Dim uDesc As SecBufferDesc
        uBuf(0).BufferType = SECBUFFER_STREAM_HEADER
        uBuf(0).cbBuffer = m_uStreamSizes.cbHeader
        uBuf(0).pvBuffer = VarPtr(baBuf(0))
        uBuf(1).BufferType = SECBUFFER_DATA
        uBuf(1).cbBuffer = lChunkSize
        uBuf(1).pvBuffer = VarPtr(baBuf(m_uStreamSizes.cbHeader))
        uBuf(2).BufferType = SECBUFFER_STREAM_TRAILER
        uBuf(2).cbBuffer = m_uStreamSizes.cbTrailer
        uBuf(2).pvBuffer = VarPtr(baBuf(m_uStreamSizes.cbHeader + lChunkSize))
        uBuf(3).BufferType = SECBUFFER_EMPTY
        uDesc.ulVersion = SECBUFFER_VERSION
        uDesc.cBuffers = 4
        uDesc.pBuffers = VarPtr(uBuf(0))
        Dim lResult As Long
        lResult = EncryptMessage(m_hCtx(0), 0, uDesc, 0)
        If lResult <> SEC_E_OK Then
            pvSetError LastDllError:=lResult
            Exit Function
        End If
        '--- Actual encrypted size
        Dim lEncSize As Long
        lEncSize = uBuf(0).cbBuffer + uBuf(1).cbBuffer + uBuf(2).cbBuffer
        ReDim Preserve baBuf(0 To lEncSize - 1)
        '--- Append to output
        pvAppendBuffer baEncrypted, baBuf
        lOffset = lOffset + lChunkSize
    Loop
    pvTlsEncryptData = True
End Function

Private Sub pvTlsProcessReceivedData()
    Const FUNC_NAME As String = "pvTlsProcessReceivedData"
    Dim lResult As Long
    Dim bHasData As Boolean
    On Error GoTo EH
    Do While pvArraySize(m_baTlsPending) > 0
        '--- Setup decrypt buffers
        Dim uBuf(0 To 3) As SecBuffer
        Dim uDesc As SecBufferDesc
        uBuf(0).BufferType = SECBUFFER_DATA
        uBuf(0).cbBuffer = pvArraySize(m_baTlsPending)
        uBuf(0).pvBuffer = VarPtr(m_baTlsPending(0))
        uBuf(1).BufferType = SECBUFFER_EMPTY
        uBuf(1).cbBuffer = 0
        uBuf(1).pvBuffer = 0
        uBuf(2).BufferType = SECBUFFER_EMPTY
        uBuf(2).cbBuffer = 0
        uBuf(2).pvBuffer = 0
        uBuf(3).BufferType = SECBUFFER_EMPTY
        uBuf(3).cbBuffer = 0
        uBuf(3).pvBuffer = 0
        uDesc.ulVersion = SECBUFFER_VERSION
        uDesc.cBuffers = 4
        uDesc.pBuffers = VarPtr(uBuf(0))
        lResult = DecryptMessage(m_hCtx(0), uDesc, 0, 0)
        Select Case lResult
        Case SEC_E_OK
            '--- Find decrypted data and extra data
            Dim i As Long
            Dim baExtra() As Byte
            baExtra = vbNullString
            For i = 0 To 3
                If uBuf(i).BufferType = SECBUFFER_DATA And uBuf(i).cbBuffer > 0 Then
                    '--- Copy decrypted data to recv buffer
                    Dim baDecrypted() As Byte
                    ReDim baDecrypted(0 To uBuf(i).cbBuffer - 1)
                    CopyMemory baDecrypted(0), ByVal uBuf(i).pvBuffer, uBuf(i).cbBuffer
                    pvAppendBuffer m_baRecvBuffer, baDecrypted
                    bHasData = True
                End If
                If uBuf(i).BufferType = SECBUFFER_EXTRA And uBuf(i).cbBuffer > 0 Then
                    ReDim baExtra(0 To uBuf(i).cbBuffer - 1)
                    CopyMemory baExtra(0), ByVal uBuf(i).pvBuffer, uBuf(i).cbBuffer
                End If
            Next i
            '--- Set pending to extra data (may be empty)
            m_baTlsPending = baExtra
            '--- Continue loop if there's more data
            If pvArraySize(m_baTlsPending) = 0 Then Exit Do
        Case SEC_E_INCOMPLETE_MESSAGE
            '--- Need more data, keep pending buffer as is
            Exit Do
        Case SEC_I_CONTEXT_EXPIRED
            '--- TLS connection closed by peer
            m_lTlsState = TLS_NONE
            '--- Dispatch any pending received data before closing
            If bHasData And pvArraySize(m_baRecvBuffer) > 0 Then
                pvDispatchReceivedData
                bHasData = False
            End If
            pvOnClose
            Exit Sub
        Case SEC_I_RENEGOTIATE
            '--- Server wants to renegotiate
            m_lTlsState = TLS_HANDSHAKE
            Dim baEmptyReneg() As Byte
            baEmptyReneg = vbNullString
            pvTlsHandshakeStep baEmptyReneg
            Exit Do
        Case Else
            '--- Error
            pvSetError LastDllError:=lResult
            Exit Do
        End Select
    Loop
    '--- Raise DataArrival if we have data
    If bHasData And pvArraySize(m_baRecvBuffer) > 0 Then
        pvDispatchReceivedData
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvTlsShutdown()
    Const FUNC_NAME As String = "pvTlsShutdown"
    On Error GoTo EH
    If Not m_bHasCtx Then Exit Sub
    '--- Apply shutdown token
    Dim lToken As Long
    lToken = SCHANNEL_SHUTDOWN_TOKEN
    Dim uBuf(0 To 0) As SecBuffer
    Dim uDesc As SecBufferDesc
    uBuf(0).BufferType = SECBUFFER_TOKEN
    uBuf(0).cbBuffer = 4
    uBuf(0).pvBuffer = VarPtr(lToken)
    uDesc.ulVersion = SECBUFFER_VERSION
    uDesc.cBuffers = 1
    uDesc.pBuffers = VarPtr(uBuf(0))
    ApplyControlToken m_hCtx(0), uDesc
    '--- Get shutdown message
    Dim uOutBuf(0 To 0) As SecBuffer
    Dim uOutDesc As SecBufferDesc
    uOutBuf(0).BufferType = SECBUFFER_TOKEN
    uOutBuf(0).cbBuffer = 0
    uOutBuf(0).pvBuffer = 0
    uOutDesc.ulVersion = SECBUFFER_VERSION
    uOutDesc.cBuffers = 1
    uOutDesc.pBuffers = VarPtr(uOutBuf(0))
    Dim lFlags As Long
    Dim lAttr As Long
    Dim tsExpiry As Currency
    lFlags = ISC_REQ_SEQUENCE_DETECT Or ISC_REQ_REPLAY_DETECT Or ISC_REQ_CONFIDENTIALITY Or _
             ISC_REQ_ALLOCATE_MEMORY Or ISC_REQ_STREAM
    '--- ISC for shutdown: pass empty input desc, receive output token
    Dim uEmptyDesc As SecBufferDesc
    uEmptyDesc.ulVersion = SECBUFFER_VERSION
    uEmptyDesc.cBuffers = 1
    uEmptyDesc.pBuffers = VarPtr(uOutBuf(0))
    InitializeSecurityContext m_hCred(0), m_hCtx(0), m_sRemoteHost, lFlags, 0, 0, _
        ByVal 0&, 0, m_hCtx(0), uEmptyDesc, lAttr, tsExpiry
    '--- Send shutdown message
    If uOutBuf(0).cbBuffer > 0 And uOutBuf(0).pvBuffer <> 0 Then
        Dim baSend() As Byte
        ReDim baSend(0 To uOutBuf(0).cbBuffer - 1)
        CopyMemory baSend(0), ByVal uOutBuf(0).pvBuffer, uOutBuf(0).cbBuffer
        FreeContextBuffer uOutBuf(0).pvBuffer
        pvRawSend baSend
    End If
    m_lTlsState = TLS_SHUTDOWN
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvTlsCleanup()
    If m_bHasCtx Then
        DeleteSecurityContext m_hCtx(0)
        m_hCtx(0) = 0
        m_hCtx(1) = 0
        m_bHasCtx = False
    End If
    If m_bHasCred Then
        FreeCredentialsHandle m_hCred(0)
        m_hCred(0) = 0
        m_hCred(1) = 0
        m_bHasCred = False
    End If
    m_lTlsState = TLS_NONE
    m_baTlsPending = vbNullString
    ZeroMemory m_uStreamSizes, LenB(m_uStreamSizes)
End Sub

Private Sub pvRawSend(baData() As Byte)
    Dim lSize As Long
    lSize = pvArraySize(baData)
    If lSize = 0 Then Exit Sub
    Dim lPos As Long
    Dim lSent As Long
    lPos = 0
    Do While lPos < lSize
        lSent = ws_send(m_hSocket, baData(lPos), lSize - lPos, 0)
        If lSent = SOCKET_ERROR Then
            Dim lErr As Long
            lErr = WSAGetLastError()
            If lErr = WSAEWOULDBLOCK Then
                '--- Brief wait and retry during handshake
                Dim dStart As Double
                dStart = Timer
                Do
                    DoEvents
                    If Timer - dStart > CDbl(m_lTimeout) / 1000# Then Exit Sub
                    lSent = ws_send(m_hSocket, baData(lPos), lSize - lPos, 0)
                    If lSent <> SOCKET_ERROR Then Exit Do
                    If WSAGetLastError() <> WSAEWOULDBLOCK Then Exit Sub
                Loop
            Else
                Exit Sub
            End If
        End If
        If lSent > 0 Then lPos = lPos + lSent
    Loop
End Sub

'=========================================================================
' Private - Async event handler (Timer-based polling)
'=========================================================================

Private Sub tmrPoll_Timer()
    Const FUNC_NAME As String = "tmrPoll_Timer"
    On Error GoTo EH
    If m_hSocket = INVALID_SOCKET Or m_hSocket = 0 Then Exit Sub
    If m_hEvent = 0 Then Exit Sub
    '--- Check WSA event (non-blocking, timeout=0)
    Dim lWait As Long
    lWait = WSAWaitForMultipleEvents(1, m_hEvent, 0, 0, 0)
    If lWait <> WSA_WAIT_EVENT_0 Then Exit Sub
    '--- Enumerate events
    Dim uNE As WSANETWORKEVENTS_TYPE
    If WSAEnumNetworkEvents(m_hSocket, m_hEvent, uNE) = SOCKET_ERROR Then Exit Sub
    '--- Process events
    If (uNE.lNetworkEvents And FD_CONNECT) <> 0 Then
        If uNE.iErrorCode(FD_CONNECT_BIT) <> 0 Then
            pvSetError LastDllError:=uNE.iErrorCode(FD_CONNECT_BIT)
        Else
            pvOnConnect
        End If
    End If
    If (uNE.lNetworkEvents And FD_ACCEPT) <> 0 Then
        If uNE.iErrorCode(FD_ACCEPT_BIT) = 0 Then
            pvOnAccept
        End If
    End If
    If (uNE.lNetworkEvents And FD_READ) <> 0 Then
        If uNE.iErrorCode(FD_READ_BIT) = 0 Then
            pvOnReceive
        End If
    End If
    If (uNE.lNetworkEvents And FD_WRITE) <> 0 Then
        If uNE.iErrorCode(FD_WRITE_BIT) = 0 Then
            pvOnSend
        End If
    End If
    If (uNE.lNetworkEvents And FD_CLOSE) <> 0 Then
        pvOnClose
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Private - Internal event handlers
'=========================================================================

Private Sub pvOnConnect()
    Const FUNC_NAME As String = "pvOnConnect"
    On Error GoTo EH
    If m_eProtocol = sckTLSProtocol Then
        '--- Start TLS handshake
        pvTlsBeginHandshake
    Else
        '--- TCP connected
        m_bConnecting = False
        pvState = sckConnected
        '--- Auto-send for HTTP/WS modes
        pvPostConnectActions
        If m_lWsState <> WS_CONNECTING Then RaiseEvent Connect
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvOnReceive()
    Const FUNC_NAME As String = "pvOnReceive"
    Dim baBuffer() As Byte
    Dim lRecv As Long
    On Error GoTo EH
    '--- Receive data from socket
    ReDim baBuffer(0 To RECV_BUFFER_SIZE - 1)
    If m_eProtocol = sckUDPProtocol Then
        Dim uFrom As SOCKADDR_IN
        Dim lFromLen As Long
        lFromLen = LenB(uFrom)
        lRecv = ws_recvfrom(m_hSocket, baBuffer(0), RECV_BUFFER_SIZE, 0, uFrom, lFromLen)
    Else
        lRecv = ws_recv(m_hSocket, baBuffer(0), RECV_BUFFER_SIZE, 0)
    End If
    If lRecv = SOCKET_ERROR Then
        Dim lErr As Long
        lErr = WSAGetLastError()
        If lErr = WSAEWOULDBLOCK Then Exit Sub
        pvSetError LastDllError:=lErr
        Exit Sub
    End If
    If lRecv = 0 Then
        '--- Connection closed gracefully
        pvOnClose
        Exit Sub
    End If
    '--- Trim buffer to actual size
    ReDim Preserve baBuffer(0 To lRecv - 1)
    '--- Process based on TLS state
    Select Case m_lTlsState
    Case TLS_HANDSHAKE
        '--- Append to TLS pending and continue handshake
        pvAppendBuffer m_baTlsPending, baBuffer
        pvTlsHandshakeStep m_baTlsPending
    Case TLS_CONNECTED
        '--- Append to TLS pending and decrypt
        pvAppendBuffer m_baTlsPending, baBuffer
        pvTlsProcessReceivedData
    Case Else
        '--- Plain TCP/UDP: append directly to recv buffer
        pvAppendBuffer m_baRecvBuffer, baBuffer
        pvDispatchReceivedData
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvOnSend()
    Const FUNC_NAME As String = "pvOnSend"
    On Error GoTo EH
    '--- Flush pending send data
    If pvArraySize(m_baSendBuffer) > 0 Then
        pvFlushSendBuffer
    End If
    '--- Continue chunked file sending if active and buffer empty
    If m_bSendingFile And pvArraySize(m_baSendBuffer) = 0 Then
        pvSendNextFileChunk
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvDispatchReceivedData()
    '--- Route received data based on active mode
    Select Case m_lWsState
    Case WS_CONNECTING
        pvWsProcessHandshakeResponse
    Case WS_OPEN, WS_CLOSING
        pvWsProcessFrames
    Case Else
        If m_lHttpState = HTTP_WAITING Then
            pvHttpProcessResponse
        Else
            RaiseEvent DataArrival(pvArraySize(m_baRecvBuffer))
        End If
    End Select
End Sub

Private Sub pvOnAccept()
    Const FUNC_NAME As String = "pvOnAccept"
    On Error GoTo EH
    '--- Accept the incoming connection
    Dim uAddr As SOCKADDR_IN
    Dim lLen As Long
    lLen = LenB(uAddr)
    Dim hNewSocket As Long
    hNewSocket = ws_accept(m_hSocket, uAddr, lLen)
    If hNewSocket = INVALID_SOCKET Then Exit Sub
    '--- Store globally for Accept pattern
    SetPropA GetDesktopWindow(), PROP_REQUEST_SOCKET, hNewSocket
    SetPropA GetDesktopWindow(), PROP_REQUEST_PROTOCOL, CLng(m_eProtocol)
    '--- Raise event
    pvState = sckConnectionPending
    RaiseEvent ConnectionRequest(hNewSocket)
    '--- Clear global
    RemovePropA GetDesktopWindow(), PROP_REQUEST_SOCKET
    RemovePropA GetDesktopWindow(), PROP_REQUEST_PROTOCOL
    pvState = sckListening
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvOnClose()
    Const FUNC_NAME As String = "pvOnClose"
    On Error GoTo EH
    '--- Guard against double-call
    If m_eState = sckClosed Then Exit Sub
    '--- Check for remaining data
    If m_lTlsState = TLS_CONNECTED Then
        '--- Try to read any remaining TLS data
        Dim baBuffer() As Byte
        Dim lRecv As Long
        ReDim baBuffer(0 To RECV_BUFFER_SIZE - 1)
        Do
            lRecv = ws_recv(m_hSocket, baBuffer(0), RECV_BUFFER_SIZE, 0)
            If lRecv > 0 Then
                ReDim Preserve baBuffer(0 To lRecv - 1)
                pvAppendBuffer m_baTlsPending, baBuffer
                pvTlsProcessReceivedData
                ReDim baBuffer(0 To RECV_BUFFER_SIZE - 1)
            End If
        Loop While lRecv > 0
    End If
    '--- HTTP: if waiting for response with no Content-Length, deliver what we have
    If m_lHttpState = HTTP_WAITING And LenB(m_sHttpRawResponse) <> 0 Then
        Dim lHdrEnd As Long
        lHdrEnd = InStr(m_sHttpRawResponse, vbCrLf & vbCrLf)
        If lHdrEnd > 0 And m_lHttpContentLength < 0 And Not m_bHttpChunked Then
            m_sHttpBody = Mid$(m_sHttpRawResponse, lHdrEnd + 4)
            m_lHttpState = HTTP_NONE
            pvFireResponse m_lHttpStatus, m_sHttpContentType, m_sHttpBody, m_sHttpHeaders
        End If
    End If
    m_lHttpState = HTTP_NONE
    m_lWsState = WS_NONE
    pvState = sckClosing
    RaiseEvent Disconnected(0, vbNullString)
    '--- Cleanup
    If m_hSocket <> INVALID_SOCKET And m_hSocket <> 0 Then
        ws_closesocket m_hSocket
        m_hSocket = INVALID_SOCKET
    End If
    If m_hEvent <> 0 Then
        WSACloseEvent m_hEvent
        m_hEvent = 0
    End If
    pvTlsCleanup
    tmrPoll.Enabled = False
    m_bListening = False
    m_bConnecting = False
    pvState = sckClosed
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Private - Error handling
'=========================================================================

Private Sub pvSetError(Optional ByVal LastDllError As Long, Optional LastError As VBA.ErrObject, Optional ByVal RaiseError As Boolean)
    Const LNG_FACILITY_WIN32 As Long = &H80070000
    Dim Number As Long
    Dim Description As String
    Dim Source As String
    Dim bCancel As Boolean
    '--- Guard: no actual error info, skip (prevents Err.Raise 0 crash in compiled EXE)
    If LastDllError = 0 And LastError Is Nothing Then Exit Sub
    pvState = sckError
    If LastDllError <> 0 Then
        Number = LastDllError
        If Number > 0 And Number < &H80000000 Then
            Number = Number Or LNG_FACILITY_WIN32
        End If
        Description = pvGetErrorDescription(LastDllError)
    ElseIf Not LastError Is Nothing Then
        Number = LastError.Number
        Source = LastError.Source
        Description = LastError.Description
    End If
    RaiseEvent Error(Number, Description, CLng(LastDllError And &HFFFF&), Source, App.HelpFile, 0, bCancel)
    If Not bCancel And RaiseError Then
        If Number <> 0 Then Err.Raise Number, Source, Description
    End If
End Sub

Private Property Let pvState(ByVal eValue As UcsStateConstants)
    m_eState = eValue
    Select Case eValue
    Case sckClosing, sckConnected, sckListening, sckOpen
        '--- preserve buffers
    Case Else
        m_baRecvBuffer = vbNullString
        m_baSendBuffer = vbNullString
        m_lSendPos = 0
    End Select
End Property

'=========================================================================
' Helper functions
'=========================================================================

Private Function pvGetErrorDescription(ByVal ErrorCode As Long) As String
    Dim sBuf As String
    sBuf = String$(1024, 0)
    Dim lLen As Long
    lLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, ErrorCode, 0, sBuf, 1024, 0)
    If lLen > 0 Then
        pvGetErrorDescription = Left$(sBuf, lLen)
        '--- Trim trailing CrLf
        Do While Right$(pvGetErrorDescription, 1) = vbCr Or Right$(pvGetErrorDescription, 1) = vbLf
            pvGetErrorDescription = Left$(pvGetErrorDescription, Len(pvGetErrorDescription) - 1)
        Loop
    Else
        pvGetErrorDescription = "Error &H" & Hex$(ErrorCode)
    End If
End Function

Private Function pvGetHostIP(ByVal pHostent As Long) As String
    If pHostent = 0 Then Exit Function
    Dim lAddrList As Long
    CopyMemory lAddrList, ByVal pHostent + 12, 4
    If lAddrList = 0 Then Exit Function
    Dim lFirstAddr As Long
    CopyMemory lFirstAddr, ByVal lAddrList, 4
    If lFirstAddr = 0 Then Exit Function
    Dim lAddr As Long
    CopyMemory lAddr, ByVal lFirstAddr, 4
    pvGetHostIP = pvInetNtoa(lAddr)
End Function

Private Function pvInetNtoa(ByVal lAddr As Long) As String
    Dim pStr As Long
    pStr = ws_inet_ntoa(lAddr)
    If pStr <> 0 Then
        Dim lLen As Long
        lLen = lstrlenA(pStr)
        If lLen > 0 Then
            '--- inet_ntoa returns ANSI string, convert to VB Unicode
            Dim baAnsi() As Byte
            ReDim baAnsi(0 To lLen - 1)
            CopyMemory baAnsi(0), ByVal pStr, lLen
            pvInetNtoa = StrConv(baAnsi, vbUnicode)
        End If
    End If
End Function

Private Function pvArraySize(baArr() As Byte) As Long
    On Error Resume Next
    pvArraySize = UBound(baArr) + 1
    If Err.Number <> 0 Then pvArraySize = 0
    On Error GoTo 0
End Function

Private Sub pvAppendBuffer(baDst() As Byte, baSrc() As Byte)
    Dim lDstSize As Long
    Dim lSrcSize As Long
    lSrcSize = pvArraySize(baSrc)
    If lSrcSize = 0 Then Exit Sub
    lDstSize = pvArraySize(baDst)
    If lDstSize = 0 Then
        baDst = baSrc
    Else
        ReDim Preserve baDst(0 To lDstSize + lSrcSize - 1)
        CopyMemory baDst(lDstSize), baSrc(0), lSrcSize
    End If
End Sub

Private Function pvToAcpArray(sText As String) As Byte()
    If LenB(sText) = 0 Then
        pvToAcpArray = vbNullString
        Exit Function
    End If
    Dim lLen As Long
    lLen = WideCharToMultiByte(0, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lLen > 0 Then
        Dim baResult() As Byte
        ReDim baResult(0 To lLen - 1)
        WideCharToMultiByte 0, 0, StrPtr(sText), Len(sText), VarPtr(baResult(0)), lLen, 0, 0
        pvToAcpArray = baResult
    Else
        pvToAcpArray = vbNullString
    End If
End Function

Private Function pvFromAcpArray(baText() As Byte) As String
    Dim lSize As Long
    lSize = pvArraySize(baText)
    If lSize = 0 Then
        pvFromAcpArray = vbNullString
        Exit Function
    End If
    Dim lLen As Long
    lLen = MultiByteToWideChar(0, 0, VarPtr(baText(0)), lSize, 0, 0)
    If lLen > 0 Then
        pvFromAcpArray = String$(lLen, 0)
        MultiByteToWideChar 0, 0, VarPtr(baText(0)), lSize, StrPtr(pvFromAcpArray), lLen
    End If
End Function

'=========================================================================
' Private - URL parser
'=========================================================================

Private Sub pvParseUrl(sUrl As String, sScheme As String, sHost As String, lPort As Long, sPath As String)
    Dim sWork As String
    Dim lPos As Long
    sWork = sUrl
    '--- Extract scheme
    lPos = InStr(sWork, "://")
    If lPos > 0 Then
        sScheme = LCase$(Left$(sWork, lPos - 1))
        sWork = Mid$(sWork, lPos + 3)
    Else
        sScheme = "http"
    End If
    '--- Extract path
    lPos = InStr(sWork, "/")
    If lPos > 0 Then
        sPath = Mid$(sWork, lPos)
        sWork = Left$(sWork, lPos - 1)
    Else
        sPath = "/"
    End If
    '--- Extract port
    lPos = InStr(sWork, ":")
    If lPos > 0 Then
        lPort = CLng(Mid$(sWork, lPos + 1))
        sHost = Left$(sWork, lPos - 1)
    Else
        sHost = sWork
        Select Case sScheme
        Case "https", "wss": lPort = 443
        Case Else: lPort = 80
        End Select
    End If
End Sub

'=========================================================================
' Private - Post-connect actions (HTTP request / WS handshake)
'=========================================================================

Private Sub pvPostConnectActions()
    '--- Send HTTP request if pending
    If m_lHttpState = HTTP_WAITING And LenB(m_sHttpBody) <> 0 Then
        Dim sReq As String
        sReq = m_sHttpBody
        m_sHttpBody = vbNullString
        SendData sReq
        Exit Sub
    End If
    '--- Send WebSocket upgrade handshake if connecting
    If m_lWsState = WS_CONNECTING Then
        Dim sHandshake As String
        sHandshake = "GET " & m_sWsPath & " HTTP/1.1" & vbCrLf & _
            "Host: " & m_sWsHost & vbCrLf & _
            "Upgrade: websocket" & vbCrLf & _
            "Connection: Upgrade" & vbCrLf & _
            "Sec-WebSocket-Key: " & m_sWsKey & vbCrLf & _
            "Sec-WebSocket-Version: 13" & vbCrLf & vbCrLf
        SendData sHandshake
    End If
End Sub

'=========================================================================
' Private - HTTP response parser
'=========================================================================

Private Sub pvHttpProcessResponse()
    Const FUNC_NAME As String = "pvHttpProcessResponse"
    On Error GoTo EH
    '--- Append received data to raw response
    Dim sChunk As String
    sChunk = pvFromAcpArray(m_baRecvBuffer)
    m_baRecvBuffer = vbNullString
    m_sHttpRawResponse = m_sHttpRawResponse & sChunk
    '--- If headers not yet parsed, try to parse them
    If m_lHttpStatus = 0 Then
        Dim lHeaderEnd As Long
        lHeaderEnd = InStr(m_sHttpRawResponse, vbCrLf & vbCrLf)
        If lHeaderEnd = 0 Then Exit Sub  '--- Need more data
        '--- Parse status line and headers
        Dim sHeaders As String
        sHeaders = Left$(m_sHttpRawResponse, lHeaderEnd - 1)
        m_sHttpBody = Mid$(m_sHttpRawResponse, lHeaderEnd + 4)
        '--- Parse status code
        Dim lSpace As Long
        lSpace = InStr(sHeaders, " ")
        If lSpace > 0 Then
            Dim lSpace2 As Long
            lSpace2 = InStr(lSpace + 1, sHeaders, " ")
            If lSpace2 > 0 Then
                m_lHttpStatus = CLng(Mid$(sHeaders, lSpace + 1, lSpace2 - lSpace - 1))
            End If
        End If
        m_sHttpHeaders = sHeaders
        '--- Extract Content-Type
        m_sHttpContentType = pvHttpGetHeader(sHeaders, "Content-Type")
        '--- Extract Content-Length
        Dim sCL As String
        sCL = pvHttpGetHeader(sHeaders, "Content-Length")
        If LenB(sCL) <> 0 Then
            m_lHttpContentLength = CLng(sCL)
        Else
            m_lHttpContentLength = -1
        End If
        '--- Check for chunked transfer encoding
        m_bHttpChunked = (InStr(1, pvHttpGetHeader(sHeaders, "Transfer-Encoding"), "chunked", vbTextCompare) > 0)
    Else
        '--- Headers already parsed, append to body
        m_sHttpBody = m_sHttpBody & sChunk
    End If
    '--- Fire download progress
    Dim lBodyLen As Long
    lBodyLen = Len(m_sHttpBody)
    If lBodyLen > 0 Then
        Dim lTotal As Long
        If m_lHttpContentLength > 0 Then
            lTotal = m_lHttpContentLength
        Else
            lTotal = lBodyLen  '--- unknown total, report received=total
        End If
        pvFireProgress CLng(lBodyLen), lTotal
    End If
    '--- Check if body is complete
    If m_lHttpContentLength >= 0 Then
        If Len(m_sHttpBody) >= m_lHttpContentLength Then
            Dim sFinal As String
            sFinal = Left$(m_sHttpBody, m_lHttpContentLength)
            m_lHttpState = HTTP_NONE
            pvFireResponse m_lHttpStatus, m_sHttpContentType, sFinal, m_sHttpHeaders
        End If
    ElseIf m_bHttpChunked Then
        Dim sDecoded As String
        If pvHttpDecodeChunked(m_sHttpBody, sDecoded) Then
            m_lHttpState = HTTP_NONE
            pvFireResponse m_lHttpStatus, m_sHttpContentType, sDecoded, m_sHttpHeaders
        End If
    End If
    '--- No Content-Length, no chunked: pvOnClose will deliver
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvFireResponse(ByVal lStatus As Long, sContentType As String, sBody As String, sHeaders As String)
    '--- If download mode, save body to file
    If m_bDownloadMode And LenB(m_sDownloadPath) <> 0 Then
        pvWriteFileFromString m_sDownloadPath, sBody
        m_bDownloadMode = False
        '--- Fire event with file path as body
        RaiseEvent Response(lStatus, sContentType, m_sDownloadPath, sHeaders)
        m_sDownloadPath = vbNullString
    Else
        RaiseEvent Response(lStatus, sContentType, sBody, sHeaders)
    End If
End Sub

Private Sub pvFireProgress(ByVal lSent As Long, ByVal lTotal As Long)
    Dim dElapsed As Double
    Dim lSpeed As Long
    Dim lETA As Long
    dElapsed = Timer - m_dTransferStart
    '--- Handle midnight rollover
    If dElapsed < 0 Then dElapsed = dElapsed + 86400#
    '--- Calculate speed (bytes per second)
    If dElapsed > 0.01 Then
        lSpeed = CLng(CDbl(lSent) / dElapsed)
    Else
        lSpeed = 0
    End If
    '--- Calculate ETA (seconds remaining)
    If lSpeed > 0 And lTotal > lSent Then
        lETA = CLng(CDbl(lTotal - lSent) / CDbl(lSpeed))
    Else
        lETA = 0
    End If
    RaiseEvent Progress(lSent, lTotal, lSpeed, lETA)
End Sub

Private Function pvHttpGetHeader(sHeaders As String, sName As String) As String
    Dim lPos As Long
    lPos = InStr(1, sHeaders, vbCrLf & sName & ": ", vbTextCompare)
    If lPos = 0 Then lPos = InStr(1, sHeaders, vbCrLf & sName & ":", vbTextCompare)
    If lPos = 0 Then Exit Function
    lPos = InStr(lPos + 2, sHeaders, ":")
    If lPos = 0 Then Exit Function
    Dim lEnd As Long
    lEnd = InStr(lPos, sHeaders, vbCrLf)
    If lEnd = 0 Then lEnd = Len(sHeaders) + 1
    pvHttpGetHeader = Trim$(Mid$(sHeaders, lPos + 1, lEnd - lPos - 1))
End Function

Private Function pvHttpDecodeChunked(sData As String, sResult As String) As Boolean
    sResult = vbNullString
    Dim sWork As String
    sWork = sData
    Do
        Dim lCrLf As Long
        lCrLf = InStr(sWork, vbCrLf)
        If lCrLf = 0 Then Exit Function  '--- incomplete
        Dim lChunkSize As Long
        lChunkSize = CLng("&H" & Left$(sWork, lCrLf - 1))
        If lChunkSize = 0 Then
            pvHttpDecodeChunked = True
            Exit Function
        End If
        sWork = Mid$(sWork, lCrLf + 2)
        If Len(sWork) < lChunkSize + 2 Then Exit Function  '--- incomplete
        sResult = sResult & Left$(sWork, lChunkSize)
        sWork = Mid$(sWork, lChunkSize + 3)  '--- skip chunk data + CrLf
    Loop
End Function

'=========================================================================
' Private - WebSocket frame encoder / decoder
'=========================================================================

Private Sub pvWsBuildFrame(ByVal lOpcode As Long, baPayload() As Byte, baFrame() As Byte)
    Dim lPayloadLen As Long
    lPayloadLen = pvArraySize(baPayload)
    '--- Calculate frame size: 1(FIN+opcode) + 1(MASK+len) + ext_len + 4(mask) + payload
    Dim lHeaderLen As Long
    If lPayloadLen <= 125 Then
        lHeaderLen = 2
    ElseIf lPayloadLen <= 65535 Then
        lHeaderLen = 4
    Else
        lHeaderLen = 10
    End If
    ReDim baFrame(0 To lHeaderLen + 4 + lPayloadLen - 1)
    '--- Byte 0: FIN=1 + opcode
    baFrame(0) = CByte(&H80 Or (lOpcode And &HF))
    '--- Byte 1: MASK=1 + payload length
    If lPayloadLen <= 125 Then
        baFrame(1) = CByte(&H80 Or lPayloadLen)
    ElseIf lPayloadLen <= 65535 Then
        baFrame(1) = CByte(&H80 Or 126)
        baFrame(2) = CByte((lPayloadLen \ 256) And &HFF)
        baFrame(3) = CByte(lPayloadLen And &HFF)
    Else
        baFrame(1) = CByte(&H80 Or 127)
        '--- 8-byte length (big-endian, high 4 bytes = 0)
        baFrame(2) = 0: baFrame(3) = 0: baFrame(4) = 0: baFrame(5) = 0
        baFrame(6) = CByte((lPayloadLen \ &H1000000) And &HFF)
        baFrame(7) = CByte((lPayloadLen \ &H10000) And &HFF)
        baFrame(8) = CByte((lPayloadLen \ &H100) And &HFF)
        baFrame(9) = CByte(lPayloadLen And &HFF)
    End If
    '--- Generate masking key
    Dim baMask(0 To 3) As Byte
    pvRandomBytes baMask(0), 4
    Dim lMaskOff As Long
    lMaskOff = lHeaderLen
    baFrame(lMaskOff) = baMask(0)
    baFrame(lMaskOff + 1) = baMask(1)
    baFrame(lMaskOff + 2) = baMask(2)
    baFrame(lMaskOff + 3) = baMask(3)
    '--- Copy and mask payload
    Dim i As Long
    For i = 0 To lPayloadLen - 1
        baFrame(lMaskOff + 4 + i) = baPayload(i) Xor baMask(i Mod 4)
    Next i
End Sub

Private Sub pvWsProcessHandshakeResponse()
    Const FUNC_NAME As String = "pvWsProcessHandshakeResponse"
    On Error GoTo EH
    '--- Accumulate response
    Dim sChunk As String
    sChunk = pvFromAcpArray(m_baRecvBuffer)
    m_baRecvBuffer = vbNullString
    m_sHttpRawResponse = m_sHttpRawResponse & sChunk
    '--- Check for complete headers
    Dim lEnd As Long
    lEnd = InStr(m_sHttpRawResponse, vbCrLf & vbCrLf)
    If lEnd = 0 Then Exit Sub   '--- Need more data
    '--- Verify 101 Switching Protocols
    If InStr(1, m_sHttpRawResponse, "101", vbTextCompare) = 0 Then
        m_lWsState = WS_NONE
        pvSetError LastDllError:=sckConnectionRefused
        Exit Sub
    End If
    '--- Verify Sec-WebSocket-Accept
    Dim sExpect As String
    sExpect = pvComputeWsAccept(m_sWsKey)
    If InStr(1, m_sHttpRawResponse, sExpect, vbTextCompare) = 0 Then
        m_lWsState = WS_NONE
        pvSetError LastDllError:=sckConnectionRefused
        Exit Sub
    End If
    '--- WebSocket is open!
    m_lWsState = WS_OPEN
    '--- Extract remaining data after headers BEFORE clearing
    Dim sRemain As String
    sRemain = Mid$(m_sHttpRawResponse, lEnd + 4)
    m_sHttpRawResponse = vbNullString
    If LenB(sRemain) <> 0 Then
        m_baWsBuffer = pvToAcpArray(sRemain)
    End If
    RaiseEvent Connect
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub pvWsProcessFrames()
    Const FUNC_NAME As String = "pvWsProcessFrames"
    On Error GoTo EH
    '--- Append new data to WS buffer
    pvAppendBuffer m_baWsBuffer, m_baRecvBuffer
    m_baRecvBuffer = vbNullString
    '--- Try to parse frames
    Do While pvArraySize(m_baWsBuffer) >= 2
        Dim lBufSize As Long
        lBufSize = pvArraySize(m_baWsBuffer)
        '--- Parse header
        Dim bFin As Boolean
        Dim lOpcode As Long
        Dim bMask As Boolean
        Dim lPayloadLen As Long
        Dim lHeaderLen As Long
        bFin = (m_baWsBuffer(0) And &H80) <> 0
        lOpcode = m_baWsBuffer(0) And &HF
        bMask = (m_baWsBuffer(1) And &H80) <> 0
        lPayloadLen = m_baWsBuffer(1) And &H7F
        lHeaderLen = 2
        If lPayloadLen = 126 Then
            If lBufSize < 4 Then Exit Do  '--- Need more
            lPayloadLen = CLng(m_baWsBuffer(2)) * 256 + CLng(m_baWsBuffer(3))
            lHeaderLen = 4
        ElseIf lPayloadLen = 127 Then
            If lBufSize < 10 Then Exit Do
            '--- Read 64-bit length (use lower 32 bits)
            lPayloadLen = CLng(m_baWsBuffer(6)) * &H1000000 + CLng(m_baWsBuffer(7)) * &H10000 + _
                           CLng(m_baWsBuffer(8)) * &H100& + CLng(m_baWsBuffer(9))
            lHeaderLen = 10
        End If
        '--- Mask key (if present)
        Dim lMaskKeyOff As Long
        lMaskKeyOff = lHeaderLen
        If bMask Then lHeaderLen = lHeaderLen + 4
        '--- Check if full frame is available
        Dim lFrameSize As Long
        lFrameSize = lHeaderLen + lPayloadLen
        If lBufSize < lFrameSize Then Exit Do  '--- Need more data
        '--- Extract payload
        Dim baPayload() As Byte
        If lPayloadLen > 0 Then
            ReDim baPayload(0 To lPayloadLen - 1)
            CopyMemory baPayload(0), m_baWsBuffer(lHeaderLen), lPayloadLen
            '--- Unmask if needed
            If bMask Then
                Dim i As Long
                For i = 0 To lPayloadLen - 1
                    baPayload(i) = baPayload(i) Xor m_baWsBuffer(lMaskKeyOff + (i Mod 4))
                Next i
            End If
        Else
            baPayload = vbNullString
        End If
        '--- Remove consumed frame from buffer
        If lBufSize > lFrameSize Then
            Dim baRemain() As Byte
            ReDim baRemain(0 To lBufSize - lFrameSize - 1)
            CopyMemory baRemain(0), m_baWsBuffer(lFrameSize), lBufSize - lFrameSize
            m_baWsBuffer = baRemain
        Else
            m_baWsBuffer = vbNullString
        End If
        '--- Dispatch based on opcode
        Select Case lOpcode
        Case WS_OPCODE_TEXT
            RaiseEvent Receive(pvFromUtf8Array(baPayload), False)
        Case WS_OPCODE_BINARY
            RaiseEvent Receive(pvFromAcpArray(baPayload), True)
        Case WS_OPCODE_PING
            '--- Auto-pong
            Dim baPong() As Byte
            pvWsBuildFrame WS_OPCODE_PONG, baPayload, baPong
            If m_lTlsState = TLS_CONNECTED Then
                Dim baEnc() As Byte
                pvTlsEncryptData baPong, baEnc
                pvAppendBuffer m_baSendBuffer, baEnc
            Else
                pvAppendBuffer m_baSendBuffer, baPong
            End If
            m_lSendPos = 0
            pvFlushSendBuffer
        Case WS_OPCODE_PONG
            '--- Ignore pong
        Case WS_OPCODE_CLOSE
            Dim lCode As Long
            Dim sReason As String
            If pvArraySize(baPayload) >= 2 Then
                lCode = CLng(baPayload(0)) * 256 + CLng(baPayload(1))
                If pvArraySize(baPayload) > 2 Then
                    Dim baR() As Byte
                    ReDim baR(0 To pvArraySize(baPayload) - 3)
                    CopyMemory baR(0), baPayload(2), pvArraySize(baPayload) - 2
                    sReason = pvFromUtf8Array(baR)
                End If
            End If
            m_lWsState = WS_NONE
            RaiseEvent Disconnected(lCode, sReason)
            Disconnect
            Exit Do
        End Select
    Loop
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Private - Crypto helpers (SHA-1, Base64, Random)
'=========================================================================

Private Function pvGenerateWsKey() As String
    Dim baKey(0 To 15) As Byte
    pvRandomBytes baKey(0), 16
    pvGenerateWsKey = pvBase64Encode(baKey)
End Function

Private Function pvComputeWsAccept(sKey As String) As String
    Dim baInput() As Byte
    baInput = pvToAcpArray(sKey & WS_MAGIC_GUID)
    Dim baHash() As Byte
    baHash = pvSHA1(baInput)
    pvComputeWsAccept = pvBase64Encode(baHash)
End Function

Private Function pvSHA1(baData() As Byte) As Byte()
    Dim hProv As Long, hHash As Long
    Dim baResult(0 To 19) As Byte
    If CryptAcquireContext(hProv, 0&, 0&, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hProv, CALG_SHA1, 0, 0, hHash) <> 0 Then
            CryptHashData hHash, baData(0), pvArraySize(baData), 0
            Dim lLen As Long
            lLen = 20
            CryptGetHashParam hHash, HP_HASHVAL, baResult(0), lLen, 0
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
    pvSHA1 = baResult
End Function

Private Function pvBase64Encode(baData() As Byte) As String
    Dim lLen As Long
    Dim lSize As Long
    lSize = pvArraySize(baData)
    If lSize = 0 Then Exit Function
    '--- Get required length
    CryptBinaryToStringA VarPtr(baData(0)), lSize, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, 0, lLen
    If lLen > 0 Then
        Dim sAnsi As String
        sAnsi = String$(lLen, 0)
        CryptBinaryToStringA VarPtr(baData(0)), lSize, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, StrPtr(sAnsi), lLen
        '--- sAnsi is ANSI in VB Unicode string buffer, convert
        Dim baAnsi() As Byte
        ReDim baAnsi(0 To lLen - 1)
        CopyMemory baAnsi(0), ByVal StrPtr(sAnsi), lLen
        pvBase64Encode = StrConv(baAnsi, vbUnicode)
        '--- Trim null/whitespace
        Do While Right$(pvBase64Encode, 1) = vbNullChar Or Right$(pvBase64Encode, 1) = vbCr Or Right$(pvBase64Encode, 1) = vbLf
            pvBase64Encode = Left$(pvBase64Encode, Len(pvBase64Encode) - 1)
        Loop
    End If
End Function

Private Sub pvRandomBytes(ByRef bBuf As Byte, ByVal lCount As Long)
    Dim hProv As Long
    If CryptAcquireContext(hProv, 0&, 0&, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        CryptGenRandom hProv, lCount, bBuf
        CryptReleaseContext hProv, 0
    Else
        '--- Fallback to VB Rnd
        Randomize Timer
        Dim i As Long
        Dim pBuf As Long
        pBuf = VarPtr(bBuf)
        For i = 0 To lCount - 1
            Dim bVal As Byte
            bVal = CByte(Int(Rnd * 256))
            CopyMemory ByVal pBuf + i, bVal, 1
        Next i
    End If
End Sub

'=========================================================================
' Private - UTF-8 conversion helpers
'=========================================================================

Private Function pvToUtf8Array(sText As String) As Byte()
    If LenB(sText) = 0 Then
        pvToUtf8Array = vbNullString
        Exit Function
    End If
    Dim lLen As Long
    lLen = WideCharToMultiByte(65001, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lLen > 0 Then
        Dim baResult() As Byte
        ReDim baResult(0 To lLen - 1)
        WideCharToMultiByte 65001, 0, StrPtr(sText), Len(sText), VarPtr(baResult(0)), lLen, 0, 0
        pvToUtf8Array = baResult
    Else
        pvToUtf8Array = vbNullString
    End If
End Function

Private Function pvFromUtf8Array(baText() As Byte) As String
    Dim lSize As Long
    lSize = pvArraySize(baText)
    If lSize = 0 Then
        pvFromUtf8Array = vbNullString
        Exit Function
    End If
    Dim lLen As Long
    lLen = MultiByteToWideChar(65001, 0, VarPtr(baText(0)), lSize, 0, 0)
    If lLen > 0 Then
        pvFromUtf8Array = String$(lLen, 0)
        MultiByteToWideChar 65001, 0, VarPtr(baText(0)), lSize, StrPtr(pvFromUtf8Array), lLen
    End If
End Function

'=========================================================================
' Private - File I/O helpers
'=========================================================================

Private Sub pvReadFile(sPath As String, baData() As Byte, lSize As Long)
    On Error GoTo EH
    lSize = 0
    baData = vbNullString
    If Dir$(sPath) = vbNullString Then Exit Sub
    lSize = FileLen(sPath)
    If lSize = 0 Then Exit Sub
    Dim hFile As Integer
    hFile = FreeFile
    Open sPath For Binary Access Read As #hFile
    ReDim baData(0 To lSize - 1)
    Get #hFile, , baData
    Close #hFile
    Exit Sub
EH:
    lSize = 0
    baData = vbNullString
End Sub

Private Sub pvWriteFile(sPath As String, baData() As Byte)
    On Error GoTo EH
    Dim lSize As Long
    lSize = pvArraySize(baData)
    If lSize = 0 Then Exit Sub
    Dim hFile As Integer
    hFile = FreeFile
    Open sPath For Binary Access Write As #hFile
    Put #hFile, , baData
    Close #hFile
    Exit Sub
EH:
    Debug.Print "pvWriteFile error: " & Err.Description
End Sub

Private Sub pvWriteFileFromString(sPath As String, sData As String)
    On Error GoTo EH
    If LenB(sData) = 0 Then Exit Sub
    Dim baData() As Byte
    baData = pvToAcpArray(sData)
    pvWriteFile sPath, baData
    Exit Sub
EH:
    Debug.Print "pvWriteFileFromString error: " & Err.Description
End Sub

Private Sub pvSendNextFileChunk()
    Const FUNC_NAME As String = "pvSendNextFileChunk"
    On Error GoTo EH
    If Not m_bSendingFile Then Exit Sub
    '--- Check if all data sent
    If m_lSendFileOffset >= m_lSendFileSize Then
        m_bSendingFile = False
        pvFireProgress m_lSendFileSize, m_lSendFileSize
        RaiseEvent SendComplete
        Exit Sub
    End If
    '--- Read next chunk from file
    Dim lToRead As Long
    lToRead = m_lSendFileChunk
    If m_lSendFileOffset + lToRead > m_lSendFileSize Then
        lToRead = m_lSendFileSize - m_lSendFileOffset
    End If
    Dim baChunk() As Byte
    ReDim baChunk(0 To lToRead - 1)
    Dim hFile As Integer
    hFile = FreeFile
    Open m_sSendFilePath For Binary Access Read As #hFile
    Seek #hFile, m_lSendFileOffset + 1  '--- VB file positions are 1-based
    Get #hFile, , baChunk
    Close #hFile
    '--- Send based on current mode
    If m_lWsState = WS_OPEN Then
        '--- WebSocket mode: wrap chunk in binary frame
        Dim baFrame() As Byte
        pvWsBuildFrame WS_OPCODE_BINARY, baChunk, baFrame
        If m_lTlsState = TLS_CONNECTED Then
            Dim baEncWs() As Byte
            pvTlsEncryptData baFrame, baEncWs
            pvAppendBuffer m_baSendBuffer, baEncWs
        Else
            pvAppendBuffer m_baSendBuffer, baFrame
        End If
    Else
        '--- Raw TCP/TLS mode: send bytes directly
        If m_lTlsState = TLS_CONNECTED Then
            Dim baEnc() As Byte
            pvTlsEncryptData baChunk, baEnc
            pvAppendBuffer m_baSendBuffer, baEnc
        Else
            pvAppendBuffer m_baSendBuffer, baChunk
        End If
    End If
    m_lSendPos = 0
    m_lSendFileOffset = m_lSendFileOffset + lToRead
    pvFlushSendBuffer
    '--- Fire progress
    pvFireProgress m_lSendFileOffset, m_lSendFileSize
    Exit Sub
EH:
    m_bSendingFile = False
    PrintError FUNC_NAME
End Sub

'=========================================================================
' UserControl events
'=========================================================================

Private Sub UserControl_Initialize()
    Const FUNC_NAME As String = "UserControl_Initialize"
    On Error GoTo EH
    m_hSocket = INVALID_SOCKET
    m_baRecvBuffer = vbNullString
    m_baSendBuffer = vbNullString
    m_baTlsPending = vbNullString
    m_baWsBuffer = vbNullString
    '--- Initialize WinSock
    Dim baWSAData(0 To 399) As Byte
    If WSAStartup(&H202, baWSAData(0)) = 0 Then
        m_bWsaInitialized = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_Terminate()
    Const FUNC_NAME As String = "UserControl_Terminate"
    On Error GoTo EH
    Disconnect
    If m_bWsaInitialized Then
        WSACleanup
        m_bWsaInitialized = False
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Size ScaleX(32, vbPixels, vbTwips), ScaleY(32, vbPixels, vbTwips)
    labLogo.Move 0, (ScaleHeight - labLogo.Height) / 2, ScaleWidth
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME As String = "UserControl_InitProperties"
    On Error GoTo EH
    labLogo.Caption = STR_LOGO
    m_lLocalPort = DEF_LOCALPORT
    m_eProtocol = DEF_PROTOCOL
    m_sRemoteHost = DEF_REMOTEHOST
    m_lRemotePort = DEF_REMOTEPORT
    m_lTimeout = DEF_TIMEOUT
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME As String = "UserControl_ReadProperties"
    On Error GoTo EH
    labLogo.Caption = STR_LOGO
    With PropBag
        m_lLocalPort = .ReadProperty("LocalPort", DEF_LOCALPORT)
        m_eProtocol = .ReadProperty("Protocol", DEF_PROTOCOL)
        m_sRemoteHost = .ReadProperty("RemoteHost", DEF_REMOTEHOST)
        m_lRemotePort = .ReadProperty("RemotePort", DEF_REMOTEPORT)
        m_lTimeout = .ReadProperty("Timeout", DEF_TIMEOUT)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME As String = "UserControl_WriteProperties"
    On Error GoTo EH
    With PropBag
        .WriteProperty "LocalPort", m_lLocalPort, DEF_LOCALPORT
        .WriteProperty "Protocol", m_eProtocol, DEF_PROTOCOL
        .WriteProperty "RemoteHost", m_sRemoteHost, DEF_REMOTEHOST
        .WriteProperty "RemotePort", m_lRemotePort, DEF_REMOTEPORT
        .WriteProperty "Timeout", m_lTimeout, DEF_TIMEOUT
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
