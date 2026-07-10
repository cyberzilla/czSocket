# czSocket

**Single-file VB6 UserControl for TCP, TLS/HTTPS, HTTP, and WebSocket (WS/WSS).**

Zero external dependencies. Drop one `.ctl` file into your project — done.

---

## Features

- **TCP/UDP** — Client and server with async I/O
- **TLS/HTTPS** — Native Windows Schannel (SSPI), TLS 1.0–1.3, client and server
- **HTTP** — GET, POST, PUT, DELETE with response parsing (chunked + Content-Length)
- **WebSocket** — Full RFC 6455: text, binary, ping/pong, close handshake
- **File Transfer** — Streaming upload (multipart, 64KB chunks), streaming download (save to disk), chunked send
- **Resume Support** — Pause/resume downloads with HTTP Range header
- **Progress Tracking** — Real-time speed (bytes/sec) and ETA for all transfers
- **Auto-Retry** — Configurable retry on connection failures
- **Universal API** — One set of methods auto-detects protocol from URL scheme
- **Zero Dependencies** — Uses only built-in Windows DLLs
- **Minimum OS** — Windows 7

## Quick Start

### 1. Add to your VB6 project

Copy `czSocket.ctl` into your project folder, then in VB6:

> **Project → Add User Control → Existing → czSocket.ctl**

Place `czSocket1` on your form. That's it.

### 2. Use it

```vb
' TCP
czSocket1.Connect "example.com", 80

' HTTPS (auto-TLS)
czSocket1.Request "GET", "https://jsonplaceholder.typicode.com/posts/1"

' WebSocket (auto from URL)
czSocket1.Connect "wss://ws.postman-echo.com/raw"
```

---

## Universal API

czSocket auto-detects the protocol from the URL scheme. You don't need to remember which method to call — the same methods work for all protocols.

### Methods

| Method | Description |
|--------|-------------|
| `Connect host, [port]` | Connect to host. Auto-detects `ws://`, `wss://`, `http://`, `https://` |
| `SendData data` | Send string or byte array. Auto: WS frame or raw TCP |
| `SendFile path, [chunkSize]` | Send file from disk in chunks. Auto: WS binary frames or raw TCP |
| `Download url, savePath, [headers], [resumeFrom]` | Download file via HTTP/HTTPS. Optional resume from byte offset |
| `Upload url, filePath, [field]` | Stream upload file via HTTP POST multipart (64KB chunks, non-blocking) |
| `Request method, url, [body]` | Send HTTP/HTTPS request (GET, POST, PUT, etc.) |
| `Ping [data]` | Send WebSocket ping frame |
| `Disconnect` | Close connection. Auto: WS graceful close + TCP close |
| `Listen` | Start listening for incoming connections (server) |
| `Accept requestID` | Accept incoming connection |
| `Bind [port], [ip]` | Bind to local port |
| `GetData data, [type], [maxLen]` | Retrieve received data |
| `PeekData data, [type], [maxLen]` | Peek at received data without consuming |

### Events

| Event | When |
|-------|------|
| `Connect()` | Connection established (TCP, TLS handshake, or WS handshake) |
| `DataArrival(bytesTotal)` | Raw data available (TCP/UDP mode) |
| `Response(Status, ContentType, Body, Headers)` | HTTP response received. In download mode, `Body` = saved file path |
| `Progress(BytesSent, BytesTotal, BytesPerSec, SecondsRemaining)` | Transfer progress with speed and ETA |
| `Receive(Data, IsBinary)` | WebSocket message received |
| `Disconnected(Code, Reason)` | Connection closed (TCP: Code=0, WS: close code) |
| `ConnectionRequest(requestID)` | Incoming connection on server |
| `SendProgress(bytesSent, bytesRemaining)` | Buffer send progress |
| `SendComplete()` | All queued/chunked data sent |
| `Error(Number, Description, ...)` | Socket or TLS error |

### Properties

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `RemoteHost` | String | R/W | Remote host name or IP |
| `RemotePort` | Long | R/W | Remote port number |
| `RemoteHostIP` | String | R | Resolved remote IP address |
| `LocalPort` | Long | R/W | Local port for Bind/Listen |
| `LocalHostName` | String | R | Local machine name |
| `LocalIP` | String | R | Local IP address |
| `Protocol` | Enum | R/W | `sckTCPProtocol` / `sckUDPProtocol` / `sckTLSProtocol` |
| `State` | Enum | R | Current socket state |
| `SocketHandle` | Long | R | Raw WinSock handle |
| `Timeout` | Long | R/W | Connection timeout (ms) |
| `StatusCode` | Long | R | Last HTTP response status code |
| `ReadyState` | Long | R | WebSocket state (0=None, 1=Connecting, 2=Open, 3=Closing) |
| `SockOpt` | Long | R/W | Raw socket options |

---

## Examples

### Raw TCP (Classic Winsock Replacement)

```vb
Private Sub Form_Load()
    czSocket1.Connect "example.com", 80
End Sub

Private Sub czSocket1_Connect()
    czSocket1.SendData "GET / HTTP/1.1" & vbCrLf & _
        "Host: example.com" & vbCrLf & _
        "Connection: close" & vbCrLf & vbCrLf
End Sub

Private Sub czSocket1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    czSocket1.GetData s
    txtOutput.Text = txtOutput.Text & s
End Sub
```

### HTTPS JSON API

```vb
Private Sub btnFetch_Click()
    czSocket1.Request "GET", "https://jsonplaceholder.typicode.com/posts/1"
End Sub

Private Sub czSocket1_Response(ByVal Status As Long, ByVal ContentType As String, _
    Body As String, Headers As String)
    Debug.Print "Status: " & Status
    Debug.Print "Body: " & Body
End Sub
```

### WebSocket

```vb
Private Sub Form_Load()
    ' Connect auto-detects wss:// → WebSocket + TLS
    czSocket1.Connect "wss://ws.postman-echo.com/raw"
End Sub

Private Sub czSocket1_Connect()
    ' SendData auto-detects WebSocket → sends as WS text frame
    czSocket1.SendData "Hello, WebSocket!"
End Sub

Private Sub czSocket1_Receive(ByVal Data As String, ByVal IsBinary As Boolean)
    MsgBox "Received: " & Data
End Sub

Private Sub czSocket1_Disconnected(ByVal Code As Long, ByVal Reason As String)
    Debug.Print "Disconnected: " & Code & " " & Reason
End Sub
```

### Web3 / JSON-RPC

```vb
' Via HTTPS
Private Sub btnGetBlock_Click()
    Dim body As String
    body = "{""jsonrpc"":""2.0"",""method"":""eth_blockNumber"",""params"":[],""id"":1}"
    czSocket1.Request "POST", "https://mainnet.infura.io/v3/YOUR_KEY", body
End Sub

' Via WebSocket (real-time subscription)
Private Sub btnSubscribe_Click()
    czSocket1.Connect "wss://mainnet.infura.io/ws/v3/YOUR_KEY"
End Sub

Private Sub czSocket1_Connect()
    czSocket1.SendData "{""jsonrpc"":""2.0"",""method"":""eth_subscribe""," & _
        """params"":[""newHeads""],""id"":1}"
End Sub

Private Sub czSocket1_Receive(ByVal Data As String, ByVal IsBinary As Boolean)
    Debug.Print "New block: " & Data
End Sub
```

### Download File with Progress

```vb
Private Sub btnDownload_Click()
    czSocket1.Download "https://example.com/largefile.zip", "C:\Downloads\largefile.zip"
End Sub

Private Sub czSocket1_Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, _
    ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    Dim pct As Long
    If BytesTotal > 0 Then pct = (BytesSent * 100) / BytesTotal
    ProgressBar1.Value = pct
    
    Dim speed As String
    If BytesPerSec > 1048576 Then
        speed = Format$(BytesPerSec / 1048576, "0.0") & " MB/s"
    ElseIf BytesPerSec > 1024 Then
        speed = Format$(BytesPerSec / 1024, "0.0") & " KB/s"
    Else
        speed = BytesPerSec & " B/s"
    End If
    
    lblStatus.Caption = pct & "% | " & speed & " | ETA: " & SecondsRemaining & "s"
End Sub

Private Sub czSocket1_Response(ByVal Status As Long, ByVal ContentType As String, _
    Body As String, Headers As String)
    MsgBox "Download complete: " & Body  ' Body = saved file path
End Sub
```

### Resume Download (Pause/Resume)

```vb
Dim lBytesDownloaded As Long

' Pause: disconnect and remember bytes
Private Sub btnPause_Click()
    lBytesDownloaded = m_lCurrentBytes
    czSocket1.Disconnect
End Sub

' Resume: use ResumeFrom parameter with Range header
Private Sub btnResume_Click()
    czSocket1.Download "https://example.com/largefile.zip", _
                       "C:\Downloads\largefile.zip", , lBytesDownloaded
End Sub
```

### Upload File

```vb
Private Sub btnUpload_Click()
    czSocket1.Upload "https://api.example.com/upload", "C:\Photos\photo.jpg"
End Sub

Private Sub czSocket1_Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, _
    ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    lblStatus.Caption = "Uploading: " & Format$(BytesSent / BytesTotal * 100, "0.0") & _
        "%  " & Format$(BytesPerSec / 1048576, "0.0") & " MB/s"
End Sub

Private Sub czSocket1_Response(ByVal Status As Long, ByVal ContentType As String, _
    Body As String, Headers As String)
    If Status = 200 Then
        MsgBox "Upload successful!"
    Else
        MsgBox "Upload failed: " & Status
    End If
End Sub
```

### Send Large File via WebSocket

```vb
Private Sub btnSend_Click()
    czSocket1.Connect "wss://fileserver.example.com/upload"
End Sub

Private Sub czSocket1_Connect()
    ' SendFile auto-detects WS mode → sends as binary frames
    czSocket1.SendFile "C:\Videos\video.mp4", 65536  ' 64KB chunks
End Sub

Private Sub czSocket1_Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, _
    ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    lblProgress.Caption = Format$(BytesSent / 1048576, "0.0") & " / " & _
        Format$(BytesTotal / 1048576, "0.0") & " MB"
End Sub

Private Sub czSocket1_SendComplete()
    MsgBox "File sent!"
End Sub
```

### TCP Server

```vb
Private Sub Form_Load()
    czSocket1.LocalPort = 8080
    czSocket1.Listen
End Sub

Private Sub czSocket1_ConnectionRequest(ByVal requestID As Long)
    czSocket1.Disconnect
    czSocket1.Accept requestID
End Sub

Private Sub czSocket1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    czSocket1.GetData s
    czSocket1.SendData "Echo: " & s
End Sub
```

### HTTPS Server (TLS)

```vb
' Uses two czSocket controls: czServer (listener) + czClient (accepted)
Private Sub Form_Load()
    czServer.LocalPort = 8443
    czServer.Listen App.Path & "\server.pfx", "password"
End Sub

Private Sub czServer_ConnectionRequest(ByVal requestID As Long)
    czClient.Disconnect
    czClient.Accept requestID  ' TLS handshake is automatic
End Sub

Private Sub czClient_Connect()
    ' TLS handshake complete — client is ready
End Sub

Private Sub czClient_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    czClient.GetData s
    ' Send HTTP response
    Dim resp As String
    resp = "HTTP/1.1 200 OK" & vbCrLf & _
           "Content-Type: text/plain" & vbCrLf & _
           "Content-Length: 13" & vbCrLf & _
           "Connection: close" & vbCrLf & vbCrLf & _
           "Hello, World!"
    czClient.SendData resp
End Sub

Private Sub czClient_SendComplete()
    czClient.Disconnect
End Sub
```

---

## Auto-Detection Logic

```
URL Scheme          →  Protocol Mode
─────────────────────────────────────
wss://...           →  WebSocket + TLS
ws://...            →  WebSocket
https://...         →  TCP + TLS
http://...          →  TCP
host, port          →  Raw TCP (backward compatible)

Method              →  Behavior When WS is Open
─────────────────────────────────────
SendData "text"     →  WS text frame
SendData bytes()    →  WS binary frame
SendFile "path"     →  WS binary frames (chunked)
Disconnect          →  WS close frame + TCP close
```

## Architecture

```
┌─────────────────────────────────────────┐
│           czSocket.ctl (Single File)    │
├─────────────────────────────────────────┤
│  Universal API Layer                    │
│  Connect / SendData / SendFile /        │
│  Download / Upload / Request / Disconnect   │
├──────────┬──────────┬───────────────────┤
│  HTTP    │ WebSocket│  Raw TCP/UDP      │
│  Engine  │  Engine  │  Engine           │
├──────────┴──────────┴───────────────────┤
│  TLS Engine (SSPI/Schannel)             │
├─────────────────────────────────────────┤
│  WinSock Engine (ws2_32.dll)            │
│  Async I/O via WSAEventSelect + Timer   │
├─────────────────────────────────────────┤
│  Windows DLLs (built-in, zero install)  │
│  ws2_32 · secur32 · advapi32 · crypt32 │
└─────────────────────────────────────────┘
```

## Requirements

| Requirement | Value |
|-------------|-------|
| IDE | Visual Basic 6.0 SP6 |
| OS | Windows 7 or later |
| Dependencies | None (zero external DLLs) |
| File | Single `czSocket.ctl` (~130 KB) |

## Performance

All file transfers use streaming I/O with 64KB chunks. Files are never loaded entirely into memory.

## Changelog
### v1.2

#### Performance
- **Fix**: O(n²) string accumulation in `pvHttpProcessResponse` — download now streams bytes directly to file via `Put #`, bypassing `m_sHttpRawResponse` accumulation after headers are parsed
- **Fix**: Single `recv` per timer tick — `pvOnReceive` now loops until `WSAEWOULDBLOCK` to drain all available data, increasing download throughput from ~3 MB/s to 175+ MB/s
- **Fix**: Upload loaded entire file into memory — `Upload` now streams file in 64KB chunks via `pvSendNextFileChunk` iterative loop
- **Fix**: `pvSendNextFileChunk` only sent one chunk per `FD_WRITE` event — now loops until `WSAEWOULDBLOCK` with periodic `DoEvents` for UI responsiveness

#### Download
- **New**: Resume support — `Download` accepts optional `ResumeFrom` parameter, sends `Range: bytes=N-` header, seeks file to correct position, handles `206 Partial Content` response
- **New**: Streaming file write — received bytes written directly to disk via file handle, never accumulated in memory

#### Upload  
- **New**: Streaming upload — HTTP headers + multipart prefix sent first, then file streamed in 64KB chunks, then multipart suffix. Replaces old approach of loading entire file into `baBody` array
- **New**: Upload progress — `Progress` event fires during upload with bytes sent, total, speed, and ETA
- **Fix**: State setup ordering — upload state set after `Connect` call (which internally calls `Disconnect` and resets state)

#### Demo: Parallel Downloader & Uploader
- **New**: Pause/Resume buttons with click-to-select on progress bar
- **New**: Stop/Cancel button for active, paused, and queued downloads
- **New**: IDM-style block mosaic progress visualization (40 blocks per file)
- **New**: Auto-retry (2 attempts) on connection failure
- **New**: Upload progress display with percentage, speed, and ETA
- **New**: Dynamic button states (Pause ↔ Resume) based on selected download
- **New**: Status bar shows Active / Queued / Done / Paused / Errors counts
- **New**: Total download time logging per file
- **Fix**: DrawProgress throttled to ~10 fps to prevent CPU saturation
- **Fix**: Unicode characters (`Chr$(&H25BA)`) causing Runtime Error 5 in compiled EXE — replaced with ASCII-safe indicators

### v1.1

- **Fix**: `Err.Raise 0` crash in compiled EXE — error handler now guards against empty error codes
- **Fix**: `Request` method (HTTP/HTTPS) not sending data — state setup moved after async `Connect` call
- **Fix**: WebSocket handshake not sent — state setup ordering corrected  
- **Fix**: HTTPS response silently dropped when TLS close arrives in same packet as response data
- **Fix**: Double `Disconnected` event firing — guard against redundant `pvOnClose` calls
- **Fix**: Double `Connect` event for WebSocket — suppress TLS-level connect when WS handshake pending
- **Fix**: Stale HTTP response data contaminating WebSocket handshake parser — proper state cleanup in `Disconnect`
- **New**: TLS Server — `Listen` now accepts PFX certificate for HTTPS/WSS server support
- **New**: `pvLoadPfxCertificate` — loads PFX/P12 certificate files via Windows CryptoAPI
- **New**: Server-side TLS handshake using `AcceptSecurityContext` (SSPI/Schannel)
- **New**: WebServer demo in `Demo/WebServer/` with self-signed certificate generator
- **Improved**: TLS feature description updated to include TLS 1.3 support
- **Improved**: HTTP response parsing description updated (chunked + Content-Length)

### v1.0

- Initial release

## License

MIT License — see source file header for details.

Inspired by [VbAsyncSocket](https://github.com/wqweto/VbAsyncSocket) by wqweto@gmail.com.
