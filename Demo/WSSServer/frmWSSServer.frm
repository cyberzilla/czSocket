VERSION 5.00
Begin VB.Form frmWSSServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czSocket Demo - WSS Server (WebSocket Secure)"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin czWSSServer.czSocket czServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin czWSSServer.czSocket czClient 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame fraServer 
      Caption         =   " WSS Server "
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9360
      Begin VB.CommandButton btnStart 
         Caption         =   "Start WSS Server"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnStop 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnBroadcast 
         Caption         =   "Broadcast Message"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtBroadcast 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Hello from VB6 WSS Server!"
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label lblStatus 
         Caption         =   "Server stopped"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lblClients 
         Caption         =   "WS Clients: 0"
         Height          =   255
         Left            =   6960
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   6375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1560
      Width           =   9360
   End
End
Attribute VB_Name = "frmWSSServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' czSocket Demo - WSS (WebSocket Secure) Multi-Client Server
'
' This demo shows:
'   1. Listen with TLS certificate (WSS = WebSocket over TLS)
'   2. Accept MULTIPLE connections via control array
'   3. Detect WebSocket upgrade request per client
'   4. Call AcceptWebSocket to complete WS handshake
'   5. Send/receive WebSocket text frames
'   6. Broadcast messages to ALL connected WebSocket clients
'
' Architecture:
'   czServer         = Listener (1 instance)
'   czClient(0..N)   = Control array, one per client connection
'   m_bWsActive(i)   = Tracks which clients are in WS mode
'=========================================================================

Private m_lNextIndex As Long        '--- Next control array index
Private m_bWsActive() As Boolean    '--- Which clients are WS-connected
Private m_lWsCount As Long          '--- Active WS client count
Private m_bStopping As Boolean      '--- True during server shutdown

Private Sub Form_Load()
    m_lNextIndex = 1
    ReDim m_bWsActive(0 To 0)
    Log "=== czSocket WSS Multi-Client Server ==="
    Log "1. Click 'Start WSS Server' to begin."
    Log "2. Open browser to https://localhost:8443/"
    Log "3. Click 'Connect to WSS Server' on the page."
    Log "4. Open MORE browser tabs for multiple clients!"
    Log "5. Use 'Broadcast' to send to ALL clients at once."
    Log ""
End Sub

'=========================================================================
' Server Controls
'=========================================================================

Private Sub btnStart_Click()
    Dim sCertFile As String
    On Error GoTo EH
    sCertFile = App.Path & "\localhost.pfx"
    If Dir$(sCertFile) = "" Then
        MsgBox "Certificate not found: " & sCertFile, vbExclamation
        Exit Sub
    End If
    czServer.LocalPort = 8443
    czServer.Listen sCertFile, "czSocket123"
    btnStart.Enabled = False
    btnStop.Enabled = True
    btnBroadcast.Enabled = True
    m_lWsCount = 0
    lblStatus.Caption = "Listening :8443"
    lblStatus.ForeColor = &H8000&
    Log "* WSS Server started on wss://localhost:8443/"
    Log "* Waiting for WebSocket connections..."
    Log ""
    Exit Sub
EH:
    Log "! Start failed: " & Err.Description
    lblStatus.Caption = "Error"
    lblStatus.ForeColor = vbRed
End Sub

Private Sub btnStop_Click()
    Dim i As Long
    m_bStopping = True
    On Error Resume Next
    '--- Disconnect all clients
    For i = czClient.UBound To 1 Step -1
        czClient(i).Disconnect
        Unload czClient(i)
    Next i
    czClient(0).Disconnect
    czServer.Disconnect
    On Error GoTo 0
    m_lNextIndex = 1
    ReDim m_bWsActive(0 To 0)
    m_lWsCount = 0
    m_bStopping = False
    btnStart.Enabled = True
    btnStop.Enabled = False
    btnBroadcast.Enabled = False
    lblStatus.Caption = "Server stopped"
    lblStatus.ForeColor = &HC0&
    lblClients.Caption = "WS Clients: 0"
    Log "* All clients disconnected. Server stopped."
    Log ""
End Sub

Private Sub btnBroadcast_Click()
    Dim i As Long
    Dim lSent As Long
    lSent = 0
    For i = 0 To UBound(m_bWsActive)
        If m_bWsActive(i) Then
            If czClient(i).ReadyState = 2 Then
                czClient(i).SendData txtBroadcast.Text
                lSent = lSent + 1
            End If
        End If
    Next i
    If lSent > 0 Then
        Log "> [Broadcast to " & lSent & " clients] " & txtBroadcast.Text
    Else
        Log "! No WebSocket clients connected."
    End If
End Sub

'=========================================================================
' czServer Events (Listener)
'=========================================================================

Private Sub czServer_ConnectionRequest(ByVal requestID As Long)
    Dim idx As Long
    idx = m_lNextIndex
    m_lNextIndex = m_lNextIndex + 1
    '--- Grow control array
    Load czClient(idx)
    '--- Grow tracking array
    If idx > UBound(m_bWsActive) Then
        ReDim Preserve m_bWsActive(0 To idx)
    End If
    m_bWsActive(idx) = False
    '--- Accept on new control
    czClient(idx).Accept requestID
    Log "* Connection #" & idx & " accepted"
End Sub

Private Sub czServer_Error(ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "! Server error " & Number & ": " & Description
End Sub

'=========================================================================
' czClient Events (Control Array - one per client)
'=========================================================================

Private Sub czClient_Connect(Index As Integer)
    '--- TLS handshake complete
    Log "* Client #" & Index & " TLS OK from " & czClient(Index).RemoteHostIP
End Sub

Private Sub czClient_DataArrival(Index As Integer, ByVal BytesTotal As Long)
    Dim sData As String
    Dim sResp As String
    Dim sBody As String
    czClient(Index).GetData sData
    '--- If not yet in WebSocket mode
    If czClient(Index).ReadyState <> 2 Then
        If InStr(1, sData, "Upgrade: websocket", vbTextCompare) > 0 Then
            '--- WebSocket upgrade
            Log "< Client #" & Index & " WebSocket upgrade request"
            czClient(Index).AcceptWebSocket sData
            m_bWsActive(Index) = True
            m_lWsCount = m_lWsCount + 1
            lblClients.Caption = "WS Clients: " & m_lWsCount
            Log "* Client #" & Index & " WebSocket connected! (" & m_lWsCount & " total)"
            Log ""
        Else
            '--- Serve HTML page
            sBody = GetWsClientPage()
            sResp = "HTTP/1.1 200 OK" & vbCrLf
            sResp = sResp & "Content-Type: text/html; charset=utf-8" & vbCrLf
            sResp = sResp & "Content-Length: " & LenB(StrConv(sBody, vbFromUnicode)) & vbCrLf
            sResp = sResp & "Connection: close" & vbCrLf
            sResp = sResp & "Server: czSocket/1.1" & vbCrLf
            sResp = sResp & vbCrLf & sBody
            czClient(Index).SendData sResp
            Log "> Served page to client #" & Index
        End If
    End If
End Sub

Private Sub czClient_Receive(Index As Integer, ByVal Data As String, ByVal IsBinary As Boolean)
    '--- WebSocket message received
    If IsBinary Then
        Log "< [#" & Index & " Binary: " & Len(Data) & " bytes]"
    Else
        Log "< [#" & Index & "] " & Data
        '--- Echo back
        Dim sReply As String
        sReply = "Echo: " & Data
        czClient(Index).SendData sReply
        Log "> [#" & Index & "] " & sReply
    End If
End Sub

Private Sub czClient_Disconnected(Index As Integer, ByVal Code As Long, ByVal Reason As String)
    If m_bStopping Then Exit Sub
    If Index > UBound(m_bWsActive) Then Exit Sub
    If m_bWsActive(Index) Then
        m_bWsActive(Index) = False
        m_lWsCount = m_lWsCount - 1
        lblClients.Caption = "WS Clients: " & m_lWsCount
        Log "* WS Client #" & Index & " disconnected (" & m_lWsCount & " remain)"
    End If
    '--- Cleanup: unload if not index 0
    If Index > 0 Then
        On Error Resume Next
        czClient(Index).Disconnect
        Unload czClient(Index)
        On Error GoTo 0
    End If
    Log ""
End Sub

Private Sub czClient_Error(Index As Integer, ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "! Client #" & Index & " error " & Number & ": " & Description
End Sub

Private Sub czClient_Close(Index As Integer)
    If m_bStopping Then Exit Sub
    If Index > UBound(m_bWsActive) Then Exit Sub
    If m_bWsActive(Index) Then
        m_bWsActive(Index) = False
        m_lWsCount = m_lWsCount - 1
        lblClients.Caption = "WS Clients: " & m_lWsCount
        Log "* Client #" & Index & " closed (" & m_lWsCount & " remain)"
    End If
    If Index > 0 Then
        On Error Resume Next
        czClient(Index).Disconnect
        Unload czClient(Index)
        On Error GoTo 0
    End If
End Sub

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    m_bStopping = True
    On Error Resume Next
    For i = czClient.UBound To 1 Step -1
        czClient(i).Disconnect
        Unload czClient(i)
    Next i
    czClient(0).Disconnect
    czServer.Disconnect
End Sub

'=========================================================================
' Helper
'=========================================================================

Private Sub Log(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Private Function GetWsClientPage() As String
    Dim s As String
    s = "<!DOCTYPE html><html><head><title>czSocket WSS Client</title>" & vbCrLf
    s = s & "<style>" & vbCrLf
    s = s & "*{margin:0;padding:0;box-sizing:border-box}" & vbCrLf
    s = s & "body{font-family:'Segoe UI',sans-serif;background:#0f0f23;color:#e0e0e0;height:100vh;display:flex;justify-content:center;align-items:center}" & vbCrLf
    s = s & ".app{background:linear-gradient(135deg,#1a1a3e,#0f2847);border-radius:20px;padding:30px;width:500px;box-shadow:0 20px 60px rgba(0,0,0,.6)}" & vbCrLf
    s = s & "h1{color:#e94560;font-size:1.5em;text-align:center;margin-bottom:5px}" & vbCrLf
    s = s & ".sub{text-align:center;color:#666;font-size:.85em;margin-bottom:15px}" & vbCrLf
    s = s & ".status{text-align:center;padding:8px;border-radius:10px;margin-bottom:15px;font-weight:600;font-size:.9em}" & vbCrLf
    s = s & ".status.off{background:#3a1111;color:#e94560}" & vbCrLf
    s = s & ".status.on{background:#0a3a0a;color:#53d769}" & vbCrLf
    s = s & "#log{background:#0a0a1a;border:1px solid #333;border-radius:10px;height:250px;overflow-y:auto;padding:12px;font-family:Consolas,monospace;font-size:.85em;margin-bottom:15px}" & vbCrLf
    s = s & ".msg{margin:4px 0;padding:4px 8px;border-radius:6px}" & vbCrLf
    s = s & ".msg.sent{background:#1a2a4a;color:#7db8f0;text-align:right}" & vbCrLf
    s = s & ".msg.recv{background:#1a3a2a;color:#53d769}" & vbCrLf
    s = s & ".msg.sys{color:#888;font-style:italic;font-size:.8em}" & vbCrLf
    s = s & ".row{display:flex;gap:8px}" & vbCrLf
    s = s & "input{flex:1;padding:10px 15px;border:1px solid #444;border-radius:10px;background:#151530;color:#fff;font-size:.95em;outline:none}" & vbCrLf
    s = s & "input:focus{border-color:#e94560}" & vbCrLf
    s = s & "button{padding:10px 20px;border:none;border-radius:10px;font-weight:600;cursor:pointer;font-size:.9em;transition:all .2s}" & vbCrLf
    s = s & "#btnSend{background:#e94560;color:#fff}#btnSend:hover{background:#ff5a7a}" & vbCrLf
    s = s & "#btnConnect{background:#2a6;color:#fff;margin-top:10px;width:100%}#btnConnect:hover{background:#3b7}" & vbCrLf
    s = s & "#btnDisconnect{background:#a33;color:#fff;margin-top:10px;width:100%;display:none}#btnDisconnect:hover{background:#c44}" & vbCrLf
    s = s & "</style></head><body>" & vbCrLf
    s = s & "<div class='app'>" & vbCrLf
    s = s & "<h1>czSocket WSS Client</h1>" & vbCrLf
    s = s & "<div class='sub'>WebSocket Secure Test Page</div>" & vbCrLf
    s = s & "<div id='status' class='status off'>Disconnected</div>" & vbCrLf
    s = s & "<div id='log'></div>" & vbCrLf
    s = s & "<div class='row'><input id='msg' placeholder='Type a message...' onkeydown='if(event.key==""Enter"")send()'><button id='btnSend' onclick='send()'>Send</button></div>" & vbCrLf
    s = s & "<button id='btnConnect' onclick='connect()'>Connect to WSS Server</button>" & vbCrLf
    s = s & "<button id='btnDisconnect' onclick='disconnect()'>Disconnect</button>" & vbCrLf
    s = s & "</div><script>" & vbCrLf
    s = s & "let ws;const log=document.getElementById('log'),st=document.getElementById('status');" & vbCrLf
    s = s & "function addMsg(t,c){const d=document.createElement('div');d.className='msg '+c;d.textContent=t;log.appendChild(d);log.scrollTop=log.scrollHeight}" & vbCrLf
    s = s & "function connect(){" & vbCrLf
    s = s & "ws=new WebSocket('wss://'+location.host);" & vbCrLf
    s = s & "ws.onopen=()=>{st.textContent='Connected';st.className='status on';addMsg('Connected to server','sys');" & vbCrLf
    s = s & "document.getElementById('btnConnect').style.display='none';document.getElementById('btnDisconnect').style.display='block'};" & vbCrLf
    s = s & "ws.onmessage=e=>addMsg(e.data,'recv');" & vbCrLf
    s = s & "ws.onclose=e=>{st.textContent='Disconnected ('+e.code+')';st.className='status off';addMsg('Disconnected','sys');" & vbCrLf
    s = s & "document.getElementById('btnConnect').style.display='block';document.getElementById('btnDisconnect').style.display='none'};" & vbCrLf
    s = s & "ws.onerror=()=>addMsg('Connection error','sys')}" & vbCrLf
    s = s & "function send(){const m=document.getElementById('msg');if(ws&&ws.readyState===1&&m.value){ws.send(m.value);addMsg(m.value,'sent');m.value=''}}" & vbCrLf
    s = s & "function disconnect(){if(ws)ws.close()}" & vbCrLf
    s = s & "</script></body></html>"
    GetWsClientPage = s
End Function
