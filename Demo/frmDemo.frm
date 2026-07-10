VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czSocket Demo - Simple HTTP & WebSocket"
   ClientHeight    =   7200
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
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin czSocketDemo.czSocket czSocket1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame fraTest 
      Caption         =   " Select Test "
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9360
      Begin VB.CommandButton btnTest1 
         Caption         =   "HTTP GET Request"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnTest2 
         Caption         =   "HTTPS JSON API"
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnTest3 
         Caption         =   "WebSocket Echo"
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnDisconnect 
         Caption         =   "Disconnect"
         Height          =   495
         Left            =   6960
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         Caption         =   "Ready"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   9000
      End
   End
   Begin VB.Frame fraInput 
      Caption         =   " Send Data (WebSocket) "
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   9360
      Begin VB.TextBox txtSend 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7095
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   240
         Width           =   1935
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
      Height          =   4695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2400
      Width           =   9360
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' czSocket Demo - Simple Usage Examples
'
' This demo shows 3 basic use cases:
'   1. HTTP GET request (raw TCP)
'   2. HTTPS JSON API request (auto-TLS)
'   3. WebSocket echo (auto from URL)
'=========================================================================

Private Sub Form_Load()
    txtSend.Text = "Hello, czSocket!"
    Log "=== czSocket Demo ==="
    Log "Click a button to test."
    Log ""
End Sub

'--- Test 1: Simple HTTP GET (raw TCP, parse response manually)
Private Sub btnTest1_Click()
    Log "--- Test 1: HTTP GET ---"
    lblStatus.Caption = "Connecting to example.com..."
    lblStatus.ForeColor = vbBlue
    czSocket1.Connect "example.com", 80
End Sub

'--- Test 2: HTTPS JSON API (auto-TLS, auto-parse)
Private Sub btnTest2_Click()
    Log "--- Test 2: HTTPS JSON API ---"
    lblStatus.Caption = "Requesting HTTPS..."
    lblStatus.ForeColor = vbBlue
    czSocket1.Request "GET", "https://jsonplaceholder.typicode.com/posts/1"
End Sub

'--- Test 3: WebSocket Echo
Private Sub btnTest3_Click()
    Log "--- Test 3: WebSocket Echo ---"
    lblStatus.Caption = "Connecting to WebSocket..."
    lblStatus.ForeColor = vbBlue
    czSocket1.Connect "wss://ws.postman-echo.com/raw"
End Sub

'--- Send data (for WebSocket test)
Private Sub btnSend_Click()
    If czSocket1.ReadyState = 2 Then  '--- WS_OPEN
        Log "> " & txtSend.Text
        czSocket1.SendData txtSend.Text
    ElseIf czSocket1.State = 7 Then   '--- sckConnected (raw TCP)
        Log "> " & txtSend.Text
        czSocket1.SendData txtSend.Text
    Else
        Log "! Not connected. Click a test button first."
    End If
End Sub

'--- Disconnect
Private Sub btnDisconnect_Click()
    czSocket1.Disconnect
    Log "* Disconnected by user."
    lblStatus.Caption = "Disconnected"
    lblStatus.ForeColor = vbRed
End Sub

'=========================================================================
' czSocket Events
'=========================================================================

Private Sub czSocket1_Connect()
    lblStatus.Caption = "Connected!"
    lblStatus.ForeColor = &H8000&  '--- Dark green
    Log "* Connected to " & czSocket1.RemoteHostIP & ":" & czSocket1.RemotePort
    
    '--- For Test 1 (raw TCP): send HTTP request manually
    If czSocket1.RemotePort = 80 And czSocket1.ReadyState = 0 Then
        Dim sReq As String
        sReq = "GET / HTTP/1.1" & vbCrLf & _
               "Host: example.com" & vbCrLf & _
               "Connection: close" & vbCrLf & vbCrLf
        czSocket1.SendData sReq
        Log "> GET / HTTP/1.1"
    End If
    
    '--- For Test 3 (WebSocket): ready to send
    If czSocket1.ReadyState = 2 Then  '--- WS_OPEN
        Log "* WebSocket ready! Type a message and click Send."
    End If
End Sub

Private Sub czSocket1_DataArrival(ByVal BytesTotal As Long)
    Dim sData As String
    czSocket1.GetData sData
    Log "< " & Left$(sData, 500)
    If Len(sData) > 500 Then Log "  ... (" & Len(sData) & " bytes total)"
End Sub

Private Sub czSocket1_Response(ByVal Status As Long, ByVal ContentType As String, _
    Body As String, Headers As String)
    Log "< HTTP " & Status & " (" & ContentType & ")"
    Log "< " & Left$(Body, 500)
    If Len(Body) > 500 Then Log "  ... (" & Len(Body) & " bytes total)"
    lblStatus.Caption = "HTTP " & Status & " received"
    lblStatus.ForeColor = IIf(Status = 200, &H8000&, vbRed)
End Sub

Private Sub czSocket1_Receive(ByVal Data As String, ByVal IsBinary As Boolean)
    If IsBinary Then
        Log "< [Binary: " & Len(Data) & " bytes]"
    Else
        Log "< " & Data
    End If
End Sub

Private Sub czSocket1_Disconnected(ByVal Code As Long, ByVal Reason As String)
    If Code > 0 Then
        Log "* Disconnected: Code=" & Code & " Reason=" & Reason
    Else
        Log "* Connection closed."
    End If
    lblStatus.Caption = "Disconnected"
    lblStatus.ForeColor = vbRed
End Sub

Private Sub czSocket1_Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, _
    ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    If BytesTotal > 0 Then
        lblStatus.Caption = "Transfer: " & Int(BytesSent * 100 / BytesTotal) & "% | " & _
            Format$(BytesPerSec / 1024, "0.0") & " KB/s | ETA: " & SecondsRemaining & "s"
    End If
End Sub

Private Sub czSocket1_Error(ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "! ERROR " & Number & ": " & Description
    lblStatus.Caption = "Error: " & Description
    lblStatus.ForeColor = vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    czSocket1.Disconnect
End Sub

'=========================================================================
' Helper
'=========================================================================

Private Sub Log(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
