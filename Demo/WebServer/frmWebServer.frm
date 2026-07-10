VERSION 5.00
Begin VB.Form frmWebServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czSocket Demo - HTTPS WebServer"
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
   Begin czWebServer.czSocket czServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin czWebServer.czSocket czClient 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame fraServer 
      Caption         =   " HTTPS Server "
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9360
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Text            =   "8443"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCertFile 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   8280
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "czSocket123"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnStart 
         Caption         =   "Start Server"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton btnStop 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   7320
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblPort 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   735
      End
      Begin VB.Label lblCert 
         Caption         =   "PFX Certificate:"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label lblPwd 
         Caption         =   "Pwd:"
         Height          =   255
         Left            =   7920
         TabIndex        =   8
         Top             =   405
         Width           =   375
      End
      Begin VB.Label lblStatus 
         Caption         =   "Server stopped"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblClients 
         Caption         =   "Clients: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   9000
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
      Height          =   6015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2280
      Width           =   9360
   End
End
Attribute VB_Name = "frmWebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' czSocket Demo - HTTPS Web Server
'
' This demo shows how to:
'   1. Load a PFX certificate
'   2. Listen on an HTTPS port with TLS
'   3. Accept incoming connections (single-client)
'   4. Respond to HTTP requests with HTML
'
' Uses czServer (listener) and czClient (accepted connection).
' Certificate: localhost.pfx (self-signed, password: czSocket123)
'=========================================================================

Private m_bRunning As Boolean
Private m_lClients As Long

Private Sub Form_Load()
    '--- Default cert path
    txtCertFile.Text = App.Path & "\localhost.pfx"
    Log "=== czSocket HTTPS WebServer Demo ==="
    Log "1. Click 'Start Server' to begin listening."
    Log "2. Open browser to https://localhost:" & txtPort.Text & "/"
    Log "3. Accept the self-signed certificate warning."
    Log ""
End Sub

'=========================================================================
' Server Controls
'=========================================================================

Private Sub btnStart_Click()
    On Error GoTo EH
    If LenB(txtCertFile.Text) = 0 Then
        MsgBox "Please select a PFX certificate file.", vbExclamation
        Exit Sub
    End If
    If Dir$(txtCertFile.Text) = "" Then
        MsgBox "Certificate file not found: " & txtCertFile.Text, vbExclamation
        Exit Sub
    End If
    Dim lPort As Long
    lPort = Val(txtPort.Text)
    If lPort < 1 Or lPort > 65535 Then
        MsgBox "Invalid port number.", vbExclamation
        Exit Sub
    End If
    '--- Start HTTPS server
    czServer.LocalPort = lPort
    czServer.Listen txtCertFile.Text, txtPassword.Text
    m_bRunning = True
    m_lClients = 0
    btnStart.Enabled = False
    btnStop.Enabled = True
    txtPort.Enabled = False
    txtCertFile.Enabled = False
    txtPassword.Enabled = False
    lblStatus.Caption = "Listening on https://localhost:" & lPort & "/"
    lblStatus.ForeColor = &H8000&
    Log "* Server started on port " & lPort
    Log "* Certificate: " & txtCertFile.Text
    Log "* Waiting for connections..."
    Log ""
    Exit Sub
EH:
    Log "! Start failed: " & Err.Description
    lblStatus.Caption = "Error: " & Err.Description
    lblStatus.ForeColor = vbRed
End Sub

Private Sub btnStop_Click()
    czClient.Disconnect
    czServer.Disconnect
    m_bRunning = False
    m_lClients = 0
    btnStart.Enabled = True
    btnStop.Enabled = False
    txtPort.Enabled = True
    txtCertFile.Enabled = True
    txtPassword.Enabled = True
    lblStatus.Caption = "Server stopped"
    lblStatus.ForeColor = &HC0&
    lblClients.Caption = "Clients: 0"
    Log "* Server stopped."
    Log ""
End Sub

Private Sub btnBrowse_Click()
    '--- Simple file dialog using CommonDialog or Shell
    Dim sFile As String
    sFile = InputBox("Enter path to PFX certificate file:", "Certificate File", txtCertFile.Text)
    If LenB(sFile) <> 0 Then
        txtCertFile.Text = sFile
    End If
End Sub

'=========================================================================
' czServer Events (Listener)
'=========================================================================

Private Sub czServer_ConnectionRequest(ByVal requestID As Long)
    '--- Accept the connection on czClient
    czClient.Disconnect
    czClient.Accept requestID
    m_lClients = m_lClients + 1
    lblClients.Caption = "Total connections: " & m_lClients
End Sub

Private Sub czServer_Error(ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "! Server error " & Number & ": " & Description
    lblStatus.Caption = "Error: " & Description
    lblStatus.ForeColor = vbRed
End Sub

'=========================================================================
' czClient Events (Accepted Connection)
'=========================================================================

Private Sub czClient_Connect()
    '--- TLS handshake complete, client is ready
    Log "* Client #" & m_lClients & " connected from " & czClient.RemoteHostIP & " (TLS OK)"
End Sub

Private Sub czClient_DataArrival(ByVal BytesTotal As Long)
    Dim sData As String
    Dim sMethod As String
    Dim sPath As String
    Dim sFirstLine As String
    Dim parts() As String
    Dim sBody As String
    Dim sContentType As String
    Dim sResponse As String
    czClient.GetData sData
    Log "< " & Left$(sData, 200)
    If Len(sData) > 200 Then Log "  ... (" & Len(sData) & " bytes total)"
    '--- Parse simple HTTP request
    sMethod = ""
    sPath = "/"
    If InStr(sData, vbCrLf) > 0 Then
        sFirstLine = Left$(sData, InStr(sData, vbCrLf) - 1)
        parts = Split(sFirstLine, " ")
        If UBound(parts) >= 1 Then
            sMethod = parts(0)
            sPath = parts(1)
        End If
    End If
    '--- Build HTML response
    sContentType = "text/html; charset=utf-8"
    If sPath = "/api" Then
        sContentType = "application/json"
        sBody = "{""status"":""ok"",""server"":""czSocket"",""tls"":true,""message"":""Hello from czSocket HTTPS Server!""}"
    ElseIf sPath = "/favicon.ico" Then
        sBody = ""
        sContentType = "image/x-icon"
    Else
        sBody = "<!DOCTYPE html>" & vbCrLf
        sBody = sBody & "<html><head><title>czSocket HTTPS Server</title>" & vbCrLf
        sBody = sBody & "<style>" & vbCrLf
        sBody = sBody & "body{font-family:'Segoe UI',sans-serif;background:#1a1a2e;color:#e0e0e0;" & vbCrLf
        sBody = sBody & "display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0}" & vbCrLf
        sBody = sBody & ".card{background:linear-gradient(135deg,#16213e,#0f3460);border-radius:20px;" & vbCrLf
        sBody = sBody & "padding:40px 60px;box-shadow:0 20px 60px rgba(0,0,0,0.5);text-align:center;max-width:500px}" & vbCrLf
        sBody = sBody & "h1{color:#e94560;font-size:2em;margin-bottom:10px}" & vbCrLf
        sBody = sBody & ".badge{display:inline-block;background:#e94560;color:#fff;padding:5px 15px;" & vbCrLf
        sBody = sBody & "border-radius:20px;font-size:.8em;margin:10px 0}" & vbCrLf
        sBody = sBody & "p{line-height:1.6;color:#a0a0b0}" & vbCrLf
        sBody = sBody & ".info{background:rgba(255,255,255,0.05);border-radius:10px;padding:15px;" & vbCrLf
        sBody = sBody & "margin-top:20px;font-family:Consolas,monospace;font-size:.9em;color:#53d769}" & vbCrLf
        sBody = sBody & "</style></head><body>" & vbCrLf
        sBody = sBody & "<div class='card'>" & vbCrLf
        sBody = sBody & "<h1>czSocket</h1>" & vbCrLf
        sBody = sBody & "<div class='badge'>HTTPS Server</div>" & vbCrLf
        sBody = sBody & "<p>This page is served over <strong>TLS</strong> by a VB6 application "
        sBody = sBody & "using <strong>czSocket</strong> with native Windows Schannel.</p>" & vbCrLf
        sBody = sBody & "<div class='info'>" & vbCrLf
        sBody = sBody & "Method: " & sMethod & "<br>" & vbCrLf
        sBody = sBody & "Path: " & sPath & "<br>" & vbCrLf
        sBody = sBody & "Server: czSocket/1.1 (VB6)<br>" & vbCrLf
        sBody = sBody & "TLS: Schannel (native Windows)" & vbCrLf
        sBody = sBody & "</div></div></body></html>"
    End If
    '--- Build HTTP response
    sResponse = "HTTP/1.1 200 OK" & vbCrLf
    sResponse = sResponse & "Content-Type: " & sContentType & vbCrLf
    sResponse = sResponse & "Content-Length: " & Len(sBody) & vbCrLf
    sResponse = sResponse & "Connection: close" & vbCrLf
    sResponse = sResponse & "Server: czSocket/1.1" & vbCrLf
    sResponse = sResponse & vbCrLf & sBody
    '--- Send response
    czClient.SendData sResponse
    Log "> HTTP 200 (" & sContentType & ") " & Len(sBody) & " bytes"
End Sub

Private Sub czClient_SendComplete()
    '--- Response sent, close connection
    czClient.Disconnect
    Log "* Client disconnected (response sent)"
    Log ""
End Sub

Private Sub czClient_Disconnected(ByVal Code As Long, ByVal Reason As String)
    If Code > 0 Then
        Log "* Client disconnected: Code=" & Code & " Reason=" & Reason
    End If
End Sub

Private Sub czClient_Error(ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log "! Client error " & Number & ": " & Description
End Sub

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Form_Unload(Cancel As Integer)
    czClient.Disconnect
    czServer.Disconnect
End Sub

'=========================================================================
' Helper
'=========================================================================

Private Sub Log(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub
