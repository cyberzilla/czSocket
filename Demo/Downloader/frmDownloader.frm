VERSION 5.00
Begin VB.Form frmDownloader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "czSocket Demo - Parallel Downloader && Uploader"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
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
   ScaleHeight     =   9600
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin czDownloader.czSocket czDl 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin czDownloader.czSocket czUpload 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame fraDownload 
      Caption         =   " Parallel Download Manager "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9960
      Begin VB.TextBox txtUrl 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Text            =   "https://labs.dev/czSocket/PowerShell-7.6.3-win-x64.msi"
         Top             =   300
         Width           =   7500
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add Download"
         Height          =   375
         Left            =   8220
         TabIndex        =   2
         Top             =   300
         Width           =   1620
      End
      Begin VB.TextBox txtSaveFolder 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   780
         Width           =   6900
      End
      Begin VB.CommandButton btnBrowseFolder 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   8220
         TabIndex        =   4
         Top             =   780
         Width           =   1620
      End
      Begin VB.TextBox txtMaxParallel 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1500
         TabIndex        =   5
         Text            =   "4"
         Top             =   1320
         Width           =   600
      End
      Begin VB.CommandButton btnClearDone 
         Caption         =   "Clear Completed"
         Height          =   375
         Left            =   8220
         TabIndex        =   6
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label lblUrl 
         Caption         =   "URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblSave 
         Caption         =   "Save Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblMax 
         Caption         =   "Max Parallel:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lblActive 
         Caption         =   "Active: 0 / Queued: 0 / Done: 0"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   1380
         Width           =   5000
      End
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4200
      Left            =   120
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   660
      TabIndex        =   7
      Top             =   2100
      Width           =   9960
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6360
      Width           =   1200
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   6360
      Width           =   1200
   End
   Begin VB.Label lblSelected 
      Caption         =   "Click a download to select it"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   6420
      Width           =   7200
   End
   Begin VB.Frame fraUpload 
      Caption         =   " File Uploader "
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   6840
      Width           =   9960
      Begin VB.TextBox txtUploadUrl 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Text            =   "https://labs.dev/czSocket/upload.php"
         Top             =   300
         Width           =   9240
      End
      Begin VB.TextBox txtUploadFile 
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   780
         Width           =   7500
      End
      Begin VB.CommandButton btnBrowseFile 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   8220
         TabIndex        =   10
         Top             =   780
         Width           =   1620
      End
      Begin VB.CommandButton btnUpload 
         Caption         =   "Upload File"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label lblUploadUrl 
         Caption         =   "URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblUploadFile 
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblUploadStatus 
         Caption         =   "Ready"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
         Width           =   7800
      End
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   8460
      Width           =   9960
   End
End
Attribute VB_Name = "frmDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================
' czSocket Demo - Parallel Downloader & Uploader
'
' Features:
'   - Add multiple download URLs
'   - Configurable max parallel downloads
'   - Visual progress bars drawn natively on PictureBox
'   - Shows filename, percentage, speed, ETA per download
'   - File upload to PHP endpoint
'=========================================================================

Private Type DownloadSlot
    sUrl        As String
    sSavePath   As String
    sFileName   As String
    lStatus     As Long     '0=empty,1=waiting,2=active,3=complete,4=error
    lBytesRecv  As Long
    lBytesTotal As Long
    lSpeed      As Long
    lETA        As Long
    sError      As String
    dStartTime  As Double   '--- Timer value when download started
    lRetryCount As Long     '--- Auto-retry counter
End Type

Private Const DL_EMPTY As Long = 0
Private Const DL_WAITING As Long = 1
Private Const DL_ACTIVE As Long = 2
Private Const DL_COMPLETE As Long = 3
Private Const DL_ERROR As Long = 4
Private Const DL_PAUSED As Long = 5

Private Const ROW_HEIGHT As Long = 52
Private Const BAR_MARGIN As Long = 4
Private Const MAX_SLOTS As Long = 50
Private Const BLOCK_COUNT As Long = 40
Private Const MAX_RETRIES As Long = 2

Private m_tSlots() As DownloadSlot
Private m_lCount As Long
Private m_lMaxParallel As Long
Private m_lActiveCount As Long
Private m_lChunkCount() As Long
Private m_dLastDraw As Double       '--- Throttle DrawProgress redraws
Private m_lSelectedSlot As Long     '--- Currently selected download (-1 = none)
Private m_lUploadChunkCount As Long '--- Upload chunk counter for logging

'--- Browse dialog API
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Form_Load()
    ReDim m_tSlots(0 To MAX_SLOTS - 1)
    ReDim m_lChunkCount(0 To MAX_SLOTS - 1)
    m_lCount = 0
    m_lMaxParallel = 4
    m_lActiveCount = 0
    m_lSelectedSlot = -1
    txtSaveFolder.Text = App.Path & "\Downloads"
    txtMaxParallel.Text = "4"
    btnPause.Enabled = False
    btnStop.Enabled = False
    Log "=== czSocket Parallel Downloader ==="
    Log "Add URLs to download. Max parallel can be changed anytime."
    DrawProgress
End Sub

'=========================================================================
' Download Controls
'=========================================================================

Private Sub btnAdd_Click()
    Dim sUrl As String
    Dim sFolder As String
    Dim sFileName As String
    sUrl = Trim$(txtUrl.Text)
    If Len(sUrl) = 0 Then
        MsgBox "Please enter a URL.", vbExclamation
        Exit Sub
    End If
    sFolder = Trim$(txtSaveFolder.Text)
    If Len(sFolder) = 0 Then
        MsgBox "Please set a save folder.", vbExclamation
        Exit Sub
    End If
    '--- Create folder if needed
    If Dir$(sFolder, vbDirectory) = "" Then
        MkDir sFolder
    End If
    '--- Extract filename from URL
    sFileName = ExtractFileName(sUrl)
    If Len(sFileName) = 0 Then sFileName = "download_" & m_lCount + 1
    '--- Add to slot
    If m_lCount >= MAX_SLOTS Then
        MsgBox "Maximum download slots reached!", vbExclamation
        Exit Sub
    End If
    Dim idx As Long
    idx = m_lCount
    m_lCount = m_lCount + 1
    m_tSlots(idx).sUrl = sUrl
    m_tSlots(idx).sSavePath = sFolder & "\" & sFileName
    m_tSlots(idx).sFileName = sFileName
    m_tSlots(idx).lStatus = DL_WAITING
    m_tSlots(idx).lBytesRecv = 0
    m_tSlots(idx).lBytesTotal = 0
    m_tSlots(idx).lSpeed = 0
    m_tSlots(idx).lETA = 0
    '--- Load czDl control
    If idx > 0 Then Load czDl(idx)
    Log "* Added: " & sFileName
    '--- Try to start
    StartNextDownloads
    DrawProgress
    UpdateStatusLabel
End Sub

Private Sub btnClearDone_Click()
    '--- Remove completed/error slots and compact
    Dim i As Long
    Dim j As Long
    j = 0
    For i = 0 To m_lCount - 1
        If m_tSlots(i).lStatus = DL_COMPLETE Or m_tSlots(i).lStatus = DL_ERROR Then
            '--- Unload control if > 0
            If i > 0 Then
                On Error Resume Next
                Unload czDl(i)
                On Error GoTo 0
            End If
        Else
            If j <> i Then
                m_tSlots(j) = m_tSlots(i)
            End If
            j = j + 1
        End If
    Next i
    '--- Clear remaining slots
    Dim lOldCount As Long
    lOldCount = m_lCount
    m_lCount = j
    For i = m_lCount To lOldCount - 1
        m_tSlots(i).sUrl = ""
        m_tSlots(i).lStatus = DL_EMPTY
    Next i
    Log "* Cleared completed downloads"
    DrawProgress
    UpdateStatusLabel
End Sub

Private Sub btnBrowseFolder_Click()
    Dim oShell As Object
    Dim oFolder As Object
    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.BrowseForFolder(Me.hWnd, "Select download folder", &H1)
    If Not oFolder Is Nothing Then
        txtSaveFolder.Text = oFolder.Self.Path
    End If
End Sub

Private Sub txtMaxParallel_Change()
    Dim lVal As Long
    If IsNumeric(txtMaxParallel.Text) Then
        lVal = CLng(txtMaxParallel.Text)
        If lVal >= 1 And lVal <= 20 Then
            m_lMaxParallel = lVal
            StartNextDownloads
        End If
    End If
End Sub

'=========================================================================
' Pause / Stop Controls
'=========================================================================

Private Sub picProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim idx As Long
    idx = CLng(Y) \ ROW_HEIGHT
    If idx >= 0 And idx < m_lCount Then
        m_lSelectedSlot = idx
        lblSelected.Caption = "Selected: " & m_tSlots(idx).sFileName
    Else
        m_lSelectedSlot = -1
        lblSelected.Caption = "Click a download to select it"
    End If
    UpdateButtons
    DrawProgress
End Sub

Private Sub UpdateButtons()
    If m_lSelectedSlot < 0 Or m_lSelectedSlot >= m_lCount Then
        btnPause.Caption = "Pause"
        btnPause.Enabled = False
        btnStop.Enabled = False
        Exit Sub
    End If
    Select Case m_tSlots(m_lSelectedSlot).lStatus
        Case DL_ACTIVE
            btnPause.Caption = "Pause"
            btnPause.Enabled = True
            btnStop.Enabled = True
        Case DL_PAUSED
            btnPause.Caption = "Resume"
            btnPause.Enabled = True
            btnStop.Enabled = True
        Case DL_WAITING
            btnPause.Caption = "Pause"
            btnPause.Enabled = False
            btnStop.Caption = "Cancel"
            btnStop.Enabled = True
        Case Else
            btnPause.Caption = "Pause"
            btnPause.Enabled = False
            btnStop.Enabled = False
    End Select
End Sub

Private Sub btnPause_Click()
    If m_lSelectedSlot < 0 Or m_lSelectedSlot >= m_lCount Then
        MsgBox "Select a download first.", vbExclamation
        Exit Sub
    End If
    Dim idx As Long
    idx = m_lSelectedSlot
    Select Case m_tSlots(idx).lStatus
        Case DL_ACTIVE
            '--- Pause: disconnect and keep state
            On Error Resume Next
            czDl(idx).Disconnect
            On Error GoTo 0
            m_tSlots(idx).lStatus = DL_PAUSED
            m_lActiveCount = m_lActiveCount - 1
            Log "* Paused: " & m_tSlots(idx).sFileName & " (" & FormatBytes(m_tSlots(idx).lBytesRecv) & ")"
            DrawProgress
            UpdateStatusLabel
            UpdateButtons
        Case DL_PAUSED
            '--- Resume: re-download from where we left off (Range header)
            Dim lResumeFrom As Long
            lResumeFrom = m_tSlots(idx).lBytesRecv
            m_tSlots(idx).lStatus = DL_ACTIVE
            m_tSlots(idx).dStartTime = Timer
            m_tSlots(idx).lSpeed = 0
            m_tSlots(idx).lETA = 0
            m_lChunkCount(idx) = 0
            m_lActiveCount = m_lActiveCount + 1
            On Error Resume Next
            czDl(idx).Download m_tSlots(idx).sUrl, m_tSlots(idx).sSavePath, , lResumeFrom
            If Err.Number <> 0 Then
                m_tSlots(idx).lStatus = DL_ERROR
                m_tSlots(idx).sError = Err.Description
                m_lActiveCount = m_lActiveCount - 1
                Log "! Resume failed: " & m_tSlots(idx).sFileName & " - " & Err.Description
            Else
                Log "* Resumed: " & m_tSlots(idx).sFileName & " from " & FormatBytes(lResumeFrom)
            End If
            On Error GoTo 0
            DrawProgress
            UpdateStatusLabel
            UpdateButtons
        Case Else
            MsgBox "This download cannot be paused.", vbInformation
    End Select
End Sub

Private Sub btnStop_Click()
    If m_lSelectedSlot < 0 Or m_lSelectedSlot >= m_lCount Then
        MsgBox "Select a download first.", vbExclamation
        Exit Sub
    End If
    Dim idx As Long
    idx = m_lSelectedSlot
    Select Case m_tSlots(idx).lStatus
        Case DL_ACTIVE
            On Error Resume Next
            czDl(idx).Disconnect
            On Error GoTo 0
            m_tSlots(idx).lStatus = DL_ERROR
            m_tSlots(idx).sError = "Cancelled"
            m_lActiveCount = m_lActiveCount - 1
            Log "* Stopped: " & m_tSlots(idx).sFileName
            StartNextDownloads
            DrawProgress
            UpdateStatusLabel
            UpdateButtons
        Case DL_PAUSED
            m_tSlots(idx).lStatus = DL_ERROR
            m_tSlots(idx).sError = "Cancelled"
            Log "* Stopped: " & m_tSlots(idx).sFileName
            DrawProgress
            UpdateStatusLabel
            UpdateButtons
        Case DL_WAITING
            m_tSlots(idx).lStatus = DL_ERROR
            m_tSlots(idx).sError = "Cancelled"
            Log "* Cancelled: " & m_tSlots(idx).sFileName
            DrawProgress
            UpdateStatusLabel
            UpdateButtons
        Case Else
            MsgBox "This download is already finished.", vbInformation
    End Select
End Sub

'=========================================================================
' Upload Controls
'=========================================================================

Private Sub btnBrowseFile_Click()
    Dim ofn As OPENFILENAME
    Dim sFile As String
    sFile = String$(260, vbNullChar)
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    ofn.lpstrFile = sFile
    ofn.nMaxFile = 260
    ofn.lpstrTitle = "Select file to upload"
    ofn.flags = &H1000 Or &H4  'OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    If GetOpenFileName(ofn) <> 0 Then
        txtUploadFile.Text = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    End If
End Sub

Private Sub btnUpload_Click()
    Dim sUrl As String
    Dim sFile As String
    sUrl = Trim$(txtUploadUrl.Text)
    sFile = Trim$(txtUploadFile.Text)
    If Len(sUrl) = 0 Then
        MsgBox "Please enter an upload URL.", vbExclamation
        Exit Sub
    End If
    If Len(sFile) = 0 Or Dir$(sFile) = "" Then
        MsgBox "Please select a valid file.", vbExclamation
        Exit Sub
    End If
    On Error GoTo EH
    lblUploadStatus.Caption = "Uploading..."
    lblUploadStatus.ForeColor = &HF0A000
    m_lUploadChunkCount = 0
    btnUpload.Enabled = False
    czUpload.Upload sUrl, sFile
    Log "* Upload started: " & sFile
    Exit Sub
EH:
    Log "! Upload error: " & Err.Description
    lblUploadStatus.Caption = "Error: " & Err.Description
    lblUploadStatus.ForeColor = vbRed
    btnUpload.Enabled = True
End Sub

'=========================================================================
' czDl Events (Download Control Array)
'=========================================================================


Private Sub czDl_Progress(Index As Integer, ByVal BytesSent As Long, ByVal BytesTotal As Long, ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    If Index < 0 Or Index >= m_lCount Then Exit Sub
    '--- Track chunk count
    If Index > UBound(m_lChunkCount) Then ReDim Preserve m_lChunkCount(0 To Index)
    m_lChunkCount(Index) = m_lChunkCount(Index) + 1
    m_tSlots(Index).lBytesRecv = BytesSent
    m_tSlots(Index).lBytesTotal = BytesTotal
    m_tSlots(Index).lSpeed = BytesPerSec
    m_tSlots(Index).lETA = SecondsRemaining
    '--- Log every 50 chunks to avoid flood
    If m_lChunkCount(Index) Mod 50 = 0 Or BytesSent >= BytesTotal Then
        Log "  [#" & Index & "] chunk=" & m_lChunkCount(Index) & _
            "  recv=" & FormatBytes(BytesSent) & "/" & FormatBytes(BytesTotal) & _
            "  speed=" & FormatSpeed(BytesPerSec) & _
            "  eta=" & FormatETA(SecondsRemaining)
    End If
    '--- Throttle DrawProgress to max ~10 fps
    If Timer - m_dLastDraw > 0.1 Or BytesSent >= BytesTotal Then
        m_dLastDraw = Timer
        DrawProgress
    End If
End Sub

Private Sub czDl_Response(Index As Integer, ByVal Status As Long, ByVal ContentType As String, Body As String, Headers As String)
    If Index < 0 Or Index >= m_lCount Then Exit Sub
    If Status >= 200 And Status < 400 Then
        m_tSlots(Index).lStatus = DL_COMPLETE
        m_tSlots(Index).lBytesRecv = m_tSlots(Index).lBytesTotal
        Dim dElapsed As Double
        dElapsed = Timer - m_tSlots(Index).dStartTime
        If dElapsed < 0 Then dElapsed = dElapsed + 86400#
        Log "* Complete: " & m_tSlots(Index).sFileName & " (" & FormatBytes(m_tSlots(Index).lBytesTotal) & ") in " & Format$(dElapsed, "0.0") & "s"
    Else
        m_tSlots(Index).lStatus = DL_ERROR
        m_tSlots(Index).sError = "HTTP " & Status
        Log "! Failed: " & m_tSlots(Index).sFileName & " - HTTP " & Status
    End If
    m_lActiveCount = m_lActiveCount - 1
    StartNextDownloads
    DrawProgress
    UpdateStatusLabel
End Sub

Private Sub czDl_Error(Index As Integer, ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    If Number = 0 Then Exit Sub
    If Index < 0 Or Index >= m_lCount Then Exit Sub
    If m_tSlots(Index).lStatus = DL_ACTIVE Then
        m_tSlots(Index).lRetryCount = m_tSlots(Index).lRetryCount + 1
        '--- Auto-retry if under limit
        If m_tSlots(Index).lRetryCount <= MAX_RETRIES Then
            Log "! Error: " & m_tSlots(Index).sFileName & " - " & Description & " (retry " & m_tSlots(Index).lRetryCount & "/" & MAX_RETRIES & ")"
            m_tSlots(Index).lStatus = DL_WAITING
            m_tSlots(Index).lBytesRecv = 0
            m_tSlots(Index).lBytesTotal = 0
            m_tSlots(Index).lSpeed = 0
            m_tSlots(Index).lETA = 0
            m_lActiveCount = m_lActiveCount - 1
            StartNextDownloads
        Else
            m_tSlots(Index).lStatus = DL_ERROR
            m_tSlots(Index).sError = Description
            m_lActiveCount = m_lActiveCount - 1
            Log "! Error: " & m_tSlots(Index).sFileName & " - " & Description & " (no more retries)"
            StartNextDownloads
        End If
        DrawProgress
        UpdateStatusLabel
    End If
End Sub

'=========================================================================
' czUpload Events
'=========================================================================

Private Sub czUpload_Response(ByVal Status As Long, ByVal ContentType As String, Body As String, Headers As String)
    btnUpload.Enabled = True
    Log "< HTTP " & Status & " (" & ContentType & ")"
    Log "< " & Left$(Body, 300)
    If Status >= 200 And Status < 300 Then
        '--- Try JSON parse if content type indicates JSON
        If InStr(1, ContentType, "json", vbTextCompare) > 0 And Len(Body) > 0 Then
            czUpload.JsonParse Body
            '--- Check if parsing actually succeeded (JsonHas returns False if m_oJsonRoot is Nothing)
            If czUpload.JsonHas("success") Then
                If czUpload.JsonBool("success") Then
                    lblUploadStatus.Caption = "Upload OK! Server: " & czUpload.JsonStr("server")
                    lblUploadStatus.ForeColor = &HAA00&
                    Log ""
                    Log "  [JSON Parsed]"
                    Log "    success   = " & czUpload.JsonStr("success")
                    Log "    server    = " & czUpload.JsonStr("server")
                    Log "    timestamp = " & czUpload.JsonStr("timestamp")
                    '--- Show uploaded file details
                    If czUpload.JsonHas("files") Then
                        Dim vFileKeys As Variant, j As Long
                        vFileKeys = czUpload.JsonGetKeys("files")
                        If IsArray(vFileKeys) Then
                            If UBound(vFileKeys) >= 0 Then
                                For j = 0 To UBound(vFileKeys)
                                    Dim sPrefix As String
                                    sPrefix = "files/" & j
                                    Log "    --- File #" & j + 1 & " ---"
                                    Log "      name = " & czUpload.JsonStr(sPrefix & "/name")
                                    Log "      size = " & czUpload.JsonStr(sPrefix & "/size_human")
                                    Log "      type = " & czUpload.JsonStr(sPrefix & "/type")
                                    Log "      path = " & czUpload.JsonStr(sPrefix & "/path")
                                Next j
                            End If
                        End If
                    End If
                Else
                    '--- Server explicitly returned success=false
                    lblUploadStatus.Caption = "Server error: " & czUpload.JsonStr("error")
                    lblUploadStatus.ForeColor = vbRed
                    Log "! Server returned success=false: " & czUpload.JsonStr("error")
                End If
            Else
                '--- JSON parse failed (m_oJsonRoot is Nothing) - fall back to raw display
                lblUploadStatus.Caption = "Upload OK! (HTTP " & Status & ") - JSON parse failed"
                lblUploadStatus.ForeColor = &HAA00&
                Log "! Note: JSON parse failed, raw body shown above"
            End If
        Else
            lblUploadStatus.Caption = "Upload OK! (HTTP " & Status & ")"
            lblUploadStatus.ForeColor = &HAA00&
        End If
    Else
        lblUploadStatus.Caption = "Failed: HTTP " & Status
        lblUploadStatus.ForeColor = vbRed
        Log "! Upload failed: HTTP " & Status
    End If
End Sub


Private Sub czUpload_Progress(ByVal BytesSent As Long, ByVal BytesTotal As Long, ByVal BytesPerSec As Long, ByVal SecondsRemaining As Long)
    m_lUploadChunkCount = m_lUploadChunkCount + 1
    Dim pct As Single
    If BytesTotal > 0 Then
        pct = (BytesSent / BytesTotal) * 100
    End If
    lblUploadStatus.Caption = "Uploading: " & Format$(pct, "0.0") & "%  " & _
        FormatBytes(BytesSent) & "/" & FormatBytes(BytesTotal) & "  " & _
        FormatSpeed(BytesPerSec)
    If SecondsRemaining > 0 Then
        lblUploadStatus.Caption = lblUploadStatus.Caption & "  ETA: " & FormatETA(SecondsRemaining)
    End If
    lblUploadStatus.ForeColor = &HF0A000
    '--- Log every 100 chunks
    If m_lUploadChunkCount Mod 100 = 0 Then
        Log "  [Upload] " & Format$(pct, "0.0") & "%  " & FormatBytes(BytesSent) & "/" & FormatBytes(BytesTotal) & _
            "  speed=" & FormatSpeed(BytesPerSec)
    End If
End Sub

Private Sub czUpload_Error(ByVal Number As Long, Description As String, _
    ByVal Scode As UcsErrorConstants, Source As String, HelpFile As String, _
    ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    If Number = 0 Then Exit Sub
    btnUpload.Enabled = True
    lblUploadStatus.Caption = "Error: " & Description
    lblUploadStatus.ForeColor = vbRed
    Log "! Upload error " & Number & ": " & Description
End Sub

'=========================================================================
' Download Engine
'=========================================================================

Private Sub StartNextDownloads()
    Dim i As Long
    m_lMaxParallel = CLng(Val(txtMaxParallel.Text))
    If m_lMaxParallel < 1 Then m_lMaxParallel = 1
    If m_lMaxParallel > 20 Then m_lMaxParallel = 20
    For i = 0 To m_lCount - 1
        If m_lActiveCount >= m_lMaxParallel Then Exit For
        If m_tSlots(i).lStatus = DL_WAITING Then
            m_tSlots(i).lStatus = DL_ACTIVE
            m_tSlots(i).dStartTime = Timer
            m_lActiveCount = m_lActiveCount + 1
            On Error Resume Next
            czDl(i).Download m_tSlots(i).sUrl, m_tSlots(i).sSavePath
            If Err.Number <> 0 Then
                m_tSlots(i).lStatus = DL_ERROR
                m_tSlots(i).sError = Err.Description
                m_lActiveCount = m_lActiveCount - 1
                Log "! Start failed: " & m_tSlots(i).sFileName & " - " & Err.Description
            Else
                Log "* Downloading: " & m_tSlots(i).sFileName
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

'=========================================================================
' Drawing - Visual Progress Bars
'=========================================================================

Private Sub DrawProgress()
    Dim i As Long
    Dim y As Long
    Dim lBgColor As Long
    Dim pct As Single
    Dim sLine1 As String
    Dim sLine2 As String
    Dim barLeft As Long
    Dim barRight As Long
    Dim barTop As Long
    Dim barBot As Long
    Dim barW As Long
    Dim blockW As Single
    Dim bx As Long
    Dim lFilled As Long
    Dim lBlockColor As Long
    
    picProgress.Cls
    barLeft = 8
    barRight = picProgress.ScaleWidth - 8
    barW = barRight - barLeft
    
    If m_lCount = 0 Then
        picProgress.ForeColor = &H666666
        picProgress.CurrentX = picProgress.ScaleWidth / 2 - 100
        picProgress.CurrentY = picProgress.ScaleHeight / 2 - 8
        picProgress.Print "No downloads. Add a URL to start."
        Exit Sub
    End If
    
    blockW = CSng(barW - (BLOCK_COUNT - 1)) / CSng(BLOCK_COUNT)
    
    For i = 0 To m_lCount - 1
        y = i * ROW_HEIGHT
        If y + ROW_HEIGHT > picProgress.ScaleHeight Then Exit For
        
        '--- Row background (alternating + selected highlight)
        If i = m_lSelectedSlot Then
            lBgColor = &H3A2820   'blue-tinted for selected
        ElseIf i Mod 2 = 0 Then
            lBgColor = &H282828
        Else
            lBgColor = &H222222
        End If
        picProgress.Line (0, y)-(picProgress.ScaleWidth, y + ROW_HEIGHT - 1), lBgColor, BF
        '--- Selected indicator
        If i = m_lSelectedSlot Then
            picProgress.Line (0, y)-(3, y + ROW_HEIGHT - 1), &HFFAA44, BF
        End If
        
        '--- Separator line
        picProgress.Line (0, y + ROW_HEIGHT - 1)-(picProgress.ScaleWidth, y + ROW_HEIGHT - 1), &H333333
        
        '--- Calculate percentage
        pct = 0
        If m_tSlots(i).lBytesTotal > 0 Then
            pct = m_tSlots(i).lBytesRecv / m_tSlots(i).lBytesTotal
            If pct > 1 Then pct = 1
        End If
        
        '=== LINE 1: Status icon + Filename + % + Speed + ETA ===
        picProgress.CurrentY = y + 3
        
        '--- Status icon
        Select Case m_tSlots(i).lStatus
            Case DL_WAITING
                picProgress.ForeColor = &H888888
                picProgress.CurrentX = barLeft
                picProgress.Print "o";
            Case DL_ACTIVE
                picProgress.ForeColor = &H44DD44
                picProgress.CurrentX = barLeft
                picProgress.Print ">";
            Case DL_PAUSED
                picProgress.ForeColor = &H44DDDD
                picProgress.CurrentX = barLeft
                picProgress.Print "=";
            Case DL_COMPLETE
                picProgress.ForeColor = &H44FF88
                picProgress.CurrentX = barLeft
                picProgress.Print Chr$(251);    'checkmark (√)
            Case DL_ERROR
                picProgress.ForeColor = &H6666FF
                picProgress.CurrentX = barLeft
                picProgress.Print "X";
        End Select
        
        '--- Filename
        picProgress.ForeColor = &HF0D080
        picProgress.CurrentX = barLeft + 16
        picProgress.CurrentY = y + 3
        picProgress.Print m_tSlots(i).sFileName;
        
        '--- Right-aligned: % + speed + ETA
        Select Case m_tSlots(i).lStatus
            Case DL_WAITING
                sLine1 = "Queued"
                picProgress.ForeColor = &H888888
            Case DL_ACTIVE
                sLine1 = Format$(pct * 100, "0.0") & "%"
                If m_tSlots(i).lSpeed > 0 Then
                    sLine1 = sLine1 & "   " & FormatSpeed(m_tSlots(i).lSpeed)
                End If
                If m_tSlots(i).lETA > 0 Then
                    sLine1 = sLine1 & "   ETA: " & FormatETA(m_tSlots(i).lETA)
                End If
                picProgress.ForeColor = &HFFFFFF
            Case DL_PAUSED
                sLine1 = "Paused  " & Format$(pct * 100, "0.0") & "%"
                picProgress.ForeColor = &H44DDDD
            Case DL_COMPLETE
                sLine1 = "Complete!"
                picProgress.ForeColor = &H44FF88
            Case DL_ERROR
                sLine1 = "Error: " & Left$(m_tSlots(i).sError, 30)
                picProgress.ForeColor = &H8888FF
        End Select
        picProgress.CurrentX = barRight - Len(sLine1) * 7
        picProgress.CurrentY = y + 3
        picProgress.Print sLine1
        
        '=== LINE 2: Block mosaic progress bar + size info ===
        barTop = y + 22
        barBot = y + ROW_HEIGHT - 6
        
        '--- Draw blocks
        If m_tSlots(i).lBytesTotal > 0 Then
            lFilled = CLng(pct * BLOCK_COUNT)
        Else
            lFilled = 0
        End If
        
        Dim bxLeft As Long
        Dim bxRight As Long
        Dim b As Long
        For b = 0 To BLOCK_COUNT - 1
            bxLeft = barLeft + CLng(b * (blockW + 1))
            bxRight = barLeft + CLng((b + 1) * (blockW + 1)) - 1
            If bxRight > barRight Then bxRight = barRight
            
            Select Case m_tSlots(i).lStatus
                Case DL_WAITING
                    lBlockColor = &H3A3A3A
                Case DL_ACTIVE
                    If b < lFilled Then
                        lBlockColor = &H00AA44&     'downloaded (green)
                    ElseIf b = lFilled Then
                        lBlockColor = &H008833&     'current block
                    Else
                        lBlockColor = &H3A3A3A     'remaining
                    End If
                Case DL_PAUSED
                    If b < lFilled Then
                        lBlockColor = &H00AAAA&     'downloaded (amber/yellow)
                    Else
                        lBlockColor = &H3A3A3A     'remaining
                    End If
                Case DL_COMPLETE
                    lBlockColor = &HCC66           'bright green
                Case DL_ERROR
                    lBlockColor = &H4444CC         'red
            End Select
            
            picProgress.Line (bxLeft, barTop)-(bxRight, barBot), lBlockColor, BF
        Next b
        
        '--- Size info text below blocks
        If m_tSlots(i).lBytesTotal > 0 Then
            sLine2 = FormatBytes(m_tSlots(i).lBytesRecv) & " / " & FormatBytes(m_tSlots(i).lBytesTotal)
        ElseIf m_tSlots(i).lBytesRecv > 0 Then
            sLine2 = FormatBytes(m_tSlots(i).lBytesRecv) & " / ?"
        Else
            sLine2 = ""
        End If
        
        If Len(sLine2) > 0 Then
            picProgress.ForeColor = &HBBBBBB
            picProgress.CurrentX = barRight - Len(sLine2) * 7
            picProgress.CurrentY = y + 22
            picProgress.Print sLine2
        End If
    Next i
    
    '--- Clear remaining area
    If m_lCount * ROW_HEIGHT < picProgress.ScaleHeight Then
        picProgress.Line (0, m_lCount * ROW_HEIGHT)-(picProgress.ScaleWidth, picProgress.ScaleHeight), &H202020, BF
    End If
End Sub

'=========================================================================
' Helpers
'=========================================================================

Private Function ExtractFileName(sUrl As String) As String
    Dim s As String
    Dim lQ As Long
    '--- Remove query string
    s = sUrl
    lQ = InStr(s, "?")
    If lQ > 0 Then s = Left$(s, lQ - 1)
    '--- Get last path segment
    Dim lSlash As Long
    lSlash = InStrRev(s, "/")
    If lSlash > 0 And lSlash < Len(s) Then
        ExtractFileName = Mid$(s, lSlash + 1)
    Else
        ExtractFileName = "download"
    End If
    '--- URL decode simple %XX
    ExtractFileName = Replace$(ExtractFileName, "%20", " ")
End Function

Private Function TruncText(s As String, lMax As Long) As String
    If Len(s) > lMax Then
        TruncText = Left$(s, lMax - 2) & ".."
    Else
        TruncText = s
    End If
End Function

Private Function FormatBytes(ByVal lBytes As Long) As String
    If lBytes < 0 Then
        '--- Handle VB6 Long overflow for > 2GB
        FormatBytes = "2+ GB"
    ElseIf lBytes < 1024 Then
        FormatBytes = lBytes & " B"
    ElseIf lBytes < 1048576 Then
        FormatBytes = Format$(lBytes / 1024, "0.0") & " KB"
    ElseIf lBytes < 1073741824 Then
        FormatBytes = Format$(lBytes / 1048576, "0.0") & " MB"
    Else
        FormatBytes = Format$(lBytes / 1073741824, "0.0") & " GB"
    End If
End Function

Private Function FormatSpeed(ByVal lBps As Long) As String
    If lBps <= 0 Then
        FormatSpeed = "---"
    Else
        FormatSpeed = FormatBytes(lBps) & "/s"
    End If
End Function

Private Function FormatETA(ByVal lSec As Long) As String
    If lSec <= 0 Then
        FormatETA = ""
    ElseIf lSec < 60 Then
        FormatETA = lSec & "s"
    ElseIf lSec < 3600 Then
        FormatETA = (lSec \ 60) & "m " & (lSec Mod 60) & "s"
    Else
        FormatETA = (lSec \ 3600) & "h " & ((lSec Mod 3600) \ 60) & "m"
    End If
End Function

Private Sub UpdateStatusLabel()
    Dim lWait As Long
    Dim lActive As Long
    Dim lDone As Long
    Dim lErr As Long
    Dim lPaused As Long
    Dim i As Long
    For i = 0 To m_lCount - 1
        Select Case m_tSlots(i).lStatus
            Case DL_WAITING: lWait = lWait + 1
            Case DL_ACTIVE: lActive = lActive + 1
            Case DL_COMPLETE: lDone = lDone + 1
            Case DL_ERROR: lErr = lErr + 1
            Case DL_PAUSED: lPaused = lPaused + 1
        End Select
    Next i
    lblActive.Caption = "Active: " & lActive & " / Queued: " & lWait & " / Done: " & lDone
    If lPaused > 0 Then lblActive.Caption = lblActive.Caption & " / Paused: " & lPaused
    If lErr > 0 Then lblActive.Caption = lblActive.Caption & " / Errors: " & lErr
End Sub

Private Sub Log(sText As String)
    txtLog.Text = txtLog.Text & sText & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
End Sub

'=========================================================================
' Cleanup
'=========================================================================

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    czUpload.Disconnect
    For i = czDl.UBound To 1 Step -1
        czDl(i).Disconnect
        Unload czDl(i)
    Next i
    czDl(0).Disconnect
End Sub
