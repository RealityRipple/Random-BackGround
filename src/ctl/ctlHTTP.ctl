VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl HTTP 
   CanGetFocus     =   0   'False
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   570
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlHTTP.ctx":0000
   ScaleHeight     =   38
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   38
   ToolboxBitmap   =   "ctlHTTP.ctx":117A
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock wsHTTP 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Enum HTTPState
  htConnecting
  htConnected
  htDisconnected
  htSending
  htReceiving
End Enum
#If False Then
  Private htConnecting
  Private htConnected
  Private htDisconnected
  Private htSending
  Private htReceiving
#End If
Event DownloadComplete(sData As String)
Event DownloadErrored(sReason As String)
Event CodeError(sSub As String, sDescription As String)
Public lTimeOut As Long
Private CFm_State   As HTTPState
Private RetStr      As String
Private sHost       As String
Private iPort       As Integer
Private sFile       As String
Private Headers()   As String
Private DataLen     As Long
Private sData       As String
Private sDelimit    As String
Private Chunked     As Boolean
Private GotHeader   As Boolean
Private Initiating  As Boolean

Public Sub About()
Attribute About.VB_UserMemId = -552
  MsgBox "RealityRipple's HTTP control ©2007-2025 " & App.CompanyName & "." & vbNewLine & vbNewLine & "Designed for " & App.ProductName & ".", vbInformation + vbOKOnly, "About RealityRipple's HTTP control"
End Sub

Private Sub AddToPacket(ByVal Data As String)
  sData = sData & Data
End Sub

Public Property Get BytesNow() As Long
  BytesNow = Len(sData)
End Property

Public Property Get BytesTotal() As Long
  BytesTotal = DataLen
End Property

Public Property Get HeaderKeys() As String()
Dim sRet() As String
Dim sKey   As String
Dim I      As Integer
  On Error GoTo Erred
  ReDim sRet(UBound(Headers) - 1)
  For I = 1 To UBound(Headers)
    sKey = Headers(I)
    If InStr(sKey, ":") > 0 Then sKey = Left$(sKey, InStr(sKey, ":") - 1)
    sRet(I - 1) = Trim$(sKey)
  Next I
  HeaderKeys = sRet
  Exit Property
Erred:
  Erase sRet
  HeaderKeys = sRet
End Property

Public Property Get HeaderValue(ByVal sFind As String) As String
Dim sKey As String
Dim I     As Integer
  On Error GoTo Erred
  For I = 1 To UBound(Headers)
    sKey = Headers(I)
    If InStr(sKey, ":") > 0 Then sKey = Left$(sKey, InStr(sKey, ":") - 1)
    sKey = Trim$(sKey)
    If LCase$(sFind) = LCase$(sKey) Then
      HeaderValue = Trim$(Mid$(Headers(I), InStr(Headers(I), ":") + 1))
      Exit For
    End If
  Next I
  Exit Property
Erred:
  HeaderValue = ""
End Property

Public Sub Disconnect()
  On Error GoTo Erred
  If wsHTTP.State <> 0 Then wsHTTP.Close
  sData = vbNullString
  sDelimit = vbCrLf
  DataLen = 0
  Chunked = False
  Erase Headers
  GotHeader = False
  CFm_State = htDisconnected
Exit Sub
Erred:
  RaiseEvent CodeError("Disconnect", Err.Description)
  RetStr = "Error: RRHTTP [Disconnect] " & Err.Description
  CFm_State = htDisconnected
End Sub

Private Sub FindDelimit(ByVal Data As String)
  If InStr(Data, vbCrLf) > 0 Then
    sDelimit = vbCrLf
  ElseIf InStr(Data, vbCr) > 0 Then
    sDelimit = vbCr
  ElseIf InStr(Data, vbLf) > 0 Then
    sDelimit = vbLf
  End If
End Sub

Private Sub GetFile()
Dim Request As String
Dim lStart  As Long
  On Error GoTo Erred
  If wsHTTP.State <> 0 Then Disconnect
  lStart = GetTickCount
  wsHTTP.Connect sHost, iPort
  CFm_State = htConnecting
  Do Until wsHTTP.State = 7
    If lTimeOut > 0 Then
      If GetTickCount - lStart > lTimeOut And wsHTTP.State = 4 Then
        RaiseEvent DownloadErrored("The attempt to resolve the host timed out.")
        RetStr = "Error: The attempt to resolve the host timed out."
        CFm_State = htDisconnected
        Exit Sub
      ElseIf GetTickCount - lStart > lTimeOut And wsHTTP.State = 6 Then
        RaiseEvent DownloadErrored("The attempt to connect timed out.")
        RetStr = "Error: The attempt to connect timed out."
        CFm_State = htDisconnected
        Exit Sub
      End If
    End If
    If CFm_State = htDisconnected Then Exit Sub
    DoEvents
  Loop
  sDelimit = vbCrLf
  Request = "GET " & sFile & " HTTP/1.1" & sDelimit & _
            "Host: " & sHost & sDelimit & _
            "User-Agent: none" & sDelimit & _
            "Accept: *" & sDelimit & _
            "Accept-Language: *" & sDelimit & _
            "Accept-Encoding: *" & sDelimit & _
            "Accept-Charset: *" & sDelimit & _
            "Keep-Alive: 300" & sDelimit & _
            "Cache-Control: no-cache" & sDelimit & _
            "Connection: keep-alive" & sDelimit & sDelimit
  wsHTTP.SendData Request
Exit Sub
Erred:
  RaiseEvent CodeError("GetFile", Err.Description)
  RetStr = "Error: RRHTTP [GetFile] " & Err.Description
  CFm_State = htDisconnected
End Sub

Public Function GetURL(ByVal URL As String, Optional ByVal Port As Integer = 80) As String
Dim lStart  As Long
Dim Retry   As Boolean
  On Error GoTo Erred
  If LCase$(Left$(URL, 7)) = "http://" Then URL = Mid$(URL, 8)
  URL = Replace$(URL, "\", "/")
  URL = Replace$(URL, " ", "%20")
  If InStr(URL, "/") > 0 Then
    RetStr = vbNullString
    sHost = Left$(URL, InStr(URL, "/") - 1)
    sFile = Mid$(URL, InStr(URL, "/"))
    iPort = Port
    Retry = False
DoAgain:
    GetFile
    DoEvents
    lStart = GetTickCount
    Initiating = True
    Do Until LenB(RetStr) > 0
      If lTimeOut > 0 Then
        If GetTickCount - lStart > lTimeOut And Initiating Then
          If Not Retry Then
            Retry = True
            GoTo DoAgain
          End If
          RaiseEvent DownloadErrored("The attempt to connect timed out.")
          RetStr = "Error: The attempt to connect timed out."
          CFm_State = htDisconnected
        End If
      End If
      If CFm_State = htDisconnected Then Exit Do
      DoEvents
    Loop
    If wsHTTP.State <> 0 Then wsHTTP.Close
    sData = vbNullString
    sDelimit = vbCrLf
    DataLen = 0
    Chunked = False
    Erase Headers
    GotHeader = False
    CFm_State = htDisconnected
    If Left$(RetStr, 7) = "Error: " Then
      On Error GoTo 0
      Err.Raise 8358, "RRHTTP", Mid$(RetStr, 8)
      On Error GoTo Erred
    End If
    GetURL = RetStr
    RetStr = vbNullString
  End If
Exit Function
Erred:
  RaiseEvent CodeError("GetURL", Err.Description)
  RetStr = "Error: RRHTTP [GetURL] " & Err.Description
  CFm_State = htDisconnected
End Function

Private Sub HandleData(ByVal Data As String)
Dim I               As Integer
Dim ChunkLen        As Long
Dim cData           As String
Dim sChunk          As String
  On Error GoTo Erred
  If Not GotHeader Then RetrieveHeader Data
  If Mid$(Headers(0), 10, 3) = "301" Then
    For I = 0 To UBound(Headers)
      If Left$(Headers(I), 10) = "Location: " Then
        sFile = Mid$(Mid$(Headers(I), 11), Len(sHost) + 8)
        GetFile
        Exit Sub
      End If
    Next I
  End If
  AddToPacket Data
  If Len(sData) < DataLen And DataLen > 0 Then Exit Sub
  If Chunked Then
    cData = sData
    sData = vbNullString
ReRun:
    ChunkLen = Val("&H" & Left$(cData, 3))
    cData = Mid$(cData, InStr(cData, sDelimit) + 2)
    If ChunkLen > 0 Then
      sChunk = Left$(cData, ChunkLen)
      AddToPacket sChunk
      cData = Mid$(cData, Len(sChunk) + 2)
      GoTo ReRun
    End If
  End If
  HandleEnd
Exit Sub
Erred:
  RaiseEvent CodeError("HandleData", Err.Description)
  RetStr = "Error: RRHTTP [HandleData] " & Err.Description
  CFm_State = htDisconnected
End Sub

Private Sub HandleEnd()
Dim ErrNo   As String
Dim errstr  As String
  On Error GoTo Erred
  If Mid$(Headers(0), 10, 3) = "200" Then
    RetStr = sData
    RaiseEvent DownloadComplete(sData)
  Else
    ErrNo = Mid$(Headers(0), 10, 3)
    errstr = Mid$(Headers(0), 14)
    RetStr = "Error: " & ErrNo & " " & errstr
    RaiseEvent DownloadErrored(ErrNo & ": " & errstr)
  End If
Exit Sub
Erred:
  RaiseEvent CodeError("HandleEnd", Err.Description)
  RetStr = "Error: RRHTTP [HandleEnd] " & Err.Description
  CFm_State = htDisconnected
End Sub

Public Sub OpenURL(ByVal URL As String, Optional ByVal Port As Integer = 80)
  On Error GoTo Erred
  If LCase$(Left$(URL, 7)) = "http://" Then URL = Mid$(URL, 8)
  URL = Replace$(URL, "\", "/")
  URL = Replace$(URL, " ", "%20")
  If InStr(URL, "/") > 0 Then
    RetStr = vbNullString
    sHost = Left$(URL, InStr(URL, "/") - 1)
    sFile = Mid$(URL, InStr(URL, "/"))
    iPort = Port
    GetFile
  End If
Exit Sub
Erred:
  RaiseEvent CodeError("OpenURL", Err.Description)
  CFm_State = htDisconnected
End Sub

Private Sub RetrieveHeader(ByRef Data As String)
Dim I       As Long
Dim DHeader As String
  On Error GoTo Erred
  FindDelimit Data
  DHeader = Data
  If InStr(DHeader, sDelimit & sDelimit) > 0 Then
    If Left$(Split(Left$(DHeader, InStr(DHeader, sDelimit & sDelimit) - 1), sDelimit)(0), 7) = "HTTP/1." Then
      Headers() = Split(Left$(DHeader, InStr(DHeader, sDelimit & sDelimit) - 1), sDelimit)
      GotHeader = True
    End If
  End If
  For I = 0 To UBound(Headers)
    If LenB(Headers(I)) > 16 Then
      If Left$(Headers(I), 16) = "Content-Length: " Then
        DataLen = Mid$(Headers(I), 17)
      End If
      If Left$(Headers(I), 26) = "Transfer-Encoding: chunked" Then
        Chunked = True
      End If
    End If
  Next I
  Data = Mid$(Data, InStr(Data, sDelimit & sDelimit) + 4)
Exit Sub
Erred:
  RaiseEvent CodeError("RetrieveHeader", Err.Description)
  RetStr = "Error: RRHTTP [RetrieveHeader] " & Err.Description
  CFm_State = htDisconnected
End Sub

Public Property Get State() As HTTPState
  State = CFm_State
End Property

Private Sub UserControl_Initialize()
  If UserControl.Width <> 570 Then UserControl.Width = 570
  If UserControl.Height <> 570 Then UserControl.Height = 570
  lTimeOut = 10000
  CFm_State = htDisconnected
End Sub

Private Sub UserControl_Paint()
  If UserControl.Width <> 570 Then UserControl.Width = 570
  If UserControl.Height <> 570 Then UserControl.Height = 570
End Sub

Private Sub UserControl_Resize()
  If UserControl.Width <> 570 Then UserControl.Width = 570
  If UserControl.Height <> 570 Then UserControl.Height = 570
End Sub

Private Sub wsHTTP_Close()
  Disconnect
End Sub

Private Sub wsHTTP_Connect()
  CFm_State = htConnected
End Sub

Private Sub wsHTTP_DataArrival(ByVal BytesTotal As Long)
Dim sData As String
  CFm_State = htReceiving
  wsHTTP.GetData sData, vbString, BytesTotal
  If sData = "0" & vbNewLine & vbNewLine Then Exit Sub
  HandleData sData
End Sub

Private Sub wsHTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  RetStr = "Error: " & Description
  CFm_State = htDisconnected
End Sub

Private Sub wsHTTP_SendComplete()
  CFm_State = htConnected
End Sub

Private Sub wsHTTP_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  CFm_State = htSending
End Sub

