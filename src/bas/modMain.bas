Attribute VB_Name = "modMain"
Option Explicit
Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Public Enum bgPOSITION
  Fill
  Fit
  Stretch
  Tile
  Center
End Enum
Public Enum bgMAXSCALE
  Unlimited = 0
  x50
  x25
  x10
  x5
  x3
  X2
  X1_5
  X1
End Enum
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function SendMessageTimeoutA Lib "user32" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef result As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal zeroOnly As Long) As Long
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Type OSVERSIONINFOEX
  dwOSVersionInfoSize         As Long
  dwMajorVersion              As Long
  dwMinorVersion              As Long
  dwBuildNumber               As Long
  dwPlatformID                As Long
  szCSDVersion                As String * 128
  dwServicePackMajor          As Integer
  dwServicePackMinor          As Integer
  wSuiteMask                  As Integer
  wProductType                As Byte
  wReserved                   As Byte
End Type

Public Function InitCommonControlsVB() As Boolean
  On Error Resume Next
  Dim iccex As tagInitCommonControlsEx
  ' Ensure CC available:
  With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_USEREX_CLASSES
  End With
  InitCommonControlsEx iccex
  InitCommonControlsVB = (Err.Number = 0)
  On Error GoTo 0
End Function

Public Sub Main()
Dim I As Long
  InitCommonControlsVB

  On Error GoTo Erred
  For I = 0 To 14
    pow2I(I) = 2 ^ I
  Next I
  pow2I(15) = &H8000

  If LenB(Command$) > 0 Then
    If Left$(Command$, 5) = "/set " Then
      If LenB(Dir$(Mid$(Command$, 6), vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then
        Load frmSet
        frmSet.NewBackground Mid$(Command$, 6)
      End If
      End
    ElseIf Command$ = "/next" Then
      Load frmSet
      frmSet.NewBackground
      End
    End If
  End If
  Load frmSet
  frmSet.Hide
  Exit Sub
Erred:
  frmNotify.Notify "Error in Startup: " & Err.Description
End Sub

Public Function HasSysAssoc() As Boolean
Dim osInfo As OSVERSIONINFOEX
  On Error GoTo Erred
  If Not GetOSInfo(osInfo) Then
    HasSysAssoc = False
    Exit Function
  End If
  HasSysAssoc = Not osInfo.dwMajorVersion < 6
  Exit Function
Erred:
  HasSysAssoc = False
End Function

Public Function CanSmooth() As Boolean
Dim osInfo As OSVERSIONINFOEX
  On Error GoTo Erred
  If Not GetOSInfo(osInfo) Then
    CanSmooth = False
    Exit Function
  End If
  If osInfo.dwMajorVersion < 6 Then
    CanSmooth = False
    Exit Function
  End If
  If osInfo.dwMajorVersion = 6 And osInfo.dwMinorVersion = 0 Then
    CanSmooth = False
    Exit Function
  End If
  CanSmooth = True
  Dim hWndow As Long
  hWndow = FindWindowA("Progman", 0)
  If hWndow = 0 Then Exit Function
  Dim result As Long
  SendMessageTimeoutA hWndow, &H52C, 0, 0, 0, 500, result
  Exit Function
Erred:
  CanSmooth = False
End Function

Public Function HasMultimon() As Boolean
Dim osInfo As OSVERSIONINFOEX
  On Error GoTo Erred
  If Not GetOSInfo(osInfo) Then
    HasMultimon = False
    Exit Function
  End If
  If osInfo.dwPlatformID > 1 Then
    HasMultimon = True
    Exit Function
  End If
  If osInfo.dwMajorVersion > 4 Then
    HasMultimon = True
    Exit Function
  End If
  HasMultimon = osInfo.dwMinorVersion > 0
  Exit Function
Erred:
  HasMultimon = False
End Function

Public Sub SetBG(Optional ByVal AltVal As String = "")
Dim OldStyle As Boolean
  On Error GoTo Erred
  SetWallpaperStyle 1
  If ReadINI("Settings", "Smooth", "N") <> "Y" Then
    OldStyle = True
  Else
    OldStyle = Not CanSmooth
  End If
  If Not OldStyle Then
    If Not ActiveDesktopSetWallpaper(SettingsFolder & "\RandomBG" & AltVal & ".bmp") Then
      OldStyle = True
      WriteINI "Settings", "Smooth", "N"
    End If
  End If
  If OldStyle Then SystemParametersInfoA 20, 0&, SettingsFolder & "\RandomBG" & AltVal & ".bmp", &H1 Or &H2
  Exit Sub
Erred:
  frmNotify.Notify "Error in SetBG: " & Err.Description
End Sub

Public Function CheckPath(ByVal Path As String) As Byte
Dim r As VbFileAttribute
  On Error GoTo Erred
  If LenB(Dir$(Path, vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then
    CheckPath = 0
    Exit Function
  End If
  r = GetAttr(Path)
  If (r And vbDirectory) = vbDirectory Then
    CheckPath = 2
  Else
    CheckPath = 1
  End If
  Exit Function
Erred:
  CheckPath = 0
End Function

Private Sub SetWallpaperStyle(ByVal Style As Integer)
  On Error GoTo Erred
  Select Case Style
    Case 0
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0 "
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0 "
    Case 1
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1 "
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "1 "
    Case 2
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0 "
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2 "
    Case Else
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0 "
      regCreate_Value_SZ HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0 "
  End Select
  Exit Sub
Erred:
  frmNotify.Notify "Error in SetWallpaperStyle: " & Err.Description
End Sub

Private Function GetOSInfo(ByRef VerInfo As OSVERSIONINFOEX) As Boolean
Dim udtVerinfo  As OSVERSIONINFOEX
Dim strInfo     As String
  On Error GoTo Erred
  udtVerinfo.szCSDVersion = Space$(128)
  udtVerinfo.dwOSVersionInfoSize = Len(udtVerinfo)
  If GetVersionExA(udtVerinfo) Then
    VerInfo = udtVerinfo
    GetOSInfo = True
  ElseIf GetVersion <> 0 Then
    strInfo = Hex$(GetVersion)
    strInfo = String$(8 - Len(strInfo), "0") & strInfo
    If Mid$(strInfo, 1, 4) <> "C000" Then udtVerinfo.dwBuildNumber = Val("&H" & Mid$(strInfo, 1, 4))
    udtVerinfo.dwMinorVersion = Val("&H" & Mid$(strInfo, 5, 2))
    udtVerinfo.dwMajorVersion = Val("&H" & Mid$(strInfo, 7, 2))
    VerInfo = udtVerinfo
    GetOSInfo = True
  Else
    GetOSInfo = False
  End If
  Exit Function
Erred:
  GetOSInfo = False
End Function

