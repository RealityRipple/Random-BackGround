Attribute VB_Name = "modScreens"
Option Explicit
Public Type Monitor
  Top     As Long
  Left    As Long
  Width   As Long
  Height  As Long
End Type

Private Type DISPLAY_DEVICE
  cb                  As Long
  DeviceName          As String * 32
  DeviceString        As String * 128
  StateFlags          As Long
  DeviceID            As String * 128
  DeviceKey           As String * 128
End Type
Private Type POINTL
  x                   As Long
  y                   As Long
End Type
Private Type DEVMODE
  dmDeviceName        As String * 32
  dmSpecVersion       As Integer
  dmDriverVersion     As Integer
  dmSize              As Integer
  dmDriverExtra       As Integer
  dmFields            As Long
  dmPosition          As POINTL
  dmScale             As Integer
  dmCopies            As Integer
  dmDefaultSource     As Integer
  dmPrintQuality      As Integer
  dmColor             As Integer
  dmDuplex            As Integer
  dmYResolution       As Integer
  dmTTOption          As Integer
  dmCollate           As Integer
  dmFormName          As String * 32
  dmLogPixels         As Integer
  dmBitsPerPel        As Long
  dmPelsWidth         As Long
  dmPelsHeight        As Long
  dmDisplayFlags      As Long
  dmDisplayFrequency  As Long
End Type
Private Type RECT
  Left                As Long
  Top                 As Long
  Right               As Long
  Bottom              As Long
End Type
Private Type MONITORINFO
  cbSize              As Long
  rcMonitor           As RECT
  rcWork              As RECT
  dwFlags             As Long
End Type
Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (ByVal lpDevice As String, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICE, dwFlags As Long) As Long
Private Declare Function EnumDisplaySettingsEx Lib "user32" Alias "EnumDisplaySettingsExA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE, dwFlags As Long) As Long
Private Declare Function MonitorFromPoint Lib "user32" (ByVal ptY As Long, ByVal ptX As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, lpmi As MONITORINFO) As Long

Public Function GetMonitors() As Monitor()
Dim dd      As DISPLAY_DEVICE
Dim dev     As Long
Dim ddMon   As DISPLAY_DEVICE
Dim devMon  As Long
Dim dm      As DEVMODE
Dim hm      As Long
Dim mi      As MONITORINFO
Dim Mons()  As Monitor
Dim mCnt    As Long
  dd.cb = Len(dd)
  dev = 0
  Do While EnumDisplayDevices(vbNullString, dev, dd, 0) <> 0
    If Not CBool(dd.StateFlags And &H8) Then
      ddMon.cb = Len(ddMon)
      devMon = 0
      Do While EnumDisplayDevices(dd.DeviceName, devMon, ddMon, 0) <> 0
        If CBool(ddMon.StateFlags And &H1) Then Exit Do
        devMon = devMon + 1
      Loop
      dm.dmSize = Len(dm)
      If EnumDisplaySettingsEx(dd.DeviceName, -1, dm, 0) = 0 Then EnumDisplaySettingsEx dd.DeviceName, -2, dm, 0
      mi.cbSize = Len(mi)
      If CBool(dd.StateFlags And &H1) Then
        hm = MonitorFromPoint(dm.dmPosition.x, dm.dmPosition.y, 0)
        If hm <> 0 Then GetMonitorInfo hm, mi
        ReDim Preserve Mons(mCnt)
        Mons(mCnt).Top = mi.rcMonitor.Top
        Mons(mCnt).Left = mi.rcMonitor.Left
        Mons(mCnt).Width = dm.dmPelsWidth
        Mons(mCnt).Height = dm.dmPelsHeight
        mCnt = mCnt + 1
      End If
    End If
    dev = dev + 1
  Loop
  GetMonitors = Mons
End Function

Public Function GetMonitorCount() As Integer
Dim Mons() As Monitor
Dim I As Integer
  I = -1
  Mons = GetMonitors
  On Error Resume Next
  I = UBound(Mons)
  On Error GoTo 0
  GetMonitorCount = I + 1
End Function

Public Function GetDisplayProfile() As String
Dim Mons()  As Monitor
Dim I       As Integer
Dim lMons   As Long
Dim sRet    As String
Dim bRet()  As Byte
  On Error GoTo Erred
  Mons = GetMonitors
  lMons = GetMonitorCount
  If lMons = 0 Then
    sRet = "1:"
    sRet = sRet & "["
    sRet = sRet & "0,"
    sRet = sRet & "0 "
    sRet = sRet & Trim$(Str$(Screen.Width / Screen.TwipsPerPixelX)) & "x"
    sRet = sRet & Trim$(Str$(Screen.Height / Screen.TwipsPerPixelY))
    sRet = sRet & "]"
  ElseIf lMons = 1 Then
    sRet = "1:"
    sRet = sRet & "["
    sRet = sRet & Trim$(Str$(Mons(0).Left)) & ","
    sRet = sRet & Trim$(Str$(Mons(0).Top)) & " "
    sRet = sRet & Trim$(Str$(Mons(0).Width)) & "x"
    sRet = sRet & Trim$(Str$(Mons(0).Height))
    sRet = sRet & "]"
  Else
    sRet = Trim$(Str$(lMons)) & ":"
    For I = 0 To lMons - 1
      If I > 0 Then sRet = sRet & ","
      sRet = sRet & "["
      sRet = sRet & Trim$(Str$(Mons(I).Left)) & ","
      sRet = sRet & Trim$(Str$(Mons(I).Top)) & " "
      sRet = sRet & Trim$(Str$(Mons(I).Width)) & "x"
      sRet = sRet & Trim$(Str$(Mons(I).Height))
      sRet = sRet & "]"
    Next I
  End If
  ReDim bRet(Len(sRet) - 1)
  For I = 0 To Len(sRet) - 1
    bRet(I) = Asc(Mid$(sRet, I + 1, 1))
  Next I
  Dim crc As New clsCRC32
  sRet = Hex(crc.GetByteArrayCrc32(bRet))
  Do While Len(sRet) < 8
    sRet = "0" & sRet
  Loop
  GetDisplayProfile = sRet
  Exit Function
Erred:
  GetDisplayProfile = "Settings"
End Function

Public Function GetDisplayDescr() As String()
Dim Mons()  As Monitor
Dim I       As Integer
Dim lMons   As Long
Dim sRet()  As String
  On Error GoTo Erred
  Mons = GetMonitors
  lMons = GetMonitorCount
  If lMons = 0 Then
    ReDim sRet(1)
    sRet(0) = "1 Monitor:"
    sRet(1) = Trim$(Str$(Screen.Width / Screen.TwipsPerPixelX)) & "x"
    sRet(1) = sRet(1) & Trim$(Str$(Screen.Height / Screen.TwipsPerPixelY)) & " at 0,0"
  ElseIf lMons = 1 Then
    ReDim sRet(1)
    sRet(0) = "1 Monitor:"
    sRet(1) = Trim$(Str$(Mons(0).Width)) & "x"
    sRet(1) = sRet(1) & Trim$(Str$(Mons(0).Height)) & " at 0,0"
  Else
    ReDim sRet(lMons)
    sRet(0) = Trim$(Str$(lMons)) & " Monitors:"
    For I = 0 To lMons - 1
      sRet(I + 1) = Trim$(Str$(Mons(I).Width)) & "x"
      sRet(I + 1) = sRet(I + 1) & Trim$(Str$(Mons(I).Height)) & " at "
      sRet(I + 1) = sRet(I + 1) & Trim$(Str$(Mons(I).Left)) & ","
      sRet(I + 1) = sRet(I + 1) & Trim$(Str$(Mons(I).Top))
    Next I
  End If
  GetDisplayDescr = sRet
  Exit Function
Erred:
  ReDim sRet(0)
  sRet(0) = "Failure"
  GetDisplayDescr = sRet
End Function
