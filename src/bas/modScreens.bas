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
      End If
      If CBool(dd.StateFlags And &H1) Then
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
