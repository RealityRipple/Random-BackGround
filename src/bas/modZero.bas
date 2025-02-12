Attribute VB_Name = "modZero"
Option Explicit
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function ZeroPrimary() As Boolean
  If LenB(Command$) > 0 Then
    If Command$ = "/0" Then
      ZeroPrimary = False
      Exit Function
    End If
  End If
  If IsTenPlus Then
    ZeroPrimary = False
    Exit Function
  End If
  ZeroPrimary = True
End Function

Private Function IsTenPlus() As Boolean
Dim WinDir As String
Dim DirLen As Long
  WinDir = Space$(&HFF)
  DirLen = GetWindowsDirectoryA(WinDir, &HFF)
  If DirLen < 1 Then
    IsTenPlus = False
    Exit Function
  End If
  WinDir = Left$(WinDir, DirLen)
  If CheckPath(WinDir & "\Web") <> 2 Then
    IsTenPlus = False
    Exit Function
  End If
  If CheckPath(WinDir & "\Web\Wallpaper") <> 2 Then
    IsTenPlus = False
    Exit Function
  End If
  If CheckPath(WinDir & "\Web\4K") = 2 Then
    IsTenPlus = True
    Exit Function
  End If
  If CheckPath(WinDir & "\Web\Wallpaper\Spotlight") = 2 Then
    IsTenPlus = True
    Exit Function
  End If
  IsTenPlus = False
End Function

Private Function CheckPath(ByVal path As String) As Byte
Dim R As VbFileAttribute
  On Error GoTo Erred
  R = GetAttr(path)
  If (R And vbDirectory) = vbDirectory Then
    CheckPath = 2
  Else
    CheckPath = 1
  End If
  Exit Function
Erred:
  CheckPath = 0
End Function

