Attribute VB_Name = "modINI"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Sub WriteINI(ByVal wiSection As String, ByVal wiKey As String, ByVal wiValue As String, Optional ByVal wiFile As String = vbNullString)
Dim sINI As String
  If LenB(wiFile) > 0 Then
    sINI = wiFile
  Else
    sINI = INIFile
  End If
  WritePrivateProfileString wiSection, wiKey, wiValue, sINI
End Sub

Public Function ReadINI(ByVal riSection As String, ByVal riKey As String, ByVal riDefault As String, Optional ByVal riFile As String = vbNullString)
Dim sRiBuffer As String
Dim sRiValue  As String
Dim sRiLong   As String
Dim sINI      As String
  If LenB(riFile) > 0 Then
    sINI = riFile
  Else
    sINI = INIFile
  End If
  If CheckPath(sINI) <> 1 Then
    ReadINI = riDefault
    Exit Function
  End If
  sRiBuffer = String(255, vbNull)
  sRiLong = GetPrivateProfileString(riSection, riKey, Chr$(1), sRiBuffer, 255, sINI)
  If Left$(sRiBuffer, 1) = Chr$(1) Then
    ReadINI = riDefault
    Exit Function
  End If
  sRiValue = Left$(sRiBuffer, sRiLong)
  If LenB(Trim$(sRiValue)) = 0 Then
    ReadINI = riDefault
    Exit Function
  End If
  ReadINI = sRiValue
End Function

Public Function INIFile()
Dim sDir As String
  sDir = Left$(App.Path, 2)
  If Right$(sDir, 1) = ":" Then
    If CheckPath(sDir & "\" & App.EXEName & ".ini") = 1 Then
      INIFile = sDir & "\" & App.EXEName & ".ini"
      Exit Function
    End If
  End If
  sDir = SettingsFolder
  If Right$(sDir, 1) = ":" Then
    INIFile = sDir & "\" & App.EXEName & ".ini"
    Exit Function
  End If
  INIFile = sDir & "\config.ini"
End Function

Public Function SettingsFolder() As String
Dim sSH As String
  sSH = SHPath(&H1A)
  If LenB(sSH) > 0 Then
    If CheckPath(sSH & "\" & App.CompanyName) <> 2 Then MkDir sSH & "\" & App.CompanyName
    If CheckPath(sSH & "\" & App.CompanyName & "\" & App.ProductName) <> 2 Then MkDir sSH & "\" & App.CompanyName & "\" & App.ProductName
    SettingsFolder = sSH & "\" & App.CompanyName & "\" & App.ProductName
    Exit Function
  End If
  sSH = SHReg("AppData")
  If LenB(sSH) > 0 Then
    If CheckPath(sSH & "\" & App.CompanyName) <> 2 Then MkDir sSH & "\" & App.CompanyName
    If CheckPath(sSH & "\" & App.CompanyName & "\" & App.ProductName) <> 2 Then MkDir sSH & "\" & App.CompanyName & "\" & App.ProductName
    SettingsFolder = sSH & "\" & App.CompanyName & "\" & App.ProductName
    Exit Function
  End If
  SettingsFolder = App.Path
  While Right$(SettingsFolder, 1)
    SettingsFolder = Left$(SettingsFolder, Len(SettingsFolder) - 1)
  Wend
End Function

Public Function PicturesFolder() As String
Dim sSH As String
  sSH = SHPath(&H27)
  If LenB(sSH) > 0 Then
    PicturesFolder = sSH
    Exit Function
  End If
  sSH = SHPath(&H5)
  If LenB(sSH) > 0 Then
    PicturesFolder = sSH
    Exit Function
  End If
  PicturesFolder = App.Path
End Function

Private Function SHPath(ByVal ID As Long) As String
Dim lngRet As Long
Dim sLoc   As String
Dim pidl   As Long
Dim iLoc   As Integer
  lngRet = SHGetSpecialFolderLocation(frmSet.hwnd, ID, pidl)
  If lngRet <> 0 Then
    SHPath = ""
    Exit Function
  End If
  sLoc = Space$(260)
  lngRet = SHGetPathFromIDList(ByVal pidl, sLoc)
  CoTaskMemFree pidl
  If lngRet = 0 Then
    SHPath = ""
    Exit Function
  End If
  If LenB(sLoc) = 0 Then
    SHPath = ""
    Exit Function
  End If
  iLoc = InStr(sLoc, vbNullChar)
  If iLoc = Null Then
    SHPath = ""
    Exit Function
  End If
  If iLoc > 0 Then sLoc = Left$(sLoc, iLoc - 1)
  If LenB(sLoc) = 0 Then
    SHPath = ""
    Exit Function
  End If
  SHPath = sLoc
End Function

Private Function SHReg(ByVal ID As String) As String
Dim sLoc As String
  sLoc = regQuery_Value_SZ(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", ID)
  If LenB(sLoc) = 0 Then
    SHReg = ""
    Exit Function
  End If
  SHReg = sLoc
End Function

Public Function ThemesPath() As String
Dim strLocation As String
  strLocation = regQuery_Value_SZ(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Themes", "CurrentTheme")
  If LenB(strLocation) = 0 Then
    ThemesPath = ""
  Else
    ThemesPath = strLocation
  End If
End Function
