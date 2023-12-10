Attribute VB_Name = "modINI"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ppidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Sub WriteINI(wiSection As String, wiKey As String, wiValue As String, wiFile As String)
Dim INIFile As String
  If LenB(Dir$(wiFile)) = 0 Then
    INIFile = SettingsFolder & "\" & wiFile
  Else
    INIFile = wiFile
  End If
  WritePrivateProfileString wiSection, wiKey, wiValue, INIFile
End Sub
Public Function ReadINI(riSection As String, riKey As String, riFile As String, riDefault As String)
Dim sRiBuffer As String
Dim sRiValue  As String
Dim sRiLong   As String
Dim INIFile   As String
  If LenB(Dir$(riFile)) = 0 Then
    INIFile = SettingsFolder & "\" & riFile
  Else
    INIFile = riFile
  End If
  If Dir(INIFile) <> "" Then
    sRiBuffer = String(255, vbNull)
    sRiLong = GetPrivateProfileString(riSection, riKey, Chr$(1), sRiBuffer, 255, INIFile)
    If Left$(sRiBuffer, 1) <> Chr$(1) Then
      sRiValue = Left$(sRiBuffer, sRiLong)
      If sRiValue <> "" Then
        ReadINI = sRiValue
      Else
        ReadINI = riDefault
      End If
    Else
      ReadINI = riDefault
    End If
  Else
    ReadINI = riDefault
  End If
End Function
Public Function SettingsFolder() As String
Dim lngRet      As Long
Dim strLocation As String
Dim pidl        As Long
  lngRet = SHGetSpecialFolderLocation(frmSet.hwnd, &H1A, pidl)
  If lngRet = 0 Then
    strLocation = Space$(260)
    lngRet = SHGetPathFromIDList(ByVal pidl, strLocation)
    If lngRet = 0 Or LenB(Trim$(strLocation)) = 0 Then
      SettingsFolder = App.Path
    Else
      SettingsFolder = Left$(strLocation, InStr(strLocation, vbNullChar) - 1) & "\RealityRipple Software\Random BackGround"
      If LenB(Dir$(Left$(strLocation, InStr(strLocation, vbNullChar) - 1) & "\RealityRipple Software\", vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then MkDir Left$(strLocation, InStr(strLocation, vbNullChar) - 1) & "\RealityRipple Software\"
      If LenB(Dir$(SettingsFolder, vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then MkDir SettingsFolder
    End If
    CoTaskMemFree pidl
  End If
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
