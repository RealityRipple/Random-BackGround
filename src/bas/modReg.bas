Attribute VB_Name = "modReg"
Option Explicit
Private m_lngRetVal             As Long
Private Const REG_SZ            As Long = 1
Private Const REG_DWORD         As Long = 4
Public Const HKEY_CLASSES_ROOT  As Long = &H80000000
Public Const HKEY_CURRENT_USER  As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const ERROR_SUCCESS     As Long = 0
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Sub regDelete_Sub_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String)
Dim lngKeyHandle As Long
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
    m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
    m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
End Sub
Public Sub regDelete_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then m_lngRetVal = RegDeleteKey(lngRootKey, strRegKeyPath)
End Sub
Public Function regQuery_Value_SZ(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String) As Variant
Dim intPosition   As Integer
Dim lngKeyHandle  As Long
Dim lngDataType   As Long
Dim lngBufferSize As Long
Dim lngBuffer     As Long
Dim strBuffer     As String
  lngKeyHandle = 0
  lngBufferSize = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
    regQuery_Value_SZ = ""
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)
  If lngKeyHandle = 0 Then
    regQuery_Value_SZ = ""
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  If lngDataType <> REG_SZ Then
    regQuery_Value_SZ = ""
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  strBuffer = Space(lngBufferSize)
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, ByVal strBuffer, lngBufferSize)
  If m_lngRetVal <> ERROR_SUCCESS Then
    regQuery_Value_SZ = ""
  Else
    intPosition = InStr(1, strBuffer, Chr(0))
    If intPosition > 0 Then
      regQuery_Value_SZ = Left(strBuffer, intPosition - 1)
    Else
      regQuery_Value_SZ = strBuffer
    End If
  End If
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function
Public Function regQuery_Value_DWORD(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String) As Variant
Dim intPosition   As Integer
Dim lngKeyHandle  As Long
Dim lngDataType   As Long
Dim lngBufferSize As Long
Dim lngBuffer     As Long
Dim strBuffer     As String
  lngKeyHandle = 0
  lngBufferSize = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
    regQuery_Value_DWORD = 0
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)
  If lngKeyHandle = 0 Then
    regQuery_Value_DWORD = 0
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  If lngDataType <> REG_DWORD Then
    regQuery_Value_DWORD = 0
    m_lngRetVal = RegCloseKey(lngKeyHandle)
    Exit Function
  End If
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngBuffer, 4&)
  If m_lngRetVal <> ERROR_SUCCESS Then
    regQuery_Value_DWORD = 0
  Else
    regQuery_Value_DWORD = lngBuffer
  End If
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function
Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, ByVal strRegKeyPath As String) As Boolean
Dim lngKeyHandle As Long
  lngKeyHandle = 0
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle = 0 Then
    regDoes_Key_Exist = False
  Else
    regDoes_Key_Exist = True
  End If
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function
Public Sub regCreate_Value_SZ(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String, varRegData As Variant)
Dim lngKeyHandle  As Long
Dim lngDataType   As Long
Dim strKeyValue   As String
  lngDataType = REG_SZ
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  strKeyValue = Trim$(varRegData) & Chr(0)
  m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal strKeyValue, Len(strKeyValue))
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Sub
Public Sub regCreate_Value_DWORD(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, ByVal strRegSubKey As String, varRegData As Long)
Dim lngKeyHandle  As Long
Dim lngDataType   As Long
Dim strKeyValue   As String
  lngDataType = REG_DWORD
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal varRegData, 4)
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Sub
Public Sub regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)
Dim lngKeyHandle As Long
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Sub
