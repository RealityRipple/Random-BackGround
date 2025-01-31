VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private crc32Table()  As Long

Private Sub Class_Initialize()
Dim dwPolynomial  As Long
Dim I             As Integer
Dim J             As Integer
Dim dwCrc         As Long
  dwPolynomial = &HEDB88320
  ReDim crc32Table(256)
  For I = 0 To 255
    dwCrc = I
    For J = 8 To 1 Step -1
      If dwCrc And 1 Then
        dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
        dwCrc = dwCrc Xor dwPolynomial
      Else
        dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
      End If
    Next J
    crc32Table(I) = dwCrc
  Next I
End Sub

Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
Dim crc32Result As Long
Dim I           As Long
Dim iLookup     As Integer
  crc32Result = &HFFFFFFFF
  For I = LBound(buffer) To UBound(buffer)
    iLookup = (crc32Result And &HFF) Xor buffer(I)
    crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
    crc32Result = crc32Result Xor crc32Table(iLookup)
  Next I
  GetByteArrayCrc32 = Not (crc32Result)
End Function
