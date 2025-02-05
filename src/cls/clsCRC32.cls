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

Const CRC_SEED As Long = &HEDB88320

Private crcTable(255) As Long

Private Sub Class_Initialize()
Dim I As Integer
Dim J As Integer
Dim L As Long
  For I = 0 To 255
    L = I
    For J = 8 To 1 Step -1
      If L And 1 Then
        L = ((L And &HFFFFFFFE) \ 2&) And MAX_LONG
        L = L Xor CRC_SEED
      Else
        L = ((L And &HFFFFFFFE) \ 2&) And MAX_LONG
      End If
    Next J
    crcTable(I) = L
  Next I
End Sub

Public Function Check(ByRef buffer() As Byte) As Long
Dim ret As Long
Dim idx As Integer
Dim I   As Long
  ret = &HFFFFFFFF
  For I = LBound(buffer) To UBound(buffer)
    idx = (ret And &HFF) Xor buffer(I)
    ret = ((ret And &HFFFFFF00) \ &H100) And &HFFFFFF
    ret = ret Xor crcTable(idx)
  Next I
  Check = Not ret
End Function
