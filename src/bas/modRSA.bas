Attribute VB_Name = "modRSA"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Type NumMeta
  Digits()   As Integer
  DigitCount As Long
End Type

Public Const HEX_CHARS As String = "0123456789abcdef0123456789ABCDEF"
Public Const MAX_LONG As Long = &H7FFFFFFF
Public Const MIN_LONG As Long = &H80000000

Public ModPowLib As New clsModPow
Public pow2I(15) As Integer

Public Function WrappedAdd(ByVal A As Long, ByVal B As Long) As Long
Dim AtoBase  As Long
Dim dAdd     As Long
Dim dBase    As Long
  If A = 0 Then
    WrappedAdd = B
    Exit Function
  End If
  If B = 0 Then
    WrappedAdd = A
    Exit Function
  End If
  If A > 0 Then
    If B < 0 Then
      WrappedAdd = A + B
      Exit Function
    End If
    AtoBase = MAX_LONG - A
    If AtoBase >= B Then
      WrappedAdd = A + B
      Exit Function
    End If
    dBase = MIN_LONG
    dAdd = B - AtoBase - 1
  Else
    If B > 0 Then
      WrappedAdd = A + B
      Exit Function
    End If
    AtoBase = MIN_LONG - A
    If AtoBase <= B Then
      WrappedAdd = A + B
      Exit Function
    End If
    dBase = MAX_LONG
    dAdd = B - AtoBase + 1
  End If
  WrappedAdd = dBase + dAdd
End Function

Public Function WrappedInt(ByVal L As Long) As Integer
  L = L And &HFFFF&
  If L < &H8000& Then
    WrappedInt = CInt(L)
    Exit Function
  End If
  L = L - &H10000
  WrappedInt = CInt(L)
End Function
