Attribute VB_Name = "modBase64"
Option Explicit

Private Function DecToBin(ByVal dec As Byte, Optional ByVal bits As Byte = 8) As String
Dim ret As String
  ret = ""
  Do While dec > 0
    ret = dec Mod 2 & ret
    dec = dec \ 2
  Loop
  ret = String$(bits - Len(ret), "0") & ret
  DecToBin = ret
End Function

Private Function BinToDec(ByVal bIn As String) As Byte
Dim I   As Integer
Dim ret As Byte
  ret = 0
  If Len(bIn) <> 8 Then
    BinToDec = 0
    Exit Function
  End If
  For I = 0 To 7
    If Mid$(bIn, 8 - I, 1) = "1" Then ret = ret Or 1 * 2 ^ I
  Next I
  BinToDec = ret
End Function

Public Function Base64Decode(ByVal b64 As String) As Byte()
Dim bChars As String
Dim bIn    As Long
Dim sBin   As String
Dim bOut() As Byte
Dim I      As Integer
Dim J      As Integer
  If LenB(b64) = 0 Then
    Base64Decode = Null
    Exit Function
  End If
  bChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
  J = 0
  For I = 0 To Len(b64) - 1
    bIn = InStr(bChars, Mid$(b64, I + 1, 1))
    If bIn = Null Then
      Base64Decode = Null
      Exit Function
    End If
    If bIn < 1 Or bIn > 65 Then
      Base64Decode = Null
      Exit Function
    End If
    If bIn = 65 Then Exit For
    sBin = sBin & DecToBin(bIn - 1, 6)
    If Len(sBin) >= 8 Then
      ReDim Preserve bOut(J)
      bOut(J) = BinToDec(Left$(sBin, 8))
      sBin = Mid$(sBin, 9)
      J = J + 1
    End If
  Next I
  sBin = ""
  Base64Decode = bOut
End Function
