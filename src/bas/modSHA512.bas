Attribute VB_Name = "modSHA512"
Option Explicit

Private Enum LongPtr
  [_]
End Enum

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr

Private Const LNG_BLOCKSZ As Long = 128
Private Const LNG_ROUNDS  As Long = 80
Private Const LNG_POW2_1  As Long = 2 ^ 1
Private Const LNG_POW2_2  As Long = 2 ^ 2
Private Const LNG_POW2_3  As Long = 2 ^ 3
Private Const LNG_POW2_4  As Long = 2 ^ 4
Private Const LNG_POW2_5  As Long = 2 ^ 5
Private Const LNG_POW2_6  As Long = 2 ^ 6
Private Const LNG_POW2_7  As Long = 2 ^ 7
Private Const LNG_POW2_8  As Long = 2 ^ 8
Private Const LNG_POW2_9  As Long = 2 ^ 9
Private Const LNG_POW2_12 As Long = 2 ^ 12
Private Const LNG_POW2_13 As Long = 2 ^ 13
Private Const LNG_POW2_14 As Long = 2 ^ 14
Private Const LNG_POW2_17 As Long = 2 ^ 17
Private Const LNG_POW2_18 As Long = 2 ^ 18
Private Const LNG_POW2_19 As Long = 2 ^ 19
Private Const LNG_POW2_22 As Long = 2 ^ 22
Private Const LNG_POW2_23 As Long = 2 ^ 23
Private Const LNG_POW2_24 As Long = 2 ^ 24
Private Const LNG_POW2_25 As Long = 2 ^ 25
Private Const LNG_POW2_26 As Long = 2 ^ 26
Private Const LNG_POW2_27 As Long = 2 ^ 27
Private Const LNG_POW2_28 As Long = 2 ^ 28
Private Const LNG_POW2_29 As Long = 2 ^ 29
Private Const LNG_POW2_30 As Long = 2 ^ 30
Private Const LNG_POW2_31 As Long = &H80000000

Private Type SAFEARRAY1D
  cDims      As Integer
  fFeatures  As Integer
  cbElements As Long
  cLocks     As Long
  pvData     As LongPtr
  cElements  As Long
  lLbound    As Long
End Type

Private Type ArrayLong16
  Item(0 To 15) As Long
End Type

Private Type ArrayLong32
  Item(0 To 31) As Long
End Type

Public Type CryptoSha512Context
  State      As ArrayLong16
  Block      As ArrayLong32
  bytes()    As Byte
  ArrayBytes As SAFEARRAY1D
  NPartial   As Long
  NInput     As Currency
  BitSize    As Long
End Type

Private LNG_K(0 To 2 * LNG_ROUNDS - 1) As Long

Private Function BSwap32(ByVal lX As Long) As Long
  BSwap32 = (lX And &H7F) * &H1000000 Or (lX And &HFF00&) * &H100 Or (lX And &HFF0000) \ &H100 Or (lX And &HFF000000) \ &H1000000 And &HFF Or -((lX And &H80) <> 0) * MIN_LONG
End Function

Private Sub pvAdd64(lAL As Long, lAH As Long, ByVal lBL As Long, ByVal lBH As Long)
Dim lSign As Long
  If (lAL Xor lBL) >= 0 Then
    lAL = ((lAL Xor MIN_LONG) + lBL) Xor MIN_LONG
  Else
    lAL = lAL + lBL
  End If
  If (lAH Xor lBH) >= 0 Then
    lAH = ((lAH Xor MIN_LONG) + lBH) Xor MIN_LONG
  Else
    lAH = lAH + lBH
  End If
  If (lAL And MIN_LONG) <> 0 Then lSign = 1
  If (lBL And MIN_LONG) <> 0 Then lSign = lSign - 1
  Select Case True
    Case lSign < 0, lSign = 0 And (lAL And MAX_LONG) < (lBL And MAX_LONG)
      If lAH >= 0 Then
        lAH = ((lAH Xor MIN_LONG) + 1) Xor MIN_LONG
      Else
        lAH = lAH + 1
      End If
  End Select
End Sub

Private Function pvSum0L(ByVal lX As Long, ByVal lY As Long) As Long
  pvSum0L = ((lX And (LNG_POW2_6 - 1)) * LNG_POW2_25 Or -((lX And LNG_POW2_6) <> 0) * MIN_LONG) _
        Xor ((lX And (LNG_POW2_1 - 1)) * LNG_POW2_30 Or -((lX And LNG_POW2_1) <> 0) * MIN_LONG) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_28 Or -(lX < 0) * LNG_POW2_3) _
        Xor ((lY And MAX_LONG) \ LNG_POW2_7 Or -(lY < 0) * LNG_POW2_24) _
        Xor ((lY And MAX_LONG) \ LNG_POW2_2 Or -(lY < 0) * LNG_POW2_29) _
        Xor ((lY And (LNG_POW2_27 - 1)) * LNG_POW2_4 Or -((lY And LNG_POW2_27) <> 0) * MIN_LONG)
End Function

Private Function pvSum1L(ByVal lX As Long, ByVal lY As Long) As Long
  pvSum1L = ((lX And (LNG_POW2_8 - 1)) * LNG_POW2_23 Or -((lX And LNG_POW2_8) <> 0) * MIN_LONG) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_14 Or -(lX < 0) * LNG_POW2_17) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_18 Or -(lX < 0) * LNG_POW2_13) _
        Xor ((lY And MAX_LONG) \ LNG_POW2_9 Or -(lY < 0) * LNG_POW2_22) _
        Xor ((lY And (LNG_POW2_13 - 1)) * LNG_POW2_18 Or -((lY And LNG_POW2_13) <> 0) * MIN_LONG) _
        Xor ((lY And (LNG_POW2_17 - 1)) * LNG_POW2_14 Or -((lY And LNG_POW2_17) <> 0) * MIN_LONG)
End Function

Private Function pvSig0L(ByVal lX As Long, ByVal lY As Long) As Long
  pvSig0L = ((lX And MAX_LONG) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_7 Or -(lX < 0) * LNG_POW2_24) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) _
        Xor ((lY And 0) * LNG_POW2_31 Or -((lY And 1) <> 0) * MIN_LONG) _
        Xor ((lY And (LNG_POW2_6 - 1)) * LNG_POW2_25 Or -((lY And LNG_POW2_6) <> 0) * MIN_LONG) _
        Xor ((lY And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lY And LNG_POW2_7) <> 0) * MIN_LONG)
End Function
  
Private Function pvSig0H(ByVal lX As Long, ByVal lY As Long) As Long
  pvSig0H = ((lX And MAX_LONG) \ LNG_POW2_1 Or -(lX < 0) * LNG_POW2_30) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_7 Or -(lX < 0) * LNG_POW2_24) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_8 Or -(lX < 0) * LNG_POW2_23) _
        Xor ((lY And 0) * LNG_POW2_31 Or -((lY And 1) <> 0) * MIN_LONG) _
        Xor ((lY And (LNG_POW2_7 - 1)) * LNG_POW2_24 Or -((lY And LNG_POW2_7) <> 0) * MIN_LONG)
End Function

Private Function pvSig1L(ByVal lX As Long, ByVal lY As Long) As Long
  pvSig1L = ((lX And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lX And LNG_POW2_28) <> 0) * MIN_LONG) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_6 Or -(lX < 0) * LNG_POW2_25) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_19 Or -(lX < 0) * LNG_POW2_12) _
        Xor ((lY And MAX_LONG) \ LNG_POW2_29 Or -(lY < 0) * LNG_POW2_2) _
        Xor ((lY And (LNG_POW2_5 - 1)) * LNG_POW2_26 Or -((lY And LNG_POW2_5) <> 0) * MIN_LONG) _
        Xor ((lY And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lY And LNG_POW2_18) <> 0) * MIN_LONG)
End Function

Private Function pvSig1H(ByVal lX As Long, ByVal lY As Long) As Long
  pvSig1H = ((lX And (LNG_POW2_28 - 1)) * LNG_POW2_3 Or -((lX And LNG_POW2_28) <> 0) * MIN_LONG) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_6 Or -(lX < 0) * LNG_POW2_25) _
        Xor ((lX And MAX_LONG) \ LNG_POW2_19 Or -(lX < 0) * LNG_POW2_12) _
        Xor ((lY And MAX_LONG) \ LNG_POW2_29 Or -(lY < 0) * LNG_POW2_2) _
        Xor ((lY And (LNG_POW2_18 - 1)) * LNG_POW2_13 Or -((lY And LNG_POW2_18) <> 0) * MIN_LONG)
End Function

Private Sub pvRound(ByVal lX00 As Long, ByVal lX01 As Long, ByVal lX02 As Long, ByVal lX03 As Long, _
                    ByVal lX04 As Long, ByVal lX05 As Long, ByRef lX06 As Long, ByRef lX07 As Long, _
                    ByVal lX08 As Long, ByVal lX09 As Long, ByVal lX10 As Long, ByVal lX11 As Long, _
                    ByVal lX12 As Long, ByVal lX13 As Long, ByRef lX14 As Long, ByRef lX15 As Long, _
                    ByRef uArray As ArrayLong32, ByVal lIdx As Long, ByVal lJdx As Long)
  pvAdd64 lX14, lX15, uArray.Item(lIdx), uArray.Item(lIdx + 1)
  pvAdd64 lX14, lX15, LNG_K(lJdx + lIdx), LNG_K(lJdx + lIdx + 1)
  pvAdd64 lX14, lX15, lX12 Xor (lX08 And (lX10 Xor lX12)), lX13 Xor (lX09 And (lX11 Xor lX13))
  pvAdd64 lX14, lX15, pvSum1L(lX08, lX09), pvSum1L(lX09, lX08)
  pvAdd64 lX06, lX07, lX14, lX15
  pvAdd64 lX14, lX15, pvSum0L(lX00, lX01), pvSum0L(lX01, lX00)
  pvAdd64 lX14, lX15, ((lX00 Or lX04) And lX02) Or (lX04 And lX00), ((lX01 Or lX05) And lX03) Or (lX05 And lX01)
End Sub

Private Sub pvStore(ByRef uArray As ArrayLong32, ByVal lIdx As Long)
Dim lTL As Long
Dim lTH As Long
Dim lUL As Long
Dim lUH As Long
  With uArray
    lTL = .Item(lIdx)
    lTH = .Item(lIdx + 1)
    pvAdd64 lTL, lTH, .Item((lIdx + 18) And &H1F), .Item((lIdx + 19) And &H1F)
    lUL = pvSig0L(.Item((lIdx + 2) And &H1F), .Item((lIdx + 3) And &H1F))
    lUH = pvSig0H(.Item((lIdx + 3) And &H1F), .Item((lIdx + 2) And &H1F))
    pvAdd64 lTL, lTH, lUL, lUH
    lUL = pvSig1L(.Item((lIdx + 28) And &H1F), .Item((lIdx + 29) And &H1F))
    lUH = pvSig1H(.Item((lIdx + 29) And &H1F), .Item((lIdx + 28) And &H1F))
    pvAdd64 lTL, lTH, lUL, lUH
    .Item(lIdx) = lTL
    .Item(lIdx + 1) = lTH
  End With
End Sub

Private Sub CryptoSha512Init(ByRef uCtx As CryptoSha512Context)
Dim pDummy As LongPtr
  If LNG_K(0) = 0 Then
    LNG_K(&H0) = &HD728AE22:  LNG_K(&H1) = &H428A2F98:  LNG_K(&H2) = &H23EF65CD:  LNG_K(&H3) = &H71374491:  LNG_K(&H4) = &HEC4D3B2F:  LNG_K(&H5) = &HB5C0FBCF:  LNG_K(&H6) = &H8189DBBC:  LNG_K(&H7) = &HE9B5DBA5
    LNG_K(&H8) = &HF348B538:  LNG_K(&H9) = &H3956C25B:  LNG_K(&HA) = &HB605D019:  LNG_K(&HB) = &H59F111F1:  LNG_K(&HC) = &HAF194F9B:  LNG_K(&HD) = &H923F82A4:  LNG_K(&HE) = &HDA6D8118:  LNG_K(&HF) = &HAB1C5ED5
    LNG_K(&H10) = &HA3030242: LNG_K(&H11) = &HD807AA98: LNG_K(&H12) = &H45706FBE: LNG_K(&H13) = &H12835B01: LNG_K(&H14) = &H4EE4B28C: LNG_K(&H15) = &H243185BE: LNG_K(&H16) = &HD5FFB4E2: LNG_K(&H17) = &H550C7DC3
    LNG_K(&H18) = &HF27B896F: LNG_K(&H19) = &H72BE5D74: LNG_K(&H1A) = &H3B1696B1: LNG_K(&H1B) = &H80DEB1FE: LNG_K(&H1C) = &H25C71235: LNG_K(&H1D) = &H9BDC06A7: LNG_K(&H1E) = &HCF692694: LNG_K(&H1F) = &HC19BF174
    LNG_K(&H20) = &H9EF14AD2: LNG_K(&H21) = &HE49B69C1: LNG_K(&H22) = &H384F25E3: LNG_K(&H23) = &HEFBE4786: LNG_K(&H24) = &H8B8CD5B5: LNG_K(&H25) = &HFC19DC6:  LNG_K(&H26) = &H77AC9C65: LNG_K(&H27) = &H240CA1CC
    LNG_K(&H28) = &H592B0275: LNG_K(&H29) = &H2DE92C6F: LNG_K(&H2A) = &H6EA6E483: LNG_K(&H2B) = &H4A7484AA: LNG_K(&H2C) = &HBD41FBD4: LNG_K(&H2D) = &H5CB0A9DC: LNG_K(&H2E) = &H831153B5: LNG_K(&H2F) = &H76F988DA
    LNG_K(&H30) = &HEE66DFAB: LNG_K(&H31) = &H983E5152: LNG_K(&H32) = &H2DB43210: LNG_K(&H33) = &HA831C66D: LNG_K(&H34) = &H98FB213F: LNG_K(&H35) = &HB00327C8: LNG_K(&H36) = &HBEEF0EE4: LNG_K(&H37) = &HBF597FC7
    LNG_K(&H38) = &H3DA88FC2: LNG_K(&H39) = &HC6E00BF3: LNG_K(&H3A) = &H930AA725: LNG_K(&H3B) = &HD5A79147: LNG_K(&H3C) = &HE003826F: LNG_K(&H3D) = &H6CA6351:  LNG_K(&H3E) = &HA0E6E70:  LNG_K(&H3F) = &H14292967
    LNG_K(&H40) = &H46D22FFC: LNG_K(&H41) = &H27B70A85: LNG_K(&H42) = &H5C26C926: LNG_K(&H43) = &H2E1B2138: LNG_K(&H44) = &H5AC42AED: LNG_K(&H45) = &H4D2C6DFC: LNG_K(&H46) = &H9D95B3DF: LNG_K(&H47) = &H53380D13
    LNG_K(&H48) = &H8BAF63DE: LNG_K(&H49) = &H650A7354: LNG_K(&H4A) = &H3C77B2A8: LNG_K(&H4B) = &H766A0ABB: LNG_K(&H4C) = &H47EDAEE6: LNG_K(&H4D) = &H81C2C92E: LNG_K(&H4E) = &H1482353B: LNG_K(&H4F) = &H92722C85
    LNG_K(&H50) = &H4CF10364: LNG_K(&H51) = &HA2BFE8A1: LNG_K(&H52) = &HBC423001: LNG_K(&H53) = &HA81A664B: LNG_K(&H54) = &HD0F89791: LNG_K(&H55) = &HC24B8B70: LNG_K(&H56) = &H654BE30:  LNG_K(&H57) = &HC76C51A3
    LNG_K(&H58) = &HD6EF5218: LNG_K(&H59) = &HD192E819: LNG_K(&H5A) = &H5565A910: LNG_K(&H5B) = &HD6990624: LNG_K(&H5C) = &H5771202A: LNG_K(&H5D) = &HF40E3585: LNG_K(&H5E) = &H32BBD1B8: LNG_K(&H5F) = &H106AA070
    LNG_K(&H60) = &HB8D2D0C8: LNG_K(&H61) = &H19A4C116: LNG_K(&H62) = &H5141AB53: LNG_K(&H63) = &H1E376C08: LNG_K(&H64) = &HDF8EEB99: LNG_K(&H65) = &H2748774C: LNG_K(&H66) = &HE19B48A8: LNG_K(&H67) = &H34B0BCB5
    LNG_K(&H68) = &HC5C95A63: LNG_K(&H69) = &H391C0CB3: LNG_K(&H6A) = &HE3418ACB: LNG_K(&H6B) = &H4ED8AA4A: LNG_K(&H6C) = &H7763E373: LNG_K(&H6D) = &H5B9CCA4F: LNG_K(&H6E) = &HD6B2B8A3: LNG_K(&H6F) = &H682E6FF3
    LNG_K(&H70) = &H5DEFB2FC: LNG_K(&H71) = &H748F82EE: LNG_K(&H72) = &H43172F60: LNG_K(&H73) = &H78A5636F: LNG_K(&H74) = &HA1F0AB72: LNG_K(&H75) = &H84C87814: LNG_K(&H76) = &H1A6439EC: LNG_K(&H77) = &H8CC70208
    LNG_K(&H78) = &H23631E28: LNG_K(&H79) = &H90BEFFFA: LNG_K(&H7A) = &HDE82BDE9: LNG_K(&H7B) = &HA4506CEB: LNG_K(&H7C) = &HB2C67915: LNG_K(&H7D) = &HBEF9A3F7: LNG_K(&H7E) = &HE372532B: LNG_K(&H7F) = &HC67178F2
    LNG_K(&H80) = &HEA26619C: LNG_K(&H81) = &HCA273ECE: LNG_K(&H82) = &H21C0C207: LNG_K(&H83) = &HD186B8C7: LNG_K(&H84) = &HCDE0EB1E: LNG_K(&H85) = &HEADA7DD6: LNG_K(&H86) = &HEE6ED178: LNG_K(&H87) = &HF57D4F7F
    LNG_K(&H88) = &H72176FBA: LNG_K(&H89) = &H6F067AA:  LNG_K(&H8A) = &HA2C898A6: LNG_K(&H8B) = &HA637DC5:  LNG_K(&H8C) = &HBEF90DAE: LNG_K(&H8D) = &H113F9804: LNG_K(&H8E) = &H131C471B: LNG_K(&H8F) = &H1B710B35
    LNG_K(&H90) = &H23047D84: LNG_K(&H91) = &H28DB77F5: LNG_K(&H92) = &H40C72493: LNG_K(&H93) = &H32CAAB7B: LNG_K(&H94) = &H15C9BEBC: LNG_K(&H95) = &H3C9EBE0A: LNG_K(&H96) = &H9C100D4C: LNG_K(&H97) = &H431D67C4
    LNG_K(&H98) = &HCB3E42B6: LNG_K(&H99) = &H4CC5D4BE: LNG_K(&H9A) = &HFC657E2A: LNG_K(&H9B) = &H597F299C: LNG_K(&H9C) = &H3AD6FAEC: LNG_K(&H9D) = &H5FCB6FAB: LNG_K(&H9E) = &H4A475817: LNG_K(&H9F) = &H6C44198C
  End If
  With uCtx
    .State.Item(&H0) = &HF3BCC908: .State.Item(&H1) = &H6A09E667: .State.Item(&H2) = &H84CAA73B: .State.Item(&H3) = &HBB67AE85
    .State.Item(&H4) = &HFE94F82B: .State.Item(&H5) = &H3C6EF372: .State.Item(&H6) = &H5F1D36F1: .State.Item(&H7) = &HA54FF53A
    .State.Item(&H8) = &HADE682D1: .State.Item(&H9) = &H510E527F: .State.Item(&HA) = &H2B3E6C1F: .State.Item(&HB) = &H9B05688C
    .State.Item(&HC) = &HFB41BD6B: .State.Item(&HD) = &H1F83D9AB: .State.Item(&HE) = &H137E2179: .State.Item(&HF) = &H5BE0CD19
    .NPartial = 0
    .NInput = 0
    .BitSize = 512
  End With
  With uCtx.ArrayBytes
    .cDims = 1
    .fFeatures = 1
    .cbElements = 1
    .cLocks = 1
    .pvData = VarPtr(uCtx.Block.Item(0))
    .cElements = LNG_BLOCKSZ \ .cbElements
  End With
  Call CopyMemory(ByVal ArrPtr(uCtx.bytes), VarPtr(uCtx.ArrayBytes), LenB(pDummy))
End Sub

Private Sub CryptoSha512Update(uCtx As CryptoSha512Context, baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
Dim lAL             As Long
Dim lAH             As Long
Dim lBL             As Long
Dim lBH             As Long
Dim lCL             As Long
Dim lCh             As Long
Dim lDL             As Long
Dim lDH             As Long
Dim lEL             As Long
Dim lEH             As Long
Dim lFL             As Long
Dim lFH             As Long
Dim lGL             As Long
Dim lGH             As Long
Dim lHL             As Long
Dim lHH             As Long
Dim lIdx            As Long
Dim lJdx            As Long
  With uCtx
    If Size < 0 Then Size = UBound(baInput) + 1 - Pos
    .NInput = .NInput + Size
    If .NPartial > 0 And Size > 0 Then
      lIdx = LNG_BLOCKSZ - .NPartial
      If lIdx > Size Then lIdx = Size
      Call CopyMemory(.bytes(.NPartial), baInput(Pos), lIdx)
      .NPartial = .NPartial + lIdx
      Pos = Pos + lIdx
      Size = Size - lIdx
    End If
    Do While Size > 0 Or .NPartial = LNG_BLOCKSZ
      If .NPartial <> 0 Then
        .NPartial = 0
      ElseIf Size >= LNG_BLOCKSZ Then
        Call CopyMemory(.bytes(0), baInput(Pos), LNG_BLOCKSZ)
        Pos = Pos + LNG_BLOCKSZ
        Size = Size - LNG_BLOCKSZ
      Else
        Call CopyMemory(.bytes(0), baInput(Pos), Size)
        .NPartial = Size
        Exit Do
      End If
      For lIdx = 0 To UBound(.Block.Item) Step 2
        lAL = BSwap32(.Block.Item(lIdx))
        .Block.Item(lIdx) = BSwap32(.Block.Item(lIdx + 1))
        .Block.Item(lIdx + 1) = lAL
      Next
      lAL = .State.Item(0): lAH = .State.Item(1)
      lBL = .State.Item(2): lBH = .State.Item(3)
      lCL = .State.Item(4): lCh = .State.Item(5)
      lDL = .State.Item(6): lDH = .State.Item(7)
      lEL = .State.Item(8): lEH = .State.Item(9)
      lFL = .State.Item(10): lFH = .State.Item(11)
      lGL = .State.Item(12): lGH = .State.Item(13)
      lHL = .State.Item(14): lHH = .State.Item(15)
      lIdx = 0
      Do While lIdx < 2 * LNG_ROUNDS
        lJdx = 0
        Do While lJdx < LNG_BLOCKSZ \ 4
          pvRound lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, .Block, lJdx + 0, lIdx
          pvRound lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, .Block, lJdx + 2, lIdx
          pvRound lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, .Block, lJdx + 4, lIdx
          pvRound lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, .Block, lJdx + 6, lIdx
          pvRound lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, lDL, lDH, .Block, lJdx + 8, lIdx
          pvRound lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, lCL, lCh, .Block, lJdx + 10, lIdx
          pvRound lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, lBL, lBH, .Block, lJdx + 12, lIdx
          pvRound lBL, lBH, lCL, lCh, lDL, lDH, lEL, lEH, lFL, lFH, lGL, lGH, lHL, lHH, lAL, lAH, .Block, lJdx + 14, lIdx
          lJdx = lJdx + 16
        Loop
        lIdx = lIdx + 32
        If lIdx >= 2 * LNG_ROUNDS Then Exit Do
        For lJdx = 0 To 30 Step 2
          pvStore .Block, lJdx
        Next
      Loop
      pvAdd64 .State.Item(0), .State.Item(1), lAL, lAH
      pvAdd64 .State.Item(2), .State.Item(3), lBL, lBH
      pvAdd64 .State.Item(4), .State.Item(5), lCL, lCh
      pvAdd64 .State.Item(6), .State.Item(7), lDL, lDH
      pvAdd64 .State.Item(8), .State.Item(9), lEL, lEH
      pvAdd64 .State.Item(10), .State.Item(11), lFL, lFH
      pvAdd64 .State.Item(12), .State.Item(13), lGL, lGH
      pvAdd64 .State.Item(14), .State.Item(15), lHL, lHH
    Loop
  End With
End Sub

Private Sub CryptoSha512Finalize(uCtx As CryptoSha512Context, baOutput() As Byte)
Static B(0 To 1)    As Long
Dim baPad()         As Byte
Dim lIdx            As Long
Dim pDummy          As LongPtr
  With uCtx
    lIdx = LNG_BLOCKSZ - .NPartial
    If lIdx < 17 Then lIdx = lIdx + LNG_BLOCKSZ
    ReDim baPad(0 To lIdx - 1) As Byte
    baPad(0) = &H80
    .NInput = .NInput / 10000@ * 8
    Call CopyMemory(B(0), .NInput, 8)
    Call CopyMemory(baPad(lIdx - 4), BSwap32(B(0)), 4)
    Call CopyMemory(baPad(lIdx - 8), BSwap32(B(1)), 4)
    CryptoSha512Update uCtx, baPad
    ReDim baOutput(0 To (.BitSize + 7) \ 8 - 1) As Byte
    .ArrayBytes.pvData = VarPtr(.State.Item(0))
    For lIdx = 0 To UBound(baOutput)
      baOutput(lIdx) = .bytes(lIdx + 7 - 2 * (lIdx And 7))
    Next
    Call CopyMemory(ByVal ArrPtr(.bytes), pDummy, LenB(pDummy))
  End With
End Sub

Private Function CryptoSha512ByteArray(baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1) As Byte()
Dim uCtx As CryptoSha512Context
  CryptoSha512Init uCtx
  CryptoSha512Update uCtx, baInput, Pos, Size
  CryptoSha512Finalize uCtx, CryptoSha512ByteArray
End Function

Private Function ToHex(baData() As Byte) As String
Dim lIdx  As Long
Dim sByte As String
  ToHex = String$(UBound(baData) * 2 + 2, 48)
  For lIdx = 0 To UBound(baData)
    sByte = LCase$(Hex$(baData(lIdx)))
    Mid$(ToHex, lIdx * 2 + 3 - Len(sByte)) = sByte
  Next
End Function

Public Function Hash(sText As String) As String
Dim uCtx As CryptoSha512Context
Dim bRet() As Byte
Dim I As Long
Dim J As Integer
Dim B As Byte
Dim bIn(LNG_BLOCKSZ - 1) As Byte
Dim Pos As Long
Dim lLen As Long
  CryptoSha512Init uCtx
  Pos = 0
  For I = 0 To Len(sText) - 1 Step LNG_BLOCKSZ
    lLen = LNG_BLOCKSZ
    If Len(sText) < I + LNG_BLOCKSZ Then
      lLen = Len(sText) - I
    End If
    For J = 0 To lLen - 1
      bIn(J) = Asc(Mid$(sText, I + J + 1, 1))
    Next J
    CryptoSha512Update uCtx, bIn, 0, lLen
  Next I
  CryptoSha512Finalize uCtx, bRet
  Hash = ToHex(bRet)
End Function
