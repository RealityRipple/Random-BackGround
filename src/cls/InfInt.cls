VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub ClearSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Arr() As Any, Optional ByVal Ptr As Long)
Private Declare Sub CopySAPtr Lib "msvbvm60.dll" Alias "GetMem4" (ByRef Src() As Any, ByRef Dst() As Any)

Private mThis As NumMeta

Friend Property Get DigitCount() As Long
Dim ret As Long
  ret = mThis.DigitCount * 2
  If ret > 0 Then If (mThis.Digits(mThis.DigitCount - 1) And &HFF00) = 0 Then ret = ret - 1
  DigitCount = ret
End Property

Public Property Get IsZero() As Boolean
  IsZero = IsMetaZero(mThis)
End Property

Private Function IsMetaZero(ByRef mTest As NumMeta) As Boolean
  IsMetaZero = mTest.DigitCount = 0
End Function

Public Function ToString() As String
Const DIGITS_PER_INTEGER As Long = 4
Dim c  As Long
Dim K  As Long
Dim I  As Long
Dim sz As Long
Dim R  As String
  On Error GoTo Erred
  sz = mThis.DigitCount * DIGITS_PER_INTEGER
  R = String$(sz, 48)
  For I = 0 To mThis.DigitCount - 1
    K = mThis.Digits(I) And &HFFFF&
    c = Len(R) - I * DIGITS_PER_INTEGER
    Do While K > 0
      If K And &HF Then Mid$(R, c, 1) = Mid$(HEX_CHARS, (K And &HF) + 1, 1)
      K = K \ &H10
      c = c - 1
    Loop
  Next I
  ToString = R
  Exit Function
Erred:
  ToString = "N/A"
End Function

Friend Function TestBit(ByVal bit As Long) As Boolean
Dim index As Long
  If bit < 0 Then
    Err.Raise 5, , "Bit cannot be negative."
    Exit Function
  End If
  If mThis.DigitCount = 0 Then Exit Function
  index = (bit And &HFFFFFFF0) \ &H10
  If index < mThis.DigitCount Then
    TestBit = mThis.Digits(index) And (pow2I(bit And &HF))
  Else
    TestBit = 0
  End If
End Function

Friend Function ShiftRight(ByVal bitCount As Long) As InfInt
Dim wordCount As Long
  If (bitCount = 0) Or IsMetaZero(mThis) Then
    Set ShiftRight = Me
    Exit Function
  ElseIf bitCount < 0 Then
    Set ShiftRight = Me
    Exit Function
  End If
  wordCount = (bitCount And &HFFFFFFF0) \ &H10
  If wordCount >= mThis.DigitCount Then
    Set ShiftRight = ModPowLib.Zero
    Exit Function
  End If
  Set ShiftRight = New InfInt
  ShiftRight.InternalInit ShiftRightNumber(mThis, bitCount)
End Function

Public Sub Init(ByRef value As Variant)
  Select Case VarType(value)
    Case vbLong, vbInteger, vbByte
      InitFromNumber value
    Case vbByte Or vbArray
      InitFromArray value
    Case Else
      Err.Raise 5, , "Cannot initialize BigInteger from specified value."
  End Select
End Sub

Friend Sub InternalInit(ByRef b() As Integer)
  CopySAPtr b, mThis.Digits
  ClearSAPtr b
  Normalize
End Sub

Friend Sub AttachNumber(ByRef N As NumMeta)
  CopyMemory N, mThis, LenB(N)
End Sub

Friend Sub DetachNumber(ByRef N As NumMeta)
  ZeroMemory N, LenB(N)
End Sub

Friend Sub CopyNumber(ByRef N As NumMeta)
  N = mThis
End Sub

Friend Property Get IsOne() As Boolean
  If mThis.DigitCount <> 1 Then Exit Property
  IsOne = mThis.Digits(0) = 1
End Property

Private Sub InitFromArray(ByRef value As Variant)
Dim bytes() As Byte
Dim L       As Long
  If IsNullArray(value) Then Exit Sub
  bytes = value
  L = UBound(bytes) - LBound(bytes) + 1
  If L = 0 Then Exit Sub
  ReDim mThis.Digits(-Int(-L / 2) - 1)
  CopyMemory mThis.Digits(0), bytes(LBound(bytes)), L
  Normalize
End Sub

Private Function IsNullArray(ByRef Arr As Variant) As Boolean
Dim I As Long
  On Error GoTo errTrap
  I = LBound(Arr)
  Exit Function
errTrap:
  IsNullArray = True
End Function

Private Sub InitFromNumber(ByVal value As Long)
  ReDim mThis.Digits(1)
  CopyMemory mThis.Digits(0), value, 4
  Normalize
End Sub

Private Sub Normalize()
Dim ub As Long
Dim I  As Long
  ub = UBound(mThis.Digits)
  If ub < 0 Then Exit Sub
  Select Case mThis.Digits(ub)
    Case 0
      For I = ub - 1 To 0 Step -1
        If mThis.Digits(I) <> 0 Then
          mThis.DigitCount = I + 1
          Exit Sub
        End If
      Next I
      mThis.DigitCount = 0
    Case Else
      mThis.DigitCount = ub + 1
  End Select
End Sub

Private Function ShiftRightNumber(ByRef N As NumMeta, ByVal bitCount As Long) As Integer()
Dim wordCount     As Long
Dim remainingBits As Long
Dim newPrecision  As Long
Dim newNumber()   As Integer
Dim lowermask     As Long
Dim uppermask     As Long
Dim rshift        As Long
Dim lshift        As Long
Dim srcInt        As Integer
Dim dstInt        As Integer
Dim I             As Long
Dim J             As Long
  wordCount = (bitCount And &HFFFFFFF0) \ &H10
  remainingBits = bitCount And &HF
  newPrecision = N.DigitCount - wordCount
  ReDim newNumber(newPrecision)
  If remainingBits = 0 Then
    CopyMemory newNumber(0), N.Digits(wordCount), newPrecision * 2
    ShiftRightNumber = newNumber
    Exit Function
  End If
  lowermask = ((pow2I(remainingBits) And &HFFFF&) - 1) And &HFFFF&
  uppermask = (Not lowermask) And &HFFFF&
  rshift = pow2I(remainingBits) And &HFFFF&
  lshift = pow2I(16 - remainingBits) And &HFFFF&
  J = wordCount
  srcInt = N.Digits(J)
  Do
    dstInt = ((srcInt And &HFFFF&) And uppermask) \ rshift
    J = J + 1
    If J = N.DigitCount Then Exit Do
    srcInt = N.Digits(J)
    newNumber(I) = WrappedInt(dstInt Or (((srcInt And &HFFFF&) And lowermask) * lshift))
    I = I + 1
  Loop
  newNumber(I) = dstInt
  ShiftRightNumber = newNumber
End Function
