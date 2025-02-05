VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModPow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SafeArrayCreateVector Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long) As Long
Private Declare Sub SetSAPtr Lib "msvbvm60.dll" Alias "PutMem4" (ByRef Arr() As Any, ByVal Ptr As Long)

Private mZero     As InfInt
Private mOne      As InfInt

Friend Property Get Zero() As InfInt
  Set Zero = mZero
End Property

Friend Property Get One() As InfInt
  Set One = mOne
End Property

Public Function ModPow(ByVal baseValue As InfInt, ByVal exponent As InfInt, ByVal m As InfInt) As InfInt
  Set ModPow = One
  Do While Not exponent.IsZero
    If exponent.TestBit(0) Then Set ModPow = Modulus(MultiplyInfInt(ModPow, baseValue), m)
    Set exponent = exponent.ShiftRight(1)
    Set baseValue = Modulus(MultiplyInfInt(baseValue, baseValue), m)
  Loop
End Function

Private Function MultiplyInfInt(ByVal X As InfInt, ByVal Y As InfInt) As InfInt
Dim u As NumMeta
Dim v As NumMeta
Dim product() As Integer
  If X Is Nothing Then
    Set MultiplyInfInt = mZero
    Exit Function
  End If
  If X.IsZero Then
    Set MultiplyInfInt = mZero
    Exit Function
  End If
  If X.IsOne Then
    Set MultiplyInfInt = Y
    Exit Function
  End If
  If Y Is Nothing Then
    Set MultiplyInfInt = mZero
    Exit Function
  End If
  If Y.IsZero Then
    Set MultiplyInfInt = mZero
    Exit Function
  End If
  If Y.IsOne Then
    Set MultiplyInfInt = X
    Exit Function
  End If
  X.AttachNumber u
  Y.AttachNumber v
  product = MultiplyMeta(u, v)
  X.DetachNumber u
  Y.DetachNumber v
  Set MultiplyInfInt = New InfInt
  MultiplyInfInt.InternalInit product
End Function

Private Function Modulus(ByVal X As InfInt, ByVal Y As InfInt) As InfInt
Dim rmd     As InfInt
Dim xNumber As NumMeta
Dim yNumber As NumMeta
Dim iRmd()  As Integer
Dim R       As Long
  If X Is Nothing Then
    Set Modulus = mZero
    Exit Function
  End If
  If X.IsZero Then
    Set Modulus = mZero
    Exit Function
  End If
  If Y Is Nothing Then Err.Raise 11
  If Y.IsZero Then Err.Raise 11
  If Y.IsOne Then
    Set Modulus = mZero
    Exit Function
  End If
  X.CopyNumber xNumber
  Y.CopyNumber yNumber
  If yNumber.DigitCount > 1 Then
    DivideMeta xNumber, yNumber, iRmd, True
  Else
    DivideSingle xNumber.Digits, xNumber.DigitCount, yNumber.Digits(0), R
    ReDim iRmd(0)
    iRmd(0) = R
  End If
  Set rmd = New InfInt
  rmd.InternalInit iRmd
  Set Modulus = rmd
End Function

Private Function MultiplyMeta(ByRef X As NumMeta, ByRef Y As NumMeta) As Integer()
Dim ret() As Integer
Dim I As Long
Dim J As Long
Dim T As Long
  ReDim ret(X.DigitCount + Y.DigitCount)
  For I = 0 To Y.DigitCount - 1
    T = 0
    For J = 0 To X.DigitCount - 1
      T = WrappedAdd(WrappedAdd(Multiply16(Y.Digits(I), X.Digits(J)), GetLong(ret(I + J))), T)
      ret(I + J) = GetInt(T)
      T = GetLong((T And &HFFFF0000) \ &H10000)
    Next J
    ret(I + J) = GetInt(T)
  Next I
  MultiplyMeta = ret
End Function

Private Sub MultiplyAssign(ByRef X As NumMeta, ByVal Y As Long)
Dim I As Long
Dim K As Long
  For I = 0 To X.DigitCount - 1
    K = Multiply16(X.Digits(I), Y) + K
    X.Digits(I) = GetInt(K)
    K = GetLong((K And &HFFFF0000) \ &H10000)
  Next I
  If K Then
    X.Digits(X.DigitCount) = GetInt(K)
    X.DigitCount = X.DigitCount + 1
  End If
End Sub

Private Function DivideMeta(ByRef X As NumMeta, ByRef Y As NumMeta, ByRef rmd() As Integer, ByVal bRmd As Boolean) As Integer()
Dim lY     As Long
Dim lDif   As Long
Dim D      As Long
Dim ret()  As Integer
Dim dgt    As Integer
Dim uMul() As Integer
Dim I      As Long
Dim J      As Long
Dim K      As Long
Dim rHat   As Long
Dim qHat   As Long
Dim uWORD  As Long
Dim debt   As Boolean
  lY = Y.DigitCount
  lDif = X.DigitCount - lY
  If lDif < 0 Then
    If bRmd Then
      ReDim rmd(X.DigitCount)
      CopyMemory rmd(0), X.Digits(0), X.DigitCount * 2
    End If
    SetSAPtr DivideMeta, SafeArrayCreateVector(vbInteger, 0, 0)
    Exit Function
  End If
  ReDim ret(lDif + 1)
  ReDim uMul(lY)
  If (X.DigitCount - 1) = UBound(X.Digits) Then ReDim Preserve X.Digits(X.DigitCount)
  X.Digits(X.DigitCount) = 0
  X.DigitCount = X.DigitCount + 1
  D = &H10000 \ (1 + GetLong(Y.Digits(lY - 1)))
  If D > 1 Then
    MultiplyAssign X, D
    MultiplyAssign Y, D
  End If
  dgt = Y.Digits(lY - 1)
  For I = lDif To 0 Step -1
    uWORD = LeftShift16(GetLong(X.Digits(I + lY))) Or GetLong(X.Digits(I + lY - 1))
    qHat = Divide32(uWORD, dgt)
    rHat = Modulus32(uWORD, dgt)
    Do
      If qHat < &H10000 Then
        If CompareDWORD(Multiply32(qHat, Y.Digits(lY - 2)), LeftShift16(rHat) + GetLong(X.Digits(I + lY - 2))) <= 0 Then Exit Do
      End If
      qHat = qHat - 1
      rHat = rHat + GetLong(dgt)
    Loop While rHat < &H10000
    K = 0
    For J = 0 To lY - 1
      K = WrappedAdd(K, Multiply32(qHat, Y.Digits(J)))
      uMul(J) = GetInt(K)
      K = GetLong((K And &HFFFF0000) \ &H10000)
    Next J
    uMul(lY) = GetInt(K)
    debt = SubtractLoop(X.Digits, I, uMul)
    If debt Then
      qHat = qHat - 1
      AddLoop X.Digits, I, Y.Digits
    End If
    ret(I) = GetInt(qHat)
  Next I
  If bRmd Then
    If D > 1 Then
      rmd = DivideSingle(X.Digits, lY, D)
    Else
      rmd = X.Digits
    End If
  End If
  DivideMeta = ret
End Function

Private Function DivideSingle(ByRef X() As Integer, ByVal lLen As Long, ByVal Y As Long, Optional ByRef rmd As Long) As Integer()
Dim R() As Integer
Dim T   As Long
Dim I   As Long
  ReDim R(lLen)
  For I = lLen - 1 To 0 Step -1
    T = T * &H10000 + GetLong(X(I))
    R(I) = GetInt(Divide32(T, Y))
    T = GetInt(Modulus32(T, Y))
  Next I
  rmd = T
  DivideSingle = R
End Function

Private Function SubtractLoop(ByRef X() As Integer, ByVal idx As Long, ByRef Y() As Integer) As Boolean
Dim R  As Long
Dim T  As Long
Dim D  As Long
Dim I  As Long
Dim N  As Long
Dim lY As Long
  lY = UBound(Y)
  For I = idx To UBound(X)
    If N <= lY Then
      D = GetLong(Y(N))
    Else
      D = 0
    End If
    T = T + (GetLong(X(I)) - D) + R
    If T < 0 Then
      T = T + &H10000
      R = -1
    Else
      R = 0
    End If
    X(I) = GetInt(T)
    T = GetLong((T And &HFFFF0000) \ &H10000)
    N = N + 1
  Next I
  SubtractLoop = R
End Function

Private Sub AddLoop(ByRef X() As Integer, ByVal idx As Long, ByRef Y() As Integer)
Dim T  As Long
Dim D  As Long
Dim I  As Long
Dim N  As Long
Dim lY As Long
  lY = UBound(Y)
  For I = idx To UBound(X)
    If N <= lY Then
      D = GetLong(Y(N))
    Else
      D = 0
    End If
    T = T + GetLong(X(I)) + D
    X(I) = GetInt(T)
    T = GetLong((T And &HFFFF0000) \ &H10000)
    N = N + 1
  Next I
End Sub

Public Function GetInt(ByVal L As Long) As Integer
  If L And &H8000& Then GetInt = &H8000
  GetInt = GetInt Or (L And &H7FFF&)
End Function

Public Function GetLong(ByVal X As Long) As Long
  GetLong = X And &HFFFF&
End Function

Private Function CompareDWORD(ByVal X As Long, ByVal Y As Long) As Long
Dim u As Currency
Dim v As Currency
  CopyMemory u, X, 4
  CopyMemory v, Y, 4
  CompareDWORD = Sgn(u - v)
End Function

Private Function Multiply32(ByVal X As Long, ByVal Y As Integer) As Long
Dim v As Currency
Dim w As Currency
  v = GetLong(Y)
  w = (v * X) * 0.0001@
  CopyMemory Multiply32, w, 4
End Function

Private Function Multiply16(ByVal X As Long, ByVal Y As Long) As Long
Dim u As Currency
Dim v As Currency
Dim w As Currency
  u = GetLong(X)
  v = GetLong(Y)
  w = (u * v) * 0.0001@
  CopyMemory Multiply16, w, 4
End Function

Private Function Divide32(ByVal X As Long, ByVal Y As Long) As Long
Dim D As Currency
  CopyMemory D, X, 4
  D = D * 10000@
  Divide32 = Int(D / GetLong(Y))
End Function

Private Function Modulus32(ByVal X As Long, ByVal Y As Long) As Long
Dim q As Currency
Dim D As Currency
Dim v As Currency
  v = GetLong(Y)
  CopyMemory D, X, 4
  D = D * 10000@
  q = Int(D / v)
  Modulus32 = D - q * v
End Function

Private Function LeftShift16(ByVal X As Long) As Long
  If X And &H8000& Then LeftShift16 = &H80000000
  LeftShift16 = LeftShift16 Or ((X And &H7FFF) * &H10000)
End Function

Private Sub Class_Initialize()
  Set mZero = New InfInt
  mZero.Init 0
  Set mOne = New InfInt
  mOne.Init 1
End Sub
