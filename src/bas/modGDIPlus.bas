Attribute VB_Name = "modGDIPlus"
Option Explicit
Private Type GUID
  Data1    As Long
  Data2    As Integer
  Data3    As Integer
  Data4(7) As Byte
End Type
Private Type PICTDESC
  size     As Long
  Type     As Long
  hBmp     As Long
  hPal     As Long
  Reserved As Long
End Type
Private Type GdiplusStartupInput
  GdiplusVersion           As Long
  DebugEventCallback       As Long
  SuppressBackgroundThread As Long
  SuppressExternalCodecs   As Long
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

Private Const PLANES = 14
Private Const BITSPIXEL = 12
Private Const PATCOPY = &HF00021
Private Const PICTYPE_BITMAP = 1
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

Public Function InitGDIPlus() As Long
Dim Token    As Long
Dim gdipInit As GdiplusStartupInput
  gdipInit.GdiplusVersion = 1
  GdiplusStartup Token, gdipInit, ByVal 0&
  InitGDIPlus = Token
End Function

Public Sub FreeGDIPlus(Token As Long)
  GdiplusShutdown Token
End Sub

Public Function LoadPictureGDIPlus(PicFile As String, Optional ByVal BackColor As Long = vbWhite) As IPicture
Dim hDC     As Long
Dim hBitmap As Long
Dim Img     As Long
Dim iRet    As Long
Dim Width   As Long
Dim Height  As Long
  iRet = GdipLoadImageFromFile(StrPtr(PicFile), Img)
  If iRet <> 0 Then
    Err.Raise 513, "GDI+ Module", "Error loading picture " & PicFile & vbNewLine & iRet & " " & Error$(iRet)
    Exit Function
  End If
  GdipGetImageWidth Img, Width
  GdipGetImageHeight Img, Height
  InitDC hDC, hBitmap, BackColor, Width, Height
  gdipDraw Img, hDC, Width, Height
  GdipDisposeImage Img
  GetBitmap hDC, hBitmap
  Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function

Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
Dim hBrush As Long
  hDC = CreateCompatibleDC(ByVal 0&)
  hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
  hBitmap = SelectObject(hDC, hBitmap)
  hBrush = CreateSolidBrush(BackColor)
  hBrush = SelectObject(hDC, hBrush)
  PatBlt hDC, 0, 0, Width, Height, PATCOPY
  DeleteObject SelectObject(hDC, hBrush)
End Sub

Private Sub gdipDraw(Img As Long, hDC As Long, Width As Long, Height As Long)
Dim Graphics   As Long
Dim DestHeight As Long
  GdipCreateFromHDC hDC, Graphics
  GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
  GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
  GdipDeleteGraphics Graphics
End Sub

Private Sub GetBitmap(hDC As Long, hBitmap As Long)
  hBitmap = SelectObject(hDC, hBitmap)
  DeleteDC hDC
End Sub

Private Function CreatePicture(hBitmap As Long) As IPicture
Dim IID_IDispatch As GUID
Dim Pic           As PICTDESC
Dim IPic          As IPicture
  IID_IDispatch.Data1 = &H20400
  IID_IDispatch.Data4(0) = &HC0
  IID_IDispatch.Data4(7) = &H46
  Pic.size = Len(Pic)
  Pic.Type = PICTYPE_BITMAP
  Pic.hBmp = hBitmap
  OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
  Set CreatePicture = IPic
End Function
