Attribute VB_Name = "modIActiveDesktop"
Option Explicit
Private Declare Function IIDFromString Lib "ole32" (ByVal lpszIID As Long, iid As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, riid As Any, ByVal ppv As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal dlen As Long)
Private Const CLSCTX_INPROC_SERVER  As Long = 1&
Private Const CLSID_ActiveDesktop   As String = "{75048700-EF1F-11D0-9888-006097DEACF9}"
Private Const IID_ActiveDesktop     As String = "{F490EB00-1240-11D1-9888-006097DEACF9}"
Private Type GUID
  Data1                   As Long
  Data2                   As Integer
  Data3                   As Integer
  Data4(7)                As Byte
End Type
Private Type IActiveDesktop
  QueryInterface          As Long
  AddRef                  As Long
  Release                 As Long
  ApplyChanges            As Long
  GetWallpaper            As Long
  SetWallpaper            As Long
  GetWallpaperOptions     As Long
  SetWallpaperOptions     As Long
  GetPattern              As Long
  SetPattern              As Long
  GetDesktopItemOptions   As Long
  SetDesktopItemOptions   As Long
  AddDesktopItem          As Long
  AddDesktopItemWithUI    As Long
  ModifyDesktopItem       As Long
  RemoveDesktopItem       As Long
  GetDesktopItemCount     As Long
  GetDesktopItem          As Long
  GetDesktopItemByID      As Long
  GenerateDesktopItemHtml As Long
  AddUrl                  As Long
  GetDesktopItemBySource  As Long
End Type
Private Enum AD_APPLY
  AD_APPLY_SAVE = &H1
  AD_APPLY_HTMLGEN = &H2
  AD_APPLY_REFRESH = &H4
  AD_APPLY_ALL = &H7
  AD_APPLY_FORCE = &H8
  AD_APPLY_BUFFERED_REFRESH = &H10
  AD_APPLY_DYNAMICREFRESH = &H20
End Enum

Private Sub ModifyTheme()
Dim CustomThemePath As String
  CustomThemePath = ThemesPath
  If LenB(CustomThemePath) = 0 Then Exit Sub
  If LenB(Dir$(CustomThemePath)) = 0 Then Exit Sub
  'WriteINI "Slideshow", "ImagesRootPIDL", "", CustomThemePath
  'WriteINI "Slideshow", "Interval", "86400000", CustomThemePath
  WriteINI "Slideshow", "Shuffle", "1", CustomThemePath
End Sub

Public Function ActiveDesktopSetWallpaper(ByVal strFile As String) As Boolean
Dim vtbl            As IActiveDesktop
Dim vtblptr         As Long
Dim classid         As GUID
Dim iid             As GUID
Dim obj             As Long
Dim hRes            As Long
  ModifyTheme
  hRes = IIDFromString(StrPtr(CLSID_ActiveDesktop), classid)
  If Not hRes = 0 Then Exit Function
  hRes = IIDFromString(StrPtr(IID_ActiveDesktop), iid)
  If Not hRes = 0 Then Exit Function
  hRes = CoCreateInstance(classid, 0, CLSCTX_INPROC_SERVER, iid, VarPtr(obj))
  If Not hRes = 0 Then Exit Function
  RtlMoveMemory vtblptr, ByVal obj, 4
  RtlMoveMemory vtbl, ByVal vtblptr, Len(vtbl)
  hRes = CallPointer(vtbl.SetWallpaper, obj, StrPtr(strFile), 0)
  If hRes = 0 Then
    hRes = CallPointer(vtbl.ApplyChanges, obj, AD_APPLY_ALL Or AD_APPLY_BUFFERED_REFRESH)
    If hRes = 0 Then ActiveDesktopSetWallpaper = True
    CallPointer vtbl.Release, obj
  End If
End Function

Private Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long
Dim btASM(&HEC00& - 1)  As Byte
Dim pASM                As Long
Dim I                   As Integer
  pASM = VarPtr(btASM(0))
  AddByte pASM, &H58                  ' POP EAX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H59                  ' POP ECX
  AddByte pASM, &H50                  ' PUSH EAX
  For I = UBound(params) To 0 Step -1
    AddPush pASM, CLng(params(I))     ' PUSH dword
  Next
  AddCall pASM, fnc                   ' CALL rel addr
  AddByte pASM, &HC3                  ' RET
  CallPointer = CallWindowProcA(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, lng As Long)
  AddByte pASM, &H68
  AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
  AddByte pASM, &HE8
  AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
  RtlMoveMemory ByVal pASM, lng, 4
  pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
  RtlMoveMemory ByVal pASM, bt, 1
  pASM = pASM + 1
End Sub

