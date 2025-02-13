VERSION 5.00
Object = "{8AB4F8F4-773A-4D7E-8DE6-3E5FD03E18CF}#1.0#0"; "RRTrayIcon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loading..."
   ClientHeight    =   255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1335
   ControlBox      =   0   'False
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlDummy 
      Left            =   600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin TrayIconOCX.TrayIcon TrayIcon 
      Left            =   0
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer tmrNewBG 
      Interval        =   1000
      Left            =   600
      Top             =   240
   End
   Begin RBG.HTTP httpDummy 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Wait..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuNewBG 
         Caption         =   "&New Background"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Random BackGround
'Andrew Sachen
'RealityRipple Software
'
'Program Created        September 17, 2005
'Program Last Modified  February 13, 2025
Option Explicit
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public FileDir      As String
Public lInterval    As Long
Public bPosition    As bgPOSITION
Public bMaxScale    As bgMAXSCALE
Public BGColor      As Long
Public Multimonitor As Long
Public Subdirs      As Boolean
Private LastTick    As Long

Private Function CreateImage(ByVal FromFile As String, ByVal Width As Long, ByVal Height As Long) As IPictureDisp
Dim bGD As Boolean
Dim p   As StdPicture
Dim mDC As Long
Dim srcW As Long
Dim srcH As Long
  On Error GoTo Erred
  pctImage.Cls
  If FromFile = "" Then
    Set CreateImage = Nothing
    Exit Function
  End If
  If LCase$(Right$(FromFile, 4)) = ".png" Then
    bGD = True
  Else
    On Error Resume Next
    Set p = LoadPicture(FromFile)
    On Error GoTo Erred
    If p Is Nothing Then bGD = True
  End If
  If bGD Then
    Dim gdiToken As Long
    gdiToken = InitGDIPlus
    Set p = LoadPictureGDIPlus(FromFile, BGColor)
    FreeGDIPlus gdiToken
  End If
  srcW = ScaleX(p.Width, vbHimetric, vbPixels)
  srcH = ScaleY(p.Height, vbHimetric, vbPixels)
  pctImage.Width = Width * Screen.TwipsPerPixelX
  pctImage.Height = Height * Screen.TwipsPerPixelY
  pctImage.BackColor = BGColor
  mDC = CreateCompatibleDC(pctImage.hDC)
  DeleteObject SelectObject(mDC, p.Handle)
  SetStretchBltMode pctImage.hDC, 4
  Dim toX As Integer, toY As Integer, toW As Integer, toH As Integer, ratio As Double
  Select Case bPosition
    Case bgPOSITION.Fill
      If (Width / srcW) < (Height / srcH) Then
        'use H Ratio, use W offset
        ratio = Height / srcH
        toW = ratio * srcW
        toH = ratio * srcH
        toX = (Width - toW) / 2
        toY = 0
      ElseIf (Width / srcW) > (Height / srcH) Then
        'use W ratio, use H offset
        ratio = Width / srcW
        toW = srcW * ratio
        toH = srcH * ratio
        toX = 0
        toY = (Height - toH) / 2
      Else
        toW = Width
        toH = Height
        toX = 0
        toY = 0
      End If

      If bMaxScale <> Unlimited Then
        Select Case bMaxScale
          Case bgMAXSCALE.X1:   ratio = 1
          Case bgMAXSCALE.X1_5: ratio = 1.5
          Case bgMAXSCALE.X2:   ratio = 2
          Case bgMAXSCALE.x3:   ratio = 3
          Case bgMAXSCALE.x5:   ratio = 5
          Case bgMAXSCALE.x10:  ratio = 10
          Case bgMAXSCALE.x25:  ratio = 25
          Case bgMAXSCALE.x50:  ratio = 50
        End Select
        If toW > srcW * ratio Then
          'force Max:1 centered
          toW = srcW * ratio
          toH = srcH * ratio
          toX = (Width / 2) - (toW / 2)
          toY = (Height / 2) - (toH / 2)
        End If
      End If

      StretchBlt pctImage.hDC, toX, toY, toW, toH, mDC, 0, 0, srcW, srcH, vbSrcCopy
    Case bgPOSITION.Fit
      If (srcW / Width) > (srcH / Height) Then
        'use W Ratio, use H offset
        ratio = srcW / Width
        toW = srcW / ratio
        toH = srcH / ratio
        toX = 0
        toY = (Height - toH) / 2
      ElseIf (srcW / Width) < (srcH / Height) Then
        'use H ratio, use W offset
        ratio = srcH / Height
        toW = srcW / ratio
        toH = srcH / ratio
        toX = (Width - toW) / 2
        toY = 0
      Else
        toW = Width
        toH = Height
        toX = 0
        toY = 0
      End If

      If bMaxScale <> Unlimited Then
        Select Case bMaxScale
          Case bgMAXSCALE.X1:   ratio = 1
          Case bgMAXSCALE.X1_5: ratio = 1.5
          Case bgMAXSCALE.X2:   ratio = 2
          Case bgMAXSCALE.x3:   ratio = 3
          Case bgMAXSCALE.x5:   ratio = 5
          Case bgMAXSCALE.x10:  ratio = 10
          Case bgMAXSCALE.x25:  ratio = 25
          Case bgMAXSCALE.x50:  ratio = 50
        End Select
        If toW > srcW * ratio Then
          'force Max:1 centered
          toW = srcW * ratio
          toH = srcH * ratio
          toX = (Width / 2) - (toW / 2)
          toY = (Height / 2) - (toH / 2)
        End If
      End If

      StretchBlt pctImage.hDC, toX, toY, toW, toH, mDC, 0, 0, srcW, srcH, vbSrcCopy
    Case bgPOSITION.Stretch
      toX = 0
      toY = 0
      toW = Width
      toH = Height
      StretchBlt pctImage.hDC, toX, toY, toW, toH, mDC, 0, 0, srcW, srcH, vbSrcCopy
    Case bgPOSITION.Tile
      For toX = 0 To Width Step srcW
        For toY = 0 To Height Step srcH
          StretchBlt pctImage.hDC, toX, toY, srcW, srcH, mDC, 0, 0, srcW, srcH, vbSrcCopy
        Next toY
      Next toX
    Case bgPOSITION.Center
      toX = (Width / 2) - (srcW / 2)
      toY = (Height / 2) - (srcH / 2)
      toW = srcW
      toH = srcH
      StretchBlt pctImage.hDC, toX, toY, toW, toH, mDC, 0, 0, srcW, srcH, vbSrcCopy
  End Select
  DeleteDC mDC
  Set p = Nothing
  Set CreateImage = pctImage.Image
Exit Function
Erred:
  Set CreateImage = Nothing
End Function

Private Function CompoundImages(ByRef Images() As IPictureDisp, ByRef Mons() As Monitor) As IPictureDisp
Dim mDC As Long, srcW As Long, srcH As Long, iTop As Long, iLeft As Long, iWidth As Long, iHeight As Long, I As Integer
  On Error GoTo Erred
  pctImage.Cls
  For I = 0 To UBound(Mons)
    If iTop > Mons(I).Top Then iTop = Mons(I).Top
    If iLeft > Mons(I).Left Then iLeft = Mons(I).Left
    If iHeight < Mons(I).Top + Mons(I).Height Then iHeight = Mons(I).Top + Mons(I).Height
    If iWidth < Mons(I).Left + Mons(I).Width Then iWidth = Mons(I).Left + Mons(I).Width
  Next I
  iWidth = (-1 * iLeft) + iWidth
  iHeight = (-1 * iTop) + iHeight
  pctImage.Width = iWidth * Screen.TwipsPerPixelX
  pctImage.Height = iHeight * Screen.TwipsPerPixelY
  pctImage.BackColor = BGColor
  mDC = CreateCompatibleDC(pctImage.hDC)
  SetStretchBltMode pctImage.hDC, 4
  For I = 0 To UBound(Mons)
    DeleteObject SelectObject(mDC, Images(I).Handle)
    srcW = ScaleX(Images(I).Width, vbHimetric, vbPixels)
    srcH = ScaleY(Images(I).Height, vbHimetric, vbPixels)
    If Mons(I).Top < 0 Or Mons(I).Left < 0 Then
      If Mons(I).Top < 0 And Mons(I).Left < 0 Then
        StretchBlt pctImage.hDC, iWidth + Mons(I).Left, iHeight + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        If Mons(I).Left + Mons(I).Width <= 0 Then
          StretchBlt pctImage.hDC, iWidth + Mons(I).Left, 0 + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        ElseIf Mons(I).Top + Mons(I).Height <= 0 Then
          StretchBlt pctImage.hDC, 0 + Mons(I).Left, iHeight + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        Else
          StretchBlt pctImage.hDC, iWidth + Mons(I).Left, 0 + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        End If
      ElseIf Mons(I).Top < 0 Then
        StretchBlt pctImage.hDC, Mons(I).Left, iHeight + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        StretchBlt pctImage.hDC, Mons(I).Left, 0 + Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
      ElseIf Mons(I).Left < 0 Then
        StretchBlt pctImage.hDC, iWidth + Mons(I).Left, Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
        StretchBlt pctImage.hDC, 0 + Mons(I).Left, Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
      End If
    Else
      StretchBlt pctImage.hDC, Mons(I).Left, Mons(I).Top, Mons(I).Width, Mons(I).Height, mDC, 0, 0, srcW, srcH, vbSrcCopy
    End If
  Next I
  DeleteDC mDC
  Set CompoundImages = pctImage.Image
  Exit Function
Erred:
  Set CompoundImages = Nothing
End Function

Private Function SplitImage(ByRef Image As IPictureDisp, ByRef Mons() As Monitor) As IPictureDisp()
Dim I     As Integer
Dim mDC   As Long
Dim lLeft As Long, lTop  As Long
Dim trueX As Long, trueY As Long
Dim ret() As IPictureDisp
  ReDim ret(UBound(Mons))
  For I = 0 To UBound(Mons)
    If lLeft > Mons(I).Left Then lLeft = Mons(I).Left
    If lTop > Mons(I).Top Then lTop = Mons(I).Top
  Next I
  For I = 0 To UBound(Mons)
    pctImage.Cls
    trueX = Abs(lLeft) + Mons(I).Left
    trueY = Abs(lTop) + Mons(I).Top
    pctImage.Width = Mons(I).Width * Screen.TwipsPerPixelX
    pctImage.Height = Mons(I).Height * Screen.TwipsPerPixelY
    pctImage.BackColor = BGColor
    mDC = CreateCompatibleDC(pctImage.hDC)
    SetStretchBltMode pctImage.hDC, 4
    DeleteObject SelectObject(mDC, Image.Handle)
    StretchBlt pctImage.hDC, 0, 0, Mons(I).Width, Mons(I).Height, mDC, trueX, trueY, Mons(I).Width, Mons(I).Height, vbSrcCopy
    DeleteDC mDC
    Set ret(I) = pctImage.Image
  Next I
  SplitImage = ret
End Function

Private Function FindFiles(ByVal Path As String, Optional ByVal Monitor As Integer = 0) As String
Dim I       As Integer
Dim BGs()   As String
Dim toShow  As Long
Static LastFile() As String
  On Error Resume Next
  BGs = Split(GetAllFiles(Path), vbNullChar)
  I = -1
  I = UBound(LastFile)
  On Error GoTo Erred
  If I = -1 Then
    ReDim Preserve LastFile(Monitor)
  Else
    If UBound(LastFile) < Monitor Then ReDim Preserve LastFile(Monitor)
  End If
  I = -1
  I = UBound(BGs)
  If I > -1 Then
    If I > 0 Then
      Do
        toShow = Int(Rnd * (I + 1))
      Loop While BGs(toShow) = LastFile(Monitor)
      If LenB(BGs(toShow)) > 0 Then FindFiles = BGs(toShow)
      LastFile(Monitor) = BGs(toShow)
    Else
      toShow = 0
      If LenB(BGs(toShow)) > 0 Then FindFiles = BGs(toShow)
      LastFile(Monitor) = BGs(toShow)
    End If
  Else
    frmSettings.Show
    'frmNotify.Notify "Random BackGround could not find any valid images in " & Path & "."
    FindFiles = vbNullString
    Erase LastFile
  End If
  Exit Function
Erred:
  frmNotify.Notify "Error in FindFiles: " & Err.Description & vbNewLine & "Path: " & Path
  FindFiles = vbNullString
End Function

Private Sub Form_Load()
Dim bShow As Boolean
  On Error GoTo Erred
  Randomize
  bShow = True
  If LenB(Command$) > 0 Then
    If Left$(Command$, 5) = "/set " Then
      bShow = False
    ElseIf Command$ = "/next" Then
      bShow = False
    End If
  End If
  If bShow Then
    Set TrayIcon.Icon = Me.Icon
    TrayIcon.ToolTipText = "Random BackGround"
    TrayIcon.ShowIcon
    If App.PrevInstance Then End
  End If
  App.TaskVisible = False
  Me.Hide
  LoadSettings
  Exit Sub
Erred:
  frmNotify.Notify "Error in Load: " & Err.Description
End Sub

Private Sub LoadProfile()
Dim sProfile  As String
Dim sColor    As String
Dim sSubdirs  As String
Dim sInterval As String
Dim lPosition As Long
Dim sPosition As String
Dim lMaxScale As Long
Dim sMaxScale As String
Dim lMultimon As Long
Dim sMultimon As String
  On Error GoTo Erred
  sProfile = GetDisplayProfile

  sColor = ReadINI(sProfile, "Color", "UNSET")
  If sColor = "UNSET" Then sColor = ReadINI("Settings", "Color", "0")
  If IsNumeric(sColor) Then
    BGColor = Val(sColor)
  Else
    BGColor = 0
  End If

  FileDir = ReadINI(sProfile, "Directory", "UNSET")
  If FileDir = "UNSET" Then FileDir = ReadINI("Settings", "Directory", "%APP%")

  sSubdirs = ReadINI(sProfile, "Subdirectories", "UNSET")
  If sSubdirs = "UNSET" Then sSubdirs = ReadINI("Settings", "Subdirectories", "N")
  Subdirs = sSubdirs = "Y"

  sInterval = ReadINI(sProfile, "Interval", "UNSET")
  If sInterval = "UNSET" Then sInterval = ReadINI("Settings", "Interval", "180")
  If IsNumeric(sInterval) Then
    lInterval = Val(sInterval)
  Else
    lInterval = 180
  End If

  sPosition = ReadINI(sProfile, "Position", "UNSET")
  If sPosition = "UNSET" Then sPosition = ReadINI("Settings", "Position", "1")
  If IsNumeric(sPosition) Then
    lPosition = Val(sPosition)
    If lPosition >= 0 And lPosition <= 5 Then
      bPosition = lPosition
    Else
      bPosition = bgPOSITION.Fit
    End If
  Else
    bPosition = bgPOSITION.Fit
  End If

  sMaxScale = ReadINI(sProfile, "MaxScale", "UNSET")
  If sMaxScale = "UNSET" Then sMaxScale = ReadINI("Settings", "MaxScale", "0")
  If IsNumeric(sMaxScale) Then
    lMaxScale = Val(sMaxScale)
    If lMaxScale >= 0 And lMaxScale <= 9 Then
      bMaxScale = lMaxScale
    Else
      bMaxScale = bgMAXSCALE.Unlimited
    End If
  Else
    bMaxScale = bgMAXSCALE.Unlimited
  End If

  sMultimon = ReadINI("Settings", "Unique", "UNSET")
  If sMultimon <> "UNSET" Then
    WriteINI "Settings", "Unique", vbNullString
    WriteINI "Settings", "Multimonitor", IIf(sMultimon = "N", "0", "1")
  End If
  sMultimon = ReadINI(sProfile, "Multimonitor", "UNSET")
  If sMultimon = "UNSET" Then sMultimon = ReadINI("Settings", "Multimonitor", "1")
  If IsNumeric(sMultimon) Then
    lMultimon = Val(sMultimon)
    If lMultimon >= 0 And lMultimon <= 2 Then
      Multimonitor = lMultimon
    Else
      Multimonitor = 1
    End If
  Else
    Multimonitor = 1
  End If

  If FileDir = "%APP%" Then FileDir = App.Path
  Do While Right$(FileDir, 1) = "\" Or Right$(FileDir, 1) = "/"
    FileDir = Left$(FileDir, Len(FileDir) - 1)
  Loop
  Exit Sub
Erred:
  frmNotify.Notify "Error in LoadProfile: " & Err.Description
  BGColor = 0
  FileDir = App.Path
  Subdirs = False
  lInterval = 180
  bPosition = bgPOSITION.Fit
  bMaxScale = bgMAXSCALE.Unlimited
  Multimonitor = 1
End Sub

Private Sub LoadSettings()
  On Error GoTo Erred
  LoadProfile
  If ReadINI("Settings", "Assoc", "N") = "Y" Then
    SetAssoc
  Else
    RemAssoc
  End If
  If ReadINI("Settings", "DesktopMenu", "N") = "Y" Then
    SetDesktopMenu
  Else
    RemDesktopMenu
  End If
  If ReadINI("Settings", "Boot", "Y") = "Y" Then
    regCreate_A_Key HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    regCreate_Value_SZ HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RBG", App.Path & "\" & App.EXEName & ".exe"
  Else
    regDelete_Sub_Key HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RBG"
  End If
  regDelete_Sub_Key HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RBG"
  If lInterval > 0 Then
    LastTick = GetTickCount - (lInterval * 1000 - 10)
  Else
    LastTick = GetTickCount - 1000
  End If
  Exit Sub
Erred:
  frmNotify.Notify "Error in LoadSettings: " & Err.Description
  LastTick = GetTickCount - (180 * 1000 - 10)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  RemDesktopMenu
  Unload frmNotify
  Unload frmSettings
  Unload frmAbout
End Sub

Private Function GetAllFiles(ByVal sPath As String) As String
Dim fName    As String
Dim lFile    As Long
Dim sFiles() As String
Dim lDir     As Long
Dim sDirs()  As String
Dim I        As Integer
  On Error GoTo Erred
  If Not Right$(sPath, 1) = "\" Then sPath = sPath & "\"
  fName = Dir$(sPath & "*.*", vbNormal Or vbDirectory)
  Do While Len(fName)
    If Not (fName = ".." Or fName = ".") Then
      If GetAttr(sPath & fName) And vbDirectory Then
        ReDim Preserve sDirs(lDir)
        sDirs(lDir) = sPath & fName
        lDir = lDir + 1
      Else
        Select Case LCase$(Mid$(fName, InStrRev(fName, ".") + 1))
          Case "bmp", "dib", "jpg", "jpeg", "jpe", "gif", "png"
            ReDim Preserve sFiles(lFile)
            sFiles(lFile) = sPath & fName
            lFile = lFile + 1
        End Select
      End If
    End If
    fName = Dir$
  Loop
  If lFile > 0 Then GetAllFiles = Join$(sFiles, vbNullChar)
  If lDir > 0 And Subdirs Then
    For I = 0 To lDir - 1
      If LenB(GetAllFiles) > 0 Then
        GetAllFiles = GetAllFiles & vbNullChar & GetAllFiles(sDirs(I))
      Else
        GetAllFiles = GetAllFiles(sDirs(I))
      End If
    Next I
  End If
  Exit Function
Erred:
  frmNotify.Notify "Error in GetAllFiles: " & Err.Description & vbNewLine & "Path: " & sPath
  GetAllFiles = vbNullString
End Function

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Public Sub mnuNewBG_Click()
  NewBackground
End Sub

Public Sub NewBackground(Optional ByVal SetBackground As String = vbNullString)
Dim Mons()   As Monitor
Dim Images() As IPictureDisp
Dim BG       As String
Dim I        As Long
Dim lL       As Long
Dim lT       As Long
Dim lR       As Long
Dim lB       As Long
Dim lWidth   As Long
Dim lHeight  As Long
Dim lMons    As Long
  On Error Resume Next
  If LastTick > GetTickCount - 2000 Then Exit Sub
  LastTick = GetTickCount
  Mons = GetMonitors
  lMons = GetMonitorCount
  On Error GoTo Erred
  If lMons = 0 Then
    lWidth = Screen.Width / Screen.TwipsPerPixelX
    lHeight = Screen.Height / Screen.TwipsPerPixelY
    If LenB(SetBackground) > 0 Then
      BG = SetBackground
    Else
      BG = FindFiles(FileDir)
    End If
    If LenB(BG) > 0 Then
      If CheckPath(SettingsFolder & "\RandomBG2.bmp") = 1 Then
        If CheckPath(SettingsFolder & "\RandomBG.bmp") = 1 Then Kill SettingsFolder & "\RandomBG.bmp"
        SavePicture CreateImage(BG, lWidth, lHeight), SettingsFolder & "\RandomBG.bmp"
        SetBG
        Kill SettingsFolder & "\RandomBG.bmp"
      ElseIf CheckPath(SettingsFolder & "\RandomBG.bmp") = 1 Then
        SavePicture CreateImage(BG, lWidth, lHeight), SettingsFolder & "\RandomBG2.bmp"
        SetBG "2"
        Kill SettingsFolder & "\RandomBG2.bmp"
      Else
        SavePicture CreateImage(BG, lWidth, lHeight), SettingsFolder & "\RandomBG.bmp"
        SetBG
      End If
      LastTick = GetTickCount
    Else
      tmrNewBG.Enabled = False
    End If
  ElseIf lMons = 1 Then
    lWidth = Mons(0).Width
    lHeight = Mons(0).Height
    If LenB(SetBackground) > 0 Then
      BG = SetBackground
    Else
      BG = FindFiles(FileDir)
    End If
    If LenB(BG) > 0 Then
      SavePicture CreateImage(BG, lWidth, lHeight), SettingsFolder & "\RandomBG.bmp"
      SetBG
      LastTick = GetTickCount
    Else
      tmrNewBG.Enabled = False
    End If
  ElseIf Multimonitor = 2 Then
    lL = 0
    lT = 0
    lR = 0
    lB = 0
    For I = 0 To lMons - 1
      If Mons(I).Left < lL Then lL = Mons(I).Left
      If Mons(I).Top < lT Then lT = Mons(I).Top
      If Mons(I).Top + Mons(I).Height > lB Then lB = Mons(I).Top + Mons(I).Height
      If Mons(I).Left + Mons(I).Width > lR Then lR = Mons(I).Left + Mons(I).Width
    Next I
    lWidth = lR - lL
    lHeight = lB - lT
    If LenB(SetBackground) > 0 Then
      BG = SetBackground
    Else
      BG = FindFiles(FileDir)
    End If
    If LenB(BG) > 0 Then
      Images = SplitImage(CreateImage(BG, lWidth, lHeight), Mons)
      SavePicture CompoundImages(Images, Mons), SettingsFolder & "\RandomBG.bmp"
      SetBG
      LastTick = GetTickCount
    Else
      tmrNewBG.Enabled = False
    End If
  Else
    ReDim Images(lMons - 1)
    If LenB(SetBackground) > 0 Then
      BG = SetBackground
    Else
      If Multimonitor = 0 Then BG = FindFiles(FileDir, I)
    End If
    For I = 0 To lMons - 1
      lWidth = Mons(I).Width
      lHeight = Mons(I).Height
      If LenB(SetBackground) = 0 And Multimonitor = 1 Then BG = FindFiles(FileDir, I)
      If LenB(BG) > 0 Then
        Set Images(I) = CreateImage(BG, lWidth, lHeight)
      Else
        tmrNewBG.Enabled = False
        Exit Sub
      End If
    Next I
    If CheckPath(SettingsFolder & "\RandomBG2.bmp") = 1 Then
      If CheckPath(SettingsFolder & "\RandomBG.bmp") = 1 Then Kill SettingsFolder & "\RandomBG.bmp"
      SavePicture CompoundImages(Images, Mons), SettingsFolder & "\RandomBG.bmp"
      SetBG
      Kill SettingsFolder & "\RandomBG2.bmp"
    ElseIf CheckPath(SettingsFolder & "\RandomBG.bmp") = 1 Then
      SavePicture CompoundImages(Images, Mons), SettingsFolder & "\RandomBG2.bmp"
      SetBG "2"
      Kill SettingsFolder & "\RandomBG.bmp"
    Else
      SavePicture CompoundImages(Images, Mons), SettingsFolder & "\RandomBG.bmp"
      SetBG
    End If
    LastTick = GetTickCount
  End If
Exit Sub
Erred:
  If LenB(BG) > 0 Then
    frmNotify.Notify "Unable to load the background " & BG & "!" & vbNewLine & "Please ensure the file is a valid BMP, DIB, JPG, GIF, or PNG image file."
    Err.Clear
  Else
    frmNotify.Notify "Error when attempting to load backgrounds." & Err.Description
    Err.Clear
  End If
End Sub

Private Sub mnuPause_Click()
  tmrNewBG.Enabled = Not tmrNewBG.Enabled
End Sub

Private Sub mnuSettings_Click()
  frmSettings.Show
End Sub

Public Sub RemAssoc()
Dim FileAss As String
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".bmp", "")
  If LenB(FileAss) > 0 Then
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".dib", "")
  If LenB(FileAss) > 0 Then
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".jpg", "")
  If LenB(FileAss) > 0 Then
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".gif", "")
  If LenB(FileAss) > 0 Then
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".png", "")
  If LenB(FileAss) > 0 Then
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regDelete_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
  End If
End Sub

Public Sub SetAssoc()
Dim FileAss As String
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".bmp", "")
  If LenB(FileAss) > 0 Then
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg", vbNullString, "Set as Background"
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /set %1"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".dib", "")
  If LenB(FileAss) > 0 Then
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg", vbNullString, "Set as Background"
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /set %1"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".jpg", "")
  If LenB(FileAss) > 0 Then
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg", vbNullString, "Set as Background"
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /set %1"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".gif", "")
  If LenB(FileAss) > 0 Then
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg", vbNullString, "Set as Background"
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /set %1"
  End If
  FileAss = regQuery_Value_SZ(HKEY_CLASSES_ROOT, ".png", "")
  If LenB(FileAss) > 0 Then
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg", vbNullString, "Set as Background"
    modReg.regCreate_A_Key HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command"
    modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, FileAss & "\shell\setbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /set %1"
  End If
End Sub

Public Sub RemDesktopMenu()
  modReg.regDelete_A_Key HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg\command"
  modReg.regDelete_A_Key HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg"
End Sub

Public Sub SetDesktopMenu()
  modReg.regCreate_A_Key HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg"
  modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg", vbNullString, "Next Random BackGround"
  modReg.regCreate_A_Key HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg\command"
  modReg.regCreate_Value_SZ HKEY_CLASSES_ROOT, "DesktopBackground\shell\nextbg\command", vbNullString, """" & App.Path & "\" & App.EXEName & ".exe"" /next"
End Sub

Private Sub tmrNewBG_Timer()
Static LastProf As String
Dim thisProf    As String
  On Error GoTo Erred
  thisProf = GetDisplayProfile
  If thisProf = LastProf Then
    If lInterval > 0 Then
        If LastTick <= GetTickCount - (lInterval * 1000) Then NewBackground
    End If
    Exit Sub
  End If
  LoadProfile
  NewBackground
  LastProf = thisProf
  Exit Sub
Erred:
  frmNotify.Notify "Error in New Background Timer: " & Err.Description
End Sub

Private Sub TrayIcon_TrayDoubleClick(Button As Integer)
  If Button = 1 Then NewBackground
End Sub

Private Sub TrayIcon_TrayMouseUp(Button As Integer)
  If Button = 2 Then
    SetForegroundWindow Me.hwnd
    mnuPause.Checked = Not tmrNewBG.Enabled
    DoEvents
    PopupMenu mnuTray, vbPopupMenuRightButton, , , mnuNewBG
  End If
End Sub
