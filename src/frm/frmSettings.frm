VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random BackGround Settings"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMultimonitor 
      Height          =   315
      ItemData        =   "frmSettings.frx":014A
      Left            =   2760
      List            =   "frmSettings.frx":0157
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Advanced multimonitor control."
      Top             =   2220
      Width           =   2055
   End
   Begin VB.Timer tmrProfChange 
      Interval        =   1000
      Left            =   60
      Top             =   2640
   End
   Begin VB.CheckBox chkSmooth 
      Caption         =   "Smooth &Transition"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CheckBox chkDesktopMenu 
      Caption         =   "Des&ktop Context Menu"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Add a menu item in the Desktop context menu to show a new Random BackGround."
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmbMaxScale 
      Height          =   315
      ItemData        =   "frmSettings.frx":0180
      Left            =   4920
      List            =   "frmSettings.frx":019F
      Style           =   2  'Dropdown List
      TabIndex        =   13
      ToolTipText     =   "Maximum ratio to upscale images to fit or fill the screen."
      Top             =   300
      Width           =   2055
   End
   Begin VB.CommandButton cmdBackground 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Color of empty background space. Particularly used by Fit and Center positioning."
      Top             =   900
      Width           =   555
   End
   Begin MSComDlg.CommonDialog cdlBGColor 
      Left            =   540
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbPosition 
      Height          =   315
      ItemData        =   "frmSettings.frx":01D3
      Left            =   2760
      List            =   "frmSettings.frx":01E6
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Background positioning style."
      Top             =   900
      Width           =   2055
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      ItemData        =   "frmSettings.frx":020C
      Left            =   2760
      List            =   "frmSettings.frx":0272
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Duration between background changes."
      Top             =   300
      Width           =   2055
   End
   Begin VB.CheckBox chkAssoc 
      Caption         =   "Assoctaite &with Images"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "Add a menu item to BMP, DIB, JPG, and GIF files in Windows Explorer to set the file as the current background image."
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame fraBG 
      Caption         =   "Background &Directory"
      Height          =   2550
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2600
      Begin VB.CheckBox chkSubDir 
         Caption         =   "&Include Subdirectories"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Backgrounds will be chosen from the selected directory or any subdirectory."
         Top             =   2220
         Width           =   1935
      End
      Begin VB.DirListBox dirBG 
         Height          =   1440
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Directory to use."
         Top             =   720
         Width           =   2400
      End
      Begin VB.DriveListBox drvBG 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Drive to use."
         Top             =   300
         Width           =   2400
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Save settings."
      Top             =   2220
      Width           =   975
   End
   Begin VB.CheckBox chkBoot 
      Caption         =   "Run on &StartUp"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      ToolTipText     =   "Run the program on system Startup for this user."
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      ToolTipText     =   "Close."
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label lblMultimonitor 
      Caption         =   "M&ultimonitor:"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1980
      Width           =   1995
   End
   Begin VB.Label lblMaxScale 
      Caption         =   "&Maximum Scale:"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label lblBackground 
      Caption         =   "Background &Color:"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label lblPosition 
      Caption         =   "Picture &Position:"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   660
      Width           =   2000
   End
   Begin VB.Label lblInterval 
      Caption         =   "Cha&nge picture every:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   60
      Width           =   2000
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPosition_Click()
  Select Case cmbPosition.ListIndex
    Case 0, 1
      lblMaxScale.Enabled = True
      cmbMaxScale.Enabled = True
    Case Else
      lblMaxScale.Enabled = False
      cmbMaxScale.Enabled = False
  End Select
End Sub

Private Sub cmdBackground_Click()
  cdlBGColor.Color = cmdBackground.BackColor
  cdlBGColor.Flags = 3
  cdlBGColor.CancelError = True
  On Error Resume Next
  cdlBGColor.ShowColor
  If Err.Number = 0 Then cmdBackground.BackColor = cdlBGColor.Color
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sProfile   As String
Dim sWriteTo() As String
Dim sDescrs()  As String
Dim I          As Integer
  sProfile = GetDisplayProfile
  If sProfile = "Settings" Then
    ReDim sWriteTo(1)
    sWriteTo(1) = "Settings"
  Else
    ReDim sWriteTo(2)
    sWriteTo(1) = "Settings"
    sWriteTo(2) = sProfile
    sDescrs = GetDisplayDescr
    WriteINI sProfile, "Info", sDescrs(0), "config.ini"
    If UBound(sDescrs) > 0 Then
      For I = 1 To UBound(sDescrs)
        WriteINI sProfile, "Info_" & Trim$(Str$(I)), sDescrs(I), "config.ini"
      Next I
    End If
  End If
  For I = 0 To UBound(sWriteTo)
    If I = 0 Then
      frmSet.FileDir = dirBG.Path
      frmSet.Subdirs = chkSubDir.Value = 1
      frmSet.lInterval = IIf(cmbTime.ListIndex < 0, 0, cmbTime.ItemData(cmbTime.ListIndex))
      frmSet.bPosition = IIf(cmbPosition.ListIndex < 0, bgPOSITION.Fit, cmbPosition.ListIndex)
      frmSet.bMaxScale = IIf(cmbMaxScale.ListIndex < 0, bgMAXSCALE.Unlimited, cmbMaxScale.ListIndex)
      frmSet.BGColor = cmdBackground.BackColor
      frmSet.Multimonitor = cmbMultimonitor.ItemData(cmbMultimonitor.ListIndex)
    Else
      WriteINI sWriteTo(I), "Directory", dirBG.Path, "config.ini"
      WriteINI sWriteTo(I), "Subdirectories", IIf(chkSubDir.Value = 1, "Y", "N"), "config.ini"
      WriteINI sWriteTo(I), "Interval", IIf(cmbTime.ListIndex < 0, "180", Trim$(Str$(cmbTime.ItemData(cmbTime.ListIndex)))), "config.ini"
      WriteINI sWriteTo(I), "Position", IIf(cmbPosition.ListIndex < 0, Trim$(Str$(bgPOSITION.Fit)), Trim$(Str$(cmbPosition.ListIndex))), "config.ini"
      WriteINI sWriteTo(I), "MaxScale", IIf(cmbMaxScale.ListIndex < 0, Trim$(Str$(bgMAXSCALE.Unlimited)), Trim$(Str$(cmbMaxScale.ListIndex))), "config.ini"
      WriteINI sWriteTo(I), "Color", Trim$(Str$(cmdBackground.BackColor)), "config.ini"
      WriteINI sWriteTo(I), "Multimonitor", cmbMultimonitor.ItemData(cmbMultimonitor.ListIndex), "config.ini"
    End If
  Next I
  frmSet.tmrNewBG.Enabled = Not frmSet.mnuPause.Checked

  If chkBoot.Value = 1 Then
    WriteINI "Settings", "Boot", "Y", "config.ini"
    regCreate_Value_SZ HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RBG", App.Path & "\" & App.EXEName & ".exe"
  Else
    WriteINI "Settings", "Boot", "N", "config.ini"
    regDelete_Sub_Key HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "RBG"
  End If
  If chkAssoc.Value = 1 Then
    WriteINI "Settings", "Assoc", "Y", "config.ini"
    frmSet.SetAssoc
  Else
    WriteINI "Settings", "Assoc", "N", "config.ini"
    frmSet.RemAssoc
  End If
  If chkDesktopMenu.Value = 1 Then
    WriteINI "Settings", "DesktopMenu", "Y", "config.ini"
    frmSet.SetDesktopMenu
  Else
    WriteINI "Settings", "DesktopMenu", "N", "config.ini"
    frmSet.RemDesktopMenu
  End If
  If chkSmooth.Value = 1 Then
    WriteINI "Settings", "Smooth", "Y", "config.ini"
  Else
    WriteINI "Settings", "Smooth", "N", "config.ini"
  End If
  Unload Me
  frmSet.NewBackground
End Sub

Private Sub drvBG_Change()
Dim Temp As String
  On Error GoTo Erred
  Temp = dirBG.Path
  dirBG.Path = drvBG.Drive
  Exit Sub
Erred:
  If Err.Number = 68 Then frmNotify.Notify "Drive " & UCase$(Left$(drvBG.Drive, 1)) & " is unavailable!"
  drvBG.Drive = Left$(Temp, 2)
  dirBG.Path = Temp
End Sub

Private Sub LoadProfile()
Dim sProfile  As String
Dim sBGDir    As String
Dim sDefDir   As String
Dim sSubdir   As String
Dim sTime     As String
Dim lPosition As Long
Dim sPosition As String
Dim lMaxScale As Long
Dim sMaxScale As String
Dim sBGColor  As String
Dim sMultimon As String
Dim bMultimon As Boolean
Dim I         As Integer
  sProfile = GetDisplayProfile

  sDefDir = PicturesFolder
  sBGDir = ReadINI(sProfile, "Directory", "config.ini", "UNSET")
  If sBGDir = "UNSET" Then sBGDir = ReadINI("Settings", "Directory", "config.ini", sDefDir & "\")
  If LenB(sBGDir) = 0 Then
    sBGDir = sDefDir & "\"
  ElseIf CheckPath(sBGDir) <> 2 Then
    sBGDir = sDefDir & "\"
  End If
  dirBG.Path = sBGDir
  drvBG.Drive = Left$(sBGDir, 3)

  sSubdir = ReadINI(sProfile, "Subdirectories", "config.ini", "UNSET")
  If sSubdir = "UNSET" Then sSubdir = ReadINI("Settings", "Subdirectories", "config.ini", "N")
  chkSubDir.Value = IIf(sSubdir = "Y", 1, 0)

  sTime = ReadINI(sProfile, "Interval", "config.ini", "UNSET")
  If sTime = "UNSET" Then sTime = ReadINI("Settings", "Interval", "config.ini", "180")
  For I = 0 To cmbTime.ListCount - 1
    If cmbTime.ItemData(I) = sTime Then
      cmbTime.ListIndex = I
      Exit For
    End If
  Next I

  sPosition = ReadINI(sProfile, "Position", "config.ini", "UNSET")
  If sPosition = "UNSET" Then sPosition = ReadINI("Settings", "Position", "config.ini", "1")
  If IsNumeric(sPosition) Then
    lPosition = Val(sPosition)
    If lPosition >= 0 And lPosition <= 5 Then
      cmbPosition.ListIndex = lPosition
    Else
      cmbPosition.ListIndex = bgPOSITION.Fit
    End If
  Else
    cmbPosition.ListIndex = bgPOSITION.Fit
  End If

  sMaxScale = ReadINI(sProfile, "MaxScale", "config.ini", "UNSET")
  If sMaxScale = "UNSET" Then sMaxScale = ReadINI("Settings", "MaxScale", "config.ini", "0")
  If IsNumeric(sMaxScale) Then
    lMaxScale = Val(sMaxScale)
    If lMaxScale >= 0 And lMaxScale <= 9 Then
      cmbMaxScale.ListIndex = lMaxScale
    Else
      cmbMaxScale.ListIndex = bgMAXSCALE.Unlimited
    End If
  Else
    cmbMaxScale.ListIndex = bgMAXSCALE.Unlimited
  End If

  sBGColor = ReadINI(sProfile, "Color", "config.ini", "UNSET")
  If sBGColor = "UNSET" Then sBGColor = ReadINI("Settings", "Color", "config.ini", "0")
  If IsNumeric(sBGColor) Then
    cmdBackground.BackColor = Val(sBGColor)
  Else
    cmdBackground.BackColor = 0
  End If

  If GetMonitorCount < 2 Then
    lblMultimonitor.Enabled = False
    cmbMultimonitor.Enabled = False
    cmbMultimonitor.ListIndex = 0
  Else
    lblMultimonitor.Enabled = True
    cmbMultimonitor.Enabled = True
    sMultimon = ReadINI("Settings", "Unique", "config.ini", "UNSET")
    If sMultimon <> "UNSET" Then
      WriteINI "Settings", "Unique", vbNullString, "config.ini"
      WriteINI "Settings", "Multimonitor", IIf(sMultimon = "N", "0", "1"), "config.ini"
    End If
    sMultimon = ReadINI(sProfile, "Multimonitor", "config.ini", "UNSET")
    If sMultimon = "UNSET" Then sMultimon = ReadINI("Settings", "Multimonitor", "config.ini", "1")
    bMultimon = False
    For I = 0 To cmbMultimonitor.ListCount - 1
      If cmbMultimonitor.ItemData(I) = Val(sMultimon) Then
        cmbMultimonitor.ListIndex = I
        bMultimon = True
        Exit For
      End If
    Next I
    If Not bMultimon Then cmbMultimonitor.ListIndex = 0
  End If
End Sub

Private Sub Form_Load()
  LoadProfile

  chkBoot.Value = IIf(ReadINI("Settings", "Boot", "config.ini", "Y") = "Y", 1, 0)
  If HasSysAssoc Then
    chkAssoc.Enabled = False
    chkAssoc.Value = 0
  Else
    chkAssoc.Enabled = True
    chkAssoc.Value = IIf(ReadINI("Settings", "Assoc", "config.ini", "N") = "Y", 1, 0)
  End If
  If Not CanSmooth Then
    chkSmooth.Enabled = False
    chkSmooth.Value = 0
    chkDesktopMenu.Enabled = False
    chkDesktopMenu.Value = 0
  Else
    chkSmooth.Enabled = True
    chkSmooth.Value = IIf(ReadINI("Settings", "Smooth", "config.ini", "N") = "Y", 1, 0)
    chkDesktopMenu.Enabled = True
    chkDesktopMenu.Value = IIf(ReadINI("Settings", "DesktopMenu", "config.ini", "N") = "Y", 1, 0)
  End If
End Sub

Private Sub tmrProfChange_Timer()
Static LastProf As String
Dim thisProf    As String
  On Error Resume Next
  If Not Me.Visible Then Exit Sub
  thisProf = GetDisplayProfile
  If thisProf = LastProf Then Exit Sub
  LoadProfile
  LastProf = thisProf
End Sub
