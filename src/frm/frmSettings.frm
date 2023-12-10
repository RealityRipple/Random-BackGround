VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random BackGround Settings"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSmooth 
      Caption         =   "Smooth &Transition"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox chkDesktopMenu 
      Caption         =   "Des&ktop Context Menu"
      Height          =   255
      Left            =   4860
      TabIndex        =   16
      ToolTipText     =   "Add a menu item in the Desktop context menu to show a new Random BackGround."
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmbMaxScale 
      Height          =   315
      ItemData        =   "frmSettings.frx":014A
      Left            =   4860
      List            =   "frmSettings.frx":0169
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Maximum ratio to upscale images to fit or fill the screen."
      Top             =   300
      Width           =   2055
   End
   Begin VB.CheckBox chkUnique 
      Caption         =   "&Unique to Screen"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      ToolTipText     =   $"frmSettings.frx":019D
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackground 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Color of empty background space. Particularly used by Fit and Center positioning."
      Top             =   900
      Width           =   555
   End
   Begin MSComDlg.CommonDialog cdlBGColor 
      Left            =   4320
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbPosition 
      Height          =   315
      ItemData        =   "frmSettings.frx":0229
      Left            =   2760
      List            =   "frmSettings.frx":023C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Background positioning style."
      Top             =   900
      Width           =   2055
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      ItemData        =   "frmSettings.frx":0262
      Left            =   2760
      List            =   "frmSettings.frx":02C8
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Duration between background changes."
      Top             =   300
      Width           =   2055
   End
   Begin VB.CheckBox chkAssoc 
      Caption         =   "Assoctaite &with Images"
      Height          =   255
      Left            =   4860
      TabIndex        =   15
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
      Left            =   4860
      TabIndex        =   17
      ToolTipText     =   "Save settings."
      Top             =   2220
      Width           =   975
   End
   Begin VB.CheckBox chkBoot 
      Caption         =   "Run on &StartUp"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "Run the program on system Startup for this user."
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   18
      ToolTipText     =   "Close."
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label lblMaxScale 
      Caption         =   "&Maximum Scale:"
      Height          =   255
      Left            =   4860
      TabIndex        =   8
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label lblBackground 
      Caption         =   "Background &Color:"
      Height          =   255
      Left            =   4860
      TabIndex        =   10
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
  frmSet.FileDir = dirBG.Path
  WriteINI "Settings", "Directory", dirBG.Path, "config.ini"
  If chkSubDir.Value = 1 Then
    WriteINI "Settings", "Subdirectories", "Y", "config.ini"
    frmSet.Subdirs = True
  Else
    WriteINI "Settings", "Subdirectories", "N", "config.ini"
    frmSet.Subdirs = False
  End If
  frmSet.tmrNewBG.Enabled = True
  If cmbTime.ListIndex >= 0 Then
    frmSet.lInterval = cmbTime.ItemData(cmbTime.ListIndex)
    WriteINI "Settings", "Interval", cmbTime.ItemData(cmbTime.ListIndex), "config.ini"
  Else
    frmSet.lInterval = 0
    WriteINI "Settings", "Interval", 0, "config.ini"
  End If
  If cmbPosition.ListIndex >= 0 Then
    frmSet.bPosition = cmbPosition.ListIndex
    WriteINI "Settings", "Position", cmbPosition.ListIndex, "config.ini"
  Else
    frmSet.bPosition = Fit
    WriteINI "Settings", "Position", 0, "config.ini"
  End If
  If cmbMaxScale.ListIndex >= 0 Then
    frmSet.bMaxScale = cmbMaxScale.ListIndex
    WriteINI "Settings", "MaxScale", cmbMaxScale.ListIndex, "config.ini"
  Else
    frmSet.bMaxScale = Unlimited
    WriteINI "Settings", "MaxScale", cmbMaxScale.ListIndex, "config.ini"
  End If
  WriteINI "Settings", "Color", cmdBackground.BackColor, "config.ini"
  frmSet.BGColor = cmdBackground.BackColor
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
  If chkUnique.Value = 1 Then
    WriteINI "Settings", "Unique", "Y", "config.ini"
    frmSet.Unique = True
  Else
    WriteINI "Settings", "Unique", "N", "config.ini"
    frmSet.Unique = False
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
  If Err.Number = 68 Then MsgBox "Drive " & UCase$(Left$(drvBG.Drive, 1)) & " is unavailable!", vbExclamation + vbOKOnly + vbSystemModal
  drvBG.Drive = Left$(Temp, 2)
  dirBG.Path = Temp
End Sub

Private Sub Form_Load()
Dim BGDir    As String
Dim DefDir   As String
Dim Position As String
Dim MaxScale As String
Dim I        As Integer
  DefDir = PicturesFolder
  BGDir = ReadINI("Settings", "Directory", "config.ini", PicturesFolder & "\")
  If LenB(Dir$(BGDir, vbDirectory)) = 0 Or LenB(BGDir) = 0 Then BGDir = PicturesFolder & "\"
  dirBG.Path = BGDir
  drvBG.Drive = Left$(BGDir, 3)
  chkSubDir.Value = IIf(ReadINI("Settings", "Subdirectories", "config.ini", "N") = "Y", 1, 0)
  For I = 0 To cmbTime.ListCount - 1
    If cmbTime.ItemData(I) = ReadINI("Settings", "Interval", "config.ini", "180") Then
      cmbTime.ListIndex = I
      Exit For
    End If
  Next I
  
  Position = ReadINI("Settings", "Position", "config.ini", "1")
  If IsNumeric(Position) Then
    Dim lPosition As Long: lPosition = Int(Position)
    If lPosition >= 0 And lPosition <= 5 Then
      cmbPosition.ListIndex = lPosition
    Else
      cmbPosition.ListIndex = bgPOSITION.Fit
    End If
  Else
    cmbPosition.ListIndex = bgPOSITION.Fit
  End If
  
  MaxScale = ReadINI("Settings", "MaxScale", "config.ini", "0")
  If IsNumeric(MaxScale) Then
    Dim lMaxScale As Long: lMaxScale = Int(MaxScale)
    If lMaxScale >= 0 And lMaxScale <= 9 Then
      cmbMaxScale.ListIndex = lMaxScale
    Else
      cmbMaxScale.ListIndex = bgMAXSCALE.Unlimited
    End If
  Else
    cmbMaxScale.ListIndex = bgMAXSCALE.Unlimited
  End If
  
  cmdBackground.BackColor = ReadINI("Settings", "Color", "config.ini", "0")
  chkBoot.Value = IIf(ReadINI("Settings", "Boot", "config.ini", "Y") = "Y", 1, 0)
  If HasSysAssoc Then
    chkAssoc.Enabled = False
    chkAssoc.Value = 0
  Else
    chkAssoc.Enabled = True
    chkAssoc.Value = IIf(ReadINI("Settings", "Assoc", "config.ini", "N") = "Y", 1, 0)
  End If
  chkDesktopMenu.Value = IIf(ReadINI("Settings", "DesktopMenu", "config.ini", "N") = "Y", 1, 0)
  If Not CanSmooth Then
    chkSmooth.Enabled = False
    chkSmooth.Value = 0
  Else
    chkSmooth.Enabled = True
    chkSmooth.Value = IIf(ReadINI("Settings", "Smooth", "config.ini", "N") = "Y", 1, 0)
  End If
  If GetMonitorCount < 2 Then
    chkUnique.Enabled = False
    chkUnique.Value = 0
  Else
    chkUnique.Enabled = True
    chkUnique.Value = IIf(ReadINI("Settings", "Unique", "config.ini", "Y") = "Y", 1, 0)
  End If
End Sub
