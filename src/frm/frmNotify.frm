VERSION 5.00
Begin VB.Form frmNotify 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   690
   ClientLeft      =   1.00050e5
   ClientTop       =   1.00050e5
   ClientWidth     =   3000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Notify"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   960
   End
   Begin VB.Timer tmrTopLoc 
      Interval        =   1
      Left            =   480
      Top             =   960
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   960
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   960
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   60
      Picture         =   "frmNotify.frx":000C
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Random BackGround"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Top             =   60
      Width           =   2325
   End
   Begin VB.Label lblNotify 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Notification Message Goes Here."
      ForeColor       =   &H80000017&
      Height          =   210
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2370
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Type RECT
  Left                       As Long
  Top                        As Long
  Right                      As Long
  Bottom                     As Long
End Type
Private ScrSize              As RECT
Private iTBHgt               As Long
Private Direction            As String
Private Const HWND_TOPMOST   As Integer = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE     As Long = &H2
Private Const SWP_NOSIZE     As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  tmrShow.Enabled = False
  tmrHide.Enabled = False
  tmrStart.Enabled = False
  tmrTopLoc.Enabled = False
End Sub
Private Sub Form_DblClick()
  tmrHide.Enabled = False
  tmrHide_Timer
End Sub
Private Sub imgIcon_DblClick()
  tmrHide.Enabled = False
  tmrHide_Timer
End Sub
Private Sub lblTitle_DblClick()
  tmrHide.Enabled = False
  tmrHide_Timer
End Sub
Private Sub lblNotify_DblClick()
  tmrHide.Enabled = False
  tmrHide_Timer
End Sub
Public Sub Notify(ByVal strMessage As String)
Dim hWndDesktop As Long
Dim iWide As Integer
  hWndDesktop = GetDesktopWindow
  GetWindowRect hWndDesktop, ScrSize
  If Me.Visible Then
    Me.Left = ScrSize.Right * Screen.TwipsPerPixelX - Me.Width
    lblNotify.Caption = strMessage
    tmrHide.Enabled = False
    tmrStart.Enabled = False
    tmrShow.Enabled = False
    tmrHide.Enabled = True
    Me.Height = lblTitle.Height + lblNotify.Height + 360
    iWide = imgIcon.Width + 120 + lblTitle.Width
    If lblNotify.Width > iWide Then iWide = lblNotify.Width
    Me.Width = iWide + 240
    Exit Sub
  End If
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
  lblNotify.Caption = strMessage
  tmrHide.Enabled = False
  tmrStart.Enabled = True
  Me.Height = lblTitle.Height + lblNotify.Height + 360
  iWide = imgIcon.Width + 120 + lblTitle.Width
  If lblNotify.Width > iWide Then iWide = lblNotify.Width
  Me.Width = iWide + 240
  Me.BackColor = vbInfoBackground
  Me.ForeColor = vbInfoText
End Sub
Private Sub tmrHide_Timer()
  Direction = "Hide"
  tmrShow.Enabled = True
End Sub
Private Sub tmrShow_Timer()
Dim hWndDesktop As Long
  hWndDesktop = GetDesktopWindow
  GetWindowRect hWndDesktop, ScrSize
  If Direction = "Hide" Then
    Me.Left = Me.Left + 16 * Screen.TwipsPerPixelX
    If Me.Left >= ScrSize.Right * Screen.TwipsPerPixelX Then
      tmrShow.Enabled = False
      Unload Me
    End If
  ElseIf Direction = "Show" Then
    Me.Left = Me.Left - 16 * Screen.TwipsPerPixelX
    If Me.Left <= ScrSize.Right * Screen.TwipsPerPixelX - Me.Width Then
      Me.Left = ScrSize.Right * Screen.TwipsPerPixelX - Me.Width
      tmrShow.Enabled = False
      tmrHide.Enabled = True
    End If
  End If
End Sub
Private Sub tmrStart_Timer()
Dim TmpName     As String
Dim I           As Integer
Dim hWndDesktop As Long
  hWndDesktop = GetDesktopWindow
  GetWindowRect hWndDesktop, ScrSize
  tmrStart.Enabled = False
  Me.Left = ScrSize.Right * Screen.TwipsPerPixelX
  Direction = "Show"
  tmrShow.Enabled = True
End Sub
Private Sub tmrTopLoc_Timer()
Dim hWndTray    As Long
Dim hWndDesktop As Long
Dim Rec         As RECT
  hWndDesktop = GetDesktopWindow
  GetWindowRect hWndDesktop, ScrSize
  hWndTray = FindWindowA("Shell_TrayWnd", vbNullString)
  GetWindowRect hWndTray, Rec
  If iTBHgt <> Rec.Top * Screen.TwipsPerPixelY Then
    iTBHgt = Rec.Top * Screen.TwipsPerPixelY
  End If
  Me.Top = iTBHgt - Me.Height
  If Me.Top > iTBHgt - Me.Height Then Me.Top = iTBHgt - Me.Height
  If Me.Top < 0 Then Me.Top = iTBHgt - Me.Height
End Sub
