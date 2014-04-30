VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About [Insert Name Here]"
   ClientHeight    =   1755
   ClientLeft      =   2340
   ClientTop       =   1905
   ClientWidth     =   5370
   ClipControls    =   0   'False
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "About"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "Close."
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdHire 
      Caption         =   "Make a Donation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   0
      ToolTipText     =   "Make a donation to RealityRipple Software."
      Top             =   840
      Width           =   1995
   End
   Begin RBG.HTTP httpUpdate 
      Left            =   3360
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Image imgRR 
      Height          =   480
      Left            =   4860
      Picture         =   "frmAbout.frx":18BA
      ToolTipText     =   "RealityRipple Software."
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblProduct 
      Caption         =   "[Insert Name Here]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "a RealityRipple Software product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   540
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version #.#.##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   2715
   End
   Begin VB.Label lblUpdates 
      AutoSize        =   -1  'True
      Caption         =   "Checking for Updates..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   180
      Picture         =   "frmAbout.frx":26FC
      ToolTipText     =   "[Insert Name Here]"
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdHire_Click()
  ShellExecute 0, "", "http://realityripple.com/donate.php?itm=Random+BackGround", "", "", vbNormalFocus
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  imgIcon.ToolTipText = App.Title
  lblProduct.Caption = App.Title
  lblCompany.Caption = "a " & App.CompanyName & " product"
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  httpUpdate.Tag = "VER"
  httpUpdate.OpenURL ("http://update.realityripple.com/Random_BackGround/ver.txt")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set frmAbout = Nothing
End Sub
Private Sub httpUpdate_DownloadComplete(sData As String)
  On Error GoTo Erred
  If httpUpdate.Tag = "VER" Then
    Dim sTmp() As String
    sTmp = Split(sData, "|", 2)
    Dim sRemoteVer() As String
    sRemoteVer = Split(sTmp(0), ".")
    Dim NewVer As Boolean
    NewVer = False
    If sRemoteVer(0) > App.Major Then
      NewVer = True
    ElseIf sRemoteVer(0) = App.Major Then
      If sRemoteVer(1) > App.Minor Then
        NewVer = True
      ElseIf sRemoteVer(1) = App.Minor Then
        If sRemoteVer(2) > App.Revision Then NewVer = True
      End If
    End If
    If NewVer Then
      lblUpdates.Caption = "New Update Available"
      DoEvents
      httpUpdate.Tag = "FILE"
      tmrUpdate.Enabled = True
      httpUpdate.OpenURL (sTmp(1))
    Else
      lblUpdates.Caption = "No New Updates"
    End If
  Else
    tmrUpdate.Enabled = False
    lblUpdates.Caption = "Download Complete"
    Dim nFile As Integer: nFile = FreeFile
    Open SettingsFolder & "\Setup.exe" For Binary Access Write As #nFile
    Put #nFile, , sData
    Close #nFile
    Shell SettingsFolder & "\Setup.exe /silent", vbNormalFocus
    End
  End If
Exit Sub
Erred:
  lblUpdates.Caption = "Update failure. " & Err.Description
End Sub

Private Sub httpUpdate_DownloadErrored(sReason As String)
  tmrUpdate.Enabled = False
  lblUpdates.Caption = sReason
End Sub

Private Sub tmrUpdate_Timer()
  If httpUpdate.BytesTotal > 0 Then
    lblUpdates.Caption = "Downloading Update (" & Format$(httpUpdate.BytesNow / httpUpdate.BytesTotal, "0%") & ")"
  Else
    lblUpdates.Caption = "Downloading Update (0%)"
  End If
End Sub
