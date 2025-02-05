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
Const RSA_EXP  As Long = &H10001
Const PAD_PRE  As String = "0001"
Const PAD_CT   As Long = 852
Const PAD_POST As String = "00"
Const PAD_ASN  As String = "3051300d060960864801650304020305000440"
Const PUB_KEY As String = "qVlYy2ZAVtkUJghbFCWan6Z68hicfPuZpoV874ti1d+GjgdaRqCuPijap/35RjEPcTCYt90hCsrWqdfa2bK2eQlBhnZiyqU8Ky/HTYq2msKIU9NEslBJEb6ZeszOfuU9DZy69c97ljDIUqLvwqUFxDPs8np83IF1UHBpWdIRuPbOtGAbZi5KuIk8CVhKoxNBTSX3weJORo6LXIp4J7W1WBafJHX8I5GlVqnQaCq1w0KYHHJyQ//FWrBMoPPaHZGR94bqWGMrEl4XEcT2I5QcuixnEhgL9nQp/QmgvPkI3/ehAcv5oBlCjCSwZBx9mGTcwSXUEaNWcjF+rPLMijDz/zSQ0Fpuq6Ta1XmEc7KPomf7Ly0XgIXAQBj8jNeoSwvwbETi8D2Ht6U85S8hcKudD/otlZy/3sjSQFLwLtjBMQLH83N+LYsFn34jRYFOySvL4MUBeBpf0zfODzpcxqpLRkvCrFY7Cxr3j3jvUxtH/VY0Y7pRTspvqtDNYAh6JiwsiFtG7pPlKkLj+CYmsdfDd/YnJPcTP7oNnVhIqZ2ZSe2EpwETKP6um/CDFXdPifFX+ViEsUuUM/uTeZYDAILAhnb7WSx7bWPc6qOGL8X5XsQ2Oa4QDlMhEvy3zFErBdMcwzKS7hhSXjMerq0Xn4G8iNVQJnVYuqTnPoIKEpvKeos"
Private checkHash As String

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
  If App.Revision = 0 Then
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor
  Else
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  End If
  httpUpdate.Tag = "VER"
  httpUpdate.OpenURL ("http://update.realityripple.com/Random_BackGround/update.ver?sha=512")
  Me.MousePointer = vbHourglass
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set frmAbout = Nothing
End Sub

Private Sub httpUpdate_DownloadComplete(sData As String)
  On Error GoTo Erred
  If httpUpdate.Tag = "VER" Then
    Dim sSig As String
    sSig = httpUpdate.HeaderValue("X-Update-Signature")
    lblUpdates.Caption = "Checking Signature..."
    DoEvents
    If Not VerifySignature(sData, sSig) Then
      Me.MousePointer = vbDefault
      lblUpdates.Caption = "Update failure. Information signature mismatch."
    Else
      Me.MousePointer = vbDefault
      Dim sTmp() As String
      sTmp = Split(sData, "|", 3)
      Dim sRemoteVer() As String
      sRemoteVer = Split(sTmp(0), ".")
      checkHash = sTmp(2)
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
        httpUpdate.OpenURL ("http:" & sTmp(1))
      Else
        lblUpdates.Caption = "No New Updates"
      End If
    End If
  Else
    tmrUpdate.Enabled = False
    lblUpdates.Caption = "Checking Integrity..."
    Me.MousePointer = vbHourglass
    DoEvents
    If modSHA512.Hash(sData) = checkHash Then
      Me.MousePointer = vbDefault
      lblUpdates.Caption = "Download Complete"
      DoEvents
      Dim nFile As Integer: nFile = FreeFile
      Open SettingsFolder & "\Setup.exe" For Binary Access Write As #nFile
      Put #nFile, , sData
      Close #nFile
      Dim sVer As String
      sVer = App.Major & "." & App.Minor
      If App.Revision > 0 Then sVer = sVer & "." & App.Revision
      Shell SettingsFolder & "\Setup.exe /silent /noicons /update=""" & sVer & """", vbNormalFocus
      End
    Else
      Me.MousePointer = vbDefault
      lblUpdates.Caption = "Update failure. Installer hash mismatch."
    End If
  End If
Exit Sub
Erred:
  lblUpdates.Caption = "Update failure. " & Err.Description
End Sub

Public Function VerifySignature(ByVal Message As String, ByVal Signature As String) As Boolean
Dim bKey() As Byte
Dim bSig() As Byte
Dim biKey  As InfInt
Dim biSig  As InfInt
Dim biMod  As InfInt
Dim biRet  As InfInt
Dim sHash  As String
Dim sMatch As String
  On Error GoTo Erred
  If LenB(Signature) = 0 Then
    VerifySignature = False
    Exit Function
  End If
  sHash = modSHA512.Hash(Message)
  If LenB(sHash) = 0 Then
    VerifySignature = False
    Exit Function
  End If
  bKey = Base64Decode(PUB_KEY)
  bSig = Base64Decode(Signature)
  If UBound(bKey) <> 511 Then
    VerifySignature = False
    Exit Function
  End If
  If UBound(bSig) <> 511 Then
    VerifySignature = False
    Exit Function
  End If
  sMatch = PAD_PRE & String$(PAD_CT, "f") & PAD_POST + PAD_ASN + sHash
  Set biKey = New InfInt
  Set biMod = New InfInt
  Set biSig = New InfInt
  biKey.Init FlipBytes(bKey)
  biMod.Init RSA_EXP
  biSig.Init FlipBytes(bSig)
  Set biRet = ModPowLib.ModPow(biSig, biMod, biKey)
  VerifySignature = sMatch = biRet.ToString
  Exit Function
Erred:
  VerifySignature = False
End Function

Private Function FlipBytes(bIn() As Byte) As Byte()
Dim bOut() As Byte
Dim I      As Long
Dim L      As Long
  L = UBound(bIn)
  ReDim bOut(L)
  For I = 0 To L
    bOut(L - I) = bIn(I)
  Next I
  FlipBytes = bOut
End Function

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
