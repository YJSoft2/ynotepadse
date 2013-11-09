VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblLastUpdated 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "버전"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   9
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label lblAbout1 
         BackStyle       =   0  '투명
         Caption         =   "아무 키나 눌러주세요!"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblUser 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "(알 수 없음)"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6795
      End
      Begin VB.Image imgLogo 
         Height          =   2625
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  '투명
         Caption         =   "Copyright  (C) 2011 YJSoFT.All rights Reserved."
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "버전"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   6360
         TabIndex        =   3
         Top             =   2700
         Width           =   504
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Windows 2k/XP"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         TabIndex        =   4
         Top             =   2340
         Width           =   2235
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "제품"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   32.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2160
         TabIndex        =   6
         Top             =   1140
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "이 제품은 다음 사용자에게 사용이 허가되었습니다."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "YJSoFT"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   5
         Top             =   705
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 1
End Sub

Private Sub Form_Load()
SaveSetting PROGRAM_KEY, "Program", "LastExecuteDate", Date
If Right(LAST_UPDATED, 1) = ")" Then
    Me.lblLastUpdated = "마지막 업데이트 날짜 : " & Left(LAST_UPDATED, Len(LAST_UPDATED) - 3)
Else
    Me.lblLastUpdated = "마지막 업데이트 날짜 : " & LAST_UPDATED '마지막 업데이트 날짜 표시
End If
If IsAboutbox Then
lblAbout1.Visible = True
'Timer1.Enabled = False
End If
    lblVersion.Caption = "버전 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = PROGRAM_NAME
Me.lblUser.Caption = Username
End Sub

Private Sub Frame1_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 2
End Sub

Private Sub Label1_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 3
End Sub

Private Sub imgLogo_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 4
End Sub

Private Sub lblAbout1_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 5
End Sub

Private Sub lblCompanyProduct_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 6
End Sub

Private Sub lblLicenseTo_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 7
End Sub

Private Sub lblPlatform_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 8
End Sub

Private Sub lblProductName_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 9
End Sub

Private Sub lblUser_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 10
End Sub

Private Sub lblVersion_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 11
End Sub

Private Sub lblWarning_Click()
On Error GoTo err_1
frmMain.Show
Unload Me
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 12
End Sub

Private Sub Timer1_Timer()
On Error GoTo err_1
Static i As Byte
i = i + 1
If i = 1 Then
frmMain.Show
GetUserName
SetForegroundWindow Me.hwnd
ElseIf i = 4 Then
Unload Me
End If
Exit Sub
err_1:
MsgBox Err.Number & Err.Description & 13 '아 버그 겨우 고쳤네 -.-
End Sub
