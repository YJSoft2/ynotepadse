VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "frmOptions"
   ClientHeight    =   2790
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.OptionButton optTitle 
      Caption         =   "파일 이름(베타!)"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   5535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "스플래시 창을 비활성화 합니다."
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Value           =   1  '확인
      Width           =   3255
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "Y's Notepad SE - 경로+파일 이름"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   5535
   End
   Begin VB.CheckBox chkOnce 
      Caption         =   "스플래시 창을 하루에 한번만 봅니다."
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Value           =   1  '확인
      Width           =   5895
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "파일 이름 - Y's Notepad SE"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   5535
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "Y's Notepad SE - 파일 이름"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   5535
   End
   Begin VB.OptionButton optTitle 
      Caption         =   "경로+파일 이름 - Y's Notepad SE"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "제목 속성 변경"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5895
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "예제 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "예제 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "예제 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "적용"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkOnce_Click()
Me.cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
SaveSetting PROGRAM_KEY, "Option", "Title", TitleMode
SaveSetting PROGRAM_KEY, "Option", "Splash", chkOnce.Value
Me.cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
SaveSetting PROGRAM_KEY, "Option", "Title", TitleMode
SaveSetting PROGRAM_KEY, "Option", "Splash", chkOnce.Value
    Unload Me
End Sub



Private Sub Form_Load()
Dim a As Byte
    '폼을 가운데에 놓습니다.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Me.chkOnce.Value = GetSetting(PROGRAM_KEY, "Option", "Splash", 1)
    'a = GetSetting(PROGRAM_KEY, "Option", "Title", 1)
    TitleMode = GetSetting(PROGRAM_KEY, "Option", "Title", 1)
    Me.optTitle(TitleMode).Value = True
    Me.Caption = "옵션"
End Sub


Private Sub Form_Unload(Cancel As Integer)
UpdateFileName frmMain, FileName_Dir
End Sub

Private Sub optTitle_Click(Index As Integer)
TitleMode = Index
Me.cmdApply.Enabled = True
End Sub
