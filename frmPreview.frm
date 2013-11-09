VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "인쇄 미리 보기"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox picPreview 
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "종료"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrt 
      Caption         =   "인쇄"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Me.cmdPrt.Top = 0
Me.cmdPrt.Left = 0
Me.cmdPrt.Width = Me.ScaleWidth / 2
Me.CmdExit.Top = 0
Me.CmdExit.Left = Me.cmdPrt.Width
Me.CmdExit.Width = Me.ScaleWidth / 2
Me.picPreview.Top = Me.CmdExit.Height
Me.picPreview.Left = 0
Me.picPreview.Width = Me.ScaleWidth
If Me.ScaleHeight - Me.CmdExit.Height > 0 Then
Me.picPreview.Height = Me.ScaleHeight - Me.CmdExit.Height
End If
End Sub
