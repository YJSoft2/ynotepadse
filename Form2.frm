VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "찾기"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   30
      TabIndex        =   3
      Top             =   330
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "대/소문자 구분"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "바꾸기"
      Height          =   285
      Left            =   2220
      TabIndex        =   1
      Top             =   630
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrFind
If FindReplace = False Then
    If FindText <> "" Then
        If Check1.Value = 0 Then
            FindStartPos = InStr(FindStartPos + 1, StrConv(frmMain.txtText, vbLowerCase), StrConv(FindText, vbLowerCase))
            FindEndPos = InStr(FindStartPos, StrConv(frmMain.txtText, vbLowerCase), StrConv(Right(FindText, 1), vbLowerCase))
        Else
            FindStartPos = InStr(FindStartPos + 1, frmMain.txtText, FindText)
            FindEndPos = InStr(FindStartPos, frmMain.txtText, Right(FindText, 1))
        End If
    End If
        frmMain.txtText.SelStart = FindStartPos - 1
        frmMain.txtText.SelLength = FindEndPos - FindStartPos + 1
Else
        If FindText <> "" Then
        If Check1.Value = 0 Then
            FindStartPos = InStr(FindStartPos + 1, StrConv(frmMain.txtText, vbLowerCase), StrConv(FindText, vbLowerCase))
            FindEndPos = InStr(FindStartPos, StrConv(frmMain.txtText, vbLowerCase), StrConv(Right(FindText, 1), vbLowerCase))
        Else
            FindStartPos = InStr(FindStartPos + 1, frmMain.txtText, FindText)
            FindEndPos = InStr(FindStartPos, frmMain.txtText, Right(FindText, 1))
        End If
    End If
        frmMain.txtText.SelStart = FindStartPos - 1
        frmMain.txtText.SelLength = FindEndPos - FindStartPos + 1
        frmMain.txtText.SelText = Text2.Text
End If
Unload Me
'frmMain.SetFocus

Exit Sub

ErrFind:
    FindStartPos = 0
    FindEndPos = 0
    Unload Me
End Sub

Private Sub Text1_Change()
FindStartPos = 0
FindText = Text1
End Sub

Private Sub Text2_Change()
ReplaceText = Text2
End Sub
