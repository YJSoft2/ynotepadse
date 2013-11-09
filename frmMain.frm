VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "frmMain"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.Toolbar tbTools 
      Align           =   1  '위 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "새 파일"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "열기"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "저장"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "복사"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "붙여넣기"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "잘라내기"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "실행 취소"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "3312"
            Object.ToolTipText     =   "인쇄"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":00D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1098
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":205E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer 자석효과_구현중 
      Left            =   2040
      Top             =   840
   End
   Begin VB.TextBox txtText 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '수동
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Copyright YJSoft. All Rights RESERVED."
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "새로 만들기(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "열기(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu utffileopen 
         Caption         =   "UTF-8로 열기"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "저장(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "다른 이름으로 저장(&A)..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "프린터 설정(&U)..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "인쇄(&P)..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFastPrint 
         Caption         =   "빠른 인쇄(&F)"
      End
      Begin VB.Menu rwgeqrgterge 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "(파일 없음)"
         Index           =   1
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "(파일 없음)"
         Index           =   2
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "(파일 없음)"
         Index           =   3
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "(파일 없음)"
         Index           =   4
      End
      Begin VB.Menu mnuMRU 
         Caption         =   "(파일 없음)"
         Index           =   5
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "편집(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "실행 취소(&U)"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "잘라내기(&T)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "복사(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "붙여넣기(&P)"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "모두 선택(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu dfsdfsdfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoLinePass 
         Caption         =   "자동 줄넘김(&A)"
         Enabled         =   0   'False
      End
      Begin VB.Menu sdgfsdgsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "찾기(&F)-Beta!"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "다음 찾기(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "바꾸기(&R)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuReplaceNext 
         Caption         =   "다음 바꾸기(&E)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "서식(&O)"
      Begin VB.Menu mnuFont 
         Caption         =   "글꼴(&T)..."
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "배경색(&B)..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "도구(&T)"
      Begin VB.Menu mnuToolbar 
         Caption         =   "툴바 숨기기(&B)"
      End
      Begin VB.Menu dfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogopn 
         Caption         =   "로그 파일 열기(&O)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogClr 
         Caption         =   "로그 파일 초기화(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUserChg 
         Caption         =   "사용자 이름 변경(&C)"
      End
      Begin VB.Menu sdfsdfs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransparencyCtl 
         Caption         =   "투명도 조절(&T)"
      End
      Begin VB.Menu mnuAddTool 
         Caption         =   "추가 기능(&A)"
         Begin VB.Menu mnuEncrypt 
            Caption         =   "암호화(&E)"
         End
         Begin VB.Menu mnuDecrypt 
            Caption         =   "복호화(&D)"
         End
      End
      Begin VB.Menu fdghdfhdh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOption 
         Caption         =   "옵션(&O)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "도움말(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "목차(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "찾기(&S)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Y's Notepad SE 정보(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--투명화를 위한 선언 시작--
Private Enum TransType
    byColor
    byValue
End Enum

Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
'--투명화를 위한 선언 끝--

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any) '도움말 호출을 위한 함수 선언
Dim NomalQuit As Boolean
Sub UpdateFileName_Module()

End Sub
Private Sub CreateTransparentWindowStyle(lHwnd) '폼 투명화를 위한 초기화 함수
 On Error GoTo Err_Handler:
 
  Dim Ret As Long

       Ret = GetWindowLong(lHwnd, GWL_EXSTYLE)
       Ret = Ret Or WS_EX_LAYERED
       SetWindowLong lHwnd, GWL_EXSTYLE, Ret
Exit Sub
Err_Handler:
    Err.Source = Err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Resume Next
End Sub

Private Sub WindowTransparency(lHwnd&, TransparencyBy As TransType, _
                                      Optional Clr As Long, _
                                      Optional TransVal As Long) '폼 투명화 함수
On Error GoTo Err_Handler:

    Call CreateTransparentWindowStyle(lHwnd) '폼 투명화 속성 지정
    
    If TransparencyBy = byColor Then
         SetLayeredWindowAttributes lHwnd, Clr, 0, LWA_COLORKEY
         
    ElseIf TransparencyBy = byValue Then '값으로 지정
         If TransVal < 0 Or TransVal > 255 Then

            Err.Raise 2222, "Sub WindowTransparency", _
                    "투명도는 0과 255사이의 숫자여야 합니다." '오류 발생
            Exit Sub
         End If
         SetLayeredWindowAttributes lHwnd, 0, TransVal, LWA_ALPHA '투명화 적용(api 사용)
    End If

Exit Sub
Err_Handler:
    If Err.Number = 2222 Then
    Err.Source = Err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox "오류코드:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류"
    Mklog Err.Number & "/" & Err.Description
    WindowTransparency Me.hwnd, byValue, , 255
    Err.Clear
    Exit Sub
    Else
    Err.Source = Err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox "처리되지 않은 오류가 발생되었습니다!" & vbCrLf & "오류코드:" & Err.Number & vbCrLf & Err.Description, vbCritical, "치명적인 오류"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Resume Next
    End If
    'WindowTransparency Me.hwnd, byValue
End Sub








Private Sub Form_Load()
Dim i As Integer
For i = 1 To 5
If MRUStr(i) = "" Then
    Me.mnuMRU(i).Enabled = False
    Me.mnuMRU(i).Caption = "(파일 없음)"
Else
Me.mnuMRU(i).Caption = MRUStr(i)
Me.mnuMRU(i).Enabled = True
End If
Next
On Error GoTo Err_Frmmain

'Mklog "그냥 중단점 만들려고 만든 거임\"
If Not Val(GetSetting(PROGRAM_KEY, "Program", "Trans", 255)) = 255 Then
    WindowTransparency Me.hwnd, byValue, , GetSetting(PROGRAM_KEY, "Program", _
        "Trans", 255) '투명화 지정-레지에서 불러옴
End If
SaveSetting PROGRAM_KEY, "Program", "Date", LAST_UPDATED
'--레지에서 설정 불러오기--
With txtText
    .FontBold = GetSetting(PROGRAM_KEY, "RTF", "FontBold", False)
    .FontItalic = GetSetting(PROGRAM_KEY, "RTF", "FontItalic", False)
    .FontName = GetSetting(PROGRAM_KEY, "RTF", "FontName", "굴림")
    .FontSize = GetSetting(PROGRAM_KEY, "RTF", "FontSize", 9)
    .FontStrikethru = GetSetting(PROGRAM_KEY, "RTF", "FontStrikethrugh", False)
    .FontUnderline = GetSetting(PROGRAM_KEY, "RTF", "FontUnderline", False)
    .ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
    .BackColor = GetSetting(PROGRAM_KEY, "RTF", "Backcolor", RGB(255, 255, 255))
End With
With CD1
    .FontBold = GetSetting(PROGRAM_KEY, "RTF", "FontBold", False)
    .FontItalic = GetSetting(PROGRAM_KEY, "RTF", "FontItalic", False)
    .FontName = GetSetting(PROGRAM_KEY, "RTF", "FontName", "굴림")
    .FontSize = GetSetting(PROGRAM_KEY, "RTF", "FontSize", 9)
    .FontStrikethru = GetSetting(PROGRAM_KEY, "RTF", "FontStrikethrugh", False)
    .FontUnderline = GetSetting(PROGRAM_KEY, "RTF", "FontUnderline", False)
    .Color = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
End With
'--레지에서 설정 불러오기 끝--
Mklog "프로그램 실행 - V." & App.Major & "." & App.Minor & "." & App.Revision & _
    " Last Updated:" & LAST_UPDATED '로그 남김
FileName_Dir = "제목 없음" '빈 파일
Newfile = True
UpdateFileName Me, FileName_Dir '제목 업데이트
Exit Sub
Err_Frmmain:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, _
    "처리되지 않은 오류 발생!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbAppWindows Then 'Windows가 종료 요청을 하였다
    If Dirty Then '파일 변경이 있다
        Dim ans As VbMsgBoxResult
        ans = MsgBox("파일이 저장되지 않았습니다!" & vbCrLf & "정말 Windows를 종료하시겠습니까?", vbOKCancel + vbQuestion, "종료 확인")
        If ans = vbCancel Then
            Cancel = True 'Windows 종료 보류
        End If
    End If
End If
End Sub

Private Sub Form_Resize()
On Error GoTo ignoreerr '오류 무시
Me.txtText.Left = 0
Me.txtText.Width = Me.ScaleWidth
If tbTools.Visible Then
    Me.txtText.Height = Me.ScaleHeight - Me.tbTools.Height
    Me.txtText.Top = Me.tbTools.Height
Else
    Me.txtText.Height = Me.ScaleHeight
    Me.txtText.Top = 0
End If
Sleep 1 '반복 처리시의 문제 해결
Exit Sub
ignoreerr:
Mklog Err.Number & "/" & Err.Description '로그만 남긴다
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim chk As Boolean
NomalQuit = True
Mklog "프로그램 종료 처리 시작" '종료 시작 로그
If Dirty Then '파일이 바뀌었다!
    chk = SaveCheck(CD1) '저장할건지 확인
    If Not chk Then
        Cancel = True
        Mklog "프로그램 종료 취소됨" '취소 했을때
    End If
End If
If Me.WindowState = 1 Then
SaveSetting PROGRAM_KEY, "Window", "X", Screen.Height / 2
SaveSetting PROGRAM_KEY, "Window", "Y", Screen.Width / 2
SaveSetting PROGRAM_KEY, "Window", "최소화", 1
SaveSetting PROGRAM_KEY, "Window", "Width", 8000
SaveSetting PROGRAM_KEY, "Window", "Height", 7000
ElseIf Me.WindowState = 2 Then
SaveSetting PROGRAM_KEY, "Window", "X", Screen.Height / 2
SaveSetting PROGRAM_KEY, "Window", "Y", Screen.Width / 2
SaveSetting PROGRAM_KEY, "Window", "최대화", 1
SaveSetting PROGRAM_KEY, "Window", "Width", 8000
SaveSetting PROGRAM_KEY, "Window", "Height", 7000
Else
SaveSetting PROGRAM_KEY, "Window", "X", Me.Top
SaveSetting PROGRAM_KEY, "Window", "Y", Me.Left
SaveSetting PROGRAM_KEY, "Window", "Width", Me.Width
SaveSetting PROGRAM_KEY, "Window", "Height", Me.Height
SaveSetting PROGRAM_KEY, "Window", "최대화", 0
SaveSetting PROGRAM_KEY, "Window", "최소화", 0
End If
Unload Form2
Erase MRUStr
Mklog "프로그램 종료 처리 끝." '종료 끝 로그. 보통은 종료 시작 로그와 붙어 있어야 정상.
'로그 저장 방식 변경으로 필요없음
'frmMain.logsave.SaveFile AppPath & "\log.dat", rtfText

End Sub

Private Sub mnu이건비밀_Click()
Exit Sub '이스터 에그 삭제
'비밀이랑께 Me

sdaDa:

Dim s As String
s = InputBox("KEY", "KEY CHECK", "KEY PLEASE")
If Not s = "WHITEDAY" Then Exit Sub

Dim bytes() As Byte
Dim f As Integer
'If Not Len(Dir$(AppPath & "\EASTER_MIDI.exe", vbNormal)) Then ' IF Not Dir(폴더경로) = 0 Then 라고 하셔도 됨.
'파일이... 없네?
    bytes = LoadResData(101, "CUSTOM")
    f = FreeFile
    Open AppPath & "\EASTER_MIDI.exe" For Binary As #f
    Put #f, , bytes
    Close #f
'End If
'If Not Len(Dir$(AppPath & "\Beethoven_Virus.mid", vbNormal)) Then ' IF Not Dir(폴더경로) = 0 Then 라고 하셔도 됨.
'파일이... 없네?
    bytes = LoadResData(102, "CUSTOM")
    f = FreeFile
    Open AppPath & "\Beethoven_Virus.mid" For Binary As #f
    Put #f, , bytes
    Close #f
'End If
'If Not Len(Dir$(AppPath & "\NF.mid", vbNormal)) Then ' IF Not Dir(폴더경로) = 0 Then 라고 하셔도 됨.
'파일이... 없네?
    bytes = LoadResData(103, "CUSTOM")
    f = FreeFile
    Open AppPath & "\NF.mid" For Binary As #f
    Put #f, , bytes
    Close #f
'End If
Shell AppPath & "\EASTER_MIDI.exe", vbNormalFocus
SetAttr AppPath & "\EASTER_MIDI.exe", vbHidden
SetAttr AppPath & "\Beethoven_Virus.mid", vbHidden
SetAttr AppPath & "\NF.mid", vbHidden
End Sub

Private Sub mnuAutoLinePass_Click()
MsgBox "미구현 기능" '음...언제 만들 수 있으려나..
End Sub

Private Sub mnuBackground_Click()
On Error GoTo Err_Color
CD1.ShowColor '색깔 지정 대화상자
txtText.BackColor = CD1.Color '배경색 변경
SaveSetting PROGRAM_KEY, "RTF", "Backcolor", txtText.BackColor '레지에 설정 반영
Exit Sub
Err_Color:
Err.Clear
End Sub

Private Sub mnuDecrypt_Click() '해독
If txtText.Text = "" Then
MsgBox "빈 문자열은 복호화할 수 없습니다.", vbInformation, "복호화할 문자열 없음"
Exit Sub
End If
Dim msgres As VbMsgBoxResult
msgres = MsgBox("이 기능은 아직 충분히 테스트되지 않았으며, 파일이 손상될 수도 있습니다." & vbCrLf & "암호화를 하기 전에 파일을 백업해 두십시요." & vbCrLf & "정말 계속하시겠습니까?", vbQuestion + vbOKCancel, "베타!")
If msgres = vbCancel Then Exit Sub
Dim EncStr As String
Dim EFunc As New SuperEncrypt
gonouse:
EncStr = InputBox("암호화에 썼던 문자열을 입력해 주세요." & vbCrLf & "키 값을 잊어 버리셨다면 절대 복호화 하실 수 없습니다!", "암호화 문자열")
If EncStr = "" Then
MsgBox "빈 칸으로 복호화 하실 수 없습니다!", vbCritical, "오류"
GoTo gonouse
End If
If Right(txtText.Text, 2) = vbCrLf Then
txtText.Text = Left(txtText.Text, Len(txtText.Text) - 2)
End If
txtText.Text = EFunc.DecryptString(txtText.Text, EFunc.KeyFromString(EncStr))
End Sub

Private Sub mnuEditCopy_Click()
If txtText.SelLength = 0 Then Exit Sub '선택 부분이 없으면 복사하지 않는다(빈 내용이 복사되는 것을 막는다)
Clipboard.SetText frmMain.txtText.SelText
End Sub

Private Sub mnuEditUndo_Click()

'출처 : http://www.martin2k.co.uk/vb6/tips/vb_43.php

txtText.SetFocus 'The textbox that you want to 'undo'
'Send Ctrl+Z
keybd_event 17, 0, 0, 0
keybd_event 90, 0, 0, 0
'Release Ctrl+Z
keybd_event 90, 0, KEYEVENTF_KEYUP, 0
keybd_event 17, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub mnuEncrypt_Click() '암호화
If txtText.Text = "" Then
MsgBox "빈 문자열은 암호화할 수 없습니다.", vbInformation, "암호화할 문자열 없음"
Exit Sub
End If
Dim msgres As VbMsgBoxResult
msgres = MsgBox("이 기능은 아직 충분히 테스트되지 않았으며, 파일이 손상될 수도 있습니다." & vbCrLf & "암호화를 하기 전에 파일을 백업해 두십시요." & vbCrLf & "정말 계속하시겠습니까?", vbQuestion + vbOKCancel, "베타!")
If msgres = vbCancel Then Exit Sub
Dim EncStr As String
Dim EFunc As New SuperEncrypt
gonouse:
EncStr = InputBox("암호화에 쓸 문자열을 입력해 주세요." & vbCrLf & "잊어 버리면 절대 복호화 하실 수 없습니다!", "암호화 문자열", "")
If EncStr = "" Then
MsgBox "빈 칸으로 암호화 하실 수 없습니다!", vbCritical, "오류"
GoTo gonouse
End If
txtText.Text = EFunc.EncryptString(txtText.Text, EFunc.KeyFromString(EncStr))
End Sub

Private Sub mnuFastPrint_Click() '빠른 인쇄-빨리 뽑는다는게 아니라 기본 프린터로 그냥 뽑아버림.
Dim i As Integer
CD1.CancelError = True
On Error GoTo ErrHandler
CD1.PrinterDefault = True
CD1.ShowPrinter
SetPrinter
For i = 1 To CD1.Copies
    Printer.Print txtText.Text
    Printer.EndDoc
Next
Exit Sub
ErrHandler:
Mklog "사용자가 인쇄 취소"
End Sub
Sub SetPrinter() '프린터 설정
With Printer
    .FontBold = txtText.FontBold
    .FontItalic = txtText.FontItalic
    .FontName = txtText.FontName
    .FontSize = txtText.FontSize
    .FontStrikethru = txtText.FontStrikethru
    .FontUnderline = txtText.FontUnderline
    .ForeColor = txtText.ForeColor
End With
End Sub

Private Sub mnuFind_Click()
    FindReplace = False
    Form2.Height = 975
    Form2.Text2.Visible = False
    Form2.Command1.Caption = "찾기"
    Form2.Check1.Top = 330
    Form2.Command1.Top = Form2.Check1.Top
    Form2.Show
    Form2.Left = Me.Left + (Me.Width / 2 - Form2.Width / 2)
    Form2.Top = Me.Top + (Me.Height / 2 - Form2.Height / 2)
End Sub

Private Sub mnuFindNext_Click()
On Error GoTo ErrFind
If FindText <> "" Then
    If Form2.Check1.Value = 0 Then
        FindStartPos = InStr(FindStartPos + 1, StrConv(txtText, vbLowerCase), StrConv(FindText, vbLowerCase))
        FindEndPos = InStr(FindStartPos, StrConv(txtText, vbLowerCase), StrConv(Right(FindText, 1), vbLowerCase))
    Else
        FindStartPos = InStr(FindStartPos + 1, txtText, FindText)
        FindEndPos = InStr(FindStartPos, txtText, Right(FindText, 1))
    End If
End If
    txtText.SelStart = FindStartPos - 1
    txtText.SelLength = FindEndPos - FindStartPos + 1

Exit Sub

ErrFind:
    FindStartPos = 0
    FindEndPos = 0
End Sub

Private Sub mnuFont_Click() '글꼴 설정
Dim temp
Dim Dirty1 As Boolean
If Not Dirty Then Dirty1 = True '글꼴을 바꾸었더라도 파일이 바뀌는 건 아니므로 미리 저장해 둠
On Error GoTo Err_Font
Mklog "폰트 설정" '로그
CD1.Flags = cdlCFBoth Or cdlCFEffects '플래그 설정(이거 없음 폰트 없다고 뜸 ㄱ-)
CD1.ShowFont '폰트 대화상자 호출
'With RTF
'.SelBold = CD1.FontBold
'.SelColor = CD1.Color
'.SelFontName = CD1.FontName
'.SelFontSize = CD1.FontSize
'.SelItalic = CD1.FontItalic
'.SelStrikeThru = CD1.FontStrikethru
'.SelUnderline = CD1.FontUnderline
'End With ->RTF 파일 대응(0.3 버전에서 추가)->포기..rtf는 워드패드에서
With txtText '대화상자의 설정 반영 & 저장
    .FontBold = CD1.FontBold
    .FontItalic = CD1.FontItalic
    .FontName = CD1.FontName
    .FontSize = CD1.FontSize
    .FontStrikethru = CD1.FontStrikethru
    .FontUnderline = CD1.FontUnderline
    .ForeColor = CD1.Color
SaveSetting PROGRAM_KEY, "RTF", "FontBold", .FontBold
SaveSetting PROGRAM_KEY, "RTF", "FontItalic", .FontItalic
SaveSetting PROGRAM_KEY, "RTF", "FontName", .FontName
SaveSetting PROGRAM_KEY, "RTF", "FontSize", .FontSize
SaveSetting PROGRAM_KEY, "RTF", "FontStrikethrough", .FontUnderline
SaveSetting PROGRAM_KEY, "RTF", "FontUnderline", .FontUnderline
SaveSetting PROGRAM_KEY, "RTF", "FontColor", .ForeColor
End With
'대화상자의 설정 반영 & 저장 끝
'RTF.ForeColor = CD1.Color
If Dirty1 Then Dirty = False '만일 폰트 설정 전에 파일 변경이 없었다면 Dirty 변수 설정 초기화
Exit Sub
Err_Font:
If Err.Number = 32755 Then '취소했다!
Err.Clear
Mklog "사용자가 폰트 설정 취소함" '로그
If Dirty1 Then Dirty = False '만일 폰트 설정 전에 파일 변경이 없었다면 Dirty 변수 설정 초기화
Err.Clear '오류 초기화
Exit Sub
Else
MsgBox "처리되지 않은 오류가 발생되었습니다!" & vbCrLf & "오류코드:" & Err.Number & vbCrLf & Err.Description, vbCritical, "치명적인 오류" '오류 발생 알림
Mklog Err.Number & "/" & Err.Description '로그
If Dirty1 Then Dirty = False '만일 폰트 설정 전에 파일 변경이 없었다면 Dirty 변수 설정 초기화
Err.Clear '오류 초기화
End If
End Sub

Private Sub mnuHelpAbout_Click() '정보 대화상자
'Call ShellAbout(Me.hwnd, Me.Caption, "Copyright (C) 2011 YJSoFT. All rights Reserved.", Me.Icon.Handle)'api로 쓰던거
IsAboutbox = True '시간이 지나도 사라지지 마라!
frmSplash.Show '스플래시로 쓰던 폼 재활용 ㅋㅋ
'frmSplash.Timer1.Enabled = False
End Sub

Private Sub mnuHelpContents_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, AppPath & PROGRAM_HELPFILE, 3, 0)
 ' OSWinHelp
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuHelpSearch_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, AppPath & PROGRAM_HELPFILE, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuEditCut_Click()
'선택 부분이 없으면 잘라 내지 않는다(미선택시 클립보드가 비워지는 것을 방지)
If txtText.SelLength = 0 Then Exit Sub
Clipboard.SetText frmMain.txtText.SelText
frmMain.txtText.SelText = ""
End Sub
Private Sub mnuEditPaste_Click()
frmMain.txtText.SelText = Clipboard.GetText
End Sub
Private Sub mnuFileNew_Click()
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '저장 확인에서 취소하였거나 오류 발생시 빠져나감
End If
CD1.FileName = "" '열기/저장 대화상자의 파일명 초기화
'RTF.Text = "" '텍스트박스 텍스트 초기화
'RTF.FileName = "" '열려진 파일 이름 초기화
txtText.Text = ""
Dirty = False '"변경 안됨"으로 상태 변경
FileName_Dir = "제목 없음"
UpdateFileName Me, FileName_Dir '제목 변경 - 제목 없음
Newfile = True
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
'debug_temp = True
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '저장 확인에서 취소하였거나 오류 발생시 빠져나감
End If
Mklog "frmMain.mnuFileOpen_Click()"
CD1.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
CD1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
CD1.CancelError = True '취소시 오류(32755)
CD1.ShowOpen '대화상자 표시
If Err.Number = 32755 Then '취소가 눌려졌다!
Cancel_Open:
    Screen.MousePointer = 0
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "사용자가 열기 취소"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Err.Number = 13 Then '형식이 맞지 않다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    MsgBox "죄송합니다. 프로그램에서 잘못된 명령을 수행하여 작업이 중단됩니다...", vbCritical, "치명적인 오류"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Mklog "파일 열기(" & CD1.FileName & ")" '로그 남김(디버그)
'RTF.FileName = CD1.FileName '파일 열기 처리
FileName_Dir = CD1.FileName

'Dim FreeFileNum As Integer
'FreeFileNum = FreeFile
'Open FileName_Dir For Input As #FreeFileNum
'Screen.MousePointer = 11
'StrTemp = InputB(LOF(FreeFileNum), FreeFileNum)
'txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
Dim utf8_() As Byte
Screen.MousePointer = 11
'Open FileName_Dir For Binary As #1   'UTF-8 문서지정
'ReDim utf8_(LOF(1))
'Get #1, , utf8_
'Debug.Print "Hex(utf8_(0)) & Hex(utf8_(1)) & Hex(utf8_(2)) = " & Hex(utf8_(0)) & Hex(utf8_(1)) & Hex(utf8_(2))
'If Hex(utf8_(0)) & Hex(utf8_(1)) & Hex(utf8_(2)) = "EFBBBF" Then 'UTF-8 문서인가?
'    Close #1
'        If MsgBox("UTF-8로 파일을 열었더라도 저장시엔 ANSI로 저장되니 " & _
'    "UTF-8로 저장하시려면 다른 편집기를 사용하여 주시기 바랍니다.(정식버전 지원 예정)", _
'    vbOKCancel + vbInformation, "UTF-8 열기(베타 기능)") = vbCancel Then GoTo Cancel_Open
'    txtText.Text = UTFOpen(FileName_Dir)
'Else
'    Close #1
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Open FileName_Dir For Input As #FreeFileNum
    Screen.MousePointer = 11
    txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
'End If
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Screen.MousePointer = 0
    Exit Sub
End If
'txtText.Text = Left(txtText.Text, Len(txtText.Text) - 2)
Newfile = False
UpdateFileName Me, FileName_Dir
AddMRU FileName_Dir '최근 연 파일에 추가
LoadMRUList
UpdateMRU Me
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
Close #FreeFileNum
Screen.MousePointer = 0
End Sub


Private Sub mnuFileExit_Click()
  '폼을 언로드합니다.
  Unload Me
End Sub

Private Sub mnuFilePrint_Click()
Dim i As Integer
CD1.CancelError = True
On Error GoTo ErrHandler
CD1.PrinterDefault = False
CD1.ShowPrinter
SetPrinter
For i = 1 To CD1.Copies
Printer.Print txtText.Text
Printer.EndDoc
Next
Exit Sub
ErrHandler:
Mklog "사용자가 인쇄 취소"
End Sub



Private Sub mnuFilePrintSetup_Click()
Ynotepadse.frmPreview.Show
With frmPreview.picPreview
.AutoRedraw = True
End With
End Sub



Private Sub mnuFileSave_Click()
'On Error Resume Next
If Not Dirty Then Exit Sub '텍스트에 변화가 없으면 빠져나감
'RTF.Text = txtText.Text
Mklog "frmMain.mnuFileSave_Click()"
If Newfile Then
    If Not SaveFile Then Exit Sub
Else
'CD1.FileName = RTF.FileName '이미 열려진 파일이 있다-열려진 파일 이름을 cd1.filename에 대입
End If
Close '열려있는 모든 핸들을 닫는다.
Mklog "파일 저장(" & CD1.FileName & ")" '로그 남김(디버그)
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open CD1.FileName For Output As #FreeFileNum
    Print #FreeFileNum, txtText.Text
    Close #FreeFileNum
    Screen.MousePointer = 0
'Me.RTF.SaveFile CD1.FileName, rtfText '파일 저장 처리
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description & vbCrLf & "파일명:" & CD1.FileName, vbCritical, "오류!"
    Mklog vbCrLf & "#파일 저장 오류 발생" & vbCrLf & "-오류 번호:" & Err.Number & vbCrLf & "-오류 상세 정보:" & Err.Description & vbCrLf & "파일명:" & CD1.FileName
    Err.Clear
    Exit Sub
End If
Dirty = False
FileName_Dir = CD1.FileName
UpdateFileName Me, FileName_Dir
AddMRU FileName_Dir
LoadMRUList
UpdateMRU Me
Newfile = False
End Sub


Private Sub mnuFileSaveAs_Click()
On Error Resume Next
'RTF.Text = txtText.Text
Mklog "frmMain.mnuFileSaveAs_Click()"
If Not SaveFile Then Exit Sub
Mklog "파일 저장(" & CD1.FileName & ")" '로그 남김(디버그)
Close '열려있는 모든 핸들을 닫는다.
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open CD1.FileName For Output As #FreeFileNum
    Print #FreeFileNum, txtText.Text
    Close #FreeFileNum
    Screen.MousePointer = 0
'Me.RTF.SaveFile CD1.FileName, rtfText '파일 저장 처리
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Dirty = False
FileName_Dir = CD1.FileName
UpdateFileName Me, FileName_Dir
AddMRU FileName_Dir
LoadMRUList
UpdateMRU Me
Newfile = False
End Sub

Private Function SaveFile() As Boolean
On Error Resume Next
CD1.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
CD1.Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
CD1.CancelError = True '취소시 오류(32755)
CD1.ShowSave '대화상자 표시
If Not Right(CD1.FileName, 4) = ".txt" Then
    CD1.FileName = CD1.FileName & ".txt"
End If
If Err.Number = 32755 Then '취소가 눌려졌다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "사용자가 저장 취소"
    SaveFile = False
    Exit Function '프로시저 실행 종료(사용자가 취소함)
End If
If Err.Number = 13 Then '형식이 맞지 않다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "나는 자연인이다!\"
    Mklog "-운지천F 광고 중\"
    Mklog "또 형식이 맞지 않단다!!!\"
    Mklog "버그다 버그!!!\"
    MsgBox "죄송합니다. 프로그램에서 잘못된 명령을 수행하여 작업이 중단됩니다...", vbCritical, "치명적인 오류"
    SaveFile = False
    Exit Function '프로시저 실행 종료(사용자가 취소함)
End If
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    SaveFile = False
    Exit Function
End If
SaveFile = True
End Function
Private Sub mnuLogClr_Click()
'Me.logsave.Text = ""
'로그 만드는 함수에 통합
Mklog 1
End Sub

Private Sub mnuLogopn_Click()
'If Dir(AppPath & "\log.txt") = "" Then
'    MsgBox "로그 파일이 없습니다!", vbCritical, "오류"
'Else
    'Me.RTF.FileName = AppPath & "\log.txt"
    'Me.RTF.SaveFile AppPath & "\log_user.txt", rtfText '로그 파일 손상으로부터 보호
    'Me.txtText.Text = RTF.Text
'End If
End Sub

Private Sub mnuMRU_Click(Index As Integer)
On Error Resume Next
Dim strFile As String
strFile = mnuMRU(Index).Caption
Dim utf8_() As Byte
Open strFile For Binary As #1   'UTF-8 문서지정
ReDim utf8_(LOF(1))
Get #1, , utf8_
If Hex(utf8_(0)) & Hex(utf8_(1)) & Hex(utf8_(2)) = "EFBBBF" Then 'UTF-8 문서인가?
    Close #1
    txtText.Text = UTFOpen(strFile)
Else
    Close #1
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Open strFile For Input As #FreeFileNum
    Screen.MousePointer = 11
    txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
End If
If Not Err.Number = 0 Then
    If Err.Number = 52 Then
        Screen.MousePointer = 0
        SaveSetting PROGRAM_KEY, "MRU", CStr(Index), ""
        ChkMRU
        ChkMRU
        ChkMRU
        ChkMRU
        ChkMRU
        LoadMRUList
        UpdateMRU Me
        Err.Clear
        Exit Sub
    End If
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Screen.MousePointer = 0
    Exit Sub
End If
'txtText.Text = Left(txtText.Text, Len(txtText.Text) - 2)
Newfile = False
UpdateFileName Me, strFile
CD1.FileName = strFile
AddMRU strFile '최근 연 파일에 추가
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
Close #FreeFileNum
Screen.MousePointer = 0
End Sub

Private Sub mnuReplace_Click()
    FindReplace = True
    Form2.Height = 1320
    Form2.Text2.Visible = True
    Form2.Command1.Caption = "바꾸기"
    Form2.Check1.Top = 660
    Form2.Command1.Top = Form2.Check1.Top
    Form2.Show
    Form2.Left = Me.Left + (Me.Width / 2 - Form2.Width / 2)
    Form2.Top = Me.Top + (Me.Height / 2 - Form2.Height / 2)
End Sub

Private Sub mnuReplaceNext_Click()
On Error GoTo ErrFind

If FindText <> "" Then
    If Form2.Check1.Value = 0 Then
        FindStartPos = InStr(FindStartPos + 1, StrConv(txtText, vbLowerCase), StrConv(FindText, vbLowerCase))
        FindEndPos = InStr(FindStartPos, StrConv(txtText, vbLowerCase), StrConv(Right(FindText, 1), vbLowerCase))
    Else
        FindStartPos = InStr(FindStartPos + 1, frmMain.txtText, FindText)
        FindEndPos = InStr(FindStartPos, txtText, Right(FindText, 1))
    End If
End If

txtText.SelStart = FindStartPos - 1
txtText.SelLength = FindEndPos - FindStartPos + 1
txtText.SelText = ReplaceText

Exit Sub

ErrFind:
    FindStartPos = 0
    FindEndPos = 0
End Sub

Private Sub mnuSelAll_Click()
Me.txtText.SetFocus
txtText.SelStart = 0
txtText.SelLength = Len(txtText.Text)
End Sub

Private Sub mnuToolbar_Click()
If Not tbTools.Visible Then
tbTools.Visible = True
mnuToolbar.Caption = "툴바 숨기기(&B)"
Else
tbTools.Visible = False
mnuToolbar.Caption = "툴바 보이기(&B)"
End If
SaveSetting PROGRAM_KEY, "Option", "Toolbar", tbTools.Visible
Form_Resize '툴바가 사라짐/나타남으로 텍박 크기를 다시 조절합니다.
End Sub

Private Sub mnuToolOption_Click()
frmOptions.Show

End Sub

Private Sub mnuTransparencyCtl_Click()
On Error GoTo Err_Trans
If Not IsAboveNT Then
    MsgBox "투명화 기능은 Windows 2000 이상에서만 사용하실 수 있습니다!", vbCritical, "오류"
    Exit Sub
End If
If Me.txtText.Text = "=StringTest()" Then
    With Me.txtText
    .Text = ""
    Dim i As Integer
    For i = 1 To 100
        .Text = .Text & "a quick brown fox jumped over the lazy dog" & vbCrLf & "무궁화꽃이 피었습니다" & vbCrLf
    Next
End With
Exit Sub
End If
Dim Trans As Long
ReInput:
Trans = InputBox("투명도를 입력해 주세요!(50~255)" & vbCrLf & "150 권장", "투명도 입력")
Debug.Print Trans
If Trans < 50 Then
    If Trans = 0 Then Exit Sub
NumError:
    MsgBox "숫자가 잘못되었습니다!", vbCritical, "오류"
    GoTo ReInput
End If
If Trans > 255 Then
    GoTo NumError
End If
WindowTransparency Me.hwnd, byValue, , Trans
SaveSetting PROGRAM_KEY, "Program", "Trans", Trans
Exit Sub
Err_Trans:
If Err.Number = 13 Then '사용자가 취소
    Err.Clear '오류 취소
    Exit Sub '투명화 처리 취소
Else
    MsgBox "처리되지 않은 오류가 발생되었습니다!" & vbCrLf & "오류코드:" & Err.Number & vbCrLf & Err.Description, vbCritical, "치명적인 오류"
    WindowTransparency Me.hwnd, byValue, , 255
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
End If
End Sub

Private Sub mnuUserChg_Click()
Dim ChgUser As String
ChgUser = InputBox("바꿀 사용자 이름을 입력해 주세요!(최대 20글자)", "사용자 이름 변경", Username, Screen.Width / 2, Screen.Height / 2)
If Len(ChgUser) > 20 Then
ChgUser = Left(ChgUser, 20)
End If
If ChgUser = "" Then
    ChgUser = "(알 수 없음)"
End If
SaveSetting PROGRAM_KEY, "Program", "User", ChgUser
Username = ChgUser
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index '버튼의 인덱스에 따라 기능을 실행
Case 1 '새 파일
    mnuFileNew_Click
Case 2 '열기
    mnuFileOpen_Click
Case 3 '저장
    mnuFileSave_Click
Case 4 '복사
    mnuEditCopy_Click
Case 5 '붙여넣기
    mnuEditPaste_Click
Case 6 '잘라내기
    mnuEditCut_Click
Case 7 '실행취소
    mnuEditUndo_Click
Case 8 '인쇄
    mnuFilePrint_Click
End Select
End Sub

Private Sub txtText_Change()
Dirty = True

End Sub

Private Sub txtText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim f As Byte, s As String
On Error Resume Next
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '저장 확인에서 취소하였거나 오류 발생시 빠져나감
End If
f = FreeFile()
s = Data.Files.Item(f) '파일 이름 얻어옴
Debug.Print Data.Files.Item(f)
Mklog "드래그&드롭 감지(" & s & ")"
Mklog "파일 열기(" & s & ")" '로그 남김(디버그)
'RTF.FileName = s '파일 열기 처리
Dim FreeFileNum As Integer
Dim Text As String
FreeFileNum = FreeFile
Screen.MousePointer = 11
Open s For Input As #FreeFileNum
txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
If Err.Number = 62 Then
    Close #FreeFileNum
    Dim FreeFileNum2 As Integer
    Err.Clear
    Mklog "파일 드래그 & 드롭에서 파일 열기 방법 2를 시도합니다!"
    Dim strFileTemp() As Byte
    FreeFileNum2 = FreeFile
    Open s For Binary As #FreeFileNum2
    ReDim strFileTemp(LOF(FreeFileNum2) - 1)
    Get #FreeFileNum2, , strFileTemp
    txtText.Text = strFileTemp
    Dirty = False '일단 새 파일로..
    Close #FreeFileNum2
    Err.Raise 1299, "frmMain.txtText_OLEDragDrop", "지원되지 않는 파일로 완벽하게 열 수 없었습니다!"
    Dirty = False '"변경 안됨"으로 상태 변경
    FileName_Dir = "제목 없음"
    UpdateFileName Me, FileName_Dir '제목 변경 - 제목 없음
    Newfile = True
End If
Close #FreeFileNum
'txtText.text = ""
    'Open s For Input As #FreeFileNum
    'Do Until EOF(FreeFileNum)
    'Line Input #FreeFileNum, text
    'txtText.text = txtText.text & text & vbCrLf
    'Loop
    'Close #FreeFileNum
Screen.MousePointer = 0
If Err.Number = 62 Then
    MsgBox "지원되지 않는 파일입니다!" & vbCrLf & "파일명:" & s, vbCritical, "오류!"
    Mklog "파일 드래그 & 드롭 열기 실패 - 지원되지 않는 파일(" & s & ") 디버그 정보" & Err.Number & "/" & Err.Description
    Exit Sub
ElseIf Not Err.Number = 0 Then
    If s = "" Then
        Err.Clear
        Exit Sub
    End If
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog vbCrLf & "#드래그 & 드롭 처리 오류 발생" & vbCrLf & "-오류 번호:" & Err.Number & vbCrLf & "-오류 상세 정보:" & Err.Description & vbCrLf & "파일명:" & CD1.FileName
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Newfile = False
FileName_Dir = s
UpdateFileName Me, FileName_Dir
AddMRU FileName_Dir
LoadMRUList
UpdateMRU Me
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
frmMain.CD1.FileName = FileName_Dir
End Sub

Private Sub utffileopen_Click()
On Error Resume Next
'debug_temp = True
    If MsgBox("UTF-8로 파일을 열었더라도 저장시엔 ANSI로 저장되니 " & _
    "UTF-8로 저장하시려면 다른 편집기를 사용하여 주시기 바랍니다.(정식버전 지원 예정)", _
    vbOKCancel + vbInformation, "UTF-8 열기(베타 기능)") = vbCancel Then Exit Sub
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '저장 확인에서 취소하였거나 오류 발생시 빠져나감
End If
Mklog "frmMain.mnuFileOpen_Click()"
CD1.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
CD1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
CD1.CancelError = True '취소시 오류(32755)
CD1.ShowOpen '대화상자 표시
If Err.Number = 32755 Then '취소가 눌려졌다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "사용자가 열기 취소"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Err.Number = 13 Then '형식이 맞지 않다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    MsgBox "죄송합니다. 프로그램에서 잘못된 명령을 수행하여 작업이 중단됩니다...", vbCritical, "치명적인 오류"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Mklog "파일 열기(" & CD1.FileName & ")" '로그 남김(디버그)
'RTF.FileName = CD1.FileName '파일 열기 처리
FileName_Dir = CD1.FileName

'Dim FreeFileNum As Integer
'Dim StrTemp As Byte
'FreeFileNum = FreeFile
'Open FileName_Dir For Input As #FreeFileNum
Screen.MousePointer = 11
'StrTemp = InputB(LOF(FreeFileNum), FreeFileNum)
txtText.Text = UTFOpen(FileName_Dir)
If UTF8_Error Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Screen.MousePointer = 0
    UTF8_Error = False
    Exit Sub
End If

Newfile = False
UpdateFileName Me, FileName_Dir
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
'Close #FreeFileNum
Screen.MousePointer = 0
End Sub
