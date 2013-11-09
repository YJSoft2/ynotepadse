VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain_Eng 
   AutoRedraw      =   -1  'True
   Caption         =   "frmMain_Eng"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer 자석효과_구현중 
      Left            =   2040
      Top             =   720
   End
   Begin VB.TextBox txtText 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '수동
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Copyright YJSoft. All Rights RESERVED."
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
      Begin VB.Menu dfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoLinePass 
         Caption         =   "자동 줄넘김(&A)"
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
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "찾기(&S)..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Y's Notepad SE 정보(&A)..."
      End
   End
   Begin VB.Menu mnu이건비밀 
      Caption         =   " "
   End
End
Attribute VB_Name = "frmMain_Eng"
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

On Error GoTo Err_frmMain_Eng

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
'로그 저장 방식 Shell Echo로 바꿔서 필요없음
'Me.logsave.Text = ""
'If Dir(AppPath & "\log.dat") = "" Then '로그 파일이 있는지 확인
'    Me.logsave.SaveFile AppPath & "\log.dat", rtfText '없다면 만들어 준다
'Else
'    Me.logsave.FileName = AppPath & "\log.dat" '있다면 불러온다
'    Debug.Print AppPath
'End If
Mklog "프로그램 실행 - V." & App.Major & "." & App.Minor & "." & App.Revision & _
    " Last Updated:" & LAST_UPDATED '로그 남김
FileName_Dir = "제목 없음" '빈 파일
Newfile = True
UpdateFileName Me, FileName_Dir '제목 업데이트
Exit Sub
Err_frmMain_Eng:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, _
    "처리되지 않은 오류 발생!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbAppWindows Then 'Windows가 종료 요청을 하였다
    If Dirty Then '파일 변경이 있다
        Dim ans As VbMsgBoxResult
        ans = MsgBox("파일이 저장되지 않았습니다!" & vbCrLf & "정말 Windows를 종료하시겠습니까?", vbOKCancel, "종료 확인")
        If ans = vbCancel Then
            Cancel = True 'Windows 종료 보류
        End If
    End If
End If
End Sub

Private Sub Form_Resize()
On Error GoTo ignoreerr '오류 무시
Me.txtText.Left = 0
Me.txtText.Top = 0
Me.txtText.Width = Me.ScaleWidth
Me.txtText.Height = Me.ScaleHeight
Sleep 0
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
Mklog "프로그램 종료 처리 끝." '종료 끝 로그. 보통은 종료 시작 로그와 붙어 있어야 정상.
'로그 저장 방식 변경으로 필요없음
'frmMain_Eng.logsave.SaveFile AppPath & "\log.dat", rtfText
End Sub

Private Sub mnu이건비밀_Click()
'비밀이랑께 Me
'If Me.txtText.text = "=StringTest()" Then
    With Me.txtText
    .Text = ""
    Dim i As Integer
    For i = 1 To 100
        .Text = .Text & "abc def" & vbCrLf
    Next
End With
'End If
End Sub

Private Sub mnuAutoLinePass_Click()
MsgBox "미구현 기능" '음...언제 만들 수 있으려나..
End Sub

Private Sub mnuBackground_Click()
CD1.ShowColor '색깔 지정 대화상자
txtText.BackColor = CD1.Color '배경색 변경
SaveSetting PROGRAM_KEY, "RTF", "Backcolor", txtText.BackColor '레지에 설정 반영
End Sub

Private Sub mnuDecrypt_Click() '해독
Dim msgres As VbMsgBoxResult
msgres = MsgBox("이 기능은 아직 충분히 테스트되지 않았으며, 파일이 손상될 수도 있습니다." & vbCrLf & "암호화를 하기 전에 파일을 백업해 두십시요." & vbCrLf & "정말 계속하시겠습니까?", vbQuestion + vbOKCancel, "베타!")
If msgres = vbCancel Then Exit Sub
txtText.Text = DeCrypt(txtText.Text)
End Sub

Private Sub mnuEncrypt_Click() '암호화
Dim msgres As VbMsgBoxResult
msgres = MsgBox("이 기능은 아직 충분히 테스트되지 않았으며, 파일이 손상될 수도 있습니다." & vbCrLf & "암호화를 하기 전에 파일을 백업해 두십시요." & vbCrLf & "정말 계속하시겠습니까?", vbQuestion + vbOKCancel, "베타!")
If msgres = vbCancel Then Exit Sub
txtText.Text = EnCrypt(txtText.Text)
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

Private Sub mnuHelpContents_Click() '도움말<파일이 없어서 FAIL
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
 ' OSWinHelp
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuHelpSearch_Click() '도움말<파일이 없어서 FAIL
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuEditCut_Click()
MsgBox "미구현 기능"
End Sub
Private Sub mnuEditPaste_Click()
MsgBox "미구현 기능"
End Sub
Private Sub mnuEditUndo_Click()
MsgBox "미구현 기능"
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
Mklog "frmMain_Eng.mnuFileOpen_Click()"
CD1.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
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
    Mklog "나는 자연인이다!\"
    Mklog "-운지천F 광고 중\"
    Mklog "또 형식이 맞지 않단다!!!\"
    Mklog "버그다 버그!!!\"
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

Dim FreeFileNum As Integer
FreeFileNum = FreeFile
Open FileName_Dir For Input As #FreeFileNum
Screen.MousePointer = 11
txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If

Newfile = False
UpdateFileName Me, FileName_Dir
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
Close #FreeFileNum
Screen.MousePointer = 0
End Sub






'Private Sub mnuFileClose_Click()
'  MsgBox "닫기 코드를 작성하십시오!"
'End Sub

Private Sub mnuFileExit_Click()
  '폼을 언로드합니다.
  Unload Me
End Sub

'Private Sub mnuFileNew_Click()
'  MsgBox "새 파일 코드를 작성하십시오!"
'End Sub



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
On Error Resume Next
If Not Dirty Then Exit Sub '텍스트에 변화가 없으면 빠져나감
'RTF.Text = txtText.Text
Mklog "frmMain_Eng.mnuFileSave_Click()"
If Newfile Then
    SaveFile
Else
'CD1.FileName = RTF.FileName '이미 열려진 파일이 있다-열려진 파일 이름을 cd1.filename에 대입
End If
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
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Dirty = False
FileName_Dir = CD1.FileName
UpdateFileName Me, FileName_Dir
Newfile = False
End Sub


Private Sub mnuFileSaveAs_Click()
On Error Resume Next
'RTF.Text = txtText.Text
Mklog "frmMain_Eng.mnuFileSaveAs_Click()"
SaveFile
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
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Dirty = False
FileName_Dir = CD1.FileName
UpdateFileName Me, FileName_Dir
Newfile = False
End Sub

Private Sub SaveFile()
CD1.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
CD1.CancelError = True '취소시 오류(32755)
CD1.ShowSave '대화상자 표시
If Not Right(CD1.FileName, 4) = ".txt" Then
    CD1.FileName = CD1.FileName & ".txt"
End If
If Err.Number = 32755 Then '취소가 눌려졌다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "사용자가 열기 취소"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Err.Number = 13 Then '형식이 맞지 않다!
    CD1.FileName = "" '열려진 파일 초기화
    Err.Clear
    Mklog "나는 자연인이다!\"
    Mklog "-운지천F 광고 중\"
    Mklog "또 형식이 맞지 않단다!!!\"
    Mklog "버그다 버그!!!\"
    MsgBox "죄송합니다. 프로그램에서 잘못된 명령을 수행하여 작업이 중단됩니다...", vbCritical, "치명적인 오류"
    Exit Sub '프로시저 실행 종료(사용자가 취소함)
End If
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
End Sub
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
        FindStartPos = InStr(FindStartPos + 1, frmMain_Eng.txtText, FindText)
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

Private Sub mnuToolOption_Click()
frmOptions.Show

End Sub

Private Sub mnuTransparencyCtl_Click()
On Error GoTo Err_Trans
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
Mklog "드래그&드롭 감지(" & s & ")"
Mklog "파일 열기(" & s & ")" '로그 남김(디버그)
'RTF.FileName = s '파일 열기 처리
Dim FreeFileNum As Integer
Dim Text As String
FreeFileNum = FreeFile
Screen.MousePointer = 11
Open s For Input As #FreeFileNum
txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
Close #FreeFileNum
'txtText.text = ""
    'Open s For Input As #FreeFileNum
    'Do Until EOF(FreeFileNum)
    'Line Input #FreeFileNum, text
    'txtText.text = txtText.text & text & vbCrLf
    'Loop
    'Close #FreeFileNum
Screen.MousePointer = 0
If Not Err.Number = 0 Then
    MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Newfile = False
FileName_Dir = s
UpdateFileName Me, FileName_Dir
txtText.ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
Dirty = False
End Sub
