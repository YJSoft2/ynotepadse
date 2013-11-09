Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' DateTime  : 2012-10-20 21:41
' Author    : YJSoft
' Purpose   : Y's Notepad SE Main Module
'---------------------------------------------------------------------------------------
'Y's Notepad SE V.0.8
'제작:유영재(yyj9411@naver.com)
'All rights RESERVED. :-)

'업데이트 로그
'12/6:프로그램 안정화 작업
'12/12:로그 파일 확장자 txt에서 dat로 변경, 로그 파일 이름 상수화(나중에 수정하기 편하게)
'2012/3/8:스플래시 폼 처음에 표시 비활성화,Logsave RTF 컨트롤 삭제(직접 open문으로 열어서 작업)
'MsgBox "frm"
Public MRUStr(5) As String
Public Dirty As Boolean '파일이 변경되었는지 여부를 저장하는 변수입니다.
Public insu As String '명령줄 인수 처리용 변수입니다.
Public FileName_File As String '파일 이름을 저장하는 변수입니다.
Public FileName_Dir As String '파일 경로를 저장하는 변수입니다.
Public Newfile As Boolean '새 파일인지 여부를 저장하는 변수입니다.
Public Username As String '사용자 이름을 저장하는 변수입니다.
Public TitleMode As Byte '타이틀 표시 모드를 저장하는 변수입니다.
Public IsAboutbox As Boolean '스플래시 폼이 초기 실행인지, 메뉴-정보 로의 실행인지를 구별하는 변수
Public NewLogFile As Boolean
Public Const PROGRAM_TITLE = "Y's Notepad SE Beta(V." '프로그램 기본 타이틀
Public Const PROGRAM_NAME = "Y's Notepad SE" '프로그램 이름
Public Const PROGRAM_KEY = "YNotepadSE" '프로그램 코드
Public Const LAST_UPDATED = "2013-02-27" '마지막 업데이트 날짜
Public Const LOGFILE = "log.dat" '로그 파일 이름
Public DEBUG_VERSION As Boolean
Public FindStartPos As Integer
Public FindEndPos As Integer
Public FindText As String
Public ReplaceText As String
Public Lang As Boolean
'Public Const YJSoft = "YJSoft"

'여기부터는 프로그램용 선언
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long '정보 대화 상자의 선언
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal deMiliseconds As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public FindReplace As Boolean
Public Sub LoadMRUList()
Dim i As Integer
For i = 1 To 5
    MRUStr(i) = GetSetting(PROGRAM_KEY, "MRU", CStr(i), "")
Next i
End Sub
Public Sub ChkMRU()
Dim i As Integer
Dim j As Integer
LoadMRUList
For i = 1 To 5
    If MRUStr(i) = "" Then
        If i = 5 Then
            SaveSetting PROGRAM_KEY, "MRU", "Index", 4
        Else
            For j = i To 4
                SaveSetting PROGRAM_KEY, "MRU", CStr(j), MRUStr(j + 1)
            Next j
            SaveSetting PROGRAM_KEY, "MRU", "5", ""
        End If
    End If
Next i
For i = 1 To 5
    If MRUStr(i) = "" Then
        SaveSetting PROGRAM_KEY, "MRU", "Index", i - 1
        Exit Sub
    End If
Next
SaveSetting PROGRAM_KEY, "MRU", "Index", 5
End Sub
Public Sub UpdateMRU(frmdta As Form)
Dim i As Integer
For i = 1 To 5
If MRUStr(i) = "" Then
    frmdta.mnuMRU(i).Enabled = False
    frmdta.mnuMRU(i).Caption = "(파일 없음)"
Else
frmdta.mnuMRU(i).Caption = MRUStr(i)
frmdta.mnuMRU(i).Enabled = True
End If
Next
End Sub

Public Sub AddMRU(MRUSting As String)
Dim intindex As Integer
Dim i As Integer
For i = 1 To 5
    If MRUSting = MRUStr(i) Then Exit Sub '중복 파일은 기록하지 않는다
Next i
intindex = CInt(GetSetting(PROGRAM_KEY, "MRU", "Index", 0))
Select Case intindex
Case 0 '새로 만든다
SaveSetting PROGRAM_KEY, "MRU", "Index", 1
SaveSetting PROGRAM_KEY, "MRU", "1", MRUSting
Case 1
SaveSetting PROGRAM_KEY, "MRU", "Index", 2
SaveSetting PROGRAM_KEY, "MRU", "2", MRUSting
Case 2
SaveSetting PROGRAM_KEY, "MRU", "Index", 3
SaveSetting PROGRAM_KEY, "MRU", "3", MRUSting
Case 3
SaveSetting PROGRAM_KEY, "MRU", "Index", 4
SaveSetting PROGRAM_KEY, "MRU", "4", MRUSting
Case 4
SaveSetting PROGRAM_KEY, "MRU", "Index", 5
SaveSetting PROGRAM_KEY, "MRU", "5", MRUSting
Case 5
SaveSetting PROGRAM_KEY, "MRU", "1", MRUStr(2)
SaveSetting PROGRAM_KEY, "MRU", "2", MRUStr(3)
SaveSetting PROGRAM_KEY, "MRU", "3", MRUStr(4)
SaveSetting PROGRAM_KEY, "MRU", "4", MRUStr(5)
SaveSetting PROGRAM_KEY, "MRU", "5", MRUSting
End Select
End Sub
Public Sub ClearMRU()
SaveSetting PROGRAM_KEY, "MRU", "Index", 0
SaveSetting PROGRAM_KEY, "MRU", "1", ""

SaveSetting PROGRAM_KEY, "MRU", "2", ""

SaveSetting PROGRAM_KEY, "MRU", "3", ""

SaveSetting PROGRAM_KEY, "MRU", "4", ""

SaveSetting PROGRAM_KEY, "MRU", "5", ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnCrypt
' DateTime  : 2012-08-05 20:05
' Author    : PC1
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function EnCrypt(ByRef sString As String) As String '암호화
    Dim n As Long, nKey As Byte
    Randomize
    nKey = Int(Rnd * 256)
    For n = 1 To Len(sString)
        EnCrypt = EnCrypt & Right$("0000" & Hex$(Oct(IntToLong(AscW(Mid$(sString, n, 1))) Xor (nKey Xor &H1234 Xor n))), 5)
    Next
    EnCrypt = StrReverse$(Right$("0" & Hex$(nKey Xor &HBB), 2) & EnCrypt)
End Function

Public Function DeCrypt(ByRef sHexString As String) As String '복호화
    Dim sTemp As String, n As Long, nKey As Byte
    sTemp = StrReverse$(sHexString)
    nKey = CByte("&H" & Left$(sTemp, 2)) Xor &HBB
    sTemp = Mid$(sTemp, 3)

    For n = 1 To Len(sTemp) Step 5
        DeCrypt = DeCrypt & ChrW$(LongToInt(CLng("&O" & CLng("&H" & Mid$(sTemp, n, 5))) Xor (nKey Xor &H1234 Xor ((n + 4) \ 5))))
    Next
End Function

Private Function IntToLong(ByVal IntNum As Integer) As Long
    RtlMoveMemory IntToLong, IntNum, 2
End Function

Private Function LongToInt(ByVal LongNum As Long) As Integer
    RtlMoveMemory LongToInt, LongNum, 2
End Function

Function FindWon(findstr As String) As Integer '문장 내에서 \의 위치를 찾아내어 그 다음 위치를 반환하는 함수입니다. \가 없다면 0이 반환됩니다.
Dim i As Integer
Dim tempstr As String * 1
If findstr = "제목 없음" And Newfile = True Then
    FindWon = 0
    Exit Function
End If
For i = Len(findstr) To 1 Step -1
    tempstr = Mid(findstr, i, 1)
    'Mklog "modMain.FindWon.tempstr = " & tempstr
    If tempstr = "\" Then
        FindWon = i
        'Mklog "modMain.FindWon - " & Chr(34) & "\" & Chr(34) & "위치 찾음(" & i & ")"
        Exit Function
    End If
Next
'Mklog "modMain.FindWon - 문장 안에 " & Chr(34) & "\" & Chr(34) & "가 없음."
FindWon = 0
End Function

'####################################################################
'#######################UpdateFileName 함수##########################
'###################제작:유영재(yyj9411@naver.com)###################
'###############################인수#################################
'###############1)Form-제목을 바꿀 폼의 이름#########################
'###############2)FileName-파일의 이름(경로 포함)####################
'#####################사용하는 외부 변수/상수########################
'###############1)TitleMode-제목 서식 반환(1,2,3,4)##################
'#####################2)PROGRAM_TITLE(상수)##########################
'########################사용하는 외부 함수##########################
'###########################1)FindWon################################
'####################################################################
Public Sub UpdateFileName(Form As Form, FileName As String)
Dim i As Integer
Select Case TitleMode
Case 1 '파일 이름과 경로가 맨 뒤에
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
        App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
    'End If
Case 2 '파일 이름과 경로가 맨 앞에
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        Form.Caption = FileName & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        App.Title = FileName & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'End If
Case 3 '파일 이름이 맨 뒤에
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "파일 이름 추출 - " & FileName_File
            Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName_File
            App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName_File
        End If
    'End If
Case 4 '파일 이름이 맨 앞에
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "파일 이름 추출 - " & FileName_File
            Form.Caption = FileName_File & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
            App.Title = FileName_File & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        End If
    'End If
Case 5 '파일 이름만-베타!
    'If FileName = "" Then
    '    Form.Caption = "제목 없음"
    '    App.Title = "제목 없음"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "파일 이름 추출 - " & FileName_File
            Form.Caption = FileName_File ' & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
            App.Title = FileName_File ' & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        End If
    'End If
End Select
End Sub
Public Sub 비밀이랑께(나랑께 As Form)
'Mklog Left(나랑께.RTF.Text, 11) & 1
'Mklog Mid(나랑께.RTF.Text, 13, 1) & 2
'On Error Resume Next
'Mklog Mid(나랑께.RTF.Text, 14, Len(나랑께.RTF.Text) - 13) & 3
If Left(나랑께.txtText.Text, 11) = "이거 누가 만든 거임" Then
    If Mid(나랑께.txtText.Text, 13, 1) = Chr(34) Then
        'Debug.Print Mid(나랑께.RTF.Text, Len(나랑께.RTF.Text), 1)
        If Mid(나랑께.txtText.Text, Len(나랑께.txtText.Text), 1) = Chr(34) Then
            Dim aaaaa_OS2 As String
            If Len(나랑께.txtText.Text) - 14 = 0 Then GoTo A11
            aaaaa_OS2 = Mid(나랑께.txtText.Text, 14, Len(나랑께.txtText.Text) - 14)
            이거_누가_만든_거임 aaaaa_OS2
        Else
            Dim i As Integer
                For i = 1 To 10
A11:                '잘못 썼다!
                    MsgBox "나랑께 빨리 문좀열어보랑께 " & i & "/10", vbCritical, "호성성님"
                Next
            End
        End If
    Else
    이거_누가_만든_거임
    End If
End If
End Sub
Function FileCheck(ChkFile As String) As Boolean
Dim a
On Error GoTo n
a = FileLen(ChkFile)
If a > 1000000 Then '로그 파일 용량이 너무 크다!
    Mklog 1 '로그 파일 초기화
End If
FileCheck = True
Exit Function
n:
FileCheck = False
Err.Clear
End Function
'#######################################################################
'###############################Sub Main()##############################
'###################제작:유영재(yyj9411@naver.com)######################
'#######################################################################
Sub Main()
Dim temp As String * 4
ChkMRU
ChkMRU
ChkMRU
ChkMRU
ChkMRU
temp = GetSetting(PROGRAM_KEY, "Install", "Language", Korean_1)
If temp = "English" Then Lang = True '영문 실행 모드(베타!)
If Val(GetSetting(PROGRAM_KEY, "Program", "Notepad", 0)) Then
    Shell "C:\Windows\notepad.exe " & Command(), vbNormalFocus
    End
End If
If Command() = "/nodebug" Then
    DEBUG_VERSION = False
Else
    DEBUG_VERSION = True
End If
On Error GoTo Err_Main
If Not FileCheck(AppPath & "\" & LOGFILE) Then
    NewLogFile = True
End If
TitleMode = GetSetting(PROGRAM_KEY, "Option", "Title", 99) '타이틀 서식을 불러옵니다.
If TitleMode = 99 Then '기본값- 처음 실행한다
    SaveSetting PROGRAM_KEY, "Option", "Title", 4
    TitleMode = 4
End If
If Not Lang Then
    Load frmMain '메인 폼을 불러들인다.
    frmMain.Top = GetSetting(PROGRAM_KEY, "Window", "X", Screen.Height / 2)
    frmMain.Left = GetSetting(PROGRAM_KEY, "Window", "Y", Screen.Width / 2)
    frmMain.Width = GetSetting(PROGRAM_KEY, "Window", "Width", 8000)
    frmMain.Height = GetSetting(PROGRAM_KEY, "Window", "Height", 7000)
    If Val(GetSetting(PROGRAM_KEY, "Window", "최대화")) Then
        frmMain.WindowState = 2
    End If
Else
    MsgBox "English Execute Mode is under construction. Sorry :(", vbInformation, "Y's Notepad SE English Team :)"
    End
End If
If Not Command() = "" Then '명령줄 인수가 있다!
    If Command() = "/nodebug" Then GoTo debugmode
    Mklog "명령줄 인수 감지(" & Command() & ")"
    If Left(Command(), 1) = Chr(34) Then '명령줄 인수에 "가 있다!(파일 이름 or 경로에 빈칸이 있으면 따옴표로 감싸져서 파일 이름이 들어옴.
                                         '하지만 우린 필요 없다! 고로 삭제!
        insu = Mid(Command(), 2, Len(Command()) - 2)
        Mklog "명령줄 인수 처리(" & insu & ")" '처리된 파일 경로 로깅
    Else
    insu = Command() '명령줄 인수에 "가 없다!(그대로 불러들임)
    End If
'frmMain.RTF.FileName = insu '파일 불러들이기!
Dim FreeFileNum As Integer
FreeFileNum = FreeFile
Open insu For Input As #FreeFileNum
Newfile = False
frmMain.txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
Close #FreeFileNum
FileName_Dir = insu
frmMain.CD1.FileName = FileName_Dir '버그 원인 #1
UpdateFileName frmMain, FileName_Dir '타이틀 변경(파일 이름으로..)
AddMRU insu '최근 연 파일에 추가
Dirty = False
Else
    Newfile = True '새 파일임
    frmMain.CD1.FileName = ""
End If

debugmode:
GetUserName
If GetSetting(PROGRAM_KEY, "Option", "Toolbar", False) Then
    frmMain.tbTools.Visible = True
Else
    frmMain.tbTools.Visible = False
    frmMain.mnuToolbar.Caption = "툴바 보이기(&B)"
End If
frmMain.Show
Exit Sub

Err_Main:
If Err.Number = 75 Then
    MsgBox "파일 " & insu & vbCrLf & "을 찾을 수 없습니다!", vbCritical, "명령줄 인수 파싱 오류"
    Mklog "#파일 열기 오류 - 명령줄 인수 처리 실패" & vbCrLf & "파일명:" & insu
    Err.Clear
Else
    MsgBox "처리되지 않은 오류가 발생되었습니다!" & vbCrLf & "오류코드:" & Err.Number & vbCrLf & Err.Description, vbCritical, "치명적인 오류"
    Mklog Err.Number & "/" & Err.Description & "/" & insu
    Err.Clear
End If
frmMain.txtText.Text = "" '텍스트 내용 제거
frmMain.CD1.FileName = ""
GetUserName
frmMain.Show
Dirty = False

End Sub
Public Sub 이거_누가_만든_거임(Optional 그냥그냥 As String = "<<피실험자 이름>>")
MsgBox "yyj9411@naver.com이 만든거다! " & 그냥그냥 & "아~", vbInformation + vbYesNo

End Sub
'사용자 이름을 구하는 함수입니다.

Public Sub GetUserName()
Username = GetSetting(PROGRAM_KEY, "Program", "User", "")
If Username = "" Then '사용자 이름을 등록할까?
    If MsgBox("등록된 사용자 이름이 없습니다! 등록하시겠습니까?", vbYesNo, "사용자 등록") = vbYes Then
        Username = InputBox("사용자 이름을 입력해 주세요. 입력하지 않을시" & vbCrLf & Chr(34) & "(알 수 없음)" & Chr(34) & "로 등록됩니다.", "사용자 등록", "", Screen.Width / 2, Screen.Height / 2)
        If Username = "" Then
            Username = "(알 수 없음)"
        End If
    Else
    Username = "(알 수 없음)"
    End If
SaveSetting PROGRAM_KEY, "Program", "User", Username
End If
End Sub
Public Sub Mklog(LogStr As String)
Dim FreeFileNum As Integer
If NewLogFile Then
    FreeFileNum = FreeFile
    Open AppPath & "\" & LOGFILE For Output As #FreeFileNum
        Print #FreeFileNum, Now() & " - " & "로그 파일이 생성되었습니다."
        Print #FreeFileNum, Now() & " - " & LogStr
    Close #FreeFileNum
    NewLogFile = False
    Exit Sub
End If
FreeFileNum = FreeFile
If Val(LogStr) = 1 Then
    FreeFileNum = FreeFile
    Open AppPath & "\" & LOGFILE For Output As #FreeFileNum
        Print #FreeFileNum, Now() & " - " & "로그 파일이 초기화되었습니다."
    Close #FreeFileNum
    NewLogFile = False
    Exit Sub
End If
Open AppPath & "\" & LOGFILE For Append As #FreeFileNum
If Right(LogStr, 1) = "\" Then 'logstr 끝에 \가 있으면 시간출력 제외,log.dat에 출력안함
    LogStr = Left(LogStr, Len(LogStr) - 1)
    Debug.Print LogStr
Else '없으면 시간 출력
    Debug.Print Now() & " - " & LogStr
    If DEBUG_VERSION Then
        'frmMain.logsave.Text = frmMain.logsave.Text & Now() & " - " & LogStr & vbCrLf
        Print #FreeFileNum, Now() & " - " & LogStr
    End If
End If
Close #FreeFileNum
End Sub
Public Function AppPath() As String
If Right(App.Path, 1) = "\" Then
    AppPath = Left(App.Path, Len(App.Path) - 1)
Else
    AppPath = App.Path
End If
End Function
'#######################################################################
'###########################SaveCheck 함수##############################
'##############파일을 저장할 것인지를 묻는 함수입니다.##################
'###################제작:유영재(yyj9411@naver.com)######################
'###############################인수####################################
'###############1)Ritf-리치 텍스트 박스 컨트롤의 이름###################
'###############2)Cd-공통 대화상자 컨트롤의 이름########################
'###########반환값(True-함수 실행 완료/False-함수 실행 취소#############
'#######################################################################
Public Function SaveCheck(Cd As CommonDialog) As Boolean
On Error Resume Next
Dim Respond As VbMsgBoxResult
Respond = MsgBox("파일이 변경되었습니다." & vbCrLf & "저장하시겠습니까?", vbExclamation + vbYesNoCancel, "파일 변경")
If Respond = vbYes Then '저장한다
    If FileName_Dir = "제목 없음" Then
        '열려진 파일이 없다(새 파일이다)
        Cd.Filter = "텍스트 파일|*.txt|모든 파일|*.*" '파일 열기 대화상자 플래그 설정
        Cd.CancelError = True '취소시 오류(32755)
        Cd.ShowSave '대화상자 표시
        If Err.Number = 32755 Then '취소가 눌려졌다!
            Cd.FileName = "" '입력된 파일 초기화
            Err.Clear
            Mklog "사용자가 저장 취소"
            SaveCheck = False
            Exit Function '프로시저 실행 종료(사용자가 취소함)
        End If
        If Err.Number = 13 Then '형식이 맞지 않다!
            Cd.FileName = "" '열려진 파일 초기화
            Err.Clear
            Mklog "또 형식이 맞지 않단다!!!\"
            Mklog "버그다 버그!!!\"
            MsgBox "죄송합니다. 프로그램에서 잘못된 명령을 수행하여 작업이 중단됩니다...", vbCritical, "치명적인 오류"
            SaveCheck = False
            Exit Function '프로시저 실행 종료(버그)
        End If
        If Not Err.Number = 0 Then
            MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
            Mklog Err.Number & "/" & Err.Description
            SaveCheck = False
            Exit Function
        End If
    Else
        Cd.FileName = FileName_Dir '이미 열려진 파일이 있다-열려진 파일 이름을 Cd.filename에 대입
    End If
    Mklog "파일 저장(" & Cd.FileName & ")" '로그 남김(디버그)
    'frmMain.RTF.Text = frmMain.txtText.Text
    'Ritf.SaveFile Cd.FileName, rtfText '파일 저장 처리
    'frmMain.txtText.Text
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open Cd.FileName For Output As #FreeFileNum
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Print #FreeFileNum, frmMain.txtText.Text
    Close #FreeFileNum
    Screen.MousePointer = 0
    If Not Err.Number = 0 Then
        MsgBox "오류 발생!" & vbCrLf & "오류 번호:" & Err.Number & vbCrLf & Err.Description, vbCritical, "오류!"
        Mklog Err.Number & "/" & Err.Description
        Err.Clear
        SaveCheck = False
        Exit Function
    End If
    Dirty = False
    SaveCheck = True
ElseIf Respond = vbNo Then
    SaveCheck = True
Else
    SaveCheck = False
End If
End Function
Public Sub 미구현()

End Sub
