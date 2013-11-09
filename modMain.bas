Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' DateTime  : 2012-10-20 21:41
' Author    : YJSoft
' Purpose   : Y's Notepad SE Main Module
'---------------------------------------------------------------------------------------
'Y's Notepad SE V.0.8
'����:������(yyj9411@naver.com)
'All rights RESERVED. :-)

'������Ʈ �α�
'12/6:���α׷� ����ȭ �۾�
'12/12:�α� ���� Ȯ���� txt���� dat�� ����, �α� ���� �̸� ���ȭ(���߿� �����ϱ� ���ϰ�)
'2012/3/8:���÷��� �� ó���� ǥ�� ��Ȱ��ȭ,Logsave RTF ��Ʈ�� ����(���� open������ ��� �۾�)
'MsgBox "frm"
Public MRUStr(5) As String
Public Dirty As Boolean '������ ����Ǿ����� ���θ� �����ϴ� �����Դϴ�.
Public insu As String '����� �μ� ó���� �����Դϴ�.
Public FileName_File As String '���� �̸��� �����ϴ� �����Դϴ�.
Public FileName_Dir As String '���� ��θ� �����ϴ� �����Դϴ�.
Public Newfile As Boolean '�� �������� ���θ� �����ϴ� �����Դϴ�.
Public Username As String '����� �̸��� �����ϴ� �����Դϴ�.
Public TitleMode As Byte 'Ÿ��Ʋ ǥ�� ��带 �����ϴ� �����Դϴ�.
Public IsAboutbox As Boolean '���÷��� ���� �ʱ� ��������, �޴�-���� ���� ���������� �����ϴ� ����
Public NewLogFile As Boolean
Public Const PROGRAM_TITLE = "Y's Notepad SE Beta(V." '���α׷� �⺻ Ÿ��Ʋ
Public Const PROGRAM_NAME = "Y's Notepad SE" '���α׷� �̸�
Public Const PROGRAM_KEY = "YNotepadSE" '���α׷� �ڵ�
Public Const LAST_UPDATED = "2013-02-27" '������ ������Ʈ ��¥
Public Const LOGFILE = "log.dat" '�α� ���� �̸�
Public DEBUG_VERSION As Boolean
Public FindStartPos As Integer
Public FindEndPos As Integer
Public FindText As String
Public ReplaceText As String
Public Lang As Boolean
'Public Const YJSoft = "YJSoft"

'������ʹ� ���α׷��� ����
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long '���� ��ȭ ������ ����
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
    frmdta.mnuMRU(i).Caption = "(���� ����)"
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
    If MRUSting = MRUStr(i) Then Exit Sub '�ߺ� ������ ������� �ʴ´�
Next i
intindex = CInt(GetSetting(PROGRAM_KEY, "MRU", "Index", 0))
Select Case intindex
Case 0 '���� �����
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
Public Function EnCrypt(ByRef sString As String) As String '��ȣȭ
    Dim n As Long, nKey As Byte
    Randomize
    nKey = Int(Rnd * 256)
    For n = 1 To Len(sString)
        EnCrypt = EnCrypt & Right$("0000" & Hex$(Oct(IntToLong(AscW(Mid$(sString, n, 1))) Xor (nKey Xor &H1234 Xor n))), 5)
    Next
    EnCrypt = StrReverse$(Right$("0" & Hex$(nKey Xor &HBB), 2) & EnCrypt)
End Function

Public Function DeCrypt(ByRef sHexString As String) As String '��ȣȭ
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

Function FindWon(findstr As String) As Integer '���� ������ \�� ��ġ�� ã�Ƴ��� �� ���� ��ġ�� ��ȯ�ϴ� �Լ��Դϴ�. \�� ���ٸ� 0�� ��ȯ�˴ϴ�.
Dim i As Integer
Dim tempstr As String * 1
If findstr = "���� ����" And Newfile = True Then
    FindWon = 0
    Exit Function
End If
For i = Len(findstr) To 1 Step -1
    tempstr = Mid(findstr, i, 1)
    'Mklog "modMain.FindWon.tempstr = " & tempstr
    If tempstr = "\" Then
        FindWon = i
        'Mklog "modMain.FindWon - " & Chr(34) & "\" & Chr(34) & "��ġ ã��(" & i & ")"
        Exit Function
    End If
Next
'Mklog "modMain.FindWon - ���� �ȿ� " & Chr(34) & "\" & Chr(34) & "�� ����."
FindWon = 0
End Function

'####################################################################
'#######################UpdateFileName �Լ�##########################
'###################����:������(yyj9411@naver.com)###################
'###############################�μ�#################################
'###############1)Form-������ �ٲ� ���� �̸�#########################
'###############2)FileName-������ �̸�(��� ����)####################
'#####################����ϴ� �ܺ� ����/���########################
'###############1)TitleMode-���� ���� ��ȯ(1,2,3,4)##################
'#####################2)PROGRAM_TITLE(���)##########################
'########################����ϴ� �ܺ� �Լ�##########################
'###########################1)FindWon################################
'####################################################################
Public Sub UpdateFileName(Form As Form, FileName As String)
Dim i As Integer
Select Case TitleMode
Case 1 '���� �̸��� ��ΰ� �� �ڿ�
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
        App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
    'End If
Case 2 '���� �̸��� ��ΰ� �� �տ�
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        Form.Caption = FileName & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        App.Title = FileName & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'End If
Case 3 '���� �̸��� �� �ڿ�
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "���� �̸� ���� - " & FileName_File
            Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName_File
            App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName_File
        End If
    'End If
Case 4 '���� �̸��� �� �տ�
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "���� �̸� ���� - " & FileName_File
            Form.Caption = FileName_File & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
            App.Title = FileName_File & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        End If
    'End If
Case 5 '���� �̸���-��Ÿ!
    'If FileName = "" Then
    '    Form.Caption = "���� ����"
    '    App.Title = "���� ����"
    'Else
        If Not Len(FileName) <= 1 Then
            i = FindWon(FileName)
            FileName_File = Mid(FileName, i + 1, Len(FileName) - i)
            Mklog "���� �̸� ���� - " & FileName_File
            Form.Caption = FileName_File ' & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
            App.Title = FileName_File ' & " - " & PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
        End If
    'End If
End Select
End Sub
Public Sub ����̶���(������ As Form)
'Mklog Left(������.RTF.Text, 11) & 1
'Mklog Mid(������.RTF.Text, 13, 1) & 2
'On Error Resume Next
'Mklog Mid(������.RTF.Text, 14, Len(������.RTF.Text) - 13) & 3
If Left(������.txtText.Text, 11) = "�̰� ���� ���� ����" Then
    If Mid(������.txtText.Text, 13, 1) = Chr(34) Then
        'Debug.Print Mid(������.RTF.Text, Len(������.RTF.Text), 1)
        If Mid(������.txtText.Text, Len(������.txtText.Text), 1) = Chr(34) Then
            Dim aaaaa_OS2 As String
            If Len(������.txtText.Text) - 14 = 0 Then GoTo A11
            aaaaa_OS2 = Mid(������.txtText.Text, 14, Len(������.txtText.Text) - 14)
            �̰�_����_����_���� aaaaa_OS2
        Else
            Dim i As Integer
                For i = 1 To 10
A11:                '�߸� ���!
                    MsgBox "������ ���� ����������� " & i & "/10", vbCritical, "ȣ������"
                Next
            End
        End If
    Else
    �̰�_����_����_����
    End If
End If
End Sub
Function FileCheck(ChkFile As String) As Boolean
Dim a
On Error GoTo n
a = FileLen(ChkFile)
If a > 1000000 Then '�α� ���� �뷮�� �ʹ� ũ��!
    Mklog 1 '�α� ���� �ʱ�ȭ
End If
FileCheck = True
Exit Function
n:
FileCheck = False
Err.Clear
End Function
'#######################################################################
'###############################Sub Main()##############################
'###################����:������(yyj9411@naver.com)######################
'#######################################################################
Sub Main()
Dim temp As String * 4
ChkMRU
ChkMRU
ChkMRU
ChkMRU
ChkMRU
temp = GetSetting(PROGRAM_KEY, "Install", "Language", Korean_1)
If temp = "English" Then Lang = True '���� ���� ���(��Ÿ!)
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
TitleMode = GetSetting(PROGRAM_KEY, "Option", "Title", 99) 'Ÿ��Ʋ ������ �ҷ��ɴϴ�.
If TitleMode = 99 Then '�⺻��- ó�� �����Ѵ�
    SaveSetting PROGRAM_KEY, "Option", "Title", 4
    TitleMode = 4
End If
If Not Lang Then
    Load frmMain '���� ���� �ҷ����δ�.
    frmMain.Top = GetSetting(PROGRAM_KEY, "Window", "X", Screen.Height / 2)
    frmMain.Left = GetSetting(PROGRAM_KEY, "Window", "Y", Screen.Width / 2)
    frmMain.Width = GetSetting(PROGRAM_KEY, "Window", "Width", 8000)
    frmMain.Height = GetSetting(PROGRAM_KEY, "Window", "Height", 7000)
    If Val(GetSetting(PROGRAM_KEY, "Window", "�ִ�ȭ")) Then
        frmMain.WindowState = 2
    End If
Else
    MsgBox "English Execute Mode is under construction. Sorry :(", vbInformation, "Y's Notepad SE English Team :)"
    End
End If
If Not Command() = "" Then '����� �μ��� �ִ�!
    If Command() = "/nodebug" Then GoTo debugmode
    Mklog "����� �μ� ����(" & Command() & ")"
    If Left(Command(), 1) = Chr(34) Then '����� �μ��� "�� �ִ�!(���� �̸� or ��ο� ��ĭ�� ������ ����ǥ�� �������� ���� �̸��� ����.
                                         '������ �츰 �ʿ� ����! ��� ����!
        insu = Mid(Command(), 2, Len(Command()) - 2)
        Mklog "����� �μ� ó��(" & insu & ")" 'ó���� ���� ��� �α�
    Else
    insu = Command() '����� �μ��� "�� ����!(�״�� �ҷ�����)
    End If
'frmMain.RTF.FileName = insu '���� �ҷ����̱�!
Dim FreeFileNum As Integer
FreeFileNum = FreeFile
Open insu For Input As #FreeFileNum
Newfile = False
frmMain.txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
Close #FreeFileNum
FileName_Dir = insu
frmMain.CD1.FileName = FileName_Dir '���� ���� #1
UpdateFileName frmMain, FileName_Dir 'Ÿ��Ʋ ����(���� �̸�����..)
AddMRU insu '�ֱ� �� ���Ͽ� �߰�
Dirty = False
Else
    Newfile = True '�� ������
    frmMain.CD1.FileName = ""
End If

debugmode:
GetUserName
If GetSetting(PROGRAM_KEY, "Option", "Toolbar", False) Then
    frmMain.tbTools.Visible = True
Else
    frmMain.tbTools.Visible = False
    frmMain.mnuToolbar.Caption = "���� ���̱�(&B)"
End If
frmMain.Show
Exit Sub

Err_Main:
If Err.Number = 75 Then
    MsgBox "���� " & insu & vbCrLf & "�� ã�� �� �����ϴ�!", vbCritical, "����� �μ� �Ľ� ����"
    Mklog "#���� ���� ���� - ����� �μ� ó�� ����" & vbCrLf & "���ϸ�:" & insu
    Err.Clear
Else
    MsgBox "ó������ ���� ������ �߻��Ǿ����ϴ�!" & vbCrLf & "�����ڵ�:" & Err.Number & vbCrLf & Err.Description, vbCritical, "ġ������ ����"
    Mklog Err.Number & "/" & Err.Description & "/" & insu
    Err.Clear
End If
frmMain.txtText.Text = "" '�ؽ�Ʈ ���� ����
frmMain.CD1.FileName = ""
GetUserName
frmMain.Show
Dirty = False

End Sub
Public Sub �̰�_����_����_����(Optional �׳ɱ׳� As String = "<<�ǽ����� �̸�>>")
MsgBox "yyj9411@naver.com�� ����Ŵ�! " & �׳ɱ׳� & "��~", vbInformation + vbYesNo

End Sub
'����� �̸��� ���ϴ� �Լ��Դϴ�.

Public Sub GetUserName()
Username = GetSetting(PROGRAM_KEY, "Program", "User", "")
If Username = "" Then '����� �̸��� ����ұ�?
    If MsgBox("��ϵ� ����� �̸��� �����ϴ�! ����Ͻðڽ��ϱ�?", vbYesNo, "����� ���") = vbYes Then
        Username = InputBox("����� �̸��� �Է��� �ּ���. �Է����� ������" & vbCrLf & Chr(34) & "(�� �� ����)" & Chr(34) & "�� ��ϵ˴ϴ�.", "����� ���", "", Screen.Width / 2, Screen.Height / 2)
        If Username = "" Then
            Username = "(�� �� ����)"
        End If
    Else
    Username = "(�� �� ����)"
    End If
SaveSetting PROGRAM_KEY, "Program", "User", Username
End If
End Sub
Public Sub Mklog(LogStr As String)
Dim FreeFileNum As Integer
If NewLogFile Then
    FreeFileNum = FreeFile
    Open AppPath & "\" & LOGFILE For Output As #FreeFileNum
        Print #FreeFileNum, Now() & " - " & "�α� ������ �����Ǿ����ϴ�."
        Print #FreeFileNum, Now() & " - " & LogStr
    Close #FreeFileNum
    NewLogFile = False
    Exit Sub
End If
FreeFileNum = FreeFile
If Val(LogStr) = 1 Then
    FreeFileNum = FreeFile
    Open AppPath & "\" & LOGFILE For Output As #FreeFileNum
        Print #FreeFileNum, Now() & " - " & "�α� ������ �ʱ�ȭ�Ǿ����ϴ�."
    Close #FreeFileNum
    NewLogFile = False
    Exit Sub
End If
Open AppPath & "\" & LOGFILE For Append As #FreeFileNum
If Right(LogStr, 1) = "\" Then 'logstr ���� \�� ������ �ð���� ����,log.dat�� ��¾���
    LogStr = Left(LogStr, Len(LogStr) - 1)
    Debug.Print LogStr
Else '������ �ð� ���
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
'###########################SaveCheck �Լ�##############################
'##############������ ������ �������� ���� �Լ��Դϴ�.##################
'###################����:������(yyj9411@naver.com)######################
'###############################�μ�####################################
'###############1)Ritf-��ġ �ؽ�Ʈ �ڽ� ��Ʈ���� �̸�###################
'###############2)Cd-���� ��ȭ���� ��Ʈ���� �̸�########################
'###########��ȯ��(True-�Լ� ���� �Ϸ�/False-�Լ� ���� ���#############
'#######################################################################
Public Function SaveCheck(Cd As CommonDialog) As Boolean
On Error Resume Next
Dim Respond As VbMsgBoxResult
Respond = MsgBox("������ ����Ǿ����ϴ�." & vbCrLf & "�����Ͻðڽ��ϱ�?", vbExclamation + vbYesNoCancel, "���� ����")
If Respond = vbYes Then '�����Ѵ�
    If FileName_Dir = "���� ����" Then
        '������ ������ ����(�� �����̴�)
        Cd.Filter = "�ؽ�Ʈ ����|*.txt|��� ����|*.*" '���� ���� ��ȭ���� �÷��� ����
        Cd.CancelError = True '��ҽ� ����(32755)
        Cd.ShowSave '��ȭ���� ǥ��
        If Err.Number = 32755 Then '��Ұ� ��������!
            Cd.FileName = "" '�Էµ� ���� �ʱ�ȭ
            Err.Clear
            Mklog "����ڰ� ���� ���"
            SaveCheck = False
            Exit Function '���ν��� ���� ����(����ڰ� �����)
        End If
        If Err.Number = 13 Then '������ ���� �ʴ�!
            Cd.FileName = "" '������ ���� �ʱ�ȭ
            Err.Clear
            Mklog "�� ������ ���� �ʴܴ�!!!\"
            Mklog "���״� ����!!!\"
            MsgBox "�˼��մϴ�. ���α׷����� �߸��� ����� �����Ͽ� �۾��� �ߴܵ˴ϴ�...", vbCritical, "ġ������ ����"
            SaveCheck = False
            Exit Function '���ν��� ���� ����(����)
        End If
        If Not Err.Number = 0 Then
            MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
            Mklog Err.Number & "/" & Err.Description
            SaveCheck = False
            Exit Function
        End If
    Else
        Cd.FileName = FileName_Dir '�̹� ������ ������ �ִ�-������ ���� �̸��� Cd.filename�� ����
    End If
    Mklog "���� ����(" & Cd.FileName & ")" '�α� ����(�����)
    'frmMain.RTF.Text = frmMain.txtText.Text
    'Ritf.SaveFile Cd.FileName, rtfText '���� ���� ó��
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
        MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
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
Public Sub �̱���()

End Sub
