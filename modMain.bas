Attribute VB_Name = "modMain"

'---------------------------------------------------------------------------------------
' Module    : modMain
' DateTime  : 2012-10-20 21:41
' Author    : YJSoft
' Purpose   : Y's Notepad SE Main Module
'Hello2
'---------------------------------------------------------------------------------------
'Y's Notepad SE V.0.8
'����:������(yyj9411@naver.com)
'All rights RESERVED. :-)

'������Ʈ �α�
'12/6:���α׷� ����ȭ �۾�
'12/12:�α� ���� Ȯ���� txt���� dat�� ����, �α� ���� �̸� ����ȭ(���߿� �����ϱ� ���ϰ�)
'2012/3/8:���÷��� �� ó���� ǥ�� ��Ȱ��ȭ,Logsave RTF ��Ʈ�� ����(���� open������ ��� �۾�)
'MsgBox "frm"
Public MRUStr(5) As String
Public Dirty As Boolean '������ �����Ǿ����� ���θ� �����ϴ� �����Դϴ�.
Public insu As String '������ �μ� ó���� �����Դϴ�.
Public FileName_File As String '���� �̸��� �����ϴ� �����Դϴ�.
Public FileName_Dir As String '���� ���θ� �����ϴ� �����Դϴ�.
Public Newfile As Boolean '�� �������� ���θ� �����ϴ� �����Դϴ�.
Public Username As String '������ �̸��� �����ϴ� �����Դϴ�.
Public TitleMode As Byte 'Ÿ��Ʋ ǥ�� ���带 �����ϴ� �����Դϴ�.
Public IsAboutbox As Boolean '���÷��� ���� �ʱ� ��������, �޴�-���� ���� ���������� �����ϴ� ����
Public NewLogFile As Boolean
Public Const PROGRAM_TITLE = "Y's Notepad SE Beta(V." '���α׷� �⺻ Ÿ��Ʋ
Public Const PROGRAM_NAME = "Y's Notepad SE" '���α׷� �̸�
Public Const PROGRAM_KEY = "YNotepadSE" '���α׷� �ڵ�
Public Const LAST_UPDATED = "2013-09-17(2)" '������ ������Ʈ ��¥
Public Const LOGFILE = "log.dat" '�α� ���� �̸�
Public Const PROGRAM_HELPFILE = "\YNOTEPADSE.chm"
Public Const DEBUG_VERSION = True
Public FindStartPos As Integer
Public FindEndPos As Integer
Public FindText As String
Public ReplaceText As String
Public Lang As Boolean
Public UTF8_Error As Boolean
'Public Const YJSoft = "YJSoft"
Public IsAboveNT As Boolean
'�������ʹ� ���α׷��� ����
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long '���� ��ȭ ������ ����
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal deMiliseconds As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public FindReplace As Boolean
'---------------------------------------------------------------------------------------
' Procedure : LoadMRUList
' DateTime  : 2013-04-03 13:36
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub LoadMRUList()
Dim i As Integer
   On Error GoTo LoadMRUList_Error

For i = 1 To 5
    MRUStr(i) = GetSetting(PROGRAM_KEY, "MRU", CStr(i), "")
Next i

   On Error GoTo 0
   Exit Sub

LoadMRUList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadMRUList of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ChkMRU
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ChkMRU()
Dim i As Integer
Dim j As Integer
   On Error GoTo ChkMRU_Error

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

   On Error GoTo 0
   Exit Sub

ChkMRU_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChkMRU of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : UpdateMRU
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub UpdateMRU(frmdta As Form)
Dim i As Integer
   On Error GoTo UpdateMRU_Error

For i = 1 To 5
If MRUStr(i) = "" Then
    frmdta.mnuMRU(i).Enabled = False
    frmdta.mnuMRU(i).Caption = "(���� ����)"
Else
frmdta.mnuMRU(i).Caption = MRUStr(i)
frmdta.mnuMRU(i).Enabled = True
End If
Next

   On Error GoTo 0
   Exit Sub

UpdateMRU_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateMRU of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddMRU
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AddMRU(MRUSting As String)
Dim intindex As Integer
Dim i As Integer
   On Error GoTo AddMRU_Error

For i = 1 To 5
    If MRUSting = MRUStr(i) Then Exit Sub '�ߺ� ������ �������� �ʴ´�
Next i
intindex = CInt(GetSetting(PROGRAM_KEY, "MRU", "Index", 0))
Select Case intindex
Case 0 '���� ������
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

   On Error GoTo 0
   Exit Sub

AddMRU_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddMRU of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ClearMRU
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ClearMRU()
   On Error GoTo ClearMRU_Error

SaveSetting PROGRAM_KEY, "MRU", "Index", 0
SaveSetting PROGRAM_KEY, "MRU", "1", ""

SaveSetting PROGRAM_KEY, "MRU", "2", ""

SaveSetting PROGRAM_KEY, "MRU", "3", ""

SaveSetting PROGRAM_KEY, "MRU", "4", ""

SaveSetting PROGRAM_KEY, "MRU", "5", ""

   On Error GoTo 0
   Exit Sub

ClearMRU_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClearMRU of Module modMain"
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
   On Error GoTo EnCrypt_Error

    Randomize
    nKey = Int(Rnd * 256)
    For n = 1 To Len(sString)
        EnCrypt = EnCrypt & Right$("0000" & Hex$(Oct(IntToLong(AscW(Mid$(sString, n, 1))) Xor (nKey Xor &H1234 Xor n))), 5)
    Next
    EnCrypt = StrReverse$(Right$("0" & Hex$(nKey Xor &HBB), 2) & EnCrypt)

   On Error GoTo 0
   Exit Function

EnCrypt_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EnCrypt of Module modMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : DeCrypt
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DeCrypt(ByRef sHexString As String) As String '��ȣȭ
   On Error GoTo DeCrypt_Error

If Right(sHexString, 2) = vbCrLf Then
sHexString = Left(sHexString, Len(sHexString) - 2)
End If
    Dim sTemp As String, n As Long, nKey As Byte
    Dim sKey As String
    sTemp = StrReverse$(sHexString)
    nKey = CByte("&H" & Left$(sTemp, 2)) Xor &HBB
    sTemp = Mid$(sTemp, 3)

    For n = 1 To Len(sTemp) Step 5
        DeCrypt = DeCrypt & ChrW$(LongToInt(CLng("&O" & CLng("&H" & Mid$(sTemp, n, 5))) Xor (nKey Xor &H1234 Xor ((n + 4) \ 5))))
    Next

   On Error GoTo 0
   Exit Function

DeCrypt_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DeCrypt of Module modMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : IntToLong
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function IntToLong(ByVal IntNum As Integer) As Long
   On Error GoTo IntToLong_Error

    RtlMoveMemory IntToLong, IntNum, 2

   On Error GoTo 0
   Exit Function

IntToLong_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IntToLong of Module modMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : LongToInt
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function LongToInt(ByVal LongNum As Long) As Integer
   On Error GoTo LongToInt_Error

    RtlMoveMemory LongToInt, LongNum, 2

   On Error GoTo 0
   Exit Function

LongToInt_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LongToInt of Module modMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : FindWon
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function FindWon(findstr As String) As Integer '���� ������ \�� ��ġ�� ã�Ƴ��� �� ���� ��ġ�� ��ȯ�ϴ� �Լ��Դϴ�. \�� ���ٸ� 0�� ��ȯ�˴ϴ�.
Dim i As Integer
Dim tempstr As String * 1
   On Error GoTo FindWon_Error

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

   On Error GoTo 0
   Exit Function

FindWon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FindWon of Module modMain"
End Function

'####################################################################
'#######################UpdateFileName �Լ�##########################
'###################����:������(yyj9411@naver.com)###################
'###############################�μ�#################################
'###############1)Form-������ �ٲ� ���� �̸�#########################
'###############2)FileName-������ �̸�(���� ����)####################
'#####################�����ϴ� �ܺ� ����/����########################
'###############1)TitleMode-���� ���� ��ȯ(1,2,3,4)##################
'#####################2)PROGRAM_TITLE(����)##########################
'########################�����ϴ� �ܺ� �Լ�##########################
'###########################1)FindWon################################
'####################################################################
'---------------------------------------------------------------------------------------
' Procedure : UpdateFileName
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub UpdateFileName(Form As Form, FileName As String)
Dim i As Integer
   On Error GoTo UpdateFileName_Error

Select Case TitleMode
Case 1 '���� �̸��� ���ΰ� �� �ڿ�
    'If FileName = "" Then
    '    Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    '    App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")"
    'Else
        Form.Caption = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
        App.Title = PROGRAM_TITLE & App.Major & "." & App.Minor & "." & App.Revision & ")" & " - " & FileName
    'End If
Case 2 '���� �̸��� ���ΰ� �� �տ�
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

   On Error GoTo 0
   Exit Sub

UpdateFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateFileName of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : FileCheck
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function FileCheck(ChkFile As String) As Boolean
Dim a
   On Error GoTo FileCheck_Error

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

   On Error GoTo 0
   Exit Function

FileCheck_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FileCheck of Module modMain"
End Function
'#######################################################################
'###############################Sub Main()##############################
'###################����:������(yyj9411@naver.com)######################
'#######################################################################
Sub Main()
'Dim temp As String * 4
'MRU �ҷ���
ChkMRU
IsAboveNT = False
Mklog "O/S ����" & vbCrLf & fGetWindowVersion
Mklog "����ȭ ���� ���� - " & IsAboveNT
'temp = GetSetting(PROGRAM_KEY, "Install", "Language", Korean_1)
'If temp = "English" Then Lang = True '���� ���� ����(��Ÿ!)
If Val(GetSetting(PROGRAM_KEY, "Program", "Notepad", 0)) Then
    Shell "C:\Windows\notepad.exe " & Command(), vbNormalFocus
    End
End If
'DEBUG_VERSION = True
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
If Not Command() = "" Then '������ �μ��� �ִ�!
    If Command() = "/nodebug" Then GoTo debugmode
    Mklog "������ �μ� ����(" & Command() & ")"
    If Left(Command(), 1) = Chr(34) Then '������ �μ��� "�� �ִ�!(���� �̸� or ���ο� ��ĭ�� ������ ����ǥ�� �������� ���� �̸��� ������.
                                         '������ �츰 �ʿ� ����! ���� ����!
        insu = Mid(Command(), 2, Len(Command()) - 2)
        Mklog "������ �μ� ó��(" & insu & ")" 'ó���� ���� ���� �α�
    Else
    insu = Command() '������ �μ��� "�� ����!(�״��� �ҷ�����)
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
    MsgBox "���� " & insu & vbCrLf & "�� ã�� �� �����ϴ�!", vbCritical, "������ �μ� �Ľ� ����"
    Mklog "#���� ���� ���� - ������ �μ� ó�� ����" & vbCrLf & "���ϸ�:" & insu
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


'---------------------------------------------------------------------------------------
' Procedure : GetUserName
' DateTime  : 2013-04-03 13:37
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub GetUserName()
   On Error GoTo GetUserName_Error

Username = GetSetting(PROGRAM_KEY, "Program", "User", "")
If Username = "" Then '������ �̸��� �����ұ�?
    If MsgBox("���ϵ� ������ �̸��� �����ϴ�! �����Ͻðڽ��ϱ�?", vbYesNo, "������ ����") = vbYes Then
        Username = InputBox("������ �̸��� �Է��� �ּ���. �Է����� ������" & vbCrLf & Chr(34) & "(�� �� ����)" & Chr(34) & "�� ���ϵ˴ϴ�.", "������ ����", "", Screen.Width / 2, Screen.Height / 2)
        If Username = "" Then
            Username = "(�� �� ����)"
        End If
    Else
    Username = "(�� �� ����)"
    End If
SaveSetting PROGRAM_KEY, "Program", "User", Username
End If

   On Error GoTo 0
   Exit Sub

GetUserName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetUserName of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Mklog
' DateTime  : 2013-04-03 13:38
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Mklog(LogStr As String)
Dim FreeFileNum As Integer
   On Error GoTo Mklog_Error

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
If Right(LogStr, 1) = "\" Then 'logstr ���� \�� ������ �ð����� ����,log.dat�� ���¾���
    LogStr = Left(LogStr, Len(LogStr) - 1)
    Debug.Print LogStr
Else '������ �ð� ����
    Debug.Print Now() & " - " & LogStr
    If DEBUG_VERSION Then
        'frmMain.logsave.Text = frmMain.logsave.Text & Now() & " - " & LogStr & vbCrLf
        Print #FreeFileNum, Now() & " - " & LogStr
    End If
End If
Close #FreeFileNum

   On Error GoTo 0
   Exit Sub

Mklog_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Mklog of Module modMain"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : AppPath
' DateTime  : 2013-04-03 13:38
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AppPath() As String
   On Error GoTo AppPath_Error

If Right(App.Path, 1) = "\" Then
    AppPath = Left(App.Path, Len(App.Path) - 1)
Else
    AppPath = App.Path
End If

   On Error GoTo 0
   Exit Function

AppPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AppPath of Module modMain"
End Function
'#######################################################################
'###########################SaveCheck �Լ�##############################
'##############������ ������ �������� ���� �Լ��Դϴ�.##################
'###################����:������(yyj9411@naver.com)######################
'###############################�μ�####################################
'###############1)Ritf-��ġ �ؽ�Ʈ �ڽ� ��Ʈ���� �̸�###################
'###############2)Cd-���� ��ȭ���� ��Ʈ���� �̸�########################
'###########��ȯ��(True-�Լ� ���� �Ϸ�/False-�Լ� ���� ����#############
'#######################################################################
Public Function SaveCheck(Cd As CommonDialog) As Boolean
On Error Resume Next
Dim Respond As VbMsgBoxResult
Respond = MsgBox("������ �����Ǿ����ϴ�." & vbCrLf & "�����Ͻðڽ��ϱ�?", vbExclamation + vbYesNoCancel, "���� ����")
If Respond = vbYes Then '�����Ѵ�
    If FileName_Dir = "���� ����" Then
        '������ ������ ����(�� �����̴�)
        Cd.Filter = "�ؽ�Ʈ ����|*.txt|���� ����|*.*" '���� ���� ��ȭ���� �÷��� ����
        Cd.CancelError = True '���ҽ� ����(32755)
        Cd.ShowSave '��ȭ���� ǥ��
        If Err.Number = 32755 Then '���Ұ� ��������!
            Cd.FileName = "" '�Էµ� ���� �ʱ�ȭ
            Err.Clear
            Mklog "�����ڰ� ���� ����"
            SaveCheck = False
            Exit Function '���ν��� ���� ����(�����ڰ� ������)
        End If
        If Err.Number = 13 Then '������ ���� �ʴ�!
            Cd.FileName = "" '������ ���� �ʱ�ȭ
            Err.Clear
            Mklog "�� ������ ���� �ʴܴ�!!!\"
            Mklog "���״� ����!!!\"
            MsgBox "�˼��մϴ�. ���α׷����� �߸��� ������ �����Ͽ� �۾��� �ߴܵ˴ϴ�...", vbCritical, "ġ������ ����"
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
    Mklog "���� ����(" & Cd.FileName & ")" '�α� ����(������)
    'frmMain.RTF.Text = frmMain.txtText.Text
    'Ritf.SaveFile Cd.FileName, rtfText '���� ���� ó��
    'frmMain.txtText.Text
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open Cd.FileName For Output As #FreeFileNum
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
    AddMRU Cd.FileName
ElseIf Respond = vbNo Then
    SaveCheck = True
Else
    SaveCheck = False
End If
End Function
Public Sub �̱���()

End Sub
