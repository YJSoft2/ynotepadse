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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer �ڼ�ȿ��_������ 
      Left            =   2040
      Top             =   720
   End
   Begin VB.TextBox txtText 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '����
      ScrollBars      =   3  '�����
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
      Caption         =   "����(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "���� �����(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "����(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "�ٸ� �̸����� ����(&A)..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "������ ����(&U)..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "�μ�(&P)..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFastPrint 
         Caption         =   "���� �μ�(&F)"
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "������(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "���� ���(&U)"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "�߶󳻱�(&T)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "�ٿ��ֱ�(&P)"
      End
      Begin VB.Menu dfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoLinePass 
         Caption         =   "�ڵ� �ٳѱ�(&A)"
      End
      Begin VB.Menu sdgfsdgsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "ã��(&F)-Beta!"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "���� ã��(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "�ٲٱ�(&R)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuReplaceNext 
         Caption         =   "���� �ٲٱ�(&E)"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "����(&O)"
      Begin VB.Menu mnuFont 
         Caption         =   "�۲�(&T)..."
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "����(&B)..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "����(&T)"
      Begin VB.Menu mnuLogopn 
         Caption         =   "�α� ���� ����(&O)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLogClr 
         Caption         =   "�α� ���� �ʱ�ȭ(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUserChg 
         Caption         =   "����� �̸� ����(&C)"
      End
      Begin VB.Menu sdfsdfs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransparencyCtl 
         Caption         =   "���� ����(&T)"
      End
      Begin VB.Menu mnuAddTool 
         Caption         =   "�߰� ���(&A)"
         Begin VB.Menu mnuEncrypt 
            Caption         =   "��ȣȭ(&E)"
         End
         Begin VB.Menu mnuDecrypt 
            Caption         =   "��ȣȭ(&D)"
         End
      End
      Begin VB.Menu fdghdfhdh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOption 
         Caption         =   "�ɼ�(&O)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "ã��(&S)..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Y's Notepad SE ����(&A)..."
      End
   End
   Begin VB.Menu mnu�̰Ǻ�� 
      Caption         =   " "
   End
End
Attribute VB_Name = "frmMain_Eng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--����ȭ�� ���� ���� ����--
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
'--����ȭ�� ���� ���� ��--

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any) '���� ȣ���� ���� �Լ� ����
Dim NomalQuit As Boolean
Sub UpdateFileName_Module()

End Sub
Private Sub CreateTransparentWindowStyle(lHwnd) '�� ����ȭ�� ���� �ʱ�ȭ �Լ�
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
                                      Optional TransVal As Long) '�� ����ȭ �Լ�
On Error GoTo Err_Handler:

    Call CreateTransparentWindowStyle(lHwnd) '�� ����ȭ �Ӽ� ����
    
    If TransparencyBy = byColor Then
         SetLayeredWindowAttributes lHwnd, Clr, 0, LWA_COLORKEY
         
    ElseIf TransparencyBy = byValue Then '������ ����
         If TransVal < 0 Or TransVal > 255 Then

            Err.Raise 2222, "Sub WindowTransparency", _
                    "������ 0�� 255������ ���ڿ��� �մϴ�." '���� �߻�
            Exit Sub
         End If
         SetLayeredWindowAttributes lHwnd, 0, TransVal, LWA_ALPHA '����ȭ ����(api ���)
    End If

Exit Sub
Err_Handler:
    If Err.Number = 2222 Then
    Err.Source = Err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox "�����ڵ�:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����"
    Mklog Err.Number & "/" & Err.Description
    WindowTransparency Me.hwnd, byValue, , 255
    Err.Clear
    Exit Sub
    Else
    Err.Source = Err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox "ó������ ���� ������ �߻��Ǿ����ϴ�!" & vbCrLf & "�����ڵ�:" & Err.Number & vbCrLf & Err.Description, vbCritical, "ġ������ ����"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Resume Next
    End If
    'WindowTransparency Me.hwnd, byValue
End Sub








Private Sub Form_Load()

On Error GoTo Err_frmMain_Eng

'Mklog "�׳� �ߴ��� ������� ���� ����\"
If Not Val(GetSetting(PROGRAM_KEY, "Program", "Trans", 255)) = 255 Then
    WindowTransparency Me.hwnd, byValue, , GetSetting(PROGRAM_KEY, "Program", _
        "Trans", 255) '����ȭ ����-�������� �ҷ���
End If
SaveSetting PROGRAM_KEY, "Program", "Date", LAST_UPDATED
'--�������� ���� �ҷ�����--
With txtText
    .FontBold = GetSetting(PROGRAM_KEY, "RTF", "FontBold", False)
    .FontItalic = GetSetting(PROGRAM_KEY, "RTF", "FontItalic", False)
    .FontName = GetSetting(PROGRAM_KEY, "RTF", "FontName", "����")
    .FontSize = GetSetting(PROGRAM_KEY, "RTF", "FontSize", 9)
    .FontStrikethru = GetSetting(PROGRAM_KEY, "RTF", "FontStrikethrugh", False)
    .FontUnderline = GetSetting(PROGRAM_KEY, "RTF", "FontUnderline", False)
    .ForeColor = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
    .BackColor = GetSetting(PROGRAM_KEY, "RTF", "Backcolor", RGB(255, 255, 255))
End With
With CD1
    .FontBold = GetSetting(PROGRAM_KEY, "RTF", "FontBold", False)
    .FontItalic = GetSetting(PROGRAM_KEY, "RTF", "FontItalic", False)
    .FontName = GetSetting(PROGRAM_KEY, "RTF", "FontName", "����")
    .FontSize = GetSetting(PROGRAM_KEY, "RTF", "FontSize", 9)
    .FontStrikethru = GetSetting(PROGRAM_KEY, "RTF", "FontStrikethrugh", False)
    .FontUnderline = GetSetting(PROGRAM_KEY, "RTF", "FontUnderline", False)
    .Color = GetSetting(PROGRAM_KEY, "RTF", "FontColor", &H0&)
End With
'--�������� ���� �ҷ����� ��--
'�α� ���� ��� Shell Echo�� �ٲ㼭 �ʿ����
'Me.logsave.Text = ""
'If Dir(AppPath & "\log.dat") = "" Then '�α� ������ �ִ��� Ȯ��
'    Me.logsave.SaveFile AppPath & "\log.dat", rtfText '���ٸ� ����� �ش�
'Else
'    Me.logsave.FileName = AppPath & "\log.dat" '�ִٸ� �ҷ��´�
'    Debug.Print AppPath
'End If
Mklog "���α׷� ���� - V." & App.Major & "." & App.Minor & "." & App.Revision & _
    " Last Updated:" & LAST_UPDATED '�α� ����
FileName_Dir = "���� ����" '�� ����
Newfile = True
UpdateFileName Me, FileName_Dir '���� ������Ʈ
Exit Sub
Err_frmMain_Eng:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, _
    "ó������ ���� ���� �߻�!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbAppWindows Then 'Windows�� ���� ��û�� �Ͽ���
    If Dirty Then '���� ������ �ִ�
        Dim ans As VbMsgBoxResult
        ans = MsgBox("������ ������� �ʾҽ��ϴ�!" & vbCrLf & "���� Windows�� �����Ͻðڽ��ϱ�?", vbOKCancel, "���� Ȯ��")
        If ans = vbCancel Then
            Cancel = True 'Windows ���� ����
        End If
    End If
End If
End Sub

Private Sub Form_Resize()
On Error GoTo ignoreerr '���� ����
Me.txtText.Left = 0
Me.txtText.Top = 0
Me.txtText.Width = Me.ScaleWidth
Me.txtText.Height = Me.ScaleHeight
Sleep 0
Exit Sub
ignoreerr:
Mklog Err.Number & "/" & Err.Description '�α׸� �����
Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim chk As Boolean
NomalQuit = True
Mklog "���α׷� ���� ó�� ����" '���� ���� �α�
If Dirty Then '������ �ٲ����!
    chk = SaveCheck(CD1) '�����Ұ��� Ȯ��
    If Not chk Then
        Cancel = True
        Mklog "���α׷� ���� ��ҵ�" '��� ������
    End If
End If
If Me.WindowState = 1 Then
SaveSetting PROGRAM_KEY, "Window", "X", Screen.Height / 2
SaveSetting PROGRAM_KEY, "Window", "Y", Screen.Width / 2
SaveSetting PROGRAM_KEY, "Window", "�ּ�ȭ", 1
SaveSetting PROGRAM_KEY, "Window", "Width", 8000
SaveSetting PROGRAM_KEY, "Window", "Height", 7000
ElseIf Me.WindowState = 2 Then
SaveSetting PROGRAM_KEY, "Window", "X", Screen.Height / 2
SaveSetting PROGRAM_KEY, "Window", "Y", Screen.Width / 2
SaveSetting PROGRAM_KEY, "Window", "�ִ�ȭ", 1
SaveSetting PROGRAM_KEY, "Window", "Width", 8000
SaveSetting PROGRAM_KEY, "Window", "Height", 7000
Else
SaveSetting PROGRAM_KEY, "Window", "X", Me.Top
SaveSetting PROGRAM_KEY, "Window", "Y", Me.Left
SaveSetting PROGRAM_KEY, "Window", "Width", Me.Width
SaveSetting PROGRAM_KEY, "Window", "Height", Me.Height
SaveSetting PROGRAM_KEY, "Window", "�ִ�ȭ", 0
SaveSetting PROGRAM_KEY, "Window", "�ּ�ȭ", 0
End If
Unload Form2
Mklog "���α׷� ���� ó�� ��." '���� �� �α�. ������ ���� ���� �α׿� �پ� �־�� ����.
'�α� ���� ��� �������� �ʿ����
'frmMain_Eng.logsave.SaveFile AppPath & "\log.dat", rtfText
End Sub

Private Sub mnu�̰Ǻ��_Click()
'����̶��� Me
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
MsgBox "�̱��� ���" '��...���� ���� �� ��������..
End Sub

Private Sub mnuBackground_Click()
CD1.ShowColor '���� ���� ��ȭ����
txtText.BackColor = CD1.Color '���� ����
SaveSetting PROGRAM_KEY, "RTF", "Backcolor", txtText.BackColor '������ ���� �ݿ�
End Sub

Private Sub mnuDecrypt_Click() '�ص�
Dim msgres As VbMsgBoxResult
msgres = MsgBox("�� ����� ���� ����� �׽�Ʈ���� �ʾ�����, ������ �ջ�� ���� �ֽ��ϴ�." & vbCrLf & "��ȣȭ�� �ϱ� ���� ������ ����� �νʽÿ�." & vbCrLf & "���� ����Ͻðڽ��ϱ�?", vbQuestion + vbOKCancel, "��Ÿ!")
If msgres = vbCancel Then Exit Sub
txtText.Text = DeCrypt(txtText.Text)
End Sub

Private Sub mnuEncrypt_Click() '��ȣȭ
Dim msgres As VbMsgBoxResult
msgres = MsgBox("�� ����� ���� ����� �׽�Ʈ���� �ʾ�����, ������ �ջ�� ���� �ֽ��ϴ�." & vbCrLf & "��ȣȭ�� �ϱ� ���� ������ ����� �νʽÿ�." & vbCrLf & "���� ����Ͻðڽ��ϱ�?", vbQuestion + vbOKCancel, "��Ÿ!")
If msgres = vbCancel Then Exit Sub
txtText.Text = EnCrypt(txtText.Text)
End Sub

Private Sub mnuFastPrint_Click() '���� �μ�-���� �̴´ٴ°� �ƴ϶� �⺻ �����ͷ� �׳� �̾ƹ���.
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
Mklog "����ڰ� �μ� ���"
End Sub
Sub SetPrinter() '������ ����
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
    Form2.Command1.Caption = "ã��"
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

Private Sub mnuFont_Click() '�۲� ����
Dim temp
Dim Dirty1 As Boolean
If Not Dirty Then Dirty1 = True '�۲��� �ٲپ����� ������ �ٲ�� �� �ƴϹǷ� �̸� ������ ��
On Error GoTo Err_Font
Mklog "��Ʈ ����" '�α�
CD1.Flags = cdlCFBoth Or cdlCFEffects '�÷��� ����(�̰� ���� ��Ʈ ���ٰ� �� ��-)
CD1.ShowFont '��Ʈ ��ȭ���� ȣ��
'With RTF
'.SelBold = CD1.FontBold
'.SelColor = CD1.Color
'.SelFontName = CD1.FontName
'.SelFontSize = CD1.FontSize
'.SelItalic = CD1.FontItalic
'.SelStrikeThru = CD1.FontStrikethru
'.SelUnderline = CD1.FontUnderline
'End With ->RTF ���� ����(0.3 �������� �߰�)->����..rtf�� �����е忡��
With txtText '��ȭ������ ���� �ݿ� & ����
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
'��ȭ������ ���� �ݿ� & ���� ��
'RTF.ForeColor = CD1.Color
If Dirty1 Then Dirty = False '���� ��Ʈ ���� ���� ���� ������ �����ٸ� Dirty ���� ���� �ʱ�ȭ
Exit Sub
Err_Font:
If Err.Number = 32755 Then '����ߴ�!
Err.Clear
Mklog "����ڰ� ��Ʈ ���� �����" '�α�
If Dirty1 Then Dirty = False '���� ��Ʈ ���� ���� ���� ������ �����ٸ� Dirty ���� ���� �ʱ�ȭ
Err.Clear '���� �ʱ�ȭ
Exit Sub
Else
MsgBox "ó������ ���� ������ �߻��Ǿ����ϴ�!" & vbCrLf & "�����ڵ�:" & Err.Number & vbCrLf & Err.Description, vbCritical, "ġ������ ����" '���� �߻� �˸�
Mklog Err.Number & "/" & Err.Description '�α�
If Dirty1 Then Dirty = False '���� ��Ʈ ���� ���� ���� ������ �����ٸ� Dirty ���� ���� �ʱ�ȭ
Err.Clear '���� �ʱ�ȭ
End If
End Sub

Private Sub mnuHelpAbout_Click() '���� ��ȭ����
'Call ShellAbout(Me.hwnd, Me.Caption, "Copyright (C) 2011 YJSoFT. All rights Reserved.", Me.Icon.Handle)'api�� ������
IsAboutbox = True '�ð��� ������ ������� ����!
frmSplash.Show '���÷��÷� ���� �� ��Ȱ�� ����
'frmSplash.Timer1.Enabled = False
End Sub

Private Sub mnuHelpContents_Click() '����<������ ��� FAIL
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
 ' OSWinHelp
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuHelpSearch_Click() '����<������ ��� FAIL
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub
Private Sub mnuEditCut_Click()
MsgBox "�̱��� ���"
End Sub
Private Sub mnuEditPaste_Click()
MsgBox "�̱��� ���"
End Sub
Private Sub mnuEditUndo_Click()
MsgBox "�̱��� ���"
End Sub





Private Sub mnuFileNew_Click()
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '���� Ȯ�ο��� ����Ͽ��ų� ���� �߻��� ��������
End If
CD1.FileName = "" '����/���� ��ȭ������ ���ϸ� �ʱ�ȭ
'RTF.Text = "" '�ؽ�Ʈ�ڽ� �ؽ�Ʈ �ʱ�ȭ
'RTF.FileName = "" '������ ���� �̸� �ʱ�ȭ
txtText.Text = ""
Dirty = False '"���� �ȵ�"���� ���� ����
FileName_Dir = "���� ����"
UpdateFileName Me, FileName_Dir '���� ���� - ���� ����
Newfile = True
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
'debug_temp = True
If Dirty Then
    If SaveCheck(CD1) = False Then Exit Sub '���� Ȯ�ο��� ����Ͽ��ų� ���� �߻��� ��������
End If
Mklog "frmMain_Eng.mnuFileOpen_Click()"
CD1.Filter = "�ؽ�Ʈ ����|*.txt|��� ����|*.*" '���� ���� ��ȭ���� �÷��� ����
CD1.CancelError = True '��ҽ� ����(32755)
CD1.ShowOpen '��ȭ���� ǥ��
If Err.Number = 32755 Then '��Ұ� ��������!
    CD1.FileName = "" '������ ���� �ʱ�ȭ
    Err.Clear
    Mklog "����ڰ� ���� ���"
    Exit Sub '���ν��� ���� ����(����ڰ� �����)
End If
If Err.Number = 13 Then '������ ���� �ʴ�!
    CD1.FileName = "" '������ ���� �ʱ�ȭ
    Err.Clear
    Mklog "���� �ڿ����̴�!\"
    Mklog "-����õF ���� ��\"
    Mklog "�� ������ ���� �ʴܴ�!!!\"
    Mklog "���״� ����!!!\"
    MsgBox "�˼��մϴ�. ���α׷����� �߸��� ����� �����Ͽ� �۾��� �ߴܵ˴ϴ�...", vbCritical, "ġ������ ����"
    Exit Sub '���ν��� ���� ����(����ڰ� �����)
End If
If Not Err.Number = 0 Then
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
Mklog "���� ����(" & CD1.FileName & ")" '�α� ����(�����)
'RTF.FileName = CD1.FileName '���� ���� ó��
FileName_Dir = CD1.FileName

Dim FreeFileNum As Integer
FreeFileNum = FreeFile
Open FileName_Dir For Input As #FreeFileNum
Screen.MousePointer = 11
txtText.Text = StrConv(InputB(LOF(FreeFileNum), FreeFileNum), vbUnicode)
If Not Err.Number = 0 Then
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
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
'  MsgBox "�ݱ� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub

Private Sub mnuFileExit_Click()
  '���� ��ε��մϴ�.
  Unload Me
End Sub

'Private Sub mnuFileNew_Click()
'  MsgBox "�� ���� �ڵ带 �ۼ��Ͻʽÿ�!"
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
Mklog "����ڰ� �μ� ���"
End Sub



Private Sub mnuFilePrintSetup_Click()
Ynotepadse.frmPreview.Show
With frmPreview.picPreview
.AutoRedraw = True
End With
End Sub



Private Sub mnuFileSave_Click()
On Error Resume Next
If Not Dirty Then Exit Sub '�ؽ�Ʈ�� ��ȭ�� ������ ��������
'RTF.Text = txtText.Text
Mklog "frmMain_Eng.mnuFileSave_Click()"
If Newfile Then
    SaveFile
Else
'CD1.FileName = RTF.FileName '�̹� ������ ������ �ִ�-������ ���� �̸��� cd1.filename�� ����
End If
Mklog "���� ����(" & CD1.FileName & ")" '�α� ����(�����)
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open CD1.FileName For Output As #FreeFileNum
    Print #FreeFileNum, txtText.Text
    Close #FreeFileNum
    Screen.MousePointer = 0
'Me.RTF.SaveFile CD1.FileName, rtfText '���� ���� ó��
If Not Err.Number = 0 Then
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
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
Mklog "���� ����(" & CD1.FileName & ")" '�α� ����(�����)
    Dim FreeFileNum As Integer
    FreeFileNum = FreeFile
    Screen.MousePointer = 11
    Open CD1.FileName For Output As #FreeFileNum
    Print #FreeFileNum, txtText.Text
    Close #FreeFileNum
    Screen.MousePointer = 0
'Me.RTF.SaveFile CD1.FileName, rtfText '���� ���� ó��
If Not Err.Number = 0 Then
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
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
CD1.Filter = "�ؽ�Ʈ ����|*.txt|��� ����|*.*" '���� ���� ��ȭ���� �÷��� ����
CD1.CancelError = True '��ҽ� ����(32755)
CD1.ShowSave '��ȭ���� ǥ��
If Not Right(CD1.FileName, 4) = ".txt" Then
    CD1.FileName = CD1.FileName & ".txt"
End If
If Err.Number = 32755 Then '��Ұ� ��������!
    CD1.FileName = "" '������ ���� �ʱ�ȭ
    Err.Clear
    Mklog "����ڰ� ���� ���"
    Exit Sub '���ν��� ���� ����(����ڰ� �����)
End If
If Err.Number = 13 Then '������ ���� �ʴ�!
    CD1.FileName = "" '������ ���� �ʱ�ȭ
    Err.Clear
    Mklog "���� �ڿ����̴�!\"
    Mklog "-����õF ���� ��\"
    Mklog "�� ������ ���� �ʴܴ�!!!\"
    Mklog "���״� ����!!!\"
    MsgBox "�˼��մϴ�. ���α׷����� �߸��� ����� �����Ͽ� �۾��� �ߴܵ˴ϴ�...", vbCritical, "ġ������ ����"
    Exit Sub '���ν��� ���� ����(����ڰ� �����)
End If
If Not Err.Number = 0 Then
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
    Exit Sub
End If
End Sub
Private Sub mnuLogClr_Click()
'Me.logsave.Text = ""
'�α� ����� �Լ��� ����
Mklog 1
End Sub

Private Sub mnuLogopn_Click()
'If Dir(AppPath & "\log.txt") = "" Then
'    MsgBox "�α� ������ �����ϴ�!", vbCritical, "����"
'Else
    'Me.RTF.FileName = AppPath & "\log.txt"
    'Me.RTF.SaveFile AppPath & "\log_user.txt", rtfText '�α� ���� �ջ����κ��� ��ȣ
    'Me.txtText.Text = RTF.Text
'End If
End Sub

Private Sub mnuReplace_Click()
    FindReplace = True
    Form2.Height = 1320
    Form2.Text2.Visible = True
    Form2.Command1.Caption = "�ٲٱ�"
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
Trans = InputBox("������ �Է��� �ּ���!(50~255)" & vbCrLf & "150 ����", "���� �Է�")
Debug.Print Trans
If Trans < 50 Then
    If Trans = 0 Then Exit Sub
NumError:
    MsgBox "���ڰ� �߸��Ǿ����ϴ�!", vbCritical, "����"
    GoTo ReInput
End If
If Trans > 255 Then
    GoTo NumError
End If
WindowTransparency Me.hwnd, byValue, , Trans
SaveSetting PROGRAM_KEY, "Program", "Trans", Trans
Exit Sub
Err_Trans:
If Err.Number = 13 Then '����ڰ� ���
    Err.Clear '���� ���
    Exit Sub '����ȭ ó�� ���
Else
    MsgBox "ó������ ���� ������ �߻��Ǿ����ϴ�!" & vbCrLf & "�����ڵ�:" & Err.Number & vbCrLf & Err.Description, vbCritical, "ġ������ ����"
    WindowTransparency Me.hwnd, byValue, , 255
    Mklog Err.Number & "/" & Err.Description
    Err.Clear
End If
End Sub

Private Sub mnuUserChg_Click()
Dim ChgUser As String
ChgUser = InputBox("�ٲ� ����� �̸��� �Է��� �ּ���!(�ִ� 20����)", "����� �̸� ����", Username, Screen.Width / 2, Screen.Height / 2)
If Len(ChgUser) > 20 Then
ChgUser = Left(ChgUser, 20)
End If
If ChgUser = "" Then
    ChgUser = "(�� �� ����)"
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
    If SaveCheck(CD1) = False Then Exit Sub '���� Ȯ�ο��� ����Ͽ��ų� ���� �߻��� ��������
End If
f = FreeFile()
s = Data.Files.Item(f) '���� �̸� ����
Mklog "�巡��&��� ����(" & s & ")"
Mklog "���� ����(" & s & ")" '�α� ����(�����)
'RTF.FileName = s '���� ���� ó��
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
    MsgBox "���� �߻�!" & vbCrLf & "���� ��ȣ:" & Err.Number & vbCrLf & Err.Description, vbCritical, "����!"
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
