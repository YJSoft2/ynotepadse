Attribute VB_Name = "modETC"
'---------------------------------------------------------------------------------------
' Module    : modETC
' DateTime  : 2012-08-05 20:06
' Author    : PC1
' Purpose   :
'---------------------------------------------------------------------------------------
'������Ʈ ����
'�� ������ ��Ÿ ��������, ����� �α� ������ �����˴ϴ�.
'�̰��� ������ ���� ��� modMain�� ��� �κп��� DEBUG_VERSION �κ��� ���� False�� �ٲ� �ֽø� �˴ϴ�.
'
'�������
'�α� ���� �������� RichTextBox ��Ʈ���� �̿��ϴ� ��Ŀ��� ���� ��� �۾��ϴ� ������� �ٲپ����ϴ�.
'���� �ؽ�Ʈ ���� ���� ��� ���� ���� ���Դϴ�.(���� ���� �������)
'!���!
'�� ���α׷��� �ҽ� �ڵ�� ��� ���� ������ �� �����ϴ�!
'
'Copyright YJSoft(yyj9411@naver.com). All rights Reserved.
'
'�Ʒ� ����� ���� ��� ���� ���۱� ������ ǥ��Ǿ� �ֽ��ϴ�.
'

'��ũ ���� ���ϴ� �Լ� By HappyBono(http://www.happybono.net/285)
'CC BY-NC-ND
'������ǥ��-�񿵸�-�������
'http://creativecommons.org/licenses/by-nc-nd/2.0/kr/ ����

Public Function GetDiskFreeSpace(strDeviceID As String) As String
Dim oWMI As Object
Dim oLDK As Object

Set oWMI = GetObject("winmgmts:")

Set oLDK = oWMI.Get("Win32_LogicalDisk.DeviceID=" _
& Chr(39) & strDeviceID & Chr(58) & Chr(39))

GetDiskFreeSpace = objLogicalDisk.FreeSpace & " byte"

Set oWMI = Nothing
Set oLDK = Nothing

End Function
