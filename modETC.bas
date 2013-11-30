Attribute VB_Name = "modETC"
'---------------------------------------------------------------------------------------
' Module    : modETC
' DateTime  : 2012-08-05 20:06
' Author    : PC1
' Purpose   :
'Hello3
'---------------------------------------------------------------------------------------
'������Ʈ ����
'�� ������ ��Ÿ ��������, ������ �α� ������ �����˴ϴ�.
'�̰��� ������ ���� ���� modMain�� ���� �κп��� DEBUG_VERSION �κ��� ���� False�� �ٲ� �ֽø� �˴ϴ�.
'
'��������
'�α� ���� ���������� RichTextBox ��Ʈ���� �̿��ϴ� ���Ŀ��� ���� ��� �۾��ϴ� �������� �ٲپ����ϴ�.
'���� �ؽ�Ʈ ���� ���� ���� ���� ���� ���Դϴ�.(���� ���� ��������)
'!����!
'�� ���α׷��� �ҽ� �ڵ��� ���� ���� ������ �� �����ϴ�!
'
'Copyright YJSoft(yyj9411@naver.com). All rights Reserved.
'
'�Ʒ� ������ ���� ���� ���� ���۱� ������ ǥ���Ǿ� �ֽ��ϴ�.
'

'����ũ ���� ���ϴ� �Լ� By HappyBono(http://www.happybono.net/285)
'CC BY-NC-ND
'������ǥ��-�񿵸�-��������
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

'---------------------------------------------------------------------------------------
' Procedure : UTF8_Encode
' DateTime  : 2013-04-03 13:35
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function UTF8_Encode(ByRef sStr() As Byte) As String
    
    Dim ii As Long, sUTF8 As String, iChar As Long, iChar2 As Long
    
   On Error GoTo UTF8_Encode_Error

    For ii = 0 To UBound(sStr)
        iChar = sStr(ii)
        
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
                iChar2 = sStr(ii + 1)
                sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                ii = ii + 1
            Else
                Dim iChar3 As Integer
                iChar2 = sStr(ii + 1)
                iChar3 = sStr(ii + 2)
                sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                ii = ii + 2
            End If
        Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next ii
    
    UTF8_Encode = sUTF8

   On Error GoTo 0
   Exit Function

UTF8_Encode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UTF8_Encode of Module modETC"
    
End Function
'[��ó] VB���� UTF-8�� ���ڿ� ���ڵ�(�迭, ���ڿ�)|�ۼ��� ����

