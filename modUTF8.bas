Attribute VB_Name = "modUTF8"
'---------------------------------------------------------------------------------------
' Module    : modUTF8
' DateTime  : 2013-04-03 13:36
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
Public Const CP_UTF8 = 65001
 
Public Declare Function MultiByteToWideChar Lib "kernel32" _
(ByVal CodePage As Long, ByVal dwFlags As Long, _
ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
 
Public Function UTFOpen(FileNameUTF As String) As String
On Error GoTo ErrClear
    Dim utf8() As Byte
    Dim ucs2 As String
    Dim chars As Long
    
    UTF8_Error = False '오류 변수 초기화
    
    Open FileNameUTF For Binary As #1   'UTF-8 문서지정
    ReDim utf8(LOF(1))
    
    Get #1, , utf8
    
    If Hex(utf8(0)) & Hex(utf8(1)) & Hex(utf8(2)) = "EFBBBF" Then 'UTF-8 BOM 문서
        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(3)), LOF(1), 0, 0)
        ucs2 = Space(chars)
    
        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(3)), LOF(1), StrPtr(ucs2), chars)
    
        UTFOpen = ucs2
    Else 'UTF-8 BOM 없는 문서
        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
        ucs2 = Space(chars)
    
    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
    
    UTFOpen = ucs2
    End If
    Close
    Exit Function
ErrClear:
    Err.Clear
    UTF8_Error = True
End Function

