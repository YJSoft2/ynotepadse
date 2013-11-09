Attribute VB_Name = "Encryption1_1v"
Function HashMaxCode(ByVal sstr As String, ByVal HXCode As String, ByVal ONCode As String, ByRef StrBuffer As String) As Long()
Dim buf() As Long, i As Long, XCode As Long, NCode As Long, maxs() As Long
ReDim buf(0 To Len(sstr) - 1&)
For i = 0 To Len(sstr) - 1&
buf(i) = Asc(Mid(sstr, i + 1, 1))
Next
ReDim maxs(0 To UBound(buf)) As Long
XCode = CLng("&H" & HXCode)
NCode = CLng("&H" & ONCode) + Len(CStr(XCode))
For i = 0 To UBound(buf)
maxs(i) = buf(i) Xor XCode
maxs(i) = maxs(i) Xor NCode
maxs(i) = ((maxs(i) + XCode) Xor NCode)
Next
HashMaxCode = maxs
For i = 0 To UBound(maxs)
StrBuffer = StrBuffer & Chr(maxs(i))
Next
End Function

Function UNMaxCode(ByVal HXCode As String, ByVal ONCode As String, ByRef LongBuffer() As Long, ByVal StrEnc As String, ByVal STRINGMODE As Boolean) As String
Dim buf() As Long, i As Long, XCode As Long, NCode As Long, maxs() As Long
XCode = CLng("&H" & HXCode)
NCode = CLng("&H" & ONCode) + Len(CStr(XCode))
If STRINGMODE Then
ReDim buf(0 To Len(StrEnc) - 1&) As Long
For i = 0 To Len(StrEnc) - 1&
LongBuffer(i) = Asc(Mid(StrEnc, i + 1, 1))
Next
End If
ReDim buf(0 To UBound(LongBuffer)) As Long
For i = 0 To UBound(LongBuffer)
buf(i) = (LongBuffer(i) Xor NCode) - XCode
buf(i) = buf(i) Xor NCode
buf(i) = buf(i) Xor XCode
Next
For i = 0 To UBound(LongBuffer)
UNMaxCode = UNMaxCode & Chr(buf(i))
Next
End Function

