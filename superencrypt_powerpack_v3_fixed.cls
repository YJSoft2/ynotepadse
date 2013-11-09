VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' SuperEncrypt PowerPack Class v3 (버그패치/보안강화)
' By. 수학쟁이

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Function EncryptString(ByVal sString As String, ByRef Key() As Byte) As String
    On Error Resume Next
    Dim bData() As Byte
    bData() = StrConv(sString, vbFromUnicode)
    EncryptString = DataToHex(EncryptData(bData(), Key()))
End Function

Public Function EncryptData(ByRef Data() As Byte, ByRef Key() As Byte) As Byte()
    On Error GoTo UErr
    Dim CData() As Byte, i As Long
    ReDim CData(UBound(Data))
    Randomize
    CopyMemory CData(0), Data(0), UBound(Data) + 1
    
    ' +) Reverse!
    If UBound(CData) Then  ' 1 바이트면 뒤집을 필요 없다.
        If UBound(CData) Mod 2 Then
            For i = 0 To UBound(CData) \ 2
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
                CData(UBound(CData) - i) = CData(i) Xor CData(UBound(CData) - i)
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
            Next
        Else
            For i = 0 To UBound(CData) \ 2 - 1
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
                CData(UBound(CData) - i) = CData(i) Xor CData(UBound(CData) - i)
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
            Next
        End If
    End If
    
    ' 1) Lookup Table Crazy Encryption!
    
    Dim LTable(), X As Long, T As Long
    LTable = Array(172, 16, 169, 243, 120, 173, 239, 216, 168, 136, 49, 215, 159, 208, 206, 110, 246, 95, 145, 74, 204, 41, 118, 38, 217, 75, 52, 175, _
                   21, 245, 157, 36, 255, 197, 199, 44, 210, 139, 209, 69, 201, 241, 103, 114, 42, 218, 22, 151, 144, 229, 73, 27, 46, 122, 143, 30, 234, _
                   160, 181, 62, 113, 67, 117, 212, 4, 90, 13, 163, 80, 198, 205, 134, 105, 83, 135, 89, 254, 121, 170, 187, 119, 184, 65, 194, 92, 236, 156, 61, 11, 127, 248, 195, 140, 53, 29, 40, 222, 97, 133, 153, 124, 158, 64, 238, 0, 155, 96, 193, 34, 161, 81, 25, 252, 213, 45, 179, 154, 174, 186, 220, 126, 132, 51, 178, 240, 59, 77, 86, 237, 10, 131, 71, 251, 146, 231, 253, 76, 100, 99, 225, 196, 57, 7, 47, 219, 233, 93, 224, 66, 116, 183, 130, 19, 207, 48, 1, 43, 98, 111, 108, 78, 63, 228, 115, 152, 87, 109, 190, 129, 102, 84, 211, 230, 150, 214, 58, 82, 6, 101, 91, 14, 18, 12, 141, 177, 23, 70, 125, 149, 188, 35, 176, 226, 8, 20, 142, 227, 5, 33, 202, 164, 166, 137, 106, 167, 28, 104, 200, 79, 2, 185, 17, 15, 165, 221, 242, 94, 162, 24, 223, 203, 107, 72, 26, 56, 9, 68, 148, 244, 37, 189, 147, 60, 182, 32, 85, 171, 39, 235, 249, 192, 247, 55, 3, 54, 128, 123, 250, 50, 232, 180, 191, 112, 138, 31, 88)
    For i = 0 To UBound(CData)
        CData(i) = LTable(CData(i))
    Next
    
    
    ' 2) Crazy CRAZY Suck XOR!
    Dim bDefaultKey(15) As Byte
    bDefaultKey(0) = 10
    bDefaultKey(1) = 1
    bDefaultKey(2) = 52
    bDefaultKey(3) = 255
    bDefaultKey(4) = 10
    bDefaultKey(5) = 3
    bDefaultKey(6) = 66
    bDefaultKey(7) = 91
    bDefaultKey(8) = 75
    bDefaultKey(9) = 13
    bDefaultKey(10) = 56
    bDefaultKey(11) = 33
    bDefaultKey(12) = 20
    bDefaultKey(13) = 36
    bDefaultKey(14) = 59
    bDefaultKey(15) = 35
    
    For i = 0 To UBound(CData)
        CData(i) = LTable(((CData(i) Xor bDefaultKey(i Mod 16)) Xor Key(i Mod (UBound(Key) + 1))))
    Next
    
    ' 3) CRAZY XOR
    For i = 0 To UBound(CData)
        CData(i) = CData(i) Xor (255 - (i Mod 256))
    Next
    
    ' 4) ADD AND SUBTRACT
    Dim ASLOOKUP()
    ASLOOKUP = Array(3, 5, 9, 20, 3, -1, -8, -99, 99, 30, -90, -29, 30, 200, 19, -200)
    For i = 0 To UBound(CData)
        X = i Mod 16
        CData(i) = ByteAdd(CData(i), ASLOOKUP(X))
        CData(i) = ByteAdd(CData(i), Key(i Mod (UBound(Key) + 1))) ' NEW!
    Next
    
    ' 5) CRAZY XOR 2 & Lookuptable
    For i = 0 To UBound(CData)
        CData(i) = LTable(CData(i) Xor (i Mod 256))
    Next
    
    ' 6) Additional XOR
    Dim PData() As Byte
    ReDim PData(UBound(CData) + 1)
    PData(0) = (Int(Rnd * 256) \ 7) * 7
    CopyMemory PData(1), CData(0), UBound(CData) + 1
    For i = 1 To UBound(PData)
        ' NEW!) Rotation
        PData(i) = LTable(RotateLeft(PData(i) Xor PData(0), i Mod 8))
    Next
    PData(0) = RotateRight(PData(0), 3)
    Erase CData
    
    EncryptData = PData()
UErr:
End Function

Public Function DecryptData(ByRef Data() As Byte, ByRef Key() As Byte) As Byte()
    On Error GoTo UErr
    Dim PData() As Byte, CData() As Byte, i As Long, pKey As Byte
    Dim DTable
    DTable = Array(104, 155, 209, 243, 64, 197, 177, 142, 193, 225, 129, 88, 182, 66, 180, 212, 1, 211, 181, 152, 194, 28, 46, 185, 218, 111, _
                   223, 51, 205, 94, 55, 254, 234, 198, 108, 190, 31, 229, 23, 237, 95, 21, 44, 156, 35, 114, 52, 143, 154, 10, 248, 122, 26, 93, _
                   244, 242, 224, 141, 175, 125, 232, 87, 59, 161, 102, 82, 148, 61, 226, 39, 186, 131, 222, 50, 19, 25, 136, 126, 160, 208, 68, 110, 176, 73, 170, 235, 127, 165, 255, 75, 65, 179, 84, 146, 216, 17, 106, 97, 157, 138, 137, 178, 169, 42, 206, 72, 203, 221, 159, 166, 15, 158, 252, 60, 43, 163, 149, 62, 22, 80, 4, 77, 53, 246, 100, 187, 120, 89, 245, 168, 151, 130, 121, 98, 71, 74, 9, 202, 253, 37, 92, 183, 195, 54, 48, 18, 133, 231, 227, 188, 173, 47, 164, 99, 116, 105, 86, 30, 101, 12, 57, 109, 217, 67, 200, 213, 201, 204, 8, 2, 78, 236, 0, 5, 117, 27, 191, 184, 123, 115, 250, 58, 233, 150, 81, 210, 118, 79, 189, 230, 167, 251, 240, 107, 83, 91, 140, 33, 69, 34, 207, 40, 199, 220, 20, 70, 14, 153, 13, 38, 36, 171, 63, 113, 174, 11, 7, 24, 45, 144, 119, 214, 96, 219, 147, 139, 192, 196, 162, 49, 172, 134, 249, 145, 56, 238, 85, 128, 103, 6, 124, 41, 215, 3, 228, 29, 16, 241, 90, 239, 247, 132, 112, 135, 76, 32)
    
    ReDim PData(UBound(Data))
    CopyMemory PData(0), Data(0), UBound(Data) + 1
    
    PData(0) = RotateLeft(PData(0), 3)
    If PData(0) Mod 7 Then GoTo UErr ' Invalid Data
    pKey = PData(0)
    For i = 1 To UBound(PData)
        ' NEW!) Rotation
        PData(i) = RotateRight(DTable(PData(i)), i Mod 8) Xor pKey
    Next
    
    ReDim CData(UBound(PData) - 1)
    CopyMemory CData(0), PData(1), UBound(PData)
    
    ' -1 step
    For i = 0 To UBound(CData)
        CData(i) = DTable(CData(i)) Xor (i Mod 256)
    Next
    
    ' -2 step
    Dim ASLOOKUP()
    ASLOOKUP = Array(3, 5, 9, 20, 3, -1, -8, -99, 99, 30, -90, -29, 30, 200, 19, -200)
    For i = 0 To UBound(CData)
        X = i Mod 16
        CData(i) = ByteSubtract(CData(i), ASLOOKUP(X))
        CData(i) = ByteSubtract(CData(i), Key(i Mod (UBound(Key) + 1)))
    Next
    
    ' -3 step
    For i = 0 To UBound(CData)
        CData(i) = CData(i) Xor (255 - (i Mod 256))
    Next
    
    ' -4 step
    Dim bDefaultKey(15) As Byte
    bDefaultKey(0) = 10
    bDefaultKey(1) = 1
    bDefaultKey(2) = 52
    bDefaultKey(3) = 255
    bDefaultKey(4) = 10
    bDefaultKey(5) = 3
    bDefaultKey(6) = 66
    bDefaultKey(7) = 91
    bDefaultKey(8) = 75
    bDefaultKey(9) = 13
    bDefaultKey(10) = 56
    bDefaultKey(11) = 33
    bDefaultKey(12) = 20
    bDefaultKey(13) = 36
    bDefaultKey(14) = 59
    bDefaultKey(15) = 35
    
    For i = 0 To UBound(CData)
        CData(i) = (((DTable(CData(i)) Xor bDefaultKey(i Mod 16)) Xor Key(i Mod (UBound(Key) + 1))))
    Next
    
    ' -1 step
    For i = 0 To UBound(CData)
        CData(i) = DTable(CData(i))
    Next

    ' +) Reverse
    If UBound(CData) Then ' 1 바이트면 뒤집을 필요 없다.
        If UBound(CData) Mod 2 Then
            For i = 0 To UBound(CData) \ 2
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
                CData(UBound(CData) - i) = CData(i) Xor CData(UBound(CData) - i)
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
            Next
        Else
            For i = 0 To UBound(CData) \ 2 - 1
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
                CData(UBound(CData) - i) = CData(i) Xor CData(UBound(CData) - i)
                CData(i) = CData(i) Xor CData(UBound(CData) - i)
            Next
        End If
    End If
    
    DecryptData = CData()
UErr:
End Function

Public Function DecryptString(ByVal sString As String, ByRef Key() As Byte) As String
    On Error Resume Next
    DecryptString = StrConv(DecryptData(HexToData(sString), Key()), vbUnicode)
End Function

Public Function KeyFromString(ByVal KeyString As String) As Byte()
    KeyFromString = StrConv(KeyString, vbFromUnicode)
End Function

Public Function KeyFromHexString(ByVal strHexadecimal As String) As Byte()
    KeyFromHexString = HexToData(strHexadecimal)
End Function

' Utility Functions

Private Function DataToHex(ByRef Data() As Byte) As String
    On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(Data)
        DataToHex = DataToHex & Right$("00" & Hex$(Data(i)), 2)
    Next
End Function

Private Function HexToData(ByVal sHex As String) As Byte()
    On Error Resume Next
    Dim bBuffer() As Byte, i As Long
    ReDim bBuffer(0 To Len(sHex) \ 2 - 1)
    For i = 0 To UBound(bBuffer)
        bBuffer(i) = CByte("&H" & Mid$(sHex, i * 2 + 1, 2))
    Next
    HexToData = bBuffer()
End Function

Private Function ByteAdd(ByVal bByte As Byte, ByVal lToAdd As Long) As Byte
    ByteAdd = CByte((CLng(bByte) + lToAdd) And &HFF)
End Function

Private Function ByteSubtract(ByVal bByte As Byte, ByVal lToSubtract As Long) As Byte
    ByteSubtract = CByte((CLng(bByte) - lToSubtract) And &HFF)
End Function

Private Function RotateLeft(ByVal Value As Byte, ByVal Bit As Byte) As Byte
    Dim ToMoved As Byte, Remained As Byte
    Remained = 2 ^ (8 - Bit) - 1
    If Bit > 7 Then RotateLeft = RotateLeft(Value, Bit Mod 8): Exit Function
    ToMoved = (Value And Not Remained) \ (2 ^ (8 - Bit))
    RotateLeft = ((Value And Remained) * (2 ^ Bit)) Or ToMoved
End Function

Private Function RotateRight(ByVal Value As Byte, ByVal Bit As Byte) As Byte
    Dim ToMoved As Byte
    If Bit > 7 Then RotateRight = RotateRight(Value, Bit Mod 8): Exit Function
    ToMoved = Value And (2 ^ Bit - 1)
    RotateRight = (Value \ (2 ^ Bit)) Or (ToMoved * (2 ^ (8 - Bit)))
End Function

