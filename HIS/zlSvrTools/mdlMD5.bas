Attribute VB_Name = "mdlMD5"
Option Explicit
'**************************
'文件获取MD5值模块
'**************************
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Enum HashAlgorithm
    MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    SHA = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA
End Enum

Public Process As Long, CurrentProcess As Long

'----------获取指定字符串的MD5值---------------------------------------------------
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Property Get RegisterA() As String
        RegisterA = State(1)
End Property

Property Get RegisterB() As String
        RegisterB = State(2)
End Property

Property Get RegisterC() As String
        RegisterC = State(3)
End Property

Property Get RegisterD() As String
        RegisterD = State(4)
End Property
'----------获取指定字符串的MD5值---------------------------------------------------


'例子 HashFile("C:\APPSOFT\Apply\zlCISKernel.dll", 2 ^ 27)


'这里标记一下 标准的无符号LONG型 是4字节32位的 可存放2^32 次
'但VB的LONG型是有符号的  只有31位用于记数 还有1位用于标记正负符号 所以VB LONG 型正位只能到 2^31 = 2147483648
'出现负数的情况就是第32位也用来存放数据了 这样的情况需要特别处理  为了适应VB 的数据类型 下面的代码会比其他语言复杂


'SIZE是每次影射的文件大小 只能是2的N次方  如: 2^27=2的27次方=128M
Public Function HashFile(ByVal szFilePath As String, ByVal Size As Long, Optional ByVal Algorithm As Long = MD5, Optional ByVal Block_Size As Long = 32768) As String
    Dim hFile As Long, hMapFile As Long, lpBaseMap As Long
    Dim hCtx As Long, lRet As Long, hHash As Long, lLen As Long
    Dim i As Long, j As Long, Point As Long
    Dim FI As LARGE_INTEGER, Current As LARGE_INTEGER, CurrentPoint As Double
    Dim Temp As Long, lBlocks As Long, lLastBlock As Long, Block() As Byte
    
    '创建文件指针
    hFile = CreateFileA(szFilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        FI.lowpart = GetFileSize(hFile, FI.highpart) '成功后 获取文件大小
        If FI.highpart > 0 Then lBlocks = ((2 ^ 32 / Size) * FI.highpart) ' 高位   为1就是 2^32次字节  也就是4字节无符号长整型数值
        If FI.lowpart < 0 Then        '低位
            lBlocks = lBlocks + (2 ^ 31 / Size) '低位为负数 必然大于2^31次方  因为不大于2^31  VB可以正常显示
            Temp = LongToUnsigned(FI.lowpart) - 2 ^ 31 '转为无符号整型减掉2^31次 VB就能正常显示和运算了
            lLastBlock = Temp \ Size
            lBlocks = lBlocks + lLastBlock
            lLastBlock = Temp - lLastBlock * Size
        Else
            Temp = FI.lowpart \ Size
            lBlocks = lBlocks + Temp
            lLastBlock = FI.lowpart - Temp * Size
        End If
        
        
        hMapFile = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, FI.highpart, FI.lowpart, 0) '创建文件映射对象
        lRet = CryptAcquireContextA(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
        If err.LastDllError = &H80090016 Then lRet = CryptAcquireContextA(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
        lRet = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
        ReDim Block(Block_Size) As Byte
        
        For i = 1 To lBlocks '成功后根据指定大小 开始影射文件到内存空间
            lpBaseMap = MapViewOfFile(hMapFile, FILE_MAP_READ, Current.highpart, Current.lowpart, Size)
            If lpBaseMap Then
                Point = lpBaseMap
                For j = 1 To Size / Block_Size ' 2的N次方  必然除尽
                    
                    lRet = CryptHashData(hHash, Point, Block_Size, 0)
                    Point = Point + Block_Size
                Next
                UnmapViewOfFile (lpBaseMap)
            End If
            CurrentPoint = CurrentPoint + Size
            Current = Currency2LargeInteger(CurrentPoint / 10000@) '设置文件高低位
        Next
            
        If lLastBlock > 0 Then '映射余数
            lpBaseMap = MapViewOfFile(hMapFile, FILE_MAP_READ, Current.highpart, Current.lowpart, lLastBlock)
            If lpBaseMap Then
                Point = lpBaseMap
                Temp = lLastBlock \ Block_Size '不一定除尽 余数在FOR 循环完再次计算
                
                For j = 1 To Temp
                    lRet = CryptHashData(hHash, Point, Block_Size, 0)
                    Point = Point + Block_Size
                Next
                Temp = lLastBlock - Temp * Block_Size
                lRet = CryptHashData(hHash, Point, Temp, 0)
                UnmapViewOfFile (lpBaseMap)
            End If
        End If
        CloseHandle (hMapFile)

        If lRet Then
            lRet = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
            If lRet Then
                ReDim hash(lLen) As Byte
                lRet = CryptGetHashParam(hHash, HP_HASHVAL, hash(0), lLen, 0)
                If lRet Then
                    For j = 0 To UBound(hash) - 1
                        HashFile = HashFile & Right$("0" & Hex$(hash(j)), 2)
                    Next
                End If
                CryptDestroyHash hHash
            End If
        End If
        CryptReleaseContext hCtx, 0
        CloseHandle (hFile)
    End If
End Function

Public Function Currency2LargeInteger(ByVal curDistance As Currency) As LARGE_INTEGER
    CopyMemory Currency2LargeInteger, curDistance, 8
End Function



'----------------------------------------------------------------------
'获取指定字符串的MD5值
'----------------------------------------------------------------------
Public Function Md5_String_Calc(SourceString As String) As String
        If SourceString = "" Then Exit Function
        MD5Init
        MD5Update LenB(StrConv(SourceString, vbFromUnicode)), StringToArray(SourceString)
        MD5Final
        Md5_String_Calc = GetValues
End Function

Public Function Md5_File_Calc(InFile As String) As String
On Error GoTo errorhandler

GoSub begin

errorhandler:
        Exit Function
    
begin:
        Dim FileO As Integer
        FileO = FreeFile
        Call FileLen(InFile)
        Open InFile For Binary Access Read As #FileO
        MD5Init
        Do While Not EOF(FileO)
            Get #FileO, , ByteBuffer
            If Loc(FileO) < LOF(FileO) Then
                ByteCounter = ByteCounter + 64
                MD5Transform ByteBuffer
            End If
        Loop
        ByteCounter = ByteCounter + (LOF(FileO) Mod 64)
        Close #FileO
        MD5Final
        Md5_File_Calc = GetValues
End Function

Private Function StringToArray(InString As String) As Byte()
        Dim i As Integer, bytBuffer() As Byte
        ReDim bytBuffer(LenB(StrConv(InString, vbFromUnicode)))
        bytBuffer = StrConv(InString, vbFromUnicode)
        StringToArray = bytBuffer
End Function

Private Function GetValues() As String
        GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function

Private Function LongToString(Num As Long) As String
            Dim A As Byte, B As Byte, C As Byte, D As Byte
            A = Num And &HFF&
            If A < 16 Then LongToString = "0" & Hex(A) Else LongToString = Hex(A)
            B = (Num And &HFF00&) \ 256
            If B < 16 Then LongToString = LongToString & "0" & Hex(B) Else LongToString = LongToString & Hex(B)
            C = (Num And &HFF0000) \ 65536
            If C < 16 Then LongToString = LongToString & "0" & Hex(C) Else LongToString = LongToString & Hex(C)
            If Num < 0 Then D = ((Num And &H7F000000) \ 16777216) Or &H80& Else D = (Num And &HFF000000) \ 16777216
            If D < 16 Then LongToString = LongToString & "0" & Hex(D) Else LongToString = LongToString & Hex(D)
End Function

Private Sub MD5Init()
        ByteCounter = 0
        State(1) = UnsignedToLong(1732584193#)
        State(2) = UnsignedToLong(4023233417#)
        State(3) = UnsignedToLong(2562383102#)
        State(4) = UnsignedToLong(271733878#)
End Sub

Private Sub MD5Final()
        Dim dblBits As Double, padding(72) As Byte, lngBytesBuffered As Long
        padding(0) = &H80
        dblBits = ByteCounter * 8
        lngBytesBuffered = ByteCounter Mod 64
        If lngBytesBuffered <= 56 Then MD5Update 56 - lngBytesBuffered, padding Else MD5Update 120 - ByteCounter, padding
        padding(0) = UnsignedToLong(dblBits) And &HFF&
        padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
        padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
        padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
        padding(4) = 0
        padding(5) = 0
        padding(6) = 0
        padding(7) = 0
        MD5Update 8, padding
End Sub

Private Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
        Dim II As Integer, i As Integer, j As Integer, K As Integer, lngBufferedBytes As Long, lngBufferRemaining As Long, lngRem As Long

        lngBufferedBytes = ByteCounter Mod 64
        lngBufferRemaining = 64 - lngBufferedBytes
        ByteCounter = ByteCounter + InputLen

        If InputLen >= lngBufferRemaining Then
            For II = 0 To lngBufferRemaining - 1
                ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
            Next II
            MD5Transform ByteBuffer
            lngRem = (InputLen) Mod 64
            For i = lngBufferRemaining To InputLen - II - lngRem Step 64
                For j = 0 To 63
                    ByteBuffer(j) = InputBuffer(i + j)
                Next j
                MD5Transform ByteBuffer
            Next i
            lngBufferedBytes = 0
        Else
          i = 0
        End If
        For K = 0 To InputLen - i - 1
            ByteBuffer(lngBufferedBytes + K) = InputBuffer(i + K)
        Next K
End Sub

Private Sub MD5Transform(Buffer() As Byte)
        Dim X(16) As Long, A As Long, B As Long, C As Long, D As Long
    
        A = State(1)
        B = State(2)
        C = State(3)
        D = State(4)
        Decode 64, X, Buffer
        FF A, B, C, D, X(0), S11, -680876936
        FF D, A, B, C, X(1), S12, -389564586
        FF C, D, A, B, X(2), S13, 606105819
        FF B, C, D, A, X(3), S14, -1044525330
        FF A, B, C, D, X(4), S11, -176418897
        FF D, A, B, C, X(5), S12, 1200080426
        FF C, D, A, B, X(6), S13, -1473231341
        FF B, C, D, A, X(7), S14, -45705983
        FF A, B, C, D, X(8), S11, 1770035416
        FF D, A, B, C, X(9), S12, -1958414417
        FF C, D, A, B, X(10), S13, -42063
        FF B, C, D, A, X(11), S14, -1990404162
        FF A, B, C, D, X(12), S11, 1804603682
        FF D, A, B, C, X(13), S12, -40341101
        FF C, D, A, B, X(14), S13, -1502002290
        FF B, C, D, A, X(15), S14, 1236535329

        GG A, B, C, D, X(1), S21, -165796510
        GG D, A, B, C, X(6), S22, -1069501632
        GG C, D, A, B, X(11), S23, 643717713
        GG B, C, D, A, X(0), S24, -373897302
        GG A, B, C, D, X(5), S21, -701558691
        GG D, A, B, C, X(10), S22, 38016083
        GG C, D, A, B, X(15), S23, -660478335
        GG B, C, D, A, X(4), S24, -405537848
        GG A, B, C, D, X(9), S21, 568446438
        GG D, A, B, C, X(14), S22, -1019803690
        GG C, D, A, B, X(3), S23, -187363961
        GG B, C, D, A, X(8), S24, 1163531501
        GG A, B, C, D, X(13), S21, -1444681467
        GG D, A, B, C, X(2), S22, -51403784
        GG C, D, A, B, X(7), S23, 1735328473
        GG B, C, D, A, X(12), S24, -1926607734

        HH A, B, C, D, X(5), S31, -378558
        HH D, A, B, C, X(8), S32, -2022574463
        HH C, D, A, B, X(11), S33, 1839030562
        HH B, C, D, A, X(14), S34, -35309556
        HH A, B, C, D, X(1), S31, -1530992060
        HH D, A, B, C, X(4), S32, 1272893353
        HH C, D, A, B, X(7), S33, -155497632
        HH B, C, D, A, X(10), S34, -1094730640
        HH A, B, C, D, X(13), S31, 681279174
        HH D, A, B, C, X(0), S32, -358537222
        HH C, D, A, B, X(3), S33, -722521979
        HH B, C, D, A, X(6), S34, 76029189
        HH A, B, C, D, X(9), S31, -640364487
        HH D, A, B, C, X(12), S32, -421815835
        HH C, D, A, B, X(15), S33, 530742520
        HH B, C, D, A, X(2), S34, -995338651

        II A, B, C, D, X(0), S41, -198630844
        II D, A, B, C, X(7), S42, 1126891415
        II C, D, A, B, X(14), S43, -1416354905
        II B, C, D, A, X(5), S44, -57434055
        II A, B, C, D, X(12), S41, 1700485571
        II D, A, B, C, X(3), S42, -1894986606
        II C, D, A, B, X(10), S43, -1051523
        II B, C, D, A, X(1), S44, -2054922799
        II A, B, C, D, X(8), S41, 1873313359
        II D, A, B, C, X(15), S42, -30611744
        II C, D, A, B, X(6), S43, -1560198380
        II B, C, D, A, X(13), S44, 1309151649
        II A, B, C, D, X(4), S41, -145523070
        II D, A, B, C, X(11), S42, -1120210379
        II C, D, A, B, X(2), S43, 718787259
        II B, C, D, A, X(9), S44, -343485551

        State(1) = LongOverflowAdd(State(1), A)
        State(2) = LongOverflowAdd(State(2), B)
        State(3) = LongOverflowAdd(State(3), C)
        State(4) = LongOverflowAdd(State(4), D)
End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
        Dim intDblIndex As Integer, intByteIndex As Integer, dblSum As Double
        For intByteIndex = 0 To Length - 1 Step 4
            dblSum = InputBuffer(intByteIndex) + InputBuffer(intByteIndex + 1) * 256# + InputBuffer(intByteIndex + 2) * 65536# + InputBuffer(intByteIndex + 3) * 16777216#
            OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
            intDblIndex = intDblIndex + 1
        Next intByteIndex
End Sub

Private Function FF(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
        A = LongOverflowAdd4(A, (B And C) Or (Not (B) And D), X, ac)
        A = LongLeftRotate(A, S)
        A = LongOverflowAdd(A, B)
End Function

Private Function GG(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
        A = LongOverflowAdd4(A, (B And D) Or (C And Not (D)), X, ac)
        A = LongLeftRotate(A, S)
        A = LongOverflowAdd(A, B)
End Function

Private Function HH(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
        A = LongOverflowAdd4(A, B Xor C Xor D, X, ac)
        A = LongLeftRotate(A, S)
        A = LongOverflowAdd(A, B)
End Function

Private Function II(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
        A = LongOverflowAdd4(A, C Xor (B Or Not (D)), X, ac)
        A = LongLeftRotate(A, S)
        A = LongOverflowAdd(A, B)
End Function

Private Function LongLeftRotate(value As Long, Bits As Long) As Long
        Dim lngSign As Long, lngI As Long
        Bits = Bits Mod 32
        If Bits = 0 Then LongLeftRotate = value: Exit Function
        For lngI = 1 To Bits
            lngSign = value And &HC0000000
            value = (value And &H3FFFFFFF) * 2
            value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
        Next
        LongLeftRotate = value
End Function

Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
        Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
        lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
        lngOverflow = lngLowWord \ 65536
        lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
        LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
        Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
        lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
        lngOverflow = lngLowWord \ 65536
        lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
        LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function UnsignedToLong(value As Double) As Long
        If value < 0 Or value >= OFFSET_4 Then Error 6
        If value <= MAXINT_4 Then UnsignedToLong = value Else UnsignedToLong = value - OFFSET_4
End Function

Private Function LongToUnsigned(value As Long) As Double
        If value < 0 Then LongToUnsigned = value + OFFSET_4 Else LongToUnsigned = value
End Function


