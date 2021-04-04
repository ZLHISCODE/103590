Attribute VB_Name = "mdlSHA1"
  Option Explicit
    
  '   TITLE:
  '   Secure   Hash   Algorithm,   SHA-1
    
  '   AUTHORS:
  '   Adapted   by   Iain   Buchan   from   Visual   Basic   code   posted   at   Planet-Source-Code   by   Peter   Girard
  '   http://www.planetsourcecode.com/xq/ASP/txtCodeId.13565/lngWId.1/qx/vb/scripts/ShowCode.htm
    
  '   PURPOSE:
  '   Creating   a   secure   identifier   from   person-identifiable   data
    
  '   The   function   SecureHash   generates   a   160-bit   (20-hex-digit)   message   digest   for   a   given   message   (String).
  '   It   is   computationally   infeasable   to   recover   the   message   from   the   digest.
  '   The   digest   is   unique   to   the   message   within   the   realms   of   practical   probability.
  '   The   only   way   to   find   the   source   message   for   a   digest   is   by   hashing   all   possible   messages   and   comparison   of   their   digests.
    
  '   REFERENCES:
  '   For   a   fuller   description   see   FIPS   Publication   180-1:
  '   http://www.itl.nist.gov/fipspubs/fip180-1.htm
    
  '   SAMPLE:
  '   Message:   "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
  '   Returns   Digest:   "84983E441C3BD26EBAAE4AA1F95129E5E54670F1"
  '   Message:   "abc"
  '   Returns   Digest:   "A9993E364706816ABA3E25717850C26C9CD0D89D"
    
  Private Type Word
    B0   As Byte
    B1   As Byte
    B2   As Byte
    B3   As Byte
  End Type
'--------------------------------------------
Private Const BASE64CHR         As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)    As String
    
  Private Function AndW(w1 As Word, w2 As Word) As Word
    AndW.B0 = w1.B0 And w2.B0
    AndW.B1 = w1.B1 And w2.B1
    AndW.B2 = w1.B2 And w2.B2
    AndW.B3 = w1.B3 And w2.B3
  End Function
    
  Private Function OrW(w1 As Word, w2 As Word) As Word
    OrW.B0 = w1.B0 Or w2.B0
    OrW.B1 = w1.B1 Or w2.B1
    OrW.B2 = w1.B2 Or w2.B2
    OrW.B3 = w1.B3 Or w2.B3
  End Function
    
  Private Function XorW(w1 As Word, w2 As Word) As Word
    XorW.B0 = w1.B0 Xor w2.B0
    XorW.B1 = w1.B1 Xor w2.B1
    XorW.B2 = w1.B2 Xor w2.B2
    XorW.B3 = w1.B3 Xor w2.B3
  End Function
    
  Private Function NotW(w As Word) As Word
    NotW.B0 = Not w.B0
    NotW.B1 = Not w.B1
    NotW.B2 = Not w.B2
    NotW.B3 = Not w.B3
  End Function
    
  Private Function AddW(w1 As Word, w2 As Word) As Word
    Dim i     As Long, w       As Word
    
    i = CLng(w1.B3) + w2.B3
    w.B3 = i Mod 256
    i = CLng(w1.B2) + w2.B2 + (i \ 256)
    w.B2 = i Mod 256
    i = CLng(w1.B1) + w2.B1 + (i \ 256)
    w.B1 = i Mod 256
    i = CLng(w1.B0) + w2.B0 + (i \ 256)
    w.B0 = i Mod 256
        
    AddW = w
  End Function
  
  Private Function CircShiftLeftW(w As Word, n As Long) As Word
    Dim d1     As Double, d2       As Double
        
    d1 = WordToDouble(w)
    d2 = d1
    d1 = d1 * (2 ^ n)
    d2 = d2 / (2 ^ (32 - n))
    CircShiftLeftW = OrW(DoubleToWord(d1), DoubleToWord(d2))
  End Function
    
  Private Function WordToHex(w As Word) As String
    WordToHex = Right$("0" & Hex$(w.B0), 2) & Right$("0" & Hex$(w.B1), 2) & Right$("0" & Hex$(w.B2), 2) & Right$("0" & Hex$(w.B3), 2)
  End Function
    
  Private Function HexToWord(H As String) As Word
    HexToWord = DoubleToWord(Val("&H" & H & "#"))
  End Function
    
  Private Function DoubleToWord(n As Double) As Word
    DoubleToWord.B0 = Int(DMod(n, 2 ^ 32) / (2 ^ 24))
    DoubleToWord.B1 = Int(DMod(n, 2 ^ 24) / (2 ^ 16))
    DoubleToWord.B2 = Int(DMod(n, 2 ^ 16) / (2 ^ 8))
    DoubleToWord.B3 = Int(DMod(n, 2 ^ 8))
  End Function
    
  Private Function WordToDouble(w As Word) As Double
    WordToDouble = (w.B0 * (2 ^ 24)) + (w.B1 * (2 ^ 16)) + (w.B2 * (2 ^ 8)) + w.B3
  End Function
    
  Private Function DMod(value As Double, divisor As Double) As Double
    DMod = value - (Int(value / divisor) * divisor)
    If DMod < 0 Then DMod = DMod + divisor
  End Function
    
  Private Function F(t As Long, B As Word, C As Word, D As Word) As Word
  Select Case t
    Case Is <= 19
      F = OrW(AndW(B, C), AndW(NotW(B), D))
    Case Is <= 39
      F = XorW(XorW(B, C), D)
    Case Is <= 59
      F = OrW(OrW(AndW(B, C), AndW(B, D)), AndW(C, D))
    Case Else
      F = XorW(XorW(B, C), D)
    End Select
  End Function
    
  Public Function sha1(inMessage As String) As String
    
  Dim inLen     As Long, inLenW       As Word, padMessage       As String, numBlocks       As Long, w(0 To 79)           As Word, blockText       As String, wordText       As String, i       As Long, t       As Long, temp       As Word, K(0 To 3)           As Word, H0       As Word, H1       As Word, H2       As Word, H3       As Word, H4       As Word, A       As Word, B       As Word, C       As Word, D       As Word, E       As Word
  inLen = Len(inMessage)
  inLenW = DoubleToWord(CDbl(inLen) * 8)
        
  padMessage = inMessage & Chr$(128) & String$((128 - (inLen Mod 64) - 9) Mod 64, Chr$(0)) & String$(4, Chr$(0)) & Chr$(inLenW.B0) & Chr$(inLenW.B1) & Chr$(inLenW.B2) & Chr$(inLenW.B3)
        
  numBlocks = Len(padMessage) / 64
        
  '   initialize   constants
  K(0) = HexToWord("5A827999")
  K(1) = HexToWord("6ED9EBA1")
  K(2) = HexToWord("8F1BBCDC")
  K(3) = HexToWord("CA62C1D6")
    
  'initialize   160-bit   (5   words)   buffer
  H0 = HexToWord("67452301")
  H1 = HexToWord("EFCDAB89")
  H2 = HexToWord("98BADCFE")
  H3 = HexToWord("10325476")
  H4 = HexToWord("C3D2E1F0")
    
  'each   512   byte   message   block   consists   of   16   words   (W)   but   W   is   expanded   to   80   words
  For i = 0 To numBlocks - 1
    blockText = Mid$(padMessage, (i * 64) + 1, 64)
    'initialize   a   message   block
    For t = 0 To 15
      wordText = Mid$(blockText, (t * 4) + 1, 4)
      w(t).B0 = Asc(Mid$(wordText, 1, 1))
      w(t).B1 = Asc(Mid$(wordText, 2, 1))
      w(t).B2 = Asc(Mid$(wordText, 3, 1))
      w(t).B3 = Asc(Mid$(wordText, 4, 1))
    Next
            
    'create   extra   words   from   the   message   block
    For t = 16 To 79
      'W(t)   =   S^1   (W(t-3)   XOR   W(t-8)   XOR   W(t-14)   XOR   W(t-16))
      w(t) = CircShiftLeftW(XorW(XorW(XorW(w(t - 3), w(t - 8)), w(t - 14)), w(t - 16)), 1)
    Next
            
    'make   initial   assignments   to   the   buffer
    A = H0
    B = H1
    C = H2
    D = H3
    E = H4
            
    'process   the   block
    For t = 0 To 79
      temp = AddW(AddW(AddW(AddW(CircShiftLeftW(A, 5), F(t, B, C, D)), E), w(t)), K(t \ 20))
      E = D
      D = C
      C = CircShiftLeftW(B, 30)
      B = A
      A = temp
    Next
            
    H0 = AddW(H0, A)
    H1 = AddW(H1, B)
    H2 = AddW(H2, C)
    H3 = AddW(H3, D)
    H4 = AddW(H4, E)
  Next
     
    sha1 = WordToHex(H0) & WordToHex(H1) & WordToHex(H2) & WordToHex(H3) & WordToHex(H4)
End Function


'---------------------------------------------
'从一个经过Base64的字符串中解码到源字符串
Public Function DecodeBase64String(str2Decode As String) As String
    DecodeBase64String = StrConv(DecodeBase64Byte(str2Decode), vbUnicode)
End Function
 
'从一个经过Base64的字符串中解码到源字节数组
Private Function DecodeBase64Byte(str2Decode As String) As Byte()
 
    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim Bits(1 To 4)    As Byte
    Dim strDecode       As String
    Dim str             As String
    Dim Output()        As Byte
    
    Dim iIndex          As Long

    Dim lFrom As Long
    Dim lTo As Long
    
    InitBase
    
    '//除去回车
    str = Replace(str2Decode, vbCrLf, "")
 
    '//每4个字符一组（4个字符表示3个字）
    For lPtr = 1 To Len(str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '//查找字符在BASE64字符串中的位置
            iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue  'A~Za~z0~9+/
                Case 1 To 64:
                    Bits(iCtr + 1) = iValue - 1
                Case 65         '=
                    iLen = iCtr
                    Exit For
                    '//没有发现
                Case 0: Exit Function
            End Select
        Next
 
        '//转换4个6比特数成为3个8比特数
        Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
        Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
        Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)
 
        '//计算数组的起始位置
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1
                
        '//重新定义输出数组
        ReDim Preserve Output(0 To lTo)
        
        For iIndex = lFrom To lTo
            Output(iIndex) = Bits(iIndex - lFrom + 1)
        Next
 
        lTo = lTo + 1
        
    Next
    DecodeBase64Byte = Output
End Function

'将一个Base64字符串解码，并写入二进制文件
'Public Sub DecodeBase64StringToFile(strBase64 As String, strFilePath As String)
'    Dim fso As New Scripting.FileSystemObject, _
'        i As Long
'
'    If fso.FileExists(strFilePath) Then
'        fso.DeleteFile strFilePath, True
'    End If
'
'    i = FreeFile
'    Open strFilePath For Binary Access Write As i
'    Put i, , DecodeBase64Byte(strBase64)
'    Close i
'    Set fso = Nothing
'End Sub

'将一个Base64编码文件解码，并写入二进制文件
'Public Sub DecodeBase64FileToFile(strBase64FilePath As String, strFilePath As String)
'    Dim fso As New Scripting.FileSystemObject
'    Dim ts As TextStream
'
'    If Not fso.FileExists(strBase64FilePath) Then Exit Sub
'
'    Set ts = fso.OpenTextFile(strBase64FilePath)
'    DecodeBase64StringToFile ts.ReadAll, strFilePath
'End Sub
 
 
'将一个字节数组进行Base64编码，并返回字符串
Public Function EncodeBase64Byte(sValue() As Byte) As String
    Dim lCtr                As Long
    Dim lPtr                As Long
    Dim lLen                As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte
    
    Dim i As Integer
    
    InitBase
 
    For lCtr = 1 To UBound(sValue) + 1 Step 3
        For i = 1 To 3
            If lCtr + i - 2 <= UBound(sValue) Then
                Bits8(i) = sValue(lCtr + i - 2)
                lLen = 3
            Else
                Bits8(i) = 0
                lLen = lLen - 1
            End If
        Next
 
        '//转换字符串为数组，然后转换为4个6位(0-63)
        Bits6(1) = (Bits8(1) And &HFC) \ 4
        Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
        Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
        Bits6(4) = Bits8(3) And &H3F
 
        '//添加4个新字符
        For lPtr = 1 To lLen + 1
            sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
        Next
    Next
 
    '//不足4位，以=填充
    Select Case lLen + 1
        Case 2: sEncoded = sEncoded & "=="
        Case 3: sEncoded = sEncoded & "="
        Case 4:
    End Select
 
    EncodeBase64Byte = sEncoded
End Function
 

'对字符串进行Base64编码并返回字符串
Public Function EncodeBase64String(str2Encode As String) As String
    Dim sValue()            As Byte
    sValue = StrConv(str2Encode, vbFromUnicode)
    EncodeBase64String = EncodeBase64Byte(sValue)
End Function

'对文件进行Base64编码并返回编码后的Base64字符串
'Public Function EncodFileToBase64String(strFileSource As String)
'    Dim lpdata() As Byte, _
'        i As Long, _
'        n As Long, _
'        fso As New Scripting.FileSystemObject
'
'    If Not fso.FileExists(strFileSource) Then Exit Function
'
'    i = FreeFile
'
'    Open strFileSource For Binary Access Read Lock Write As i
'
'    n = LOF(i) - 1
'
'    ReDim lpdata(0 To n)
'    Get i, , lpdata
'    Close i
'
'    EncodFileToBase64String = EncodeBase64Byte(lpdata)
'End Function

'对文件进行Base64编码，并将编码后的内容直接写入一个文本文件中
'Public Sub EncodFileToBase64File(strFileSource As String, strFileBase64Desti As String)
'    Dim fso As New FileSystemObject, _
'        ts As TextStream
'
'    Set ts = fso.CreateTextFile(strFileBase64Desti, True)
'    ts.Write (EncodFileToBase64String(strFileSource))
'    ts.Close
'    Set ts = Nothing
'    Set fso = Nothing
'End Sub


Private Sub InitBase()
    Dim iPtr    As Integer
    '初始化 BASE64数组
    For iPtr = 0 To 63
        psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
    Next
End Sub



