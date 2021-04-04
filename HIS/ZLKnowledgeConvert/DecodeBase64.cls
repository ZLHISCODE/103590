VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecodeBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BASE64CHR         As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)    As String

'从一个经过Base64的字符串中解码到源字符串
Public Function DecodeBase64String(str2Decode As String) As String
    DecodeBase64String = StrConv(DecodeBase64Byte(str2Decode), vbUnicode)
End Function
 
'从一个经过Base64的字符串中解码到源字节数组
Public Function DecodeBase64Byte(str2Decode As String) As Byte()
 
    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
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
                    bits(iCtr + 1) = iValue - 1
                Case 65         '=
                    iLen = iCtr
                    Exit For
                    '//没有发现
                Case 0: Exit Function
            End Select
        Next
 
        '//转换4个6比特数成为3个8比特数
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
 
        '//计算数组的起始位置
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1
                
        '//重新定义输出数组
        ReDim Preserve Output(0 To lTo)
        
        For iIndex = lFrom To lTo
            Output(iIndex) = bits(iIndex - lFrom + 1)
        Next
 
        lTo = lTo + 1
        
    Next
    DecodeBase64Byte = Output
End Function


 
 
'将一个字节数组进行Base64编码，并返回字符串
Public Function EncodeBase64Byte(sValue() As Byte) As String
    Dim lCtr                As Long
    Dim lPtr                As Long
    Dim lLen                As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte
    
    Dim I As Integer
    
    InitBase
 
    For lCtr = 1 To UBound(sValue) + 1 Step 3
        For I = 1 To 3
            If lCtr + I - 2 <= UBound(sValue) Then
                Bits8(I) = sValue(lCtr + I - 2)
                lLen = 3
            Else
                Bits8(I) = 0
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




Private Sub InitBase()
    Dim iPtr    As Integer
    '初始化 BASE64数组
    For iPtr = 0 To 63
        psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
    Next
End Sub



