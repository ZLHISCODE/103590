VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPipeClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type


Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Const NMPWAIT_WAIT_FOREVER = &HFFFF

Private Const CP_ACP = 0  '  default to ANSI code page


Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_WRITE = &H2

Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function ConnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function WaitNamedPipe Lib "kernel32" Alias "WaitNamedPipeA" (ByVal lpNamedPipeName As String, ByVal nTimeOut As Long) As Long
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'vb通道名称，与服务中的通道名称对应
Private Const DEFAULT_PIPE_NAME As String = "\\.\PIPE\ZLPACSSVRCENTER\VB\0FD1A47CBDF143298EBC3F03877CB94F"


Private mlngPipeHandle As Long

Public PipeName As String

Public Function SendData(data As clsPipeData) As Boolean
    
    Dim strData As String
    Dim stuSecurity As SECURITY_ATTRIBUTES
    Dim stuOVERLAPPED As OVERLAPPED
    Dim lngResult As Long
    Dim bytesWrite As Long
    Dim curByte() As Byte
    'Dim curStr As String
    Dim i As Long
    
    SendData = False
    
    strData = data.SaveToXml()
    
    '默认15秒超时时间
    lngResult = WaitNamedPipe(PipeName, 15 * 1000) 'NMPWAIT_WAIT_FOREVER :表示无限等待
    If (lngResult = False) Then
        
        Exit Function
    End If
    
    mlngPipeHandle = CreateFile(PipeName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_WRITE, stuSecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If mlngPipeHandle = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    
    '使用utf8进行传输
    curByte() = Unicode8Encode(strData) 'UnicodeEncode(strData, False)
    bytesWrite = UBound(curByte) - LBound(curByte) + 1
    Call WriteFile(mlngPipeHandle, curByte(0), bytesWrite, bytesWrite, stuOVERLAPPED)
    
    If bytesWrite <= 0 Then
        CloseHandle mlngPipeHandle
        Exit Function
    End If

    Call FlushFileBuffers(mlngPipeHandle)
    
    CloseHandle mlngPipeHandle
    
    SendData = True
End Function

'Ansi纯文本文件转换为Unicode(Little Endian)文本,该方式存在问题，如果经过该处理后，将“项”字发往.net，解析后的字母将变成其他字
Private Function AnsiToULE(ByVal sAnsi As String) As String
        Dim retLen As Long, FileNumber As Long
        Dim sUnicodeBuffer As String

        On Error Resume Next

        AnsiToULE = ""
        
        retLen = MultiByteToWideChar(CP_ACP, 0, sAnsi, -1, vbNullChar, 0) '取得转换后需要的空间大小retLen
        sUnicodeBuffer = String$(LenB(sAnsi), vbNullChar)  '设置缓冲区大小
        If retLen > 0 Then
           retLen = MultiByteToWideChar(CP_ACP, 0, sAnsi, -1, sUnicodeBuffer, retLen) '开始转换
        Else
           AnsiToULE = False: Exit Function
        End If
        
        AnsiToULE = sUnicodeBuffer
End Function

Public Function UnicodeEncode(bTemp As String, Optional BigEndian As Boolean = False) As Byte()
'编码UNICODE UTF-16
    Dim i As Long
    Dim k As Long
    Dim strTotal() As Byte
    Dim Code As Long
    Dim bNo As Long
    
    k = Len(bTemp)
    ReDim strTotal(k * 2)
    bNo = 0
    If BigEndian Then
        For i = 1 To k
            Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
            strTotal(bNo) = Code \ 256
            strTotal(bNo + 1) = Code And 255
            bNo = bNo + 2
        Next
    Else
        For i = 1 To k
            Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
            strTotal(bNo) = Code And 255
            strTotal(bNo + 1) = Code \ 256
            bNo = bNo + 2
        Next
    End If
    
    ReDim Preserve strTotal(bNo - 1)
    UnicodeEncode = strTotal
End Function


Public Function UnicodeDecode(bTemp() As Byte, Optional BigEndian As Boolean = False) As String
'解码UNICODE UTF-16

    Dim i As Long
    Dim strTotal() As Byte
    Dim strReturn As String
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim bNo As Long
    
    bNo = 0
    
    ReDim strTotal(UBound(bTemp))
    
    If BigEndian Then
        For i = LBound(bTemp) To UBound(bTemp) Step 2
            Code1 = bTemp(i)
            Code2 = bTemp(i + 1)
            Code = Code1 * 256 + Code2
            If Code > 255 Then
                Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
                strTotal(bNo) = Code And 255
                strTotal(bNo + 1) = Code \ 256
                bNo = bNo + 2
            Else
                strTotal(bNo) = Code
                bNo = bNo + 1
            End If
            
        Next
    Else
        For i = LBound(bTemp) To UBound(bTemp) Step 2
            Code1 = bTemp(i)
            Code2 = bTemp(i + 1)
            Code = Code2 * 256 + Code1
            If Code > 255 Then
                Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
                strTotal(bNo) = Code And 255
                strTotal(bNo + 1) = Code \ 256
                bNo = bNo + 2
            Else
                strTotal(bNo) = Code
                bNo = bNo + 1
            End If
        Next
    End If
    
    ReDim Preserve strTotal(bNo - 1)
    strReturn = StrConv(strTotal, vbUnicode)
    
    UnicodeDecode = strReturn
End Function

Public Function Unicode8Decode(bTemp() As Byte) As String
'解码UNICODE UTF-8
    Dim i As Long
    Dim k As Long
    Dim strReturn As String
    Dim strTmp() As Byte
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim Code3 As Long
    Dim Code4 As Long
    Dim bNo As Long
    
    k = UBound(bTemp)
    ReDim strTmp(k * 2)
    bNo = 0
    
    For i = 0 To k
        If (bTemp(i) And 128) = 0 Then
            strTmp(bNo) = bTemp(i)
            bNo = bNo + 1
        ElseIf (bTemp(i) And 252) = 252 Then
            '11111100
            Code1 = (bTemp(i) And 1) * 64 + bTemp(i + 1) And 63
            Code2 = (bTemp(i + 2) And 63) * 4 + (bTemp(i + 3) And 48) \ 16
            Code3 = (bTemp(i + 3) And 15) * 16 + (bTemp(i + 4) And 60) \ 4
            Code4 = (bTemp(i + 4) And 3) * 64 + (bTemp(i + 5) And 63)
            Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
            Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 5
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 248) = 248 Then '11111000
            Code1 = (bTemp(i) And 3)
            Code2 = (bTemp(i + 1) And 63) * 4 + (bTemp(i + 2) And 48) \ 16
            Code3 = (bTemp(i + 2) And 15) * 16 + (bTemp(i + 3) And 60) \ 4
            Code4 = (bTemp(i + 3) And 3) * 64 + (bTemp(i + 4) And 63)
            Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
            Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 4
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 240) = 240 Then '11110000
            Code1 = (bTemp(i) And 7) * 8 + (bTemp(i + 1) And 48) \ 16
            Code2 = (bTemp(i + 1) And 15) * 16 + (bTemp(i + 2) And 60) \ 4
            Code3 = (bTemp(i + 2) And 3) * 64 + (bTemp(i + 3) And 63)
            Code = (Code1 * 256 + Code2) * 256 + Code3
            Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 3
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            strTmp(bNo + 1) = Code \ 65536
            strTmp(bNo + 1) = Code \ 16777216
            bNo = bNo + 4
        ElseIf (bTemp(i) And 224) = 224 Then '11100000
            Code1 = (bTemp(i) And 15) * 16 + (bTemp(i + 1) And 60) \ 4
            Code2 = (bTemp(i + 1) And 3) * 64 + (bTemp(i + 2) And 63)
            Code = Code1 * 256 + Code2
            Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 2
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            bNo = bNo + 2
        ElseIf (bTemp(i) And 192) = 192 Then '11000000
            Code1 = (bTemp(i) And 28) \ 4
            Code2 = (bTemp(i) And 3) * 64 + (bTemp(i + 1) And 63)
            Code = Code1 * 256 + Code2
            Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
            i = i + 1
            strTmp(bNo) = Code And 255
            strTmp(bNo + 1) = Code \ 256
            bNo = bNo + 2
        End If
    Next
        
    ReDim Preserve strTmp(bNo - 1)
    strReturn = StrConv(strTmp, vbUnicode)
    Unicode8Decode = strReturn
End Function

Public Function Unicode8Encode(bTemp As String) As Byte()
'编码UNICODE UTF-8
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strTotal() As Byte
    Dim strTmp As String
    Dim Code As Long
    Dim Code1 As Long
    Dim Code2 As Long
    Dim Code3 As Long
    Dim Code4 As Long
    Dim Code5 As Long
    Dim Code6 As Long  '已生成的字节数
    Dim bNo As Long
    
    k = Len(bTemp)
    bNo = 0
    
    ReDim strTotal(k * 3)
    For i = 1 To k
        Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
        If Code < 128& Then
            strTotal(bNo) = Code
            bNo = bNo + 1
            If bNo > 422386 Then
                Debug.Print Code
            End If
        ElseIf Code < 2048& Then
            Code1 = ((Code And 1984&) \ 32&) + 192
            Code2 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            bNo = bNo + 2
        ElseIf Code < 65536 Then
            Code1 = ((Code And 61440) \ 4096&) + 224
            Code2 = ((Code And 4032&) \ 64&) + 128
            Code3 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            bNo = bNo + 3
        ElseIf Code < 2097152 Then
            Code1 = ((Code And 1835008) \ 262144) + 240
            Code2 = ((Code And 258048) \ 4096&) + 128
            Code3 = ((Code And 4032&) \ 64&) + 128
            Code4 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            bNo = bNo + 4
        ElseIf Code < 67108864 Then
            Code1 = ((Code And 50331648) \ 16777216) + 248
            Code2 = ((Code And 16515072) \ 262144) + 128
            Code3 = ((Code And 258048) \ 4096&) + 128
            Code4 = ((Code And 4032&) \ 64&) + 128
            Code5 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            bNo = bNo + 5
        Else
            Code1 = IIf(Code And 1073741824 = 1073741824, 253&, 252&)
            Code2 = ((Code And 1056964608) \ 16777216) + 128
            Code3 = ((Code And 16515072) \ 262144) + 128
            Code4 = ((Code And 258048) \ 4096&) + 128
            Code5 = ((Code And 4032&) \ 64&) + 128
            Code6 = (Code And 63&) + 128
            strTotal(bNo) = Code1
            strTotal(bNo + 1) = Code2
            strTotal(bNo + 2) = Code3
            strTotal(bNo + 3) = Code4
            strTotal(bNo + 4) = Code5
            strTotal(bNo + 5) = Code6
            bNo = bNo + 6
        End If
    Next
    
    ReDim Preserve strTotal(bNo - 1)
    Unicode8Encode = strTotal
End Function



Private Sub Class_Initialize()
    PipeName = DEFAULT_PIPE_NAME
End Sub





























