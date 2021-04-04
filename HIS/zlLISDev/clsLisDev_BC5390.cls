VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BC5390"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血液分析仪 BC5390
'厂商:迈瑞(Mindary)
'作者:黄金桥
'用户:阆中市人民医院

Private mstrCode As String  '待解析数据
Const mstrItem  As String = ",WBC,BAS#,BAS%,NEU#,NEU%,EOS#,EOS%,LYM#,LYM%,MON#,MON%,*ALY#,*ALY%,*LIC#,*LIC%,RBC,HGB,MCV,MCH,MCHC,RDW-CV,RDW-SD,HCT,PLT,MPV,PDW,PCT,PLCC,PLCR,"
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const BITMAPTYPE = &H4D42
Private Const INVALID_HANDLE_VALUE = (-1)
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long
                             
Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
                             
Private Declare Function SetPixelV _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long
                             
Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long) As Long
                             
Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     ByVal lpSecurityAttributes As Long, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long
                                     
Private Declare Function WriteFile _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpBuffer As Any, _
                                ByVal nNumberOfBytesToWrite As Long, _
                                lpNumberOfBytesWritten As Long, _
                                ByVal lpOverlapped As Long) As Long
                                
Private Declare Function CreateDIBSection _
                Lib "gdi32" (ByVal hDC As Long, _
                             pBitmapInfo As BITMAPINFO, _
                             ByVal un As Long, _
                             lplpVoid As Long, _
                             ByVal handle As Long, _
                             ByVal dw As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



'--------------------------------------------------------------------------------------------------------------------
'以下是必须的接口函数
Public Function GetAnswerCmd() As String
    '自动应答指令
    GetAnswerCmd = Chr(&H6)
End Function

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_字段分隔符 As String
    Dim str_组件分隔符 As String
    Dim str_子组件隔符 As String
    Dim str_重复分隔符 As String
    Dim str_转义字符 As String
    
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    Dim str_唯一标识 As String '用于标识一个消息
    Dim str_版本 As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_质控 As String
    Dim str_项目 As String
    Dim str_结果 As String
    Dim str_Return As String
    
    Dim str_Base64 As String, str_WBC As String, str_RBC As String, str_PLT As String, str_DIFF As String, str_img As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        
        'Call writelislog("原始数据:" & str_Code)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '消息头
                str_字段分隔符 = Mid(str_Segment, 5, 1)
                str_组件分隔符 = Mid(str_Segment, 6, 1)
                str_子组件隔符 = Mid(str_Segment, 7, 1)
                str_重复分隔符 = Mid(str_Segment, 8, 1)
                str_转义字符 = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_唯一标识 = var_Filed(9)
                str_质控 = var_Filed(10)
                str_版本 = var_Filed(11)
                If str_版本 <> "2.3.1" Then
                    WriteLog "版本不符", "", ""
                End If
                str_Return = "": str_标本号 = "": str_日期 = ""
            ElseIf InStr(str_Segment, "OBR|") > 0 Then
                '标本号，日期
                If Val(str_标本号) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
                    str_img = ""
                    If str_WBC <> "" Then str_img = str_img & "^BASO;103;0;" & str_WBC
                    If str_DIFF <> "" Then str_img = str_img & "^DIFF;103;0;" & str_DIFF
                    If str_RBC <> "" Then str_img = str_img & "^RBC;103;0;" & str_RBC
                    If str_PLT <> "" Then str_img = str_img & "^PLT;103;0" & str_PLT
                    If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
                End If
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_标本号 = var_Filed(3)
                str_日期 = var_Filed(7)
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                str_Return = ""
            ElseIf InStr(str_Segment, "OBX|") > 0 Then
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_项目 = var_Filed(3)
                If InStr(mstrItem, "," & Split(str_项目, str_组件分隔符)(1) & ",") > 0 Then
                    str_Return = str_Return & "|" & Split(str_项目, str_组件分隔符)(1) & "|" & Trim(var_Filed(5))
                ElseIf InStr(",15056,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'RBC Histogram(直方图)
                    'Call writelislog("解析数据:" & str_Return)
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    'str_Base64 = Base64Decode(str_Base64)
                    str_RBC = SaveBase64Bmp("RBC", str_标本号, str_Base64) 'Graph_Text(str_Base64)
                ElseIf InStr(",15116,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'PLT Histogram(直方图)
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    'str_Base64 = Base64Decode(str_Base64)
                    str_PLT = SaveBase64Bmp("PLT", str_标本号, str_Base64) 'Graph_Text(str_Base64)
                ElseIf InStr(",15200,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'WBC DIFF(散点图)
                    'Call writelislog("str_WBC: " & Split(Trim(var_Filed(5)), str_组件分隔符)(4))
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    str_DIFF = SaveBase64Bmp("DIFF", str_标本号, str_Base64)
                ElseIf InStr(",15008,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'Baso Scattergram(散点图)
                    'Call writelislog("str_DIFF: " & Split(Trim(var_Filed(5)), str_组件分隔符)(4))
                    str_Base64 = Split(var_Filed(5), str_组件分隔符)(4)
                    Call writelislog("解析数据:" & str_Base64)
                    str_WBC = SaveBase64Bmp("BASO", str_标本号, str_Base64)
                End If
            
            End If
        Next
       ' Call writelislog("分段数据: " & "str_Return: " & str_Return & "str_标本号: " & str_标本号 & "str_日期: " & str_日期)
        If Val(str_标本号) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
            str_img = ""
            If str_WBC <> "" Then str_img = str_img & "^BASO;103;0;" & str_WBC
            If str_DIFF <> "" Then str_img = str_img & "^DIFF;103;0;" & str_DIFF
            If str_RBC <> "" Then str_img = str_img & "^RBC;103;0;" & str_RBC
            If str_PLT <> "" Then str_img = str_img & "^PLT;103;0;" & str_PLT
            If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
      'Call writelislog("strResult: " & strResult)
    Exit Sub
errHandle:
    WriteLog "BC5800_analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function

'---------------------------------------------------------------------------------------------------------------
'以下是内部函数
Private Sub TransparentBmp(ByVal strPicPath As String)
    Dim hmemDC         As Long
    Dim hmemBMP        As Long
    Dim lpmemBits      As Long
    Dim bmp_info       As BITMAPINFO
    Dim hFile          As Long
    Dim bmpfile_info   As BITMAPFILEHEADER
    Dim lpBytesWritten As Long
    Dim membmp         As Long, memdc As Long
    Dim x              As Integer, Y As Integer
    Dim cr             As Long, R As Integer, g As Integer, b As Integer
    Dim prR
    frmGraph.Picture1.AutoSize = True
    frmGraph.Picture1.ScaleMode = 3 'Pixel
    If Trim(strPicPath) <> "" Then frmGraph.Picture1.Picture = LoadPicture(strPicPath) 'Load Image
    memdc = CreateCompatibleDC(frmGraph.Picture1.hDC) 'Create Compatible DC
    SelectObject memdc, frmGraph.Picture1.Picture.handle 'Select Image
    
    'Turn Colors:
    For x = 0 To (frmGraph.Picture1.ScaleWidth - 1)
        For Y = 0 To (frmGraph.Picture1.ScaleHeight - 1)
            cr = GetPixel(memdc, x, Y) 'Get Colors
            R = cr Mod 256
            g = (cr \ 256) Mod 256
            b = cr \ 65536
'            If R = 255 And g = 255 And b = 255 Then
'                SetPixelV memdc, x, Y, RGB(255, 255, 255)
'            Else
            '将黑色变成白色
            If R = 0 And g = 0 And b = 0 Then
                SetPixelV memdc, x, Y, RGB(255 - R, 255 - g, 255 - b)
            End If
            '显示方式1     将白色变成黑色  其他颜色不改变 ，能够打印彩色时可以使用
            If R = 255 And g = 255 And b = 255 Then
                SetPixelV memdc, x, Y, RGB(0, 0, 0)
            End If
            '显示方式2     '所有的颜色都以黑色显示  ，对图像显示和打印都要求以黑色就可以采用这种方式
'            If R <> 0 Or g <> 0 Or b <> 0 Then
'                SetPixelV memdc, x, Y, RGB(0, 0, 0)
'            End If
            'Reset Colors:
        
            
            DoEvents 'Have a rest
        Next
    Next
    'Draw Out
    BitBlt frmGraph.Picture1.hDC, 0, 0, frmGraph.Picture1.Width, frmGraph.Picture1.Height, memdc, 0, 0, vbSrcCopy

    frmGraph.Picture1.ScaleMode = vbPixels
    With bmp_info.bmiHeader
        .biSize = LenB(bmp_info.bmiHeader)
        .biWidth = frmGraph.Picture1.ScaleWidth
        .biHeight = frmGraph.Picture1.ScaleHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = .biHeight * (((.biWidth * .biBitCount + 31) And &HFFFFFFE0) \ 8)

    End With
   
    frmGraph.Picture1.Refresh
    
    hmemDC = CreateCompatibleDC(frmGraph.Picture1.hDC)
    hmemBMP = CreateDIBSection(frmGraph.Picture1.hDC, bmp_info, DIB_RGB_COLORS, lpmemBits, 0, 0)
    SelectObject hmemDC, hmemBMP
    BitBlt hmemDC, 0, 0, bmp_info.bmiHeader.biWidth, bmp_info.bmiHeader.biHeight, frmGraph.Picture1.hDC, 0, 0, vbSrcCopy
    '保存图片
    hFile = CreateFile(strPicPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

    If hFile <> INVALID_HANDLE_VALUE Then

        With bmpfile_info
            .bfType = BITMAPTYPE
            .bfOffBits = 14 + bmp_info.bmiHeader.biSize
            .bfSize = .bfOffBits + bmp_info.bmiHeader.biSizeImage

        End With

        WriteFile hFile, bmpfile_info.bfType, 2, lpBytesWritten, 0
        WriteFile hFile, bmpfile_info.bfSize, 12, lpBytesWritten, 0
        WriteFile hFile, bmp_info.bmiHeader, bmp_info.bmiHeader.biSize, lpBytesWritten, 0
        WriteFile hFile, ByVal lpmemBits, bmp_info.bmiHeader.biSizeImage, lpBytesWritten, 0
        CloseHandle hFile

    End If

    DeleteObject hmemBMP
    DeleteDC hmemDC
End Sub

Private Function SaveBase64Bmp(ByVal strType As String, ByVal str标本号 As String, ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    Dim strDecode       As String
    Dim str             As String
    Dim ByteData()      As Byte
    Dim OutPut()        As Byte
    Dim lFrom           As Long
    Dim lTo             As Long
    Dim i               As Long
    Dim lngCount As Long, strFilename As String, lngFileNum
    
    '除去回车
    str = Replace(str2Decode, vbCrLf, "")

    '每4个字符一组（4个字符表示3个字）
    For lPtr = 1 To Len(str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '查找字符在BASE64字符串中的位置
            iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                '没有发现
                Case 0: Exit Function
            End Select
        Next

        '转换4个6比特数成为3个8比特数
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
        
        '计算数组的起始位置
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1

        '重新定义输出数组
        ReDim Preserve OutPut(0 To lTo)

        For lngCount = lFrom To lTo
            OutPut(lngCount) = bits(lngCount - lFrom + 1)
        Next
        lTo = lTo + 1
    Next
    ReDim ByteData(UBound(OutPut))
    For i = 0 To UBound(OutPut)
        ByteData(i) = OutPut(i)
    Next
    
   '
        For i = 1078 To UBound(OutPut) - 2 Step 2
            If OutPut(18) = &H18 And OutPut(19) = &H1 Then '直方图
'               If Output(i) <> byt1 Or byte2 <> Output(i + 1) Then
'                   byt1 = Output(i): byte2 = Output(i + 1)
''                   Debug.Print byt1, byte2
'               End If
                If OutPut(i) = OutPut(i - 2) And OutPut(i + 1) = OutPut(i - 1) And OutPut(i) = OutPut(i + 2) And OutPut(i + 1) = OutPut(i + 3) And ((OutPut(i) = 0 And OutPut(i + 1) = 124) Or (OutPut(i) = 255 And OutPut(i + 1) = 127) Or (OutPut(i) = 224 And OutPut(i + 1) = 3)) Then
                    If ((i - 54) \ 560 > 0 And (i - 54) \ 560 <> 159) Then
                        'If Output(i) = Output(i - 560) And Output(i + 1) = Output(i - 559) Then 'And Output(i) = Output(i + 560) And Output(i + 1) = Output(i + 561)
                        If OutPut(i) = OutPut(i + 560) And OutPut(i + 1) = OutPut(i + 561) Then
                            ByteData(i) = 0: ByteData(i + 1) = 0
    '                        Debug.Print (i - 54) \ 560
                        End If
                    End If
                End If
            End If
    
            'output2(i) = 255 - output2(i): output2(i + 1) = 255 - output2(i + 1)
        Next i

    strFilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    TransparentBmp strFilename
    SaveBase64Bmp = strFilename

End Function

Private Function Hex_To_Dec(ByVal Hex As String) As Long
  Dim i As Integer
  Dim b As Long
  Hex = UCase(Hex)
  For i = 1 To Len(Hex)
    Select Case Mid(Hex, Len(Hex) - i + 1, 1)
        Case "0": b = b + 16 ^ (i - 1) * 0
        Case "1": b = b + 16 ^ (i - 1) * 1
        Case "2": b = b + 16 ^ (i - 1) * 2
        Case "3": b = b + 16 ^ (i - 1) * 3
        Case "4": b = b + 16 ^ (i - 1) * 4
        Case "5": b = b + 16 ^ (i - 1) * 5
        Case "6": b = b + 16 ^ (i - 1) * 6
        Case "7": b = b + 16 ^ (i - 1) * 7
        Case "8": b = b + 16 ^ (i - 1) * 8
        Case "9": b = b + 16 ^ (i - 1) * 9
        Case "A": b = b + 16 ^ (i - 1) * 10
        Case "B": b = b + 16 ^ (i - 1) * 11
        Case "C": b = b + 16 ^ (i - 1) * 12
        Case "D": b = b + 16 ^ (i - 1) * 13
        Case "E": b = b + 16 ^ (i - 1) * 14
        Case "F": b = b + 16 ^ (i - 1) * 15
    End Select
  Next i
  Hex_To_Dec = b
End Function

'解析直方图数据
Private Function BC5600(ByVal strIn As String) As String
    Dim strImg As String
    Dim lngI As Long, lng_len As Long, Data As Long
    Dim str As String, DataStr As String
    
    lng_len = Len(strIn)
    If lng_len < 1 Then
        BC5600 = ""
        Exit Function
    End If
    str = strIn
    strImg = ""
    
    For lngI = 1 To Len(strIn)
        Data = AscB(Mid(strIn, lngI, 1))
        'If Chr(Mid(strIn, lngI, 1)) <> Chr(32) Then
          If strImg <> "" Then strImg = strImg & ";"
          strImg = strImg & Format(CStr(Data), "000")
       'End If
    Next
    BC5600 = strImg
End Function

Private Function Graph_Text(ByVal str_bin As String) As String
    '生成图形文本数据
  Dim a() As Byte, c As String, i As Integer, Tmp As String
  Dim varLine As Variant
  Dim str_Line As Long
  c = ""
  a = StrConv(str_bin, vbFromUnicode)
  For i = LBound(a) To UBound(a)
    Tmp = Hex(a(i))
      If Trim(c) <> "" Then c = c & ";"
      c = c & Tmp
  Next
  Tmp = ""
  varLine = Split(c, ";")
  For i = 1 To UBound(varLine)
    str_Line = Hex_To_Dec(varLine(i))
    If (str_Line <> 63) And (str_Line <> 32) Then
      If Tmp <> "" Then Tmp = Tmp & ";"
      Tmp = Tmp & CStr(str_Line)
    End If
  Next
  Graph_Text = Tmp
End Function

'Base64解码
Private Function Base64Decode(strEncoded As String) As String '解码

    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT, bRet() As Byte
    Dim i As Long, J As Long
    arrB = StrConv(strEncoded, vbFromUnicode)
    J = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(J - J \ 4 - 1)
    For i = 0 To J Step 4
        Erase bTmp
        bTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

'-------------------------------------------------------------------------------------------------------------------
'保存二进制数据为BMP图形文件
Private Function SaveBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str标本号 As String) As String
    '保存二进制数据为BMP图形文件
    Dim strFilename As String, lngFileNum As Long, lngLoop As Long
    Dim strByte As String, ByteData() As Byte
    Dim strBinData As String
    Dim lngCount As Long
    
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    ReDim ByteData(lngCount)
    
    For lngLoop = 0 To lngCount - 1
        strByte = Left(strBinData, 3)
        strBinData = Mid(strBinData, 4)
        ByteData(lngLoop) = CByte("&H" & Mid(strByte, 2))
    Next
    
    strFilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBmpToFile = strFilename
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFilename As String
    strFilename = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function





