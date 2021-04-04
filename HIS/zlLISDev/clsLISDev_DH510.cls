VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DH510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:DH510
'厂商:嘉善加斯戴克医疗器械有限公司
'作者:敬秦
'用户:天府新区人民医院
'日期:2015/10/30
Private mstrCode As String
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private Const cntIMAGE64CHR As String = "WBC,RBC,MUCS"
Private Declare Function intsetbackgroundcolor Lib "SsmDraw.dll" Alias "intSetBackgroundColor" (ByVal strgirfilename As String, ByVal intType As Integer) As Integer
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim lngS As Long, lngE As Long
    Dim strSample As String, strChannel As String, strValue As String, strDate As String, strStat As String, strQc As String
    Dim arrCode() As String, strLine As String, strCode As String, strItem As String
    Dim strImg As String, strImgName As String
    Dim intLoop As Integer
    On Error GoTo ErrorHand:
    If strSource <> "" Then mstrCode = mstrCode & strSource
'    lngS = InStr(mstrCode, Chr(&HB))   '仪器
    lngS = InStr(mstrCode, "MSH")
    lngE = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Do While lngS > 0 And lngE > 0
        strCode = Mid(mstrCode, lngS, lngE + 1)
        mstrCode = Mid(mstrCode, lngE + 2)
        lngS = InStr(mstrCode, "MSH")
        lngE = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
        arrCode = Split(strCode, Chr(&HD))
        For intLoop = 0 To UBound(arrCode)
            strLine = arrCode(intLoop)
            If InStr(strLine, "OBR|") > 0 Then
                strSample = Split(strLine, "|")(3)
                strDate = Split(strLine, "|")(7)
                strDate = Format(Left(strDate, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Right(strDate, 2), "yyyy-mm-dd hh:mm:ss")
                If UCase(Split(strLine, "|")(15)) = "QC" Then
                    strQc = "1"
                End If
                If UCase(Split(strLine, "|")(27)) = "N" Then
                    strStat = "0"
                Else
                    strStat = "1"
                End If
            End If
            If InStr(strLine, "OBX|") > 0 Then
                If InStr(UCase(Split(strLine, "|")(4)), "IMG") > 0 Then '处理图像数据
                    strImgName = ""
'                    strImgName = SaveBase64Bmp(Split(strLine, "|")(4), strSample, Split(strLine, "|")(5))
                    strImgName = SaveBase64Bmp(Split(strLine, "|")(4), strSample, Split(strLine, "|")(5))
                    TransparentBmp (strImgName)
                    strImg = strImg & "^" & Split(strLine, "|")(4) & ";103;0;" & strImgName
                Else
                    strChannel = Split(strLine, "|")(4)
                    strValue = Split(strLine, "|")(5)
                    If strChannel <> "" And strValue <> "" Then
                        strItem = strItem & "|" & strChannel & "|" & strValue
                    End If
                End If
            End If
        Next
        If strItem <> "" Then
            strResult = strDate & "|" & strSample & "^" & strStat & "^| |血清|" & IIf(strQc = "1", "1", "0") & strItem
            strItem = ""
            Call WriteLog("不带图像数据", strSource, strResult)
            If strImg <> "" Then '带上图片
                strResult = strResult & vbCrLf & Mid(strImg, 2)
                strImg = ""
                Call WriteLog("带图像数据", strSource, strResult)
            End If
        End If
    Loop
    Exit Sub
ErrorHand:
    Call WriteLog("调试日志", strSource, Err.Description)
    Err.Clear
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
    Dim lngCount As Long, strFileName As String, lngFileNum
    
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
   
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , OutPut
    Close lngFileNum
    
    SaveBase64Bmp = strFileName

End Function
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
    Dim cr             As Long, r As Integer, g As Integer, B As Integer
    
    frmGraph.Picture1.AutoSize = True
    frmGraph.Picture1.ScaleMode = 3 'Pixel
    If Trim(strPicPath) <> "" Then frmGraph.Picture1.Picture = LoadPicture(strPicPath) 'Load Image
    memdc = CreateCompatibleDC(frmGraph.Picture1.hDC) 'Create Compatible DC
    SelectObject memdc, frmGraph.Picture1.Picture.handle 'Select Image

    'Turn Colors:黑色的背景处理为白色的背景图片
    For x = 0 To (frmGraph.Picture1.ScaleWidth - 1)
        For Y = 0 To (frmGraph.Picture1.ScaleHeight - 1)
            cr = GetPixel(memdc, x, Y) 'Get Colors
            r = cr Mod 256
            g = (cr \ 256) Mod 256
            B = cr \ 65536
            'Reset Colors:
            SetPixelV memdc, x, Y, RGB(255 - r, 255 - g, 255 - B)
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


