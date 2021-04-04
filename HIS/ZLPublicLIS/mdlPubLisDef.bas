Attribute VB_Name = "mdlPubLisDef"
Option Explicit

Public strChart(9)          As Variant
Public glngModual           As Long         '模块号
Public gblnNewLis           As Boolean      '是否为新版LIS
Public gstrHospital         As String
Public gstrFilePath         As String       '图像保存路径
Public gstrSignPath         As String
Public gbln显示图片         As Boolean
Public glngTop              As Long
Public glngLeft             As Long
Public gbtyModel            As Integer
Public gblnNew           As Boolean      '是否为新版LIS
Public Const gstr图片格式   As String = ".cht|.GIF|.gif|.jpg|.JPG|.bmp|.BMP|.JPEG|.jpeg|.png|.PNG"

'API 常数:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

Public Enum COLOR
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
End Enum

'----- 保存为JPG格式的图片
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
End Type

'结构类型:
Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal Filename As Long, clsIDEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, ID As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Function Between(X, a, b) As Boolean
    '******************************************************************************************************************
    '功能：判断x是否在a和b之间
    '******************************************************************************************************************
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    '******************************************************************************************************************
    '功能：将VB的系统颜色转换为RGB色
    '参数：
    '返回：
    '******************************************************************************************************************
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function SeveLoadImg(ByVal lng_标本ID As Long)
    On Error GoTo ErrH
    Dim rsTmp   As ADODB.Recordset
    Dim strSql  As String
    Dim intLoop As Integer
    
'    strSQL = "select ID from 检验图像结果 where 标本ID = [1] order by ID"
    If gblnNewLis Then
        strSql = "select ID from 检验报告图像 where 标本ID = [1] order by ID"
    Else
        strSql = "select ID from 检验图像结果 where 标本ID = [1] order by ID"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取检验图像", lng_标本ID)

    For intLoop = 1 To 9
        strChart(intLoop) = ""
    Next
    intLoop = 1
    Do Until rsTmp.EOF
        If intLoop > 9 Then Exit Do
        strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
        Debug.Print strChart(intLoop)
        Call LoadImageData(App.Path & "\", rsTmp("ID"), 1, "")
        
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    Exit Function
ErrH:
    Call ErrLog("mdlimg", "SeveLoadImg", "图片加载错误", err.Description)
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
'Public Function zlFileUnzip(ByVal strZipFile As String) As String
'    Dim strZipPath As String
'    Dim clsUnzip As New clsUnzip
'
'    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
'    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
'
'    With clsUnzip
'        .ZipFile = strZipFile
'        .UnzipFolder = strZipPath
'        .Unzip
'    End With
'    If Dir(strZipPath) <> "" Then
'        zlFileUnzip = strZipPath & Dir(strZipPath)
'    Else
'        zlFileUnzip = ""
'    End If
'End Function
'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
'Public Function zlFileZip(ByVal strFile As String, ByVal strFilename As String) As String
'    Dim strZipFile As String, lngCount As Long
'    Dim clsZip  As New clsUnzip
'
'    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
'
'    lngCount = 0
'    Do While True
'        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
'        If Dir(strZipFile) = "" Then Exit Do
'        lngCount = lngCount + 1
'    Loop
'
'    With mclsZip
'        .Encrypt = False: .AddComment = False
'        .ZipFile = strZipFile
'        .StoreFolderNames = False
'        .RecurseSubDirs = False
'        .ClearFileSpecs
'        .AddFileSpec strFile
'        .Zip
'        If (.Success) Then
'            zlFileZip = .ZipFile
'        Else
'            zlFileZip = ""
'        End If
'    End With
'End Function

'*************************************************************************
'**函 数 名：PrintRotText
'**输    入：ByVal hDC(Long)          -
'**        ：ByVal Text(String)       -  要打印的文字
'**        ：ByVal CenterX(Long)      -  X中心点的文字像素
'**        ：ByVal CenterY(Long)      -  Y中心点的文字像素
'**        ：ByVal RotDegrees(Single) -  旋转角度(0.0 至 359.9999999) 反顺时针，0=水平(不旋转)
'**输    出：(Boolean) -
'**功能描述：在一个对象上以中心X,中心Y坐标轴上以角度绘制旋转文字
'**全局变量：
'**调用模块：
'*************************************************************************
Public Function PrintRotText(ByVal hDC As Long, ByVal Text As String, ByVal CenterX As Long, ByVal CenterY As Long, ByVal RotDegrees As Single) As Boolean

Dim bOkSoFar    As Boolean      '继续标识.
Dim hFontOld    As Long         '原字体句柄
Dim hFontNew    As Long         '新字体句柄
Dim lfFont      As LOGFONT      'LOGFONT 新字体结构.
Dim ptOrigin    As POINTAPI     '文字绘制原点
Dim ptCenter    As POINTAPI     '文字中心点.
Dim szText      As SizeStruct   '文字宽度和高度

    '从设备中得到当前 LOGFONT 结构.
    hFontOld = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    
    '如果从设备得到的字体成功...
    If hFontOld <> 0 Then
        
        '从字体获取 LOGFONT 结构
        bOkSoFar = (GetObjectAPI(hFontOld, Len(lfFont), lfFont) <> 0)
        
        '把原字体重载
        Call SelectObject(hDC, hFontOld)
        
        '复位稍后使用
        hFontOld = 0
    End If
    
    '如果成功获得 LOGFONT 结构，继续.
    If bOkSoFar Then
    
        '改变字体方向和出口
        lfFont.lfEscapement = RotDegrees * 10
        lfFont.lfOrientation = lfFont.lfEscapement
        lfFont.lfQuality = ANTIALIASED_QUALITY
        
        '从 LOGFONT 结构中创建新字体对象
        hFontNew = CreateFontIndirect(lfFont)
        
        '字体创建成功
        If hFontNew <> 0 Then
            
            'Select the ne选择新的字体到该设备
            hFontOld = SelectObject(hDC, hFontNew)
            
            '成功
            If hFontOld <> 0 Then
                
                '获取文字逻辑单位大小(像素)
                bOkSoFar = (GetTextExtentPoint32(hDC, Text, LenB(StrConv(Text, vbFromUnicode)), szText) <> 0)
                
                '成功
                If bOkSoFar Then
                    
                    '计算文字水平原点
                    With ptOrigin
                        .X = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    '转换 CenterX, CenterY 到点结构
                    '(需要调用 RotatePoint).
                    With ptCenter
                        .X = CenterX
                        .Y = CenterY
                    End With
                    
                    '以原点选择以匹配预期选择
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    '现在打印旋转文本并返回成功/失败
                    PrintRotText = (TextOut(hDC, ptOrigin.X, _
                      ptOrigin.Y, Text, LenB(StrConv(Text, vbFromUnicode))) <> 0)
                
                End If
                
                '恢复字体到原先设备
                hFontNew = SelectObject(hDC, hFontOld)
            
            End If
            
            '清除内存并删除创建的字体
            Call DeleteObject(hFontNew)
        
        End If
        
    End If
            
End Function

'*************************************************************************
'**    作    者 ：    laviewpbt
'**    函 数 名 ：    SavePic
'**    输    入 ：    pic(StdPicture)        -   图象句柄
'**             ：    FileName(String)       -   保存路径
'**             ：    Quality(Byte)          -   JPG图象质量
'**             ：    TIFF_ColorDepth(Long)  -   TTF格式的颜色深度
'**             ：    TIFF_Compression(Long) -   TTF格式的压缩比
'**    输    出 ：    无
'**    功能描述 ：    把图象保存为JPG、TIFF、PNG、GIF、BMP格式
'**    日    期 ：
'**    修 改 人 ：    laviewpbt
'**    日    期 ：    2005-10-23 14.43.52
'**    版    本 ：    Version 1.2.1
'*************************************************************************
Public Sub SavePic(ByVal pict As StdPicture, ByVal Filename As String, PicType As String, _
                    Optional ByVal Quality As Byte = 100, _
                    Optional ByVal TIFF_ColorDepth As Long = 24, _
                    Optional ByVal TIFF_Compression As Long = 6)
100    Screen.MousePointer = vbHourglass
       Dim tSI As GdiplusStartupInput
       Dim lRes As Long
       Dim lGDIP As Long
       Dim lBitmap As Long
       Dim aEncParams() As Byte
       On Error GoTo errHandle:
102    tSI.GdiplusVersion = 1   ' 初始化 GDI+
104    lRes = GdiplusStartup(lGDIP, tSI)
106    If lRes = 0 Then     ' 从句柄创建 GDI+ 图像
108       lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
110       If lRes = 0 Then
             Dim tJpgEncoder As GUID
             Dim tParams As EncoderParameters    '初始化解码器的GUID标识
112          Select Case UCase(PicType)
             Case ".JPG", "JPG", ".JPEG", "JPEG"
114             CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
116             tParams.count = 1                               ' 设置解码器参数
118             With tParams.Parameter ' Quality
120                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID    ' 得到Quality参数的GUID标识
122                .NumberOfValues = 1
124                .type = 4
126                .Value = VarPtr(Quality)
                End With
128             ReDim aEncParams(1 To Len(tParams))
130             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
132         Case ".PNG", "PNG"
134              CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
136              ReDim aEncParams(1 To Len(tParams))
138         Case ".GIF", "GIF"
140              CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
142              ReDim aEncParams(1 To Len(tParams))
144         Case ".TIFF", "TIFF"
146              CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
148              tParams.count = 2
150              ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
152              With tParams.Parameter
154                 .NumberOfValues = 1
156                 .type = 4
158                  CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID    ' 得到ColorDepth参数的GUID标识
160                 .Value = VarPtr(TIFF_Compression)
                End With
162             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
164             With tParams.Parameter
166                 .NumberOfValues = 1
168                 .type = 4
170                  CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID    ' 得到Compression参数的GUID标识
172                 .Value = VarPtr(TIFF_ColorDepth)
                End With
174             Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
176         Case ".BMP", "BMP"                                              '可以提前写保存为BMP的代码，因为并没有用GDI+
178             SavePicture pict, Filename
180             Screen.MousePointer = vbDefault
                Exit Sub
            End Select
182          lRes = GdipSaveImageToFile(lBitmap, StrPtr(Filename), tJpgEncoder, aEncParams(1))             '保存图像
184          GdipDisposeImage lBitmap       ' 销毁GDI+图像
          End If
186       GdiplusShutdown lGDIP              '销毁 GDI+
       End If
188    Screen.MousePointer = vbDefault
190    Erase aEncParams
       Exit Sub
errHandle:
192     Screen.MousePointer = vbDefault
194     WriteLog "mdlPublic.SavePic", CStr(Erl()) & "行", err.Description
End Sub


'*************************************************************************
'**函 数 名：RotatePoint
'**输    入：ptAxis(PointAPI)   -
'**        ：ptRotate(PointAPI) -
'**        ：fDegrees(Single)   -
'**输    出：无
'**功能描述：从前fdegrees当前坐标选择ptRotate左右的ptAxis
'**全局变量：
'**调用模块：
'*************************************************************************
Private Sub RotatePoint(ptAxis As POINTAPI, ptRotate As POINTAPI, fDegrees As Single)

' ***************************************************
' *                 RotatePoint                     *
' *                                                 *
' *  Created by: Rocky Clark (Kath-Rock Software)   *
' *                                                 *
' *  Rotate ptRotate around ptAxis, fDegrees from   *
' *  its current position.                          *
' *                                                 *
' * This procedure may be used and distributed, as  *
' * is, in your code, as long as these credits and  *
' * the code itself remain unchanged.               *
' *                                                 *
' ***************************************************

Dim fDX     As Single   'X坐标
Dim fDY     As Single   'Y坐标
Dim fRads   As Single   '弧度
Const dPi   As Double = 3.14159265358979  'Pi 圆周率


    '转换角度为弧度
    fRads = fDegrees * (dPi / 180#)
    
    '从中心点计算入口
    fDX = ptRotate.X - ptAxis.X
    fDY = ptRotate.Y - ptAxis.Y
    
    '旋转点
    ptRotate.X = ptAxis.X + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub



Public Function DeleteImge()

On Error GoTo ErrH
    Dim strFilename         As String
    Dim intLoop             As Integer
    Dim objFso              As New FileSystemObject
    Dim objFile             As File
    Dim strCreateTime       As String
    Dim strLastVisitTime    As String
    Dim strLastModityTime   As String
    Dim intLostDay          As Integer
    Dim currDate            As Date
    
    intLostDay = 20

    For intLoop = LBound(Split(gstr图片格式, "|")) To UBound(Split(gstr图片格式, "|"))
        strFilename = Dir(gstrFilePath & "\*" & Split(gstr图片格式, "|")(intLoop), vbNormal)  ' 找寻第一项。
        
        Do While strFilename <> ""   ' 开始循环。
            ' 跳过当前的目录及上层目录。
            If strFilename <> ".." Then
                Set objFile = objFso.GetFile(gstrFilePath & "\" & strFilename)
                strCreateTime = Format(objFile.DateCreated(), "yyyy-MM-dd")
                strLastVisitTime = Format(objFile.DateLastAccessed(), "yyyy-MM-dd")
                strLastModityTime = Format(objFile.DateLastModified(), "yyyy-MM-dd")
                If CDate(strCreateTime) < CDate(Date - intLostDay) Then
                    Kill gstrFilePath & "\" & strFilename
                End If
            End If
            strFilename = Dir   ' 查找下一个目录。
        Loop
    Next
    Set objFile = Nothing
    Set objFso = Nothing
    Exit Function
ErrH:
    WriteLog "deleteimge", err.Description, ""
End Function

Public Sub ErrLog(strObj As String, strEvent As String, strErrNum As String, strErrDesc As String)
On Error GoTo ErrH
    '将调试信息写入文件中
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    If strErrNum = "9999" Then
        strFile = App.Path & "\检验报告打印\错误日志\ErrLog" & Format(Date, "YYYYMMDD") & ".txt"
    Else
        strFile = App.Path & "\检验报告打印\错误日志\ErrLog" & Format(Date, "YYYYMMDD") & ".Log"
    End If
    If Not objFile.FolderExists(App.Path & "\检验报告打印") Then
        objFile.CreateFolder (App.Path & "\检验报告打印")
    End If
    If Not objFile.FolderExists(App.Path & "\检验报告打印\错误日志") Then
        objFile.CreateFolder (App.Path & "\检验报告打印\错误日志")
    End If
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine "出错对象：" & strObj
    objText.WriteLine "事件对象：" & strEvent
    objText.WriteLine "错误号：" & strErrNum
    objText.WriteLine "错误描述：" & strErrDesc
    objText.Close
    Set objText = Nothing
    Set objFile = Nothing
    Exit Sub
ErrH:
    err.Clear
    Exit Sub
End Sub
