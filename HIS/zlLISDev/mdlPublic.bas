Attribute VB_Name = "mdlPublic"
Option Explicit


' ***************************************************
' *             文本旋转模块                        *
' *                                                 *
' ***************************************************

Public uDisplayDescript  As Boolean      '选中时显示详细描述

'API 常数:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

'结构类型:
Private Type PointAPI
    x   As Long
    Y   As Long
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
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

'API 声明:
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'----- 保存为JPG格式的图片
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
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

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal Filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public dX As Long, dy As Long          ' distance XY = size of snapping zone
Public X1 As Long, X2 As Long          ' co鰎dinates snapping zone
Public Y1 As Long, Y2 As Long

'--------------------------------------------------------------
Dim lngTime
'       解码函数 和 返回串格式说明
'    ResultFromFile 函数  以字符串数组方式返回解码结果, 一个数组元素包含一组检验结果;
'    Analyse        函数  以字符串方式返回解码结果,每组检验结果以||分隔
'    每组检验结果的元素之间以|分隔,下面详细说明

    '   第0个元素：检验时间
    '   第1个元素：样本序号[^是否急诊^条码]
    '              ^是否急诊^条码 是可选项,仪器返回有条码时才使用.
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    '       酶标结果格式:定性结果^OD^CutOff^SCO
    '
    'Analyse strCmd 参数：如果需要，可返回向设备发送的命令
    '   1.向仪器指令不包含“|”：
    '       a.以二进制串方式返回自动应答指令如要固定应答06则写为,06
    '   2.向仪器发送的指令包含“|”：
    '       a.在需要发送的指令前添加“1|”，此类指令表示仪器请求获取标本信息，并非检验结果
    '       b.在需要发送的指令前添加“0|”，此类指令属于特殊情况，加“0|”是为了和“1|”进行区分（HL7协议向仪器发送的指令一般都包含“|”）
    '
    
    
    '-- 返回图形数据时的格式:
    '补充图像的方式：
    '                   1.图像数据跟随指标数据后，使用回车换行符来分隔。
    '                   2.有多个图像数据时使用"^"来分隔
    '                   3.单个图像数据格式: 图像画法 0=直方图  1=散点图  2=血流变粘度特征曲线  3=血沉曲线 4=PLT双曲线图 5=带界标支持双曲线，XY座标刻度值的直方图 100以上为图片数据
    '                     0) 直方图: 图像名称;图像画法(0=直方图  1=散点图);Y1;Y2;Y3;Y4;Y5...
    '                     1) 散点图: 图像名称;图像画法(0=直方图  1=散点图):
    '                        例:00000100001000010000100010;00000100001000010000100010;
    '                        说明: 1.散点图以点阵方式保存每一行使用分号来分隔.
    '                              2.有多少个分号就有多少行
    '                              3.每一行有多少个点由每一行的长度来确定
    '                              4.画图的方向是从最上边向下画，如有65*65的图就是从65行开始画(最上边开始画)
    '                     2) 粘度特征曲线:图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据
    '                                   其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
    '                                   曲线及描点数据:粘度曲线1的高点和低点座标|粘度曲线2的高点和低点座标~低切点坐标,中切点坐标,高切点坐标
    '                                   坐标轴标题数据:Y坐标标题文字,X坐标,Y座标~X坐标标题文字,X坐标,Y座标
    '                        例:粘度特征曲线;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;VIS(mPa.s),25,20~SHR(1/S),195,1

    '                     3) 血沉曲线:图像名称;图像画法;座标数据;描点数据;坐标轴标题数据
    '                                   其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
    '                                   描点数据:血沉值1,血沉值2,....血沉值30
    '                                   坐标轴标题数据:Y坐标标题文字,X坐标,Y座标~X坐标标题文字,X坐标,Y座标
    '                        例:血沉曲线;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;.5,.5,1,1,1,1.5,1.5,2,2,2,2.5,3,3,3.5,4,4.5,5.5,6.5,8,9,10.5,11.5,12.5,13.5,14.5,15.5,16.5,18,19,20;血沉值(mm),5,36~时间(m),55,1
    '                     4) PLT图：图像名称;图像画法;座标数据;描点数据
    '                               其中 坐标数据：Y长度,X长度,X座标-X座标显示的数字,....[|Y座标-Y座标显示的文字,.....]
    '                                    描点数据: Y1,Y2,Y3,......|Y1,Y2,Y3,......[~Y轴标题文字,X座标,Y座标|X轴标题文字,X座标,Y座标]
    '                        例:PLT;4;200,262;0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,4,4,7,7,12,12,17,17,20,20,25,25,30,30,33,33,36,36,41,41,43,43,44,44,46,46,47,47,47,47,47,47,46,46,46,46,44,44,44,44,43,43,41,41,39,39,38,38,36,36,35,35,33,33,31,31,30,30,28,28,27,27,25,25,23,23,22,22,22,22,20,20,19,19,17,17,15,15,15,15,14,14,12,12,12,12,11,11,11,11,9,9,9,9,9,9,7,7,7,7,7,7,6,6,6,6,6,6,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,3,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,7,7,9,9,8,8,9,9,12,12,16,16,22,22,26,26,30,30,35,35,36,36,37,37,39,39,42,42,44,44,46,46,46,46,44,44,43,43,40,40,37,37,37,37,37,37,39,39,37,37,36,36,32,32,29,29,25,25,23,23,22,22,22,22,21,21,19,19,18,18,16,16,16,16,15,15,15,15,15,15,14,14,12,12,11,11,9,9,9,9,8,8,8,8,7,7,7,7,7,7,7,7,7,7,8,8,7,7,7,7,5,5,4,4,4,4,2,2,4,4,4,4,2,2,2,2,4,4
    
    '                     5) 直方图 (带界标支持双曲线，X,Y座标刻度值)：标题;图像类型;Y高度,X长度;上下左右边框留白（用于画刻度）;X轴刻度[|Y刻度];曲线1数据[|曲线2数据...][;界标数据]
    '                                其中:曲线数据: 是y座标数据,以,分隔,多条曲线数据以|分隔
    '                                    :界标数据: 是x座标数据,以,号分隔
    '
    '                        例：   RBC;5;260,310;10,50,50,10;0-0,50-50,100-100,150-150,200-200,250-250,300-fL|50-50,100-100,150-150,200-200;
    '                               000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,001,001,001,001,002,002,001,001,001,002,002,002,003,004,005,006,008,011,014,018,022,030,038,048,058,072,089,107,124,145,162,180,196,208,221,230,239,247,248,251,255,247,246,233,229,221,204,199,188,180,169,156,150,141,130,125,116,111,104,097,093,088,085,079,074,071,067,063,061,059,056,054,053,051,048,045,044,040,038,037,034,033,031
    '                               ,030,028,026,025,022,020,019,018,017,016,015,015,013,013,012,011,011,011,011,010,010,009,009,008,008,008,008,008,008,008,007,008,008,007,008,007,007,007,007,007,007,007,007,006,006,006,005,005,005,005,005,005,005,005,005,004,004,005,004,004,004,004,004,004,003,003,003,003,003,002,002,002,002,002,002,002,002,002,002,002,002,002,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,001,001,001,001,001,000,000,000,000,000,000,000,000;55,90
    '                     6) 直方图 (同一个图形上画出多条曲线)：标题;图像类型;坐标数据;~(第一条曲线)Y1;Y2;Y3;Y4;Y5...~(第二条曲线)Y1;Y2;Y3;Y4;Y5..
    '                                其中 坐标数据：纵轴高度,横轴长度,刻度1-显示值,刻度2-显示值,...
    '                                     曲线区分：每条曲线数据以 '~' 号开始以便区分不同曲线
    '                        例：WBC;6;0,80;~0;0;0;0;1;1;2;3;4;6;9;13;18;23;27;31;32;30;28;24;21;17;13;11;8;6;5;4;3;2;1;1;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0~0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;1;2;2;3;4;5;5;6;6;6;6;5;4;4;3;2;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0~0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;1;1;2;2;3;5;7;9;13;17;21;25;28;31;34;37;38;38;38;37;35;33;31;29;26;24;22;19;16;13;11;8;6;4;3;2;1;1;1;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0
    '
    '                   100) 图片数据:图像名称;图像画法;[读取数据后是否删除];全路径
    '                        例:WBC Fsc;100;1;C:\tempfile.gif
    '
    '                   101-227) 图片数据:图像名称;图像画法;[读取数据后是否删除];全路径
    '                            现在支持三种图形格式,BMP,JPG,GIF
    '                            BMP格式图片采用 100-107编号
    '                            JPG格式图片采用 110-117
    '                            GIF格式图片采用 120-127
    '
    '                            '2开始的就是压缩过的ZIP图形
    '                            BMP格式图片采用 200-207编号
    '                            JPG格式图片采用 210-217
    '                            GIF格式图片采用 220-227
    '
    '                            ??1-??7是图片数据的补充格式，用于显示图片时的对齐方式设置。
    '                            用于指定Chart控件的 .ChartArea.Interior.Image.Layout 属性
    '                            101= oc2dImageCentered 102=oc2dImageTiled 103=oc2dImageFitted 104=oc2dImageStretched
    '                            105=oc2dImageStretchedToWidth 106=oc2dImageStretchedToHeight 107=oc2dImageCropFitted
   
'    GetAnswerCmd        函数  以二进制串方式返回自动应答指令如要固定应答06则写为,06
  '
Public Sub WriteLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strfilename As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject
    
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Val(GetSetting("ZLSOFT", "公共模块\ZlLISSrv", "清空接收日志", 1)) = 1 Then
        If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    End If
    strfilename = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strfilename) Then Call objFileSystem.CreateTextFile(strfilename)
    Set objStream = objFileSystem.OpenTextFile(strfilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "≡"))
    objStream.WriteLine ("执行时间:" & strDate & "版本:" & App.major & "." & App.minor & "." & App.Revision)
    objStream.WriteLine ("驱动:" & strFunc)
    objStream.WriteLine ("  :" & StrInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

Public Function GetStr_Section(ByVal strSource As String, ByVal strStart As String, ByVal strEnd As String) As String
    '功能：取两个字符之间的内容返回,开始字符和结束字符可以相同
    'strSource: 源字符串
    'strStart : 开始字符
    'strEnd   ：结束字符
    '
    Dim lngLength As Long, strTmp As String, strTmpStart As String, i As Integer
    
    If strStart <> strEnd Then
        lngLength = InStr(strSource, strEnd) - InStr(strSource, strStart) + 1
    Else
        For i = -22350 To -22310
            strTmpStart = Chr(i)
            If InStr(strSource, strTmpStart) <= 0 And strStart <> strTmpStart Then
                Exit For
            End If
        Next
        strTmp = Mid(strSource, 1, InStr(strSource, strStart) - 1) & strTmpStart & Mid(strSource, InStr(strSource, strStart) + 1)
        lngLength = InStr(strTmp, strEnd) - InStr(strTmp, strTmpStart) + 1
    End If
    
    If lngLength < 0 Then
        GetStr_Section = Mid(strSource, InStr(strSource, strStart) + lngLength, Abs(lngLength))
    Else
        GetStr_Section = Mid(strSource, InStr(strSource, strStart), lngLength)
    End If
End Function
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Mid_bin(ByVal str_bin As String, lng_S As Long, Optional lng_len As Long = 0, Optional blnChar As Boolean = True) As String
    '十六进制串的MID函数
    'str_Bin :传入的二进制数据，格式为,FF,AA,03 以，号开始结束无，号
    'lng_S   :开始位置
    'lng_Len :取的长度
    'blnChar :是否转换为字符格式返回
    
    Dim varBin As Variant
    Dim lng_Loop As Long
    Dim str_Return As String

    If lng_len < 0 Then Exit Function
    If lng_S <= 0 Then Exit Function
    
    varBin = Split(str_bin, ",")
    
    If lng_S + lng_len - 1 > UBound(varBin) Then
        '传入的串没得这么长
        Mid_bin = ""
        Exit Function
    End If

    If lng_len = 0 Then
        If blnChar Then
            For lng_Loop = lng_S To UBound(varBin)
                str_Return = str_Return & Chr("&H" & varBin(lng_Loop))
            Next
        Else
            str_Return = Mid(str_bin, lng_S * 3 - 2)
        End If

    Else
        If blnChar Then
            For lng_Loop = lng_S To lng_S + lng_len - 1
                str_Return = str_Return & Chr("&H" & varBin(lng_Loop))
            Next
        Else
            str_Return = Mid(str_bin, lng_S * 3 - 2, lng_len * 3)
        End If

    End If
    If str_Return <> "" Then Mid_bin = str_Return
    
End Function

Public Function Len_Bin(ByVal str_bin As String) As Long
    '十六进制串的Len函数
'    Dim varBin As Variant
'    varBin = Split(str_bin, ",")
'    Len_Bin = UBound(varBin)
    Len_Bin = Len(str_bin) / 3
End Function

Public Function Instr_Bin(ByVal str_bin As String, ByVal strChar As String, Optional ByVal lngStart As Long) As Long
    '十六进制的 Instr函数
    Dim varBin As Variant
    Dim strFindChar As String
    Dim lngS As Long
    Dim i As Integer
    Dim strHex As String
    If Len(strChar) <= 0 Then Exit Function
    strFindChar = ""
    For i = 1 To Len(strChar)
        strHex = Hex(Asc(Mid(strChar, i, 1)))
        strFindChar = strFindChar & "," & IIf(Len(strHex) = 1, "0" & strHex, strHex)
    Next
    If lngStart > 0 Then
        lngS = InStr(lngStart + 2, str_bin, strFindChar)
    Else
        lngS = InStr(str_bin, strFindChar)
    End If
    If lngS > 0 Then
        lngS = lngS + 2
        strFindChar = Mid(str_bin, 1, lngS)
'        len(strfindchar)/3
'        varBin = Len(strFindChar) / 3 'Split(strFindChar, ",")
        Instr_Bin = Len(strFindChar) / 3
    End If
End Function

Public Function Replace_Bin(ByVal str_bin As String, ByVal strFind As String, ByVal strReplace As String) As String
    '十六进制的　Replace
    Dim strFindBin As String
    Dim strReplaceBin As String
    Dim i As Long
    If str_bin = "" Then Exit Function
    
    If Len(strFindBin) <= 0 Then
        Replace_Bin = str_bin
        Exit Function
    End If
    For i = 1 To Len(strFind)
        strFindBin = strFindBin & "," & Asc(Mid(strFind, i, 1))
    Next
        
    If strReplace <> "" Then
        For i = 1 To Len(strReplace)
            strReplaceBin = strReplaceBin & "," & Asc(Mid(strReplace, i, 1))
        Next
    Else
        strReplaceBin = ""
    End If
    
    Replace_Bin = Replace(str_bin, strFindBin, strReplaceBin)
    
    
End Function


Public Sub Pause(ByVal PauseTime)
    '延时,单位秒
    Dim Start As Currency
    Start = Timer   ' 设置开始暂停的时刻。
    Do While Timer < Start + PauseTime
       DoEvents   ' 将控制让给其他程序。
    Loop
    
End Sub

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
Dim ptOrigin    As PointAPI     '文字绘制原点
Dim ptCenter    As PointAPI     '文字中心点.
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
                        .x = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    '转换 CenterX, CenterY 到点结构
                    '(需要调用 RotatePoint).
                    With ptCenter
                        .x = CenterX
                        .Y = CenterY
                    End With
                    
                    '以原点选择以匹配预期选择
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    '现在打印旋转文本并返回成功/失败
                    PrintRotText = (TextOut(hDC, ptOrigin.x, _
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
'**函 数 名：RotatePoint
'**输    入：ptAxis(PointAPI)   -
'**        ：ptRotate(PointAPI) -
'**        ：fDegrees(Single)   -
'**输    出：无
'**功能描述：从前fdegrees当前坐标选择ptRotate左右的ptAxis
'**全局变量：
'**调用模块：
'*************************************************************************
Private Sub RotatePoint(ptAxis As PointAPI, ptRotate As PointAPI, fDegrees As Single)

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
    fDX = ptRotate.x - ptAxis.x
    fDY = ptRotate.Y - ptAxis.Y
    
    '旋转点
    ptRotate.x = ptAxis.x + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub




Public Function ReadIni(strItem As String, strKey As String, strPath As String, Optional strDefault As String = "") As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, strDefault, GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function
Public Function DelSapce(strLine As String) As String
    '功能       删除多余的空格
    Dim intLoop  As Integer
    Dim strNow As String
    strNow = strLine
    For intLoop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intLoop), Space(1))
    Next
    DelSapce = strNow
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
108       lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap)
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
194     WriteLog "mdlPublic.SavePic", CStr(Erl()) & "行", Err.Description
End Sub


'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath) <> "" Then
        zlFileUnzip = strZipPath & Dir(strZipPath)
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String, ByVal strfilename As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function


Public Function GetIniKeyValue(ByVal strPathAndFileName As String, ByVal strItem As String, ByVal strKey As String, Optional ByVal strDefault As String) As String
        '读取Ini文件中的值
        '配置文件不存在则创建，并写入默认值
        Dim objFile As New FileSystemObject
        On Error GoTo hErr
100     If Not objFile.FileExists(strPathAndFileName) Then
102         Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
104         GetIniKeyValue = strDefault
        Else
106         GetIniKeyValue = ReadIni(strItem, strKey, strPathAndFileName)
            If GetIniKeyValue = "" And strDefault <> "" Then
                Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
                GetIniKeyValue = strDefault
            End If
        End If
        Exit Function
hErr:
108     WriteLog "获取" & strPathAndFileName & "中的设置," & CStr(Erl()) & "行, " & Err.Description, "", ""
End Function

