Attribute VB_Name = "mdlLedVoice"
Option Explicit '要求变量声音

Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gobjLED As Object
'-----------------------------------------------------------------------------------------------------------------------
'SYC XII
Public Declare Function dsbdllNt Lib "7CADSBNT.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public Declare Function dsbdll98 Lib "7CACKY95.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public Declare Function dsbdll16 Lib "7CACKY16.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'SYC Q9
Public Declare Function SYC_Q9 Lib "CKY95H.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'SHY-II
Public Declare Function shydsbdllNt Lib "CKY32.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'TDKJ_BJ_I/II
Public Declare Function TDKJ_BJ_FUN Lib "CKY95H.DLL" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer

'TDKJ_BJ_2008版
Public Declare Function TDKJ_BJ_2008 Lib "TdBjq.dll" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer


'Dev_MDT_SD01 麦迪特SD-01语音显示器

Public Declare Function InitService Lib "Service.dll" () As Boolean '启动服务函数，用来初始化驱动接口，必需在其他函数调用以前调用
Public Declare Function InitDevice Lib "Service.dll" (ByVal lPort As Long) As Boolean '初始化串口函数，在调用"启动服务"函数以后、调用其他函熟以前使用该函数。入口参数port：1-4。
Public Declare Function CloseService Lib "Service.dll" () As Boolean '关闭服务函数，用来关闭驱动接口，在所有的调用结束后调用。
Public Declare Function CloseDevice Lib "Service.dll" () As Boolean '关闭串口函数
Public Declare Function Clear_Screen Lib "Service.dll" () As Boolean '清屏
Private Declare Function Display Lib "Service.dll" (ByVal lStrNum As Long) As Boolean
                                 '与类模块重名,所以定义为私有,重新包装,显示函数，用来显示预置的汉字点阵。
Public Declare Function Voices Lib "Service.dll" (ByVal sCommand As String) As Boolean
                                 '播音函数，用来播放预先录制的声音序列。入口参数Command：代表声音的字符串。每两位代表一段声音
Public Declare Function Price Lib "Service.dll" (ByVal sMoney As String) As Boolean
                                  '应收函数，用来播放并显示应收的金额。入口参数Money：代表金额的字符串。
Public Declare Function GetPrice Lib "Service.dll" Alias "Get" (ByVal sMoney As String) As Boolean
                                 '实收函数，用来播放并显示应收的金额。入口参数Money：代表金额的字符串。
Public Declare Function Check Lib "Service.dll" (ByVal sMoney As String) As Boolean
                                 '找零函数，用来播放并显示应收的金额。入口参数Money：代表金额的字符串。
Public Declare Function Medincine Lib "Service.dll" (ByVal sNumber As String) As Boolean
                                  '取药函数, 显示并播放人员到某窗口取药。入口参数Number代表窗口号,必须为数字
Public Declare Function Display_Line Lib "Service.dll" (ByVal sTest As String, ByVal lSize As Long, ByVal lRow As Long) As Boolean
                                  '行显示函数，用来在指定行显示指定字体和大小的汉字。sTest代表需要显示的汉字；lSize=代表文字号(0-6),lRow=行号(0-3)

'SURPASS处理玉溪市人民医院门诊收费室语音及LED(金之裕)
Public Declare Function SetComNo Lib "Fgc01" (ByVal No As Long) As Long
Public Declare Sub SetQuickSwitch Lib "Fgc01" (ByVal Switch As Long)
Public Declare Sub SetHandleType Lib "Fgc01" (ByVal Handle As Long)
Public Declare Sub AllClear Lib "Fgc01" ()
Public Declare Sub PartClear Lib "Fgc01" (ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Sub AllDisplay Lib "Fgc01" (ByVal Handle As Long)
Public Declare Sub PartDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Sub StringDisplay Lib "Fgc01" (ByVal Str As String, ByVal Mode As Long, ByVal Delay As Long)
Public Declare Sub SetFontName Lib "Fgc01" (ByVal Name As String)
Public Declare Sub SetFontSize Lib "Fgc01" (ByVal Size As Long)
Public Declare Sub SetFontStyle Lib "Fgc01" (ByVal Style As Long)
Public Declare Sub LocStringDisplay Lib "Fgc01" (ByVal X As Long, ByVal Y As Long, ByVal Str As String)
Public Declare Sub PictureDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal Length As Long, ByVal Mode As Long, ByVal Delay As Long)
Public Declare Sub MagicDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal Mode As Long)
Public Declare Sub MagicClear Lib "Fgc01" (ByVal Mode As Long)
Public Declare Sub PickDisplay Lib "Fgc01" (ByVal Handle As Long, ByVal X0 As Long, ByVal Y0 As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long)
Public Declare Function PlayWaves Lib "Fgc01" (ByVal filenames As String) As Long
Public Declare Function RMB2Wav Lib "Fgc01" (ByVal VALDGT As Double) As Boolean
Public Declare Function Val2RMB Lib "Fgc01" (ByVal VALDGT As Double) As String
Public Declare Sub ClearWaves Lib "Fgc01" ()

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'**Begin********富顺FS-YL01型LED点阵+数码显示器(辽宁省鞍山市)***************
'**********2009-05-25 ZHQ                               ***************
Public Declare Sub opencomm Lib "datasend" (ByVal port As Integer)
'extern "C" __declspec(dllimport) int __stdcall opencomm(int port)
'功能：根据用户需要打开相应的串口。如果要进行数据传数就要先打开串口否则无法传送数据。如opencomm(1)就是打开串口 1

Public Declare Sub SendPray Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendPray(double money)
'功能：发送应收金额（即收费员所要收取的金额）。
'要求：金额范围为0.00-9999999元

Public Declare Sub SendYs Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendYs(double money)
'功能: 发送实收现金
'要求：金额范围为0.00-9999999元

Public Declare Sub SendChange Lib "datasend" (ByVal money As Double)
'extern "C" __declspec(dllimport) int __stdcall SendChange(double money)
'功能: 发送找零金额
'要求：金额范围为0.00-9999999元

Public Declare Sub SendName Lib "datasend" (ByVal bufName As String, ByVal Length As Integer)
'extern "C" __declspec(dllimport) int __stdcall SendName(unsigned char *buf,int Length)
'功能: 发送姓名
'要求: 姓名的长度不能超过24个字节

'Public Declare Sub SendCard Lib "datasend" (ByVal Handle As Long)
'extern "C" __declspec(dllimport) int __stdcall sendcard(void)
'功能: 发送提示 "出示诊疗卡"    (经咨询对方工程师，本设备不支持，需要取消)
'要求：要提示音时，直接调用发送

'Public Declare Sub SendWid Lib "datasend" (ByVal wid As Integer)
'extern "C" __declspec(dllimport) int __stdcall sendwid(int wid)
'功能: 发送窗口号               (经咨询对方工程师，由于设备仅支持数字,所以发药窗口不调用此程序)
'要求：提示请到XX号窗口取药，参数要整型，但值范围1~99

'**End**********富顺FS-YL01型LED点阵+数码显示器(辽宁省鞍山市)***************

'2010-02-24 ZHQ 一汽总医院 TDKJ_BJ_IV语音报价器
Public Declare Function TDKJ_BJ_IV Lib "TdBjq.dll" Alias "dsbdll" (ByVal port As Integer, ByVal OutString As String) As Integer
Public gBlnPic As Boolean   '记录是否首次初始化

'-----------------------------------------------------------------------------------------------------------------------
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128           '为了 PSS 的用途而维护串
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public gintOS As Integer                   '操作系统  0：Win32；1:Win98/95；2：Win2000/NT
'-----------------------------------------------------------------------------------------------------------------------
Public Enum LEDDevice
    Dev_SYC_XII = 1      'SYC XII 语音显示器
    Dev_LK822 = 2        'LK822 汉字液晶显示终端
    Dev_SHY_II = 3       'SHY-II 数码语言报价器
    Dev_NJF_VH = 4       'NJF-VH 智能语音显示器
    Dev_TDKJ_BJ = 5      'TDKJ_BJ_I/II语音报价器
    Dev_MDT_SD01 = 6     '麦迪特SD-01语音显示器
    Dev_surpass = 7      'SURPASS智能语音报价显示器
    Dev_FS_YL01 = 8      'FS-YL01型LED点阵+数码显示器
    Dev_TDKJ_BJ_2008 = 9 'TDKJ-BJ_2008版 语音报价器
    Dev_TDKJ_BJ_IV = 10  'TDKJ_BJ_IV 语音报价器
    Dev_SYC_Q9 = 11      'SYC-Q9语音报价器
    Dev_DDisplay = 99    '双屏显示器
End Enum

Public ctlComm As Object                   'MsComm控件
Public gintDevice As LEDDevice             '设备号
Public gintPort As Long                    '端口      1：COM1;2：COM2;3：COM3;4：COM4
Public gstrSpeed As String                 '传输速率
Public gblnDDisplay As Boolean             '使用双屏显示器

Public gblnHaveBottom As Integer           '显示底行的标志 1：显示，0：不显示
Public gstrBottom As String                '底行显示的内容
'曾明春（2005-10-13）
Public gblnNewDev As Boolean               '是否使用新型SHY-II型设备
Public gbln个帐余额 As Boolean             '是否对病人的个人帐户余额进行语音提示

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function
Public Sub Dev_MDT_SD01_Display(ByVal strShow As Long)
    'display 显示预定义文字
    '11  您好!，第一行    '12  请稍等，第二行    '13  请当面点清，第一行    '18  祝你早日康复，第二行
    Display (strShow)
End Sub

Public Sub Dev_MDT_SD01_Speak(ByVal strSpeak As String)
'功能：处理Dev_MDT_SD01语音报价系统的语音命令,声音播放，同时LED显示
'参数：strSpeak=公共指令格式，需要转换为该设备支持的命令
'说明：
    'Display_Line(char *String,int m,int n) 在指定行显示指定字体和大小的汉字
    'm字号               n行号
    '0   16X16   宋体    4(0-3)
    '1   24X24   仿宋体  2(0-1)
    '2   24X24   黑体    2(0-1)
    '3   24X24   楷体    2(0-1)
    '4   24X24   宋体    2(0-1)
    '5   40X40   宋体    1(0)
    '6   48X48   宋体    1(0)
    '24*24以下每行最多10个汉字,40*40以上每行最多5个汉字
    
    'Voices 发音指令  内置32段声音
    '01  您好        '02  欢迎光临    '04  请稍等      '08  祝你早日康复
    '10  谢谢        '20  应收        '40  实收        '80  找零
    '03  请当面点清  '06  再见        '0c  叮咚        '18  请到
    '30  中药房      '60  西药房      'C0  号窗口取药  '81  人民币
    '07 0e 1c 38 70 E0 C1 83 0f 1e 3c 78 F0 E1 C3 87
    '元 角 分 千 百 十 0  1  2  3  4  5  6  7  8  9
    
    Dim strMoney As String
    
    On Error Resume Next
    Clear_Screen '需要先清屏,否则不能完全覆盖历史显示
    
    Select Case strSpeak
    Case "#0"
        Display_Line "请输入密码", 5, 0 ''无此功能的发音,仅用文字显示
    Case "#1"  '--您好,请稍等"
        Display 11
        Display 12
        Voices "0104"
    Case "#2"
        Display_Line "谢谢", 5, 0
        Voices 10
    Case "#3"  '--请当面点清, 谢谢!"
        Display 13
        Display_Line "谢谢", 5, 0
        Voices "0310"
    Case "#4"
            Display_Line "请问您的姓名", 4, 0
    Case "#5"
            Display_Line "请您出示磁卡", 4, 0
    Case "#6"
            Display_Line "请您到中药房批价", 4, 0
    Case "#7"
            Display_Line "请您到X光室批价", 4, 0
    Case "#8"
            Display_Line "请您到注射室做皮试", 4, 0
    Case "#9"
            Display_Line "请您到门诊办公室", 4, 0
            Display_Line "审核盖章", 4, 1
    Case "#10"
            Display_Line "请您到挂号室", 4, 0
            Display_Line "输入门诊号", 4, 1
    Case "#11"
            Display_Line "请您出示身份证和", 4, 0
            Display_Line "医保凭证", 4, 1
    Case "#12"
            Display_Line "请您出示身份证和", 4, 0
            Display_Line "公费医疗凭证", 4, 1
    Case "#13"
            Display_Line "请您出示医保凭证和", 4, 0
            Display_Line "公费医疗凭证", 4, 1
    Case "#14"
            Display_Line "请问您挂什么科", 4, 0
    Case "#15"
            Display_Line "请问您是初诊还是复诊", 4, 0
    Case "#16"
            Display_Line "请问您挂专家门诊", 4, 0
            Display_Line "还是普通门诊?", 4, 1
    Case "#17"
            Display_Line "请您先预检", 4, 0
            Display_Line "然后再挂号", 4, 1
    Case "#18"
            Display_Line "请您先填好病历卡", 4, 0
    Case "#19"
            Display_Line "请您出示病历卡", 4, 0
    Case "#20"
            Display_Line "请您到B超室批价", 4, 0
    Case "#50"
            Display_Line "请您出示医保凭证", 4, 0
    Case Else
        '#21 1234.56   --请您付款一千二百三十四点五六元  J
        '#22 1234.56   --预收一千二百三十四点五六元 Y
        '#23 1234.56   --找零一千二百三十四点五六元 Z
    '07 0e 1c 38 70 E0 C1 83 0f 1e 3c 78 F0 E1 C3 87
    '元 角 分 千 百 十 0  1  2  3  4  5  6  7  8  9
        strMoney = Trim(Mid(strSpeak, 4))
        strSpeak = Left(strSpeak, 3)
        If strSpeak = "#21" Then '应收
            Price strMoney
        ElseIf strSpeak = "#22" Then '实收
            GetPrice strMoney
        ElseIf strSpeak = "#23" Then '找零
            Check strMoney
        End If
    End Select

End Sub

Public Sub Dev_TDKJ_BJ_2008_Speak(ByVal strSpeak As String)
'功能：处理TDKJ_BJ_2008语音报价系统的语音命令
'参数：strSpeak=公共指令格式，需要转换为该设备支持的命令
'      bytType-0-挂号,1-收费
'说明：根据配置文件中应用模式不同,相同命令的作用不同

    Dim strMoney As String
    Dim strInFor As String
    
    On Error Resume Next
    Select Case strSpeak
    Case "#0"  '--请输入密码
        '无此功能,用文字显示
        Call TDKJ_BJ_2008(gintPort, "&Sc$")
        Call TDKJ_BJ_2008(gintPort, "&C21请输入密码...$")
    Case "#1"  '--您好,请稍等"
        Call TDKJ_BJ_2008(gintPort, "W")
    Case "#2"  '--谢谢"
        Call TDKJ_BJ_2008(gintPort, "X")
    Case "#3"  '--请当面点清, 谢谢!"
        Call TDKJ_BJ_2008(gintPort, "D")
    Case "#4"  '--请问您的姓名"
        strInFor = Trim(GetSetting("ZLSOFT", "公共全局", "挂号提示", "请问您的姓名"))
        If strInFor = "请问您的姓名" Then
            Call TDKJ_BJ_2008(gintPort, "k")
        ElseIf strInFor = "请问您孩子的姓名" Then
            Call TDKJ_BJ_2008(gintPort, "w")
        Else
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & strInFor & "$")
        End If
    Case "#5"  '--请您出示磁卡"
        Call TDKJ_BJ_2008(gintPort, "b")
    Case "#6"  '--请您到中药房批价"
        Call TDKJ_BJ_2008(gintPort, "i")
    Case "#7"  '--请您到X光室批价"
        Call TDKJ_BJ_2008(gintPort, "j")
    Case "#8"  '--请您到注射室做皮试"
        Call TDKJ_BJ_2008(gintPort, "e")
    Case "#9"  '--请您到门诊办公室审核盖章"
        Call TDKJ_BJ_2008(gintPort, "g")
    Case "#10" '--请您到挂号室输入门诊号"
        Call TDKJ_BJ_2008(gintPort, "h")
    Case "#11" '--请您出示身份证和医保凭证"
        Call TDKJ_BJ_2008(gintPort, "d")
    Case "#12" '--请您出示身份证和公费医疗凭证"
        '--
    Case "#13" '--请您出示医保凭证和公费医疗凭证"
        '--
    Case "#14" '--请问您挂什么科"
        Call TDKJ_BJ_2008(gintPort, "e")
    Case "#15" '--请问您是初诊还是复诊"
        Call TDKJ_BJ_2008(gintPort, "g")
    Case "#16" '--请问您挂专家门诊还是普通门诊"
        Call TDKJ_BJ_2008(gintPort, "h")
    Case "#17" '--请您先预检, 然后再挂号"
        Call TDKJ_BJ_2008(gintPort, "i")
    Case "#18" '--请您先填好病历卡"
        
    Case "#19" '--请您出示病历卡"
        Call TDKJ_BJ_2008(gintPort, "j")
    Case "#20" '--请您到B超室批价"
        '--
    Case "#50"
        Call TDKJ_BJ_2008(gintPort, "a")
    Case "#51"  '门诊收费，请问你的姓名
        strInFor = Trim(GetSetting("ZLSOFT", "公共全局", "收费提示", "请问您的姓名"))
        If strInFor = "请出示您的挂号票" Then
            Call TDKJ_BJ_2008(gintPort, "x")
        ElseIf strInFor = "请问您的姓名" Then
            Call TDKJ_BJ_2008(gintPort, "k")
        ElseIf strInFor = "请问您孩子的姓名" Then
            Call TDKJ_BJ_2008(gintPort, "w")
        ElseIf strInFor = "不播报" Then
        Else
            Call TDKJ_BJ_2008(gintPort, "&Sc$")
            Call TDKJ_BJ_2008(gintPort, "&C21" & strInFor & "$")
        End If
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '请您付款
            Call TDKJ_BJ_2008(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then '预收
            Call TDKJ_BJ_2008(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '找零
            Call TDKJ_BJ_2008(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Dev_TDKJ_BJ_Speak(ByVal strSpeak As String)
'功能：处理TDKJ_BJ_I/II语音报价系统的语音命令
'参数：strSpeak=公共指令格式，需要转换为该设备支持的命令
'说明：根据配置文件中应用模式不同,相同命令的作用不同
    Dim strMoney As String
    
    On Error Resume Next
    
    Select Case strSpeak
    Case "#0"  '--请输入密码
        '无此功能,用文字显示
        Call TDKJ_BJ_FUN(gintPort, "&Sc$")
        Call TDKJ_BJ_FUN(gintPort, "&C21请输入密码...$")
    Case "#1"  '--您好,请稍等"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "W")
    Case "#2"  '--谢谢"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "X")
    Case "#3"  '--请当面点清, 谢谢!"
        Call TDKJ_BJ_FUN(gintPort, "D")
    Case "#4"  '--请问您的姓名"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "k")
    Case "#5"  '--请您出示磁卡"
        DoEvents
        Call TDKJ_BJ_FUN(gintPort, "b")
    Case "#6"  '--请您到中药房批价"
        Call TDKJ_BJ_FUN(gintPort, "i")
    Case "#7"  '--请您到X光室批价"
        Call TDKJ_BJ_FUN(gintPort, "j")
    Case "#8"  '--请您到注射室做皮试"
        Call TDKJ_BJ_FUN(gintPort, "e")
    Case "#9"  '--请您到门诊办公室审核盖章"
        Call TDKJ_BJ_FUN(gintPort, "g")
    Case "#10" '--请您到挂号室输入门诊号"
        Call TDKJ_BJ_FUN(gintPort, "h")
    Case "#11" '--请您出示身份证和医保凭证"
        Call TDKJ_BJ_FUN(gintPort, "d")
    Case "#12" '--请您出示身份证和公费医疗凭证"
        '--
    Case "#13" '--请您出示医保凭证和公费医疗凭证"
        '--
    Case "#14" '--请问您挂什么科"
        Call TDKJ_BJ_FUN(gintPort, "e")
    Case "#15" '--请问您是初诊还是复诊"
        Call TDKJ_BJ_FUN(gintPort, "g")
    Case "#16" '--请问您挂专家门诊还是普通门诊"
        Call TDKJ_BJ_FUN(gintPort, "h")
    Case "#17" '--请您先预检, 然后再挂号"
        Call TDKJ_BJ_FUN(gintPort, "i")
    Case "#18" '--请您先填好病历卡"
        
    Case "#19" '--请您出示病历卡"
        Call TDKJ_BJ_FUN(gintPort, "j")
    Case "#20" '--请您到B超室批价"
        '--
    Case "#50"
        Call TDKJ_BJ_FUN(gintPort, "a")
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '请您付款
            Call TDKJ_BJ_FUN(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then '预收
            Call TDKJ_BJ_FUN(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '找零
            Call TDKJ_BJ_FUN(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Dev_TDKJ_BJ_IV_Speak(ByVal strSpeak As String)
'功能：处理TDKJ_BJ_IV 语音报价系统的语音命令
'参数：strSpeak=公共指令格式，需要转换为该设备支持的命令
'说明：根据配置文件中应用模式不同,相同命令的作用不同
    Dim strMoney As String
    
    On Error Resume Next
    
    Select Case strSpeak
    Case "#0"  '--请输入密码
        '无此功能,用文字显示
        Call TDKJ_BJ_IV(gintPort, "&Sc$")
        Call TDKJ_BJ_IV(gintPort, "&C21请输入密码...$")
    Case "#1"  '--您好,请稍等"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "W")
    Case "#2"  '--谢谢"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "X")
    Case "#3"  '--请当面点清, 谢谢!"
        Call TDKJ_BJ_IV(gintPort, "D")
    Case "#4"  '--请问您的姓名"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "k")
    Case "#5"  '--请您出示磁卡"
        DoEvents
        Call TDKJ_BJ_IV(gintPort, "b")
    Case "#6"  '--请您到中药房批价"
        Call TDKJ_BJ_IV(gintPort, "i")
    Case "#7"  '--请您到X光室批价"
        Call TDKJ_BJ_IV(gintPort, "j")
    Case "#8"  '--请您到注射室做皮试"
        Call TDKJ_BJ_IV(gintPort, "e")
    Case "#9"  '--请您到门诊办公室审核盖章"
        Call TDKJ_BJ_IV(gintPort, "g")
    Case "#10" '--请您到挂号室输入门诊号"
        Call TDKJ_BJ_IV(gintPort, "h")
    Case "#11" '--请您出示身份证和医保凭证"
        Call TDKJ_BJ_IV(gintPort, "d")
    Case "#12" '--请您出示身份证和公费医疗凭证"
        '--
    Case "#13" '--请您出示医保凭证和公费医疗凭证"
        '--
    Case "#14" '--请问您挂什么科"
        Call TDKJ_BJ_IV(gintPort, "e")
    Case "#15" '--请问您是初诊还是复诊"
        Call TDKJ_BJ_IV(gintPort, "g")
    Case "#16" '--请问您挂专家门诊还是普通门诊"
        Call TDKJ_BJ_IV(gintPort, "h")
    Case "#17" '--请您先预检, 然后再挂号"
        Call TDKJ_BJ_IV(gintPort, "i")
    Case "#18" '--请您先填好病历卡"
        
    Case "#19" '--请您出示病历卡"
        Call TDKJ_BJ_IV(gintPort, "j")
    Case "#20" '--请您到B超室批价"
        '--
    Case "#50"
        Call TDKJ_BJ_IV(gintPort, "a")
    Case Else
        strMoney = Trim(Mid(strSpeak, 4))
        If Left(strSpeak, 3) = "#21" Then '请您付款
            Call TDKJ_BJ_IV(gintPort, strMoney & "J")
        ElseIf Left(strSpeak, 3) = "#22" Then '预收
            Call TDKJ_BJ_IV(gintPort, strMoney & "Y")
        ElseIf Left(strSpeak, 3) = "#23" Then '找零
            Call TDKJ_BJ_IV(gintPort, strMoney & "Z")
        End If
    End Select
End Sub

Public Sub Contrast_NJF_VH(ByVal strCommand As String)
'功能：处理NJF-VH 智能语音显示器
'参数：strCommand=SYC-X的语音命令(最初始程序是这样传的),需要转化为NJF-VH的命令
'说明："#序号"或"#序号 金额"
    Dim intNum As Integer, strMoney As String
    Dim strDisp As String, strVoice As String
        
    If InStr(strCommand, " ") > 0 Then
        strCommand = Replace(strCommand, "  ", " ")
        strCommand = Replace(strCommand, "  ", " ")
        intNum = Val(Mid(Split(strCommand, " ")(0), 2))
        strMoney = Split(strCommand, " ")(1)
    Else
        intNum = Val(Mid(strCommand, 2))
    End If
    
    On Error Resume Next
    Set gobjLED = CreateObject("CTSVR.Bjq")
    If Not gobjLED Is Nothing Then
        Select Case intNum
            Case 1
                strDisp = "~您好,请稍等!"
                strVoice = "_H"
            Case 2
                strDisp = "~谢谢!"
                strVoice = "_T"
            Case 3
                strDisp = "~找零请当面点清,谢谢!"
                strVoice = "_C"
            Case 4
                strDisp = "~请问您的姓名?"
                strVoice = "eY"
            Case 5
                strDisp = "~请您出示磁卡!"
                strVoice = "gY"
            Case 6
                strDisp = "~请您到中药房批价!"
                strVoice = "bX"
            Case 7
                strDisp = "~请您到X光室批价!"
                strVoice = "eX"
            Case 8
                strDisp = "~请您到注射室做皮试!"
                strVoice = ""
            Case 9
                strDisp = "~请您到门诊办公室审核盖章!"
                strVoice = "pY"
            Case 10
                strDisp = "~请您到挂号室输入门诊号!"
                strVoice = "aX"
            Case 11
                strDisp = "~请您出示身份证和医保凭证!"
                strVoice = "hY" '请您出示身份证
            Case 12
                strDisp = "~请您出示身份证和公费医疗凭证!"
                strVoice = "hY" '请您出示身份证
            Case 13
                strDisp = "~请您出示医保凭证和公费医疗凭证!"
                strVoice = "hY" '请您出示身份证
            Case 14
                strDisp = "~请问您挂什么科?"
                strVoice = "bY"
            Case 15
                strDisp = "~请问您是初诊还是复诊?"
                strVoice = "cY"
            Case 16
                strDisp = "~请问您挂专家门诊还是普通门诊?"
                strVoice = "dY"
            Case 17
                strDisp = "~请您先预诊，然后再挂号!"
                strVoice = "mY"
            Case 18
                strDisp = "~请您先填好病历卡!"
                strVoice = "lY"
            Case 19
                strDisp = "~请您出示病历卡!"
                strVoice = "jY"
            Case 20
                strDisp = "~请您到B超室批价!"
                strVoice = "gX"
            Case 21
                strDisp = "~请您付款:" & strMoney & "元"
                strVoice = strMoney & "_P"
            Case 22
                strDisp = "~收您:" & strMoney & "元"
                strDisp = "" '已调用其它显示
                strVoice = strMoney & "_k"
            Case 23
                If Val(strMoney) <> 0 Then
                    strDisp = "~找您:" & strMoney & "元"
                    strDisp = "" '已调用其它显示
                    strVoice = strMoney & "_b"
                Else
                    strDisp = ""
                    strVoice = ""
                End If
            Case Else
        End Select
        If strDisp <> "" Or strVoice <> "" Then
            gobjLED.Comport = gintPort
            gobjLED.DispMode = 0
            If strDisp <> "" Then gobjLED.Display strDisp
            If strVoice <> "" Then gobjLED.stdSpeak strVoice
        End If
        Set gobjLED = Nothing
    End If
End Sub
Public Sub ContrastSYC_Q9(ByVal strCommand As String)
    '功能：与SYC-XII语音显示器的命令对应起来
    '参数：strcommand 实际的命令
    '作者：冉勇
    Dim strFront As String, strLast As String, strMoney As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SycVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    
    Select Case strFront
        Case "#0"  '--请输入密码
            '无此功能,用文字显示
            Call SYC_Q9(gintPort, "*")
            Call SYC_Q9(gintPort, "#请输入密码...#")
        Case "#1"   ': W --您好, 请稍等
            Call SYC_Q9(gintPort, "*")
            Call SYC_Q9(gintPort, "W")
        Case "#2"   ':X  --谢谢
            Call SYC_Q9(gintPort, "X")
        Case "#3"   'D  --请当面点清, 谢谢!
            Call SYC_Q9(gintPort, "D")
        Case "#4"   'j  --请问您的姓名
            Call SYC_Q9(gintPort, "j")
        Case "#5"   'b  --请您出示磁卡
            Call SYC_Q9(gintPort, "b")
        Case "#6"   'h  --请您到中药房批价
            Call SYC_Q9(gintPort, "h")
        Case "#7"   'i  --请您到X光室批价
            Call SYC_Q9(gintPort, "i")
        Case "#8"   'e  --请您到注射室做皮试
            Call SYC_Q9(gintPort, "e")
        Case "#9"   'g  --请您到门诊办公室审核盖章
            Call SYC_Q9(gintPort, "g")
        Case "#10"  'h --请您到挂号室输入门诊号
            Call SYC_Q9(gintPort, "h")
        Case "#11"  '# c --请您出示身份证和医保凭证
             Call SYC_Q9(gintPort, "c")
        Case "#12"  'c --请您出示身份证和公费医疗凭证
            Call SYC_Q9(gintPort, "c")
        Case "#13"  'k --请您出示医保凭证和公费医疗凭证
            Call SYC_Q9(gintPort, "k")
        Case "#14"  'l --请问您挂什么科
            Call SYC_Q9(gintPort, "l")
        Case "#15"  'm --请问您是初诊还是复诊
            Call SYC_Q9(gintPort, "m")
        Case "#16"  'n --请问您挂专家门诊还是普通门诊
             Call SYC_Q9(gintPort, "n")
        Case "#17"  'o --请您先预检, 然后再挂号
            Call SYC_Q9(gintPort, "o")
        Case "#18"  'p --请您先填好病历卡
            Call SYC_Q9(gintPort, "p")
        Case "#19"  'q --请您出示病历卡
            Call SYC_Q9(gintPort, "q")
        Case "#20"  'q --请您到B超室批价
            Call SYC_Q9(gintPort, "q")
        Case "#21"  '1234.56   --请您付款一千二百三十四点五六元  J
            Call SYC_Q9(gintPort, strLast & "J")
        Case "#22"  '1234.56   --预收一千二百三十四点五六元 Y
            Call SYC_Q9(gintPort, strLast & "Y")
        Case "#23"  ' 1234.56   --找零一千二百三十四点五六元 Z
            Call SYC_Q9(gintPort, strLast & "Z")
        Case "#30"   '出示就诊卡
            Call SYC_Q9(gintPort, "#请出示就诊卡！#")
        Case "#50" '请出示社会保障卡
            Call SYC_Q9(gintPort, "a")
        Case "#51" '请问您的姓名
            Call SYC_Q9(gintPort, "j")
        Case Else
    End Select
End Sub
Public Sub ContrastSYC_XII(ByVal strCommand As String)
    '功能：与SYC-XII语音显示器的命令对应起来
    '参数：strcommand 实际的命令
        
    Dim strFront As String, strLast As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SycVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    
    Select Case strFront
        Case "#1"   ': W --您好, 请稍等
            SycVoice "W"
        Case "#2"   ':X  --谢谢
            SycVoice "X"
        Case "#3"   'D  --请当面点清, 谢谢!
            SycVoice "D"
        Case "#4"   'a  --请问您的姓名
            SycVoice "a"
        Case "#5"   'b  --请您出示磁卡
            SycVoice "b"
        Case "#6"   'c  --请您到中药房批价
            SycVoice "c"
        Case "#7"   'd  --请您到X光室批价
            SycVoice "d"
        Case "#8"   'e  --请您到注射室做皮试
            SycVoice "e"
        Case "#9"   'g  --请您到门诊办公室审核盖章
            SycVoice "g"
        Case "#10"  'h --请您到挂号室输入门诊号
            SycVoice "h"
        Case "#11"  '# i --请您出示身份证和医保凭证
            SycVoice "i"
        Case "#12"  'j --请您出示身份证和公费医疗凭证
            SycVoice "j"
        Case "#13"  'k --请您出示医保凭证和公费医疗凭证
            SycVoice "k"
        Case "#14"  'l --请问您挂什么科
            SycVoice "l"
        Case "#15"  'm --请问您是初诊还是复诊
            SycVoice "m"
        Case "#16"  'n --请问您挂专家门诊还是普通门诊
            SycVoice "n"
        Case "#17"  'o --请您先预检, 然后再挂号
            SycVoice "o"
        Case "#18"  'p --请您先填好病历卡
            SycVoice "p"
        Case "#19"  'q --请您出示病历卡
            SycVoice "q"
        Case "#20"  'r --请您到B超室批价
            SycVoice "r"
        Case "#21"  '1234.56   --请您付款一千二百三十四点五六元  J
            SycVoice strLast & "J"
        Case "#22"  '1234.56   --预收一千二百三十四点五六元 Y
            SycVoice strLast & "Y"
        Case "#23"  ' 1234.56   --找零一千二百三十四点五六元 Z
            SycVoice strLast & "Z"
        Case "#30"   '  请出示就诊卡(贵医要求:):32663
            SycVoice "p"
        Case Else
    End Select
End Sub

Public Sub ContrastSHY_II(ByVal strCommand As String)
    '功能：与SHY-II数码语言报价器的命令对应起来
    '参数：strcommand 实际的命令
        
    Dim strFront As String, strLast As String
    Dim intLocation As Integer
    
    If Left(strCommand, 1) = "~" Then
        SHYVoice Mid(strCommand, 2)
        Exit Sub
    End If
    intLocation = InStr(1, strCommand, " ")
    If intLocation <> 0 Then
        strFront = Left(strCommand, intLocation - 1)
        strLast = Trim(Mid(strCommand, intLocation + 1))
    Else
        strFront = strCommand
        strLast = ""
    End If
    '曾明春（2005-10-12） 部分功能只有新设备支持，所以对其区分判断
    If gblnNewDev Then '使用新型设备，支持医保功能
        Select Case strFront
            Case "#0" 'g  --请输入密码
                SHYVoice "g"
            Case "#1"   ': W --您好, 请稍等
                SHYVoice "W"
            Case "#2"   ':X  --谢谢
                SHYVoice "X"
            Case "#3"   'D  --请当面点清, 谢谢!
                SHYVoice "D"
            Case "#4"   'a  --请问您的姓名
                SHYVoice "a"
            Case "#5"   'b  --请您刷卡
                SHYVoice "b"
            Case "#6"   'b  --请您到中药房批价
                'SHYVoice "b"
            Case "#7"   'c  --请您到X光室批价
                'SHYVoice "c"
            Case "#8"   'd  --请您到注射室做皮试
                'SHYVoice "d"
            Case "#9"   'e  --请您到门诊办公室审核盖章
                'SHYVoice "e"
            Case "#10"  'f --请您到挂号室输入门诊号
                'SHYVoice "f"
            Case "#11"  '# i --请您出示身份证和医保凭证
                'SHYVoice "i"
            Case "#12"  'j --请您出示身份证和公费医疗凭证
                'SHYVoice "j"
            Case "#13"  'k --请您出示医保凭证和公费医疗凭证
                'SHYVoice "k"
            Case "#14"  'd --请问您挂什么科
                SHYVoice "d"
            Case "#15"  'h --请问您是初诊还是复诊
                SHYVoice "h"
            Case "#16"  'f --请问您挂专家门诊还是普通门诊
                SHYVoice "f"
            Case "#17"  'o --请您先预检, 然后再挂号
                'SHYVoice "o"
            Case "#18"  'p --请您先填好病历卡
                'SHYVoice "p"
            Case "#19"  'h --请您出示病历卡
                'SHYVoice "h"
            Case "#20"  'g --请您到B超室批价
                'SHYVoice "g"
            Case "#21"  '1234.56   --请您付款一千二百三十四点五六元  J
                SHYVoice strLast & "J"
            Case "#22"  '1234.56   --预收一千二百三十四点五六元 Y
                SHYVoice strLast & "Y"
            Case "#23"  ' 1234.56   --找零一千二百三十四点五六元 Z
                SHYVoice strLast & "Z"
            Case "#24"   'c --请你出示社保卡
                SHYVoice "c"
            Case "#25"  'e --你的费用为***元
                If gbln个帐余额 Then SHYVoice strLast & "e"
            Case "#26"  'f --卡上余额***元
                If gbln个帐余额 Then SHYVoice strLast & "f"
            Case "#27"  'i --你的卡上余额不足请付现金**元
                If gbln个帐余额 Then SHYVoice strLast & "i"
            Case "#28"  'X --请你先做医保身份鉴别
                SHYVoice "X"
            Case Else
            
        End Select
    Else
        Select Case strFront
            Case "#0" 'g  --请输入密码
                ' SHYVoice "g"
            Case "#1"   ': W --您好, 请稍等
                SHYVoice "W"
            Case "#2"   ':X  --谢谢
                SHYVoice "X"
            Case "#3"   'D  --请当面点清, 谢谢!
                SHYVoice "D"
            Case "#4"   'a  --请问您的姓名
                SHYVoice "a"
            Case "#5"   'b  --请您出示磁卡
                ' SHYVoice "b"
            Case "#6"   'b  --请您到中药房批价
                SHYVoice "b"
            Case "#7"   'c  --请您到X光室批价
                SHYVoice "c"
            Case "#8"   'd  --请您到注射室做皮试
                SHYVoice "d"
            Case "#9"   'e  --请您到门诊办公室审核盖章
                SHYVoice "e"
            Case "#10"  'f --请您到挂号室输入门诊号
                SHYVoice "f"
            Case "#11"  '# g --请您出示身份证和医保凭证
                SHYVoice "g"
            Case "#12"  'j --请您出示身份证和公费医疗凭证
                'SHYVoice "j"
            Case "#13"  'k --请您出示医保凭证和公费医疗凭证
                'SHYVoice "k"
            Case "#14"  'b --请问您挂什么科
                SHYVoice "b"
            Case "#15"  'c --请问您是初诊还是复诊
                SHYVoice "c"
            Case "#16"  'd --请问您挂专家门诊还是普通门诊
                SHYVoice "d"
            Case "#17"  'e --请您先预检, 然后再挂号
                SHYVoice "e"
            Case "#18"  'p --请您先填好病历卡
               ' SHYVoice "p"
            Case "#19"  'h --请您出示病历卡
                SHYVoice "h"
            Case "#20"  'g --请您到B超室批价
                ' SHYVoice "g"
            Case "#21"  '1234.56   --请您付款一千二百三十四点五六元  J
                SHYVoice strLast & "J"
            Case "#22"  '1234.56   --预收一千二百三十四点五六元 Y
                SHYVoice strLast & "Y"
            Case "#23"  ' 1234.56   --找零一千二百三十四点五六元 Z
                SHYVoice strLast & "Z"
            Case Else
            
        End Select
    End If
End Sub
    
Public Function SycVoice(ByVal OutString As String) As Integer
    '功能：SYC XII语音显示器的语音提示
    '参数：
    'outstring:命令 F,W,X,J,Y,Z,D必须大写
    'F --复位清零
    'W          --您好,请稍等
    'X  --谢谢
    '1234.56J   --请您付款一千二百三十四点五六元
    '1234.56Y   --预收一千二百三十四点五六元
    '1234.56Z   --找零一千二百三十四点五六元
    'D - -请当面点清, 谢谢!
    'a - -请问您的姓名
    'b - -请您出示磁卡
    'c - -请您到中药房批价
    'd - -请您到X光室批价
    'e - -请您到注射室做皮试
    'g - -请您到门诊办公室审核盖章
    'h - -请您到挂号室输入门诊号
    'i - -请您出示身份证和医保凭证
    'j - -请您出示身份证和公费医疗凭证
    'k - -请您出示医保凭证和公费医疗凭证
    'l - -请问您挂什么科
    'm - -请问您是初诊还是复诊
    'n - -请问您挂专家门诊还是普通门诊
    'o - -请您先预检, 然后再挂号
    'p - -请您先填好病历卡
    'q - -请您出示病历卡
    'r - -请您到B超室批价
    '！--换行
    '*--清屏
    '如果需要在第一行显示工号或其他汉字则:
    '$1  ($2为第二行)
    '#__今日当班挂号:1234#(__为空格)
    
    '返回：-1:不成功；其他：成功
        
    Dim intresult As Integer
'    On Error goto errhandel
    
    intresult = 0
    Select Case gintOS
        Case 0
            intresult = dsbdll16(gintPort, OutString)
        Case 1
            intresult = dsbdll98(gintPort, OutString)
        Case 2
            intresult = dsbdllNt(gintPort, OutString)
    End Select
    
    SycVoice = intresult
    Exit Function
errHandle:
    SycVoice = -1
End Function

Public Function SHYVoice(ByVal OutString As String) As Integer
    '功能：SYC XII语音显示器的语音提示
    '参数：
    'outstring:命令 F,W,X,J,Y,Z,D必须大写
    'F --复位清零
    'W          --您好,请稍等
    'X  --谢谢
    '1234.56J   --请您付款一千二百三十四点五六元
    '1234.56Y   --预收一千二百三十四点五六元
    '1234.56Z   --找零一千二百三十四点五六元
    'D - -请当面点清, 谢谢!
    'a - -请问您的姓名
    'b - -请您到中药房批价
    'c - -请您到X光室批价
    'd - -请您到注射室做皮试
    'e - -请您到门诊办公室审核盖章
    'f - -请您到挂号室输入门诊号
    'g - -请输入密码
    'i - -请您出示身份证和医保凭证
    'j - -请您出示身份证和公费医疗凭证
    'k - -请您出示医保凭证和公费医疗凭证
    'l - -请问您挂什么科
    'm - -请问您是初诊还是复诊
    'n - -请问您挂专家门诊还是普通门诊
    'o - -请您先预检, 然后再挂号
    'p - -请您先填好病历卡
    'q - -请您出示病历卡
    'r - -请您到B超室批价
    
    'DSBDLL(1,'a')  --请问您的姓名。
'    DSBDLL(1,'b')  --请您到中药房批价。
'    DSBDLL(1,'c')  --请您到X光室批价。
'    DSBDLL(1,'d')  --请您到注射室做皮试。
'    DSBDLL(1,'e')  --请您到门诊办公室审核盖章。
'    DSBDLL(1,'f')  --请您到挂号室输入门诊号。
'    DSBDLL(1,'g')  --请输入密码。
'    DSBDLL(1,'h')  --请您把病历卡拿出来。
'    DSBDLL(1,'i')  --找零请当面点清,谢谢。
    
    '返回：-1:不成功；其他：成功
        
    Dim intresult As Integer
    
    intresult = 0
    Select Case gintOS
        Case 0
            intresult = shydsbdllNt(gintPort, OutString)
        Case 1
            intresult = shydsbdllNt(gintPort, OutString)
        Case 2
            intresult = shydsbdllNt(gintPort, OutString)
    End Select
    
    SHYVoice = intresult
    Exit Function
errHandle:
    SHYVoice = -1
End Function


Public Sub Dev_surpass_speak(ByVal strSpeak As String)
'功能：处理Dev_surpass语音报价系统的语音命令
'参数：strSpeak=公共指令格式，需要转换为该设备支持的命令
'说明：根据配置文件中应用模式不同,相同命令的作用不同
    Dim filenames As String
    Dim strMoney As String
    Dim str应收 As String, str实收 As String, str找您 As String, str找零 As String, str请付款 As String
    Dim dbl合计 As Double
    On Error Resume Next

    str应收 = "应收.wav"
    str实收 = "预收.wav"
    str找零 = "找零.wav"
    str请付款 = "请您付款.wav"
    str找您 = "找零请当面点清谢谢.wav"
    
    Select Case strSpeak
           Case "#50"
                Call AllClear
                Call LocStringDisplay(2, 22, "您好，请稍等！")
                Call PlayWaves(App.Path & "\请出示医保卡.wav")
           Case "#1"
                Call AllClear
                Call PlayWaves(App.Path & "\请稍等.wav")
           Case Else
                strMoney = Trim(Mid(strSpeak, 4))
                If Left(strSpeak, 3) = "#21" Then '请您付款
                    Call AllClear
                    Call LocStringDisplay(2, 2, "应收：" & Format(strMoney, "0.00") & "元" + Chr(0))
                    str请付款 = "请您付款.wav"
                    Call PlayWaves(App.Path & "\" & str请付款)
                    Call RMB2Wav(strMoney)
                ElseIf Left(strSpeak, 3) = "#22" Then '预收
                    Call LocStringDisplay(2, 22, "预收：" & Format(strMoney, "0.00") & "元" + Chr(0))
                    Call PlayWaves(App.Path & "\" & str实收)
                    Call RMB2Wav(strMoney)
                ElseIf Left(strSpeak, 3) = "#23" Then '找零
                    If strMoney > 0 Then
                        Call LocStringDisplay(2, 42, "找零：" & Format(strMoney, "0.00") & "元" + Chr(0))
                        Call PlayWaves(App.Path & "\" & str找零)
                        Call RMB2Wav(strMoney)
                        Call PlayWaves(App.Path & "\" & str找您)
                    End If
                End If
                             
    End Select
End Sub

Public Sub Dev_FS_YL01_Voice(ByVal varTemp As Variant, ByVal intType As Byte, ByVal lngSec As Long)
'功能：处理Dev_FS_YL01语音报价系统的语音命令和屏显
'参数：varTemp  可能是字符(姓名),也能是数字(金额)
'      intType  传入类型(0-姓名;1-应收金额;2-实收金额;3-找补金额
'      lngSec   间隔时间,以秒为单位,为0表示不停顿
'说明：根据传入参数不同显示并发音,由于该设备是命令直接输出,
'      所以两条命令之间如果没有停顿,将造成之前没有说完的内容会被后面命令截断,所以两条语音命令之间需要作特殊处理.
'编制：2009-05-25 ZHQ

    Dim dtNow As Variant
    
    Select Case intType
    Case 0  '姓名
        Call SendName(varTemp, LenB(varTemp))
    Case 1  '应收金额
        Call SendPray(Round(varTemp, 2))
    Case 2  '实收金额
        Call SendYs(Round(varTemp, 2))
    Case 3  '找补金额
        Call SendChange(Round(varTemp, 2))
    End Select
    
    dtNow = Time
    Do While True
        If Int((Time - dtNow) * 24 * 60 * 60) >= lngSec Then Exit Do
    Loop
End Sub

Public Sub ShowLED(ByVal strRow1 As String, ByVal strRow2 As String, ByVal strRow3 As String, ByVal strRow4 As String)
'---------------------------------------------------------------------
'设计人:周海全
'   1999-8-22
'---------------------------------------------------------------------
'功能：根据传入的三个值将其显示在LED上
'参数：strRow1-strRow4:各行信息(30个字符,15个汉字)
'返回：
'---------------------------------------------------------------------
    On Error Resume Next
    With ctlComm
        .output = Chr(27) + "@"
        .output = Chr(27) + "CLR"
        
        .output = Chr(27) + "l" + Chr(1) + Chr(1)
        .output = strRow1
        .output = Chr(27) + "l" + Chr(1) + Chr(2)
        .output = strRow2
        .output = Chr(27) + "l" + Chr(1) + Chr(3)
        .output = strRow3
        If gblnHaveBottom = 1 Then
            .output = Chr(27) + "l" + Chr(1) + Chr(4)
            .output = strRow4
        End If
    End With
End Sub

Public Function SetLength(ByVal strText As String, ByVal lngLen As Long) As String
'功能：设置字符串的最大长度
'参数：lngLen=以字节为单位的最大长度
    Dim strTmp As String, i As Long
    
    If zlCommFun.ActualLen(strText) <= lngLen Then
        SetLength = strText
    Else
        For i = 1 To Len(strText)
            If zlCommFun.ActualLen(strTmp & Mid(strText, i, 1)) <= lngLen Then
                strTmp = strTmp & Mid(strText, i, 1)
            End If
        Next
        SetLength = strTmp
    End If
End Function
