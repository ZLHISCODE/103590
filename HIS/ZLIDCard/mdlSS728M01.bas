Attribute VB_Name = "mdlSS728M01"
Option Explicit
Private mblnReadCard        As Boolean

Public Enum led_Color
    红色 = 0
    绿色 = 1
    黄色 = 2
End Enum

Public Enum led_Status
    亮 = 1 '01H（亮）
    灭 = 2 '02H（灭）
    闪烁 = 3 '03H（闪烁）
End Enum

Public Enum Net_Ip
    本机IP = 1
    网关IP = 2
End Enum

Public Type typ_SS728M01
    '2.3.3取姓名
    ss_id_query_name        As String
    '2.3.4取性别代码
    ss_id_query_sexid       As String
    '2.3.5取性别名称
    ss_id_query_sex         As String
    '2.3.6取民族代码
    ss_id_query_folkid      As String
    '2.3.7取民族名称
    ss_id_query_folk        As String
    '2.3.8取出生日期
    ss_id_query_birth       As String
    '2.3.9取住址
    ss_id_query_address     As String
    '2.3.10取身份号码
    ss_id_query_number      As String
    '2.3.11取签发机关
    ss_id_query_organ       As String
    '2.3.12取有效期限起始日期
    ss_id_query_beginterm   As String
    '2.3.13取有效期限截止日期
    ss_id_query_endterm     As String
    '2.3.14取照片
    ss_id_query_photofile   As String
    '2.3.15 取最新住址
    ss_id_query_newaddr     As String
End Type

Public SS728M01 As typ_SS728M01

'====================================================================================================================================================
'2.2 通用函数
'====================================================================================================================================================

'2.2.1获取接口库信息
'long __stdcall ss_lib_version(char* pszVerInfo);
Public Declare Function ss_lib_version Lib "SS728M01.dll" _
    (ByVal pszVerInfo As String) As Long

'2.2.2打开本地终端
'long __stdcall ss_dev_open(char* szDevCom);
Public Declare Function ss_dev_open Lib "SS728M01.dll" _
    (ByVal szDevCom As String) As Long

'2.2.3关闭本地终端
'long __stdcall ss_dev_close(long icdev);
Public Declare Function ss_dev_close Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.2.4注册网络终端
'long __stdcall ss_dev_login(char* szDevIP);
Public Declare Function ss_dev_login Lib "SS728M01.dll" _
    (ByVal szDevIP As String) As Long
    
'2.2.5注销网络终端
'long __stdcall ss_dev_logout(long icdev);
Public Declare Function ss_dev_logout Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.2.6获取终端版本
'long __stdcall ss_dev_version(long icdev, char* pszVerInfo);
Public Declare Function ss_dev_version Lib "SS728M01.dll" _
    (ByVal icdev As Long, ss_dev_version As String) As Long

'2.2.7蜂鸣器
'long __stdcall ss_dev_beep(long icdev, unsigned short _Amount, unsigned short _Msec);
Public Declare Function ss_dev_beep Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal Amount As Long, ByVal Msec As Long) As Long
'2.2.8指示灯
'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。
'_Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
Public Declare Function ss_dev_led Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal Color As led_Color, ByVal Status As led_Status, ByVal Amount As Long, ByVal Msec As Long) As Long
'2.2.9读取网络参数
'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
Public Declare Function ss_dev_getnet Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal ipType As Net_Ip, ByVal pszParam As String) As Long
'2.2.10 设置网络参数
'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
Public Declare Function ss_dev_setnet Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal ipType As Net_Ip, ByVal pszParam As String) As Long
    
'====================================================================================================================================================
'2.3 第二代居民身份证函数
'====================================================================================================================================================

'2.3.1寻卡
'long __stdcall ss_id_find_card(long icdev);
Public Declare Function ss_id_find_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.3.2读取卡信息
'long __stdcall ss_id_read_card(long icdev);
Public Declare Function ss_id_read_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.3.3取姓名
'long __stdcall ss_id_query_name(long icdev, char* pszText);
Public Declare Function ss_id_query_name Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.4取性别代码
'long __stdcall ss_id_query_sexid(long icdev, char* pszText);
Public Declare Function ss_id_query_sexid Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.5取性别名称
'long __stdcall ss_id_query_sex(long icdev, char* pszText);
Public Declare Function ss_id_query_sex Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.6取民族代码
'long __stdcall ss_id_query_folkid(long icdev, char* pszText);
Public Declare Function ss_id_query_folkid Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.7取民族名称
'long __stdcall ss_id_query_folk(long icdev, char* pszText);
Public Declare Function ss_id_query_folk Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.8取出生日期
'long __stdcall ss_id_query_birth(long icdev, char* pszText);
Public Declare Function ss_id_query_birth Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.9取住址
'long __stdcall ss_id_query_address(long icdev, char* pszText);
Public Declare Function ss_id_query_address Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.10取身份号码
'long __stdcall ss_id_query_number(long icdev, char* pszText);
Public Declare Function ss_id_query_number Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.11取签发机关
'long __stdcall ss_id_query_organ(long icdev, char* pszText);
Public Declare Function ss_id_query_organ Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.12取有效期限起始日期
'long __stdcall ss_id_query_beginterm(long icdev, char* pszText);
Public Declare Function ss_id_query_beginterm Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.13取有效期限截止日期
'long __stdcall ss_id_query_endterm(long icdev, char* pszText);
Public Declare Function ss_id_query_endterm Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.14取照片
'long __stdcall ss_id_query_photofile(long icdev, unsigned char _Format, char* szImagePath);
Public Declare Function ss_id_query_photofile Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal typeFile As Long, ByVal szImagePath As String) As Long

'2.3.15 取最新住址
'long __stdcall ss_id_query_newaddr(long icdev, char* pszText);
Public Declare Function ss_id_query_newaddr Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.16 清空卡信息
'long __stdcall ss_id_free_card(long icdev);
Public Declare Function ss_id_free_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

Public Property Let ReadCardSS728M01(ByVal vNewValue As Boolean)
    mblnReadCard = vNewValue
End Property

Public Function ReadSS728M01() As Boolean
    Dim bytReadType     As Byte
    Dim strPort         As String
    Dim strServerIP     As String
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim icdev           As Long
    Dim szDevIP         As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim pszParam        As String
    Dim strSFZH         As String
    Dim pszText         As String
On Error GoTo ErrH
    bytReadType = IIf(UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
    strPort = UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "Port", "AUTO"))
    strServerIP = GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "NetIp", "192.168.31.169")
    
    '打开设备
    If bytReadType = 1 Then
        '2.2.2打开本地终端
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = "AUTO"
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    ElseIf bytReadType = 2 Then
        '2.2.4注册网络终端
        szDevIP = strServerIP
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    End If
    
    icdev = lngReturn
    ''2.2.7蜂鸣器
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(icdev, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
        
    '寻卡
    mblnReadCard = True
    
    If mblnReadCard Then
        '一直读卡
        Do While True
            DoEvents
            If Not mblnReadCard Then
                '2.2.8指示灯
                'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
                'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
                Amount = 1
                Msec = 2
                lngReturn = ss_dev_led(icdev, 绿色, 亮, Amount, Msec)
                If ss_error(lngReturn, False) Then
                    Exit Function
                End If
                If bytReadType = 1 Then
                    '2.2.3关闭本地终端
                    'long __stdcall ss_dev_close(long icdev);
                    lngReturn = ss_dev_close(icdev)
                    Call ss_error(lngReturn, False)
                    Exit Function
                Else
                    '2.2.5注销网络终端
                    lngReturn = ss_dev_logout(icdev)
                    Call ss_error(lngReturn, False)
                    Exit Function
                End If
            End If
            '2.2.8指示灯
            'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
            'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
            Amount = 1
            Msec = 1
            lngReturn = ss_dev_led(icdev, 红色, 闪烁, Amount, Msec)
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
            '2.3.1寻卡
            'long __stdcall ss_id_find_card(long icdev);
            lngReturn = ss_id_find_card(icdev)
            If lngReturn = 0 Then
                '寻卡成功
                Exit Do
            Else
                '寻卡失败，不处理。继续循环寻卡
                '2.2.8指示灯
                'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
                'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
                Amount = 1
                Msec = 2
                lngReturn = ss_dev_led(icdev, 绿色, 亮, Amount, Msec)
                If ss_error(lngReturn, False) Then
                    Exit Function
                End If
            End If
        Loop
    Else
        '只读一次
        '2.2.8指示灯
        'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
        'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
        Amount = 1
        Msec = 1
        lngReturn = ss_dev_led(icdev, 红色, 闪烁, Amount, Msec)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
        '2.3.1寻卡
        'long __stdcall ss_id_find_card(long icdev);
        lngReturn = ss_id_find_card(icdev)
        If lngReturn = 0 Then
            '寻卡成功
        Else
            '2.2.8指示灯
            'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
            'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_led(icdev, 绿色, 亮, Amount, Msec)
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
            '寻卡失败
            If bytReadType = 1 Then
                '2.2.3关闭本地终端
                'long __stdcall ss_dev_close(long icdev);
                lngReturn = ss_dev_close(icdev)
                Call ss_error(lngReturn, False)
                Exit Function
            Else
                '2.2.5注销网络终端
                lngReturn = ss_dev_logout(icdev)
                Call ss_error(lngReturn, False)
                Exit Function
            End If
        End If
    End If
    
    '2.3.2读取卡信息
    'long __stdcall ss_id_read_card(long icdev);
    lngReturn = ss_id_read_card(icdev)
    DoEvents
    If ss_error(lngReturn) Then
        If bytReadType = 1 Then
            '2.2.3关闭本地终端
            'long __stdcall ss_dev_close(long icdev);
            lngReturn = ss_dev_close(icdev)
            Call ss_error(lngReturn, False)
            Exit Function
        Else
            '2.2.5注销网络终端
            lngReturn = ss_dev_logout(icdev)
            Call ss_error(lngReturn, False)
            Exit Function
        End If
    End If
    '2.2.7蜂鸣器[读卡成功]
    Amount = 2
    Msec = 1
    lngReturn = ss_dev_beep(icdev, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    
    '2.3.3取姓名
    pszText = Space(30)
    lngReturn = ss_id_query_name(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_name = TruncZero(pszText)
    
    '2.3.4取性别代码
    pszText = Space(1)
    lngReturn = ss_id_query_sexid(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_sexid = TruncZero(pszText)
     
    '2.3.5取性别名称
    pszText = Space(2)
    lngReturn = ss_id_query_sex(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_sex = TruncZero(pszText)
    '2.3.6取民族代码
    pszText = Space(2)
    lngReturn = ss_id_query_folkid(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_folkid = TruncZero(pszText)
    '2.3.7取民族名称
    pszText = Space(10)
    lngReturn = ss_id_query_folk(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_folk = TruncZero(pszText)
    '2.3.8取出生日期
    pszText = Space(8)
    lngReturn = ss_id_query_birth(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_birth = TruncZero(pszText)
    '2.3.9取住址
    pszText = Space(70)
    lngReturn = ss_id_query_address(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_address = TruncZero(pszText)
    '2.3.10取身份号码
    pszText = Space(18)
    lngReturn = ss_id_query_number(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_number = TruncZero(pszText)
    '2.3.11取签发机关
    pszText = Space(30)
    lngReturn = ss_id_query_organ(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_organ = TruncZero(pszText)
    '2.3.12取有效期限起始日期
    pszText = Space(8)
    lngReturn = ss_id_query_beginterm(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_beginterm = TruncZero(pszText)
    '2.3.13取有效期限截止日期
    pszText = Space(8)
    lngReturn = ss_id_query_endterm(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_endterm = TruncZero(pszText)
    '取照片
    '2.3.14取照片
    lngReturn = ss_id_query_photofile(icdev, 2, "c:\" & SS728M01.ss_id_query_number)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_photofile = "c:\" & SS728M01.ss_id_query_number & ".bmp"
    
    '2.3.15 取最新住址
    pszText = Space(70)
    lngReturn = ss_id_query_newaddr(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_newaddr = TruncZero(pszText)
    
    '2.3.16 清空卡信息
    'long __stdcall ss_id_free_card(long icdev);
    lngReturn = ss_id_free_card(icdev)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
        
    '2.2.8指示灯
    'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
    'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_led(icdev, 绿色, 亮, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    If bytReadType = 1 Then
        '2.2.3关闭本地终端
        'long __stdcall ss_dev_close(long icdev);
        lngReturn = ss_dev_close(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    Else
        '2.2.5注销网络终端
        lngReturn = ss_dev_logout(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    End If
    ReadSS728M01 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Public Function ClosSS728M01() As Boolean
    Dim bytReadType     As Byte
    Dim strPort         As String
    Dim strServerIP     As String
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim icdev           As Long
    Dim szDevIP         As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim pszParam        As String
    Dim strSFZH         As String
    Dim pszText         As String
On Error GoTo ErrH
    bytReadType = IIf(UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
    strPort = UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "Port", "AUTO"))
    strServerIP = GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "Port", "192.168.31.169")
    '打开设备
    If bytReadType = 1 Then
        '2.2.2打开本地终端
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = "AUTO"
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    ElseIf bytReadType = 2 Then
        '2.2.4注册网络终端
        szDevIP = strServerIP
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    End If
    '2.2.8指示灯
    'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
    'icdev [in]：设备标识符。 _Color [in]：指示灯颜色。其中00H（红色）；01H（绿色）；02H（黄色）。 _Status [in]：指示灯状态。其中01H（亮）、02H（灭）、03H（闪烁）。 _Amount [in]：闪烁次数。 _Msec [in]：闪烁时长，单位100毫秒。
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_led(icdev, 绿色, 亮, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    If bytReadType = 1 Then
        '2.2.3关闭本地终端
        'long __stdcall ss_dev_close(long icdev);
        lngReturn = ss_dev_close(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    Else
        '2.2.5注销网络终端
        lngReturn = ss_dev_logout(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'根据错误号判断给予提示。
Public Function ss_error(lngErrNum As Long, Optional blnShowErr As Boolean = True) As Boolean
    Dim strErrMsg           As String
On Error GoTo ErrH
    Select Case lngErrNum
        Case 0 '操作成功
            ss_error = False
            Exit Function
        Case -1 '读卡器连接错
            strErrMsg = "读卡器连接错"
        Case -2 '未建立连接
            strErrMsg = "未建立连接"
        Case -3 '命令参数错
            strErrMsg = "命令参数错"
        Case -4 '通讯错误
            strErrMsg = "通讯错误"
        Case -11 '无卡
            strErrMsg = "无卡"
        Case -12 '卡片类型不对
            strErrMsg = "卡片类型不对"
        Case -13 '有卡未上电
            strErrMsg = "有卡未上电"
        Case -14 '卡片无应答
            strErrMsg = "卡片无应答"
        Case -21 '调用外部dll失败
            strErrMsg = "调用外部dll失败"
        Case -31 '照片解码失败
            strErrMsg = "照片解码失败"
        Case -100 '执行失败或未知错误
            strErrMsg = "执行失败或未知错误"
        Case Else
            strErrMsg = CStr(lngErrNum) & "未知的错误信息"
    End Select
    If blnShowErr Then MsgBox strErrMsg, vbExclamation, "神思读卡错误提示"
    ss_error = True
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, "系统消息"
    Err.Clear
    ss_error = True
End Function
 
