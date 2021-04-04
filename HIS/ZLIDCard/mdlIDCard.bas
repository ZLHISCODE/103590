Attribute VB_Name = "mdlIDCard"
Option Explicit '要求变量声明
Public gblnIDCard As Boolean    '设备是否可用
Public glngType As Long         '设备类型编号

Public editPort As Long         '记录通用接口的端口号
Public blnUsbPort As Boolean    ' 记录通用接口的端口类型

'目前来看,几种设备都使用相同的接口文件termb.dll,方法是相同的,现在做成只支持我们认可的设备,可以限制未经中联授权的设备商接入系统
Public Enum IDCardType
    CVR100U = 1                 '深圳华视CVR-100U
    CVR100D = 2
    SS_V1 = 3                   '神思
    XZX_KDQ = 4                 '新中新
    GTICR100 = 5                '成都国腾实业集团有限公司 GTICR100
    HX_FDX9 = 6                 '华旭金卡股份有限公司
    GTICR100_1 = 7              '成都国腾实业集团         GTICR100——产品升级
    CVR100U_1 = 8               '深圳华视CVR-100U——产品升级
    CVR100D_1 = 9               '深圳华视CVR-100D——产品升级
    DKQ_116D = 10               '新中新DKQ116D
    GTICR100_01 = 11            '成都国腾实业集团有限公司 GTICR100-01
    CVR100 = 12                 '深圳华视CVR-100
    COMMON = 13                 '通用接口
    SS728M01_B01C = 14          '神思SS72801 版号B01C
End Enum

Public Type MIDCardInfor    '定义结构体变量，保存身份证信息
    Name As String * 32  '姓名
    sex As String * 4  '性别
    nation As String * 6  '名族
    born As String * 18  '出生日期
    address As String * 72  '住址
    IDcardno As String * 38  '身份证号
    grantdept As String * 32  '发证机关
    UserLifeBegin As String * 18  '有效开始日期
    UserLifeEnd As String * 18  '有效截止日期
    reserved As String * 38  '保留
    PhotoFileName As String * 255  '照片路径
End Type

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public IDCardInfor As MIDCardInfor

Public Declare Function InitComm Lib "termb.dll" (ByVal Port As Integer) As Integer
Public Declare Function InitCommExt Lib "termb.dll" () As Integer
Public Declare Function Authenticate Lib "termb.dll" () As Integer
Public Declare Function AuthenticateExt Lib "termb.dll" () As Integer
Public Declare Function Read_Content_Path Lib "termb.dll" (ByVal fileName As String, ByVal Active As Integer) As Integer
Public Declare Function Read_Content Lib "termb.dll" (ByVal Active As Integer) As Integer
Public Declare Function CloseComm Lib "termb.dll" () As Integer
Public Declare Function GetSAMID Lib "termb.dll" () As String

Public Declare Function SDT_GetCOMBaud Lib "sdtapi.dll" (ByVal iPort As Integer, puiBaudRate As Long) As Integer
Public Declare Function SDT_SetCOMBaud Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal uiCurrBaud As Long, ByVal uiSetBaud) As Integer
Public Declare Function SDT_OpenPort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer
Public Declare Function SDT_ClosePort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

Public Declare Function SDT_ResetSAM Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_SetMaxRFByte Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal ucByte As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMStatus Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMID Lib "sdtapi.dll" (ByVal iPort As Integer, pucSAMID As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_GetSAMIDToStr Lib "sdtapi.dll" (ByVal iPort As Integer, pcSAMID As String, ByVal iIfOpen As Integer) As Integer

Public Declare Function SDT_StartFindIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, CardPUCIIN As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_SelectIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, pucManaMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadMngInfo Lib "sdtapi.dll" (ByVal iPort As Integer, pucManageMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadBaseMsg Lib "sdtapi.dll" (ByVal iPort As Integer, pucCHMsg As String, puiCHMsgLen As Long, _
         pucPHMsg As String, puiPHMsgLen As Long, ByVal iIfOpen As Integer) As Integer
Public Declare Function SDT_ReadBaseMsgToFile Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal pucCHFile As String, puiCHMsgLen As Long, _
         ByVal pucPHFile As String, puiPHMsgLen As Long, ByVal iIfOpen As Integer) As Integer


Public Declare Function GetBmp Lib "WltRS.dll" (ByVal WLTFile As String, ByVal intf As Integer) As Integer

'深圳华视CVR-100U——产品升级动态库函数
Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer
Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer
Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'新中新DKQ调用API
'********************** 端口类API *************************
Public Declare Function Syn_OpenPort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer) As Integer
Public Declare Function Syn_ClosePort Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer) As Integer

'********************** SAM类API **************************
Public Declare Function Syn_GetSAMStatus Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_ResetSAM Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer) As Integer

'******************* 身份证卡类API ************************
Public Declare Function Syn_StartFindIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, pucManaInfo As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_SelectIDCard Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, pucManaMsg As String, ByVal iIfOpen As Integer) As Integer
Public Declare Function Syn_ReadMsg Lib "Syn_IDCardRead.dll" (ByVal iPortID As Integer, ByVal iIfOpen As Integer, pIDCardData As MIDCardInfor) As Integer
Public Declare Function Syn_ReadCard Lib "Syn_IDCardRead.dll" (pIDCardData As MIDCardInfor, ByVal Rmode As Integer) As Integer
'Rmode 整数，1：为连续读卡；2：为一次一次读卡。
'******************* 附加类API ************************
Public Declare Function Syn_SendSound Lib "Syn_IDCardRead.dll" (ByVal iCmdNo As Integer) As Integer '说明: 发送语音
Public Declare Sub Syn_DelPhotoFileLib Lib "Syn_IDCardRead.dll" () '说明: 删除临时照片文件

Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Function GetPID() As Long
   GetWindowThreadProcessId GetForegroundWindow, GetPID
End Function
Public Function Init(ByVal objIDCard As clsIDCard) As Boolean
    Dim i As Long, lngTmp As Long
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

    On Error GoTo ErrH
    gblnIDCard = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "启用", 0)) = 1
    If gblnIDCard Then

        glngType = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "设备种类", 0))
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_1, IDCardType.GTICR100_01

                i = InitComm(1001) '初始化端口
                If i <> 1 Then i = InitCommExt
                gblnIDCard = (i = 1)
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                i = CVR_InitComm(1001)
                If i <> 1 Then i = CVR_InitComm(1)
                gblnIDCard = (i = 1)
            Case IDCardType.CVR100D, IDCardType.HX_FDX9
                i = InitComm(1)
                If i <> 1 Then i = InitCommExt
                gblnIDCard = (i = 1)

                If gblnIDCard = False Then
                    i = SDT_OpenPort(1)
                    gblnIDCard = (i = CByte(&H90))
                End If
            Case IDCardType.DKQ_116D
                i = Syn_OpenPort(1001)
                gblnIDCard = (i = 0)
                Call Syn_ClosePort(1001)   '防止多窗口初始化失败
            Case IDCardType.CVR100
                i = SDT_OpenPort(1001)
                If i <> CByte(&H90) Then i = SDT_OpenPort(1)
                gblnIDCard = (i = CByte(&H90))

            Case IDCardType.COMMON
                gblnIDCard = False
                '检测usb口的机具连接，必须先检测usb
                Dim iPort As Long
                For iPort = 1001 To 1016
                    i = SDT_OpenPort(iPort)
                    If i = CByte(&H90) Then
                        editPort = iPort
                        blnUsbPort = True
                        gblnIDCard = True
                        Exit For
                    End If
                Next

                '检测串口的机具连接
                If Not blnUsbPort Then
                    For iPort = 1 To 16
                        i = SDT_OpenPort(iPort)
                        If i = CByte(&H90) Then
                            editPort = iPort
                            blnUsbPort = False
                            gblnIDCard = True
                            Exit For
                        End If
                    Next
                End If
            Case IDCardType.SS728M01_B01C
                bytReadType = IIf(UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
                strPort = UCase(GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "Port", "AUTO"))
                strServerIP = GetSetting("ZLSOFT", "公共全局\IDCard\SS728M01", "NetIp", "192.168.31.169")

                '检测设备连接是否正常
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
                '关闭设备
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
                gblnIDCard = True
        End Select

        If gblnIDCard Then
            Init = True
        Else
            MsgBox "身份证设备端口初始化失败,请检查!", vbInformation, "ZLSoft"
            Init = False
        End If
    End If
    Exit Function

ErrH:
    gblnIDCard = False
End Function

Public Sub Terminate()
    If gblnIDCard Then
        Select Case glngType
            Case IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_1, IDCardType.GTICR100_01
                CloseComm
            Case IDCardType.CVR100U_1, IDCardType.CVR100D_1
                CVR_CloseComm
            Case IDCardType.HX_FDX9
                Call SDT_ClosePort(1)
            Case IDCardType.DKQ_116D
                Call Syn_ClosePort(1001)
            Case IDCardType.CVR100
                Call SDT_ClosePort(1001)
            Case IDCardType.COMMON
                Call SDT_ClosePort(editPort)
            Case IDCardType.SS728M01_B01C
                Call ClosSS728M01
        End Select
    End If
End Sub

Public Sub SetAutoRead(ByVal timer As timer, blnEnabled As Boolean)
    If gblnIDCard Then
        Select Case glngType
            Case IDCardType.CVR100, IDCardType.CVR100U, IDCardType.CVR100D, IDCardType.SS_V1, IDCardType.XZX_KDQ, IDCardType.GTICR100, IDCardType.GTICR100_01, IDCardType.GTICR100_1, IDCardType.HX_FDX9, IDCardType.CVR100U_1, IDCardType.CVR100D_1, IDCardType.DKQ_116D, IDCardType.COMMON
                timer.Enabled = blnEnabled
            Case IDCardType.SS728M01_B01C
                timer.Enabled = blnEnabled
                '自助读取并且找卡
                ReadCardSS728M01 = blnEnabled
        End Select
    End If
End Sub

Public Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格
Dim Nstr As String
    If InStr(str, Chr(0)) > 0 Then
        Nstr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        Nstr = Trim(str)
    End If
    TrimStr = Replace(Replace(Replace(Nstr, Chr(13), vbCr), vbLf, ""), vbTab, "")
End Function

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Sub ClearIDCardInfor()
    IDCardInfor.address = ""
    IDCardInfor.born = ""
    IDCardInfor.grantdept = ""
    IDCardInfor.IDcardno = ""
    IDCardInfor.Name = ""
    IDCardInfor.nation = ""
    IDCardInfor.PhotoFileName = ""
    IDCardInfor.reserved = ""
    IDCardInfor.sex = ""
    IDCardInfor.UserLifeBegin = ""
    IDCardInfor.UserLifeEnd = ""
End Sub

Public Function GetParentHwnd(ByVal ChildHwnd As Long) As Long
    Dim i As Long, lngHwnd As Long
    On Error GoTo errHand
    i = ChildHwnd

    Do While i > 0
        lngHwnd = i
        i = GetParent(i)
    Loop
    GetParentHwnd = lngHwnd
    Exit Function
errHand:
    Err.Clear
End Function


Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
