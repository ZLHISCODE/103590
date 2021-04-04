VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_URF_35H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Declare Function Open_USB Lib "mwhrf_bj.dll" () As Long
Private Declare Function Close_USB Lib "mwhrf_bj.dll" (ByVal icdev As Long) As Integer

Private Declare Function rf_beep Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal time1 As Integer) As Integer

Private Declare Function rf_card Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal mode%, Snr As Long) As Integer
Private Declare Function rf_load_key Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByRef nkey As Byte) As Integer
Private Declare Function rf_authentication Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal mode%, ByVal scenr%) As Integer
Private Declare Function rf_read_hex Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal Adr%, ByVal sdata$) As Integer
Private Declare Function rf_HL_readhex Lib "mwhrf_bj.dll" (ByVal icdev As Long, ByVal mode%, ByVal Adr%, Snr As Long, ByVal sdata$, Rsnr As Long) As Integer

Private mlngICDev As Long    '设备标识符
Private mint扇区 As Integer
Private mint块 As Integer
Private mintStart As Integer
Private mintLen As Integer

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim intReturn As Integer
    
    On Error GoTo errH
    intReturn = Close_USB(mlngICDev)
    clsCardDev_CloseCard = intReturn = 0
    
    Call WritLog("CardDev_CloseCard", "", IIf(clsCardDev_CloseCard, "连接成功!", "连接失败!"))
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    On Error GoTo errH
    mlngICDev = Open_USB()
    clsCardDev_initCard = mlngICDev > 0
    
    mint扇区 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & 9, "扇区", 1))
    mint块 = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & 9, "块", 0))
    mintStart = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & 9, "起始地址", 1))
    mintLen = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & 9, "长度", 10))
    Call WritLog("CardDev_initCard", "0,9600", IIf(clsCardDev_initCard, "断开成功!", "断开失败!"))
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim intResult As Integer
    Dim Snr As Long, Rsnr As Long
    Dim data As String * 32
    Dim aKey(5) As Byte
    On Err GoTo errHandle:
    
    aKey(0) = &HFF
    aKey(1) = &HFF
    aKey(2) = &HFF
    aKey(3) = &HFF
    aKey(4) = &HFF
    aKey(5) = &HFF
    intResult = rf_card(ByVal mlngICDev, 1, Snr)
    If intResult Then
        Call ReadRecord("rf_card", intResult)
    End If
    
    intResult = rf_load_key(ByVal mlngICDev, 0, mint扇区, aKey(0))
    If intResult Then
        Call ReadRecord("rf_load_key", intResult)
    End If
        
    intResult = rf_authentication(ByVal mlngICDev, 0, mint扇区)
    If intResult Then
        Call ReadRecord("rf_authentication", intResult)
    End If
    
    intResult = rf_HL_readhex(ByVal mlngICDev, 1, (mint扇区 * 4 + mint块), Snr, data, Rsnr)
    If intResult Then
        Call ReadRecord("rf_HL_readhex", intResult)
    End If
    
    If Trim(data) <> "" Then
        intResult = rf_beep(mlngICDev, 30)
        If intResult Then
            Call ReadRecord("rf_beep", intResult)
        End If
    End If
    
    clsCardDev_ReadCard = Trim(Mid(data, mintStart, mintLen))
    
    Call WritLog("读取成功", "", "卡号：" & data)
    Exit Function
errHandle:
    clsCardDev_CloseCard
    MsgBox Err.Description & Err.Number
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameter_URF_35H.ShowMe(9)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function

Private Sub ReadRecord(strMethodName As String, intReturn As Integer)
    Select Case intReturn
    Case 1
        Call WritLog(strMethodName, "0,9600", "无卡")
    Case 2
        Call WritLog(strMethodName, "0,9600", "CRC校验错")
    Case 3
        Call WritLog(strMethodName, "0,9600", "值溢出")
    Case 4
        Call WritLog(strMethodName, "0,9600", "未验证密码")
    Case 5
        Call WritLog(strMethodName, "0,9600", "奇偶校验错")
    Case 6
        Call WritLog(strMethodName, "0,9600", "通讯出错")
    Case 8
        Call WritLog(strMethodName, "0,9600", "错误的序列号")
    Case 10
        Call WritLog(strMethodName, "0,9600", "验证密码失败")
    Case 11
        Call WritLog(strMethodName, "0,9600", "接收的数据位错误")
    Case 12
        Call WritLog(strMethodName, "0,9600", "接收的数据字节错误")
    Case 14
        Call WritLog(strMethodName, "0,9600", "TransFor错误")
    Case 15
        Call WritLog(strMethodName, "0,9600", "写失败")
    Case 16
        Call WritLog(strMethodName, "0,9600", "加值失败")
    Case 17
        Call WritLog(strMethodName, "0,9600", "减值失败")
    Case &H10
        Call WritLog(strMethodName, "0,9600", "读失败")
    Case &H11
        Call WritLog(strMethodName, "0,9600", "PC与读写器通讯失败")
    Case &H20
        Call WritLog(strMethodName, "0,9600", "通讯超时")
    Case &H24
        Call WritLog(strMethodName, "0,9600", "打开通信口失败")
    Case &H30
        Call WritLog(strMethodName, "0,9600", "串口已被占用")
    Case &H31
        Call WritLog(strMethodName, "0,9600", "地址格式错误")
    Case &H32
        Call WritLog(strMethodName, "0,9600", "该块数据不是值格式")
    Case &H40
        Call WritLog(strMethodName, "0,9600", "长度错误")
    Case &H50
        Call WritLog(strMethodName, "0,9600", "值操作失败")
    Case &H60
        Call WritLog(strMethodName, "0,9600", "卡中的值不够减")
    Case &H61
        Call WritLog(strMethodName, "0,9600", "未找到读写器")
    Case Else
        Call WritLog("", "", "未知错误")
    End Select
End Sub

