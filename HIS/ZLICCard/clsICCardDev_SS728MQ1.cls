VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_SS728MQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'名称 打开串口设备
'函数原型  long ss_reader_open(char *szDevCom);
'功能      打开与机具所对应的终端设备端口，如果调用成功则返回设备句柄。
'参数      <1> szDevCom [in] 设备名称，取值“AUTO ”、“COMn ”，其中“n ”的取值为1~9。
'返回值      若正常，返回值为不小于0 的设备标识符。
Private Declare Function ss_reader_open Lib "SS728M01.dll" (ByVal szDevCom As String) As Long

'名称 关闭串口设备
'函数原型  long ss_reader_close(long icdev);
'功能      关闭由ss_reader_open 函数所打开的终端设备端口。
'参数      <1> icdev [in] 设备标识符。
'返回值 执行成功返回值为0?
Private Declare Function ss_reader_close Lib "SS728M01.dll" (ByVal icdev As Long) As Long

'名称 控制蜂鸣器
'函数原型
'long ss_dev_beep( long icdev,unsigned short _Amount,unsigned short _Msec);
'功能 控制读卡终端的蜂鸣器鸣响?
'参数
'<1> icdev [in] 设备标识符。
'<2> _Amount [in] 鸣响次数。
'<3> _Msec [in] 鸣响时间，时间以100ms为单位。
'返回值 执行成功返回值为0?
Private Declare Function ss_dev_beep Lib "SS728M01.dll" (ByVal icdev As Long, ByVal amount As Integer, ByVal msec As Integer) As Long

'名称 控制指示灯
'函数原型
'long ss_dev_led(long icdev,
'unsigned char _Color,
'unsigned char _Status,
'unsigned short _Amount,
'unsigned short _Msec);
'功能  控制读卡终端的指示灯动作，目前仅控制“工作”灯。
'参数
'<1> icdev [in] 设备标识符。
'<2> _Color [in] 颜色，其值范围0~2。其中：0表示红色；1表示绿色；2表示黄色。
'<3> _Status [in] 状态，其值范围1~3。其中：1表示亮；2表示灭；3表示闪烁。
'<4> _Amount [in] 闪烁次数。
'<5> _Msec [in] 闪烁时间，时间以100ms为单位。
'返回值 执行成功返回值为0?
Private Declare Function ss_dev_led Lib "SS728M01.dll" (ByVal icdev As Long, ByVal color As Integer, ByVal status As Integer, ByVal amount As Integer, ByVal msec As Integer) As Long

'名称 获取磁卡数据
'函数原型
'long ss_readmagcard(long icdev,
'char *pszTrack1,
'char *pszTrack2,
'char *pszTrack3,
'unsigned char WaitTime );
'适用范围 磁卡
'功能 在指定的时间内读取磁卡上3个磁道的数据?
'参数
'<1> icdev [in] 设备标识符。
'<2> pszTrack1 [out] 磁道1数据，255个字节以内的字符串，以'\0'结束。
'<3> pszTrack2 [out] 磁道2数据，255个字节以内的字符串，以'\0'结束。
'<4> pszTrack3 [out] 磁道3数据，255个字节以内的字符串，以'\0'结束。
'<5> WaitTime [in] 等待时间，单位为秒，其值范围10~60。
'返回值 执行成功返回值为0?
Private Declare Function ss_readmagcard Lib "SS728M01.dll" (ByVal icdev As Long, ByVal pszTrack1 As String, ByVal pszTrack2 As String, ByVal pszTrack3 As String, ByVal WaitTime As Byte) As Long



Private Declare Function ss_id_find_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long
Private Declare Function ss_id_read_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long
Private Declare Function ss_id_query_name Lib "SS728M01.dll" (ByVal icdev As Long, ByVal szName As String) As Long
Private Declare Function ss_id_query_address Lib "SS728M01.dll" (ByVal icdev As Long, ByVal szAddr As String) As Long
Private Declare Function ss_id_free_card Lib "SS728M01.dll" (ByVal icdev As Long) As Long

Private micdev As Long
Private lngReturn As Long
Private mintStart As Integer
Private mintLength As Integer

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    lngReturn = ss_reader_close(micdev)
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    micdev = ss_reader_open("AUTO")
    clsCardDev_initCard = micdev > 0
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim pszTrack1 As String * 255, pszTrack2 As String * 255, pszTrack3 As String * 255
    Dim WaitTime As Byte
    WaitTime = 10
    
    mintStart = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\13", "起始地址", 2))
    mintStart = IIf(mintStart = 0, 2, mintStart)
    mintLength = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\14", "长度", 16))
    
    lngReturn = ss_readmagcard(micdev, pszTrack1, pszTrack2, pszTrack3, WaitTime)
    If lngReturn = 0 Then
        clsCardDev_ReadCard = Trim(Mid(pszTrack2, mintStart, mintLength))
    Else
        clsCardDev_ReadCard = ""
    End If
'    Dim rtn As Long
'    Dim szName As String * 30
'    Dim szAddr As String * 70
'
'    rtn = ss_id_find_card(micdev)
'    If rtn = 0 Then
'        rtn = ss_id_read_card(micdev)
'        If rtn = 0 Then
'            rtn = ss_id_query_name(micdev, szName)
'            rtn = ss_id_query_address(micdev, szAddr)
'            Debug.Print szName
'            Debug.Print szAddr
'        End If
'    End If
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(13)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean
    
End Function
