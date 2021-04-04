VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_ZT606"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

'建立串口连接
Private Declare Function ZtDevice_CR_Open Lib "CardReader.DLL" (ByVal nComPort As Integer, ByVal nBaudRate As Integer) As Integer
'关闭串口连接
Private Declare Function ZtDevice_CR_Close Lib "CardReader.DLL" () As Integer
'复位
Private Declare Function ZtDevice_CR_Reset Lib "CardReader.DLL" () As Integer
'读2磁道数据
Private Declare Function ZtDevice_CR_ReadTrack2 Lib "CardReader.DLL" (ByVal sdata$, nDataTwoLen As Integer, ByVal nTimeOut As Integer) As Integer


Private mbytPort As Integer
Private mintStart As Integer
Private mintLength As Integer

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim intResult As Integer
    
    clsCardDev_initCard = True
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim intResult As Integer
    Dim intDataLen As Integer
    Dim nTimeOut As Integer
    nTimeOut = 15
    
    
    Dim strData As String * 32
    
    
    '初始化
    mbytPort = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "端口", 0)) + 1
    mintStart = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "起始地址", 1))
    mintStart = IIf(mintStart = 0, 1, mintStart)
    mintLength = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "长度", 8))
    
    intResult = ZtDevice_CR_Open(mbytPort, 9600)
    Call WritLog("clsCardDev_initCard", "", IIf(intResult, "连接成功!", "连接失败!"))
    
    '读卡
    intResult = ZtDevice_CR_ReadTrack2(strData, intDataLen, nTimeOut)
    If intResult = 1 Then
        clsCardDev_ReadCard = Trim(Mid(strData, mintStart, mintLength))
        Call WritLog("clsCardDev_initCard", "", "读卡成功！")
    Else
        clsCardDev_ReadCard = ""
        Call WritLog("clsCardDev_initCard", "", "读卡失败！")
    End If
    
    '重置
    intResult = ZtDevice_CR_Reset
    Call WritLog("clsCardDev_initCard", "", IIf(intResult, "重置成功!", "重置失败!"))
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(11)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim intResult As Integer
    intResult = ZtDevice_CR_Close
    clsCardDev_CloseCard = intResult = 1
    Call WritLog("CardDev_CloseCard", "", IIf(clsCardDev_CloseCard, "断开成功!", "断开失败!"))
End Function
