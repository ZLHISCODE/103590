VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_CQPubCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev
Private mobjOneCard As Object

Private Sub Class_Initialize()
    On Error Resume Next
    Set mobjOneCard = CreateObject("OneCardCom.OneCardSys.1")
    If mobjOneCard Is Nothing Then
        gLastErr = "创建对象失败!请检查接口文件是否存在."
    End If
End Sub

Private Sub Class_Terminate()
    Set mobjOneCard = Nothing
End Sub

Private Function clsCardDev_GetSpare() As Currency
'参数:strInterfaceType-接口类型名称在某中结算方式上输金额时可以确定
    Dim dbBalance As Double, intCardType As Integer
    Dim strCardNO As String * 12, strReMark As String * 512, lngErrCode As Long
    
    'clsCardDev_GetSpare = 200: Exit Function
    
    Call mobjOneCard.fnQueryMon(dbBalance, intCardType, strCardNO, strReMark, lngErrCode)
    If lngErrCode <> 0 Then
        gLastErr = Trim(strReMark)
    Else
        clsCardDev_GetSpare = dbBalance
    End If
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, ByRef dblBalance As Double, ByRef intCardType As Integer, _
    ByVal lngHOSCode As Long, ByRef strCardNO As String, ByRef strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    
    Dim dblMoneyOut As Double, lngErrCode As Long
    Dim strReMark As String * 512
    
'    dblBalance = 180
'    intCardType = 1
'    strCardNO = "600216"
'    strTransFlow = "20080509"
'    clsCardDev_PaymentSwap = True: Exit Function
   
    Call mobjOneCard.fnDeductSgl(dblMoney, dblMoneyOut, dblBalance, intCardType, lngHOSCode, strCardNO, strTransFlow, strBalanceID, strPatiID, strReMark, lngErrCode)
        
    If lngErrCode = 0 Then
        strCardNO = Trim(strCardNO)
        strTransFlow = Trim(strTransFlow)
    Else
        gLastErr = Trim(strReMark)
    End If
    clsCardDev_PaymentSwap = lngErrCode = 0
End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim intCardType As Integer, strCardNO As String * 12, strReMark As String * 512, lngErrCode As Long
    
    'clsCardDev_ReadCard = "600216": Exit Function
    
    On Error Resume Next
    Call mobjOneCard.fnGetCardNo(intCardType, strCardNO, strReMark, lngErrCode)
    
    If lngErrCode <> 0 Then
        gLastErr = Trim(strReMark)
    ElseIf intCardType = 0 Or intCardType = 1 Then
        clsCardDev_ReadCard = Trim(strCardNO)   '卡类型不用返回，在扣款交易时会返回
    Else
        gLastErr = "不支持该类别的卡！"
    End If
End Function


Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_initCard = Not mobjOneCard Is Nothing
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    Dim strReMark As String * 512, lngErrCode As Long, strIP As String, strHOSCode As String
    'strIP不传,由一卡通接口部件去配置和读取
    
    'clsCardDev_ReturnSwap = True: Exit Function
    
    strHOSCode = lngHOSCode
    Call mobjOneCard.fnRollBack(strCardNO, strIP, strHOSCode, strTransFlow, dblMoney, strReMark, lngErrCode)
    
    If lngErrCode <> 0 Then gLastErr = Trim(strReMark)
    clsCardDev_ReturnSwap = lngErrCode = 0
End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)

End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function
