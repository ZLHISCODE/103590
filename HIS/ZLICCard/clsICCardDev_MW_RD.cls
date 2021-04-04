VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_MW_RD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Declare Function auto_init Lib "mwic_32.dll" (ByVal port As Integer, ByVal baud As Long) As Integer
Private Declare Function ic_exit Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Private Declare Function get_status Lib "mwic_32.dll" (ByVal icdev As Long, ByRef status As Integer) As Integer
Private Declare Function srd_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer As String) As Integer
Private Declare Function chk_4442 Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Private Declare Function dv_beep Lib "mwic_32.dll" (ByVal icdev As Long, ByVal time As Integer) As Integer

Private mintICTarget As Long   '设备标识符

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim i As Integer, blnicexit As Boolean
    
    On Error GoTo errH
    i = ic_exit(mintICTarget)
    clsCardDev_CloseCard = i = 0
  
    Call WritLog("CardDev_CloseCard", "", IIf(clsCardDev_CloseCard, "成功!", "失败!"))
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    Dim i As Integer
    
    i = Val(GetSetting("ZLSOFT", "公共全局\ICCard\4", "端口", 0))
    If i < 0 Or i > 7 Then i = 0
    
    On Error GoTo errH
    mintICTarget = auto_init(i, 9600) '端口为0
    clsCardDev_initCard = mintICTarget > 0
    
 
    Call WritLog("CardDev_initCard", "0,9600", IIf(clsCardDev_initCard, "成功!", "失败!"))
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function



Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim i As Integer, intState As Integer, intSAddr As Integer, intLen As Integer
    Dim strResult As String * 16    '接口函数srd_4442要求必须为定长字符串
    
    On Error GoTo errH
  
    i = get_status(mintICTarget, intState)
    If i = 0 Then
        If intState = 0 Then
            MsgBox "请插卡!", vbInformation
            Exit Function
        End If
         
        If chk_4442(mintICTarget) < 0 Then
            MsgBox "无效卡", vbInformation
            Exit Function
        End If
    Else
        MsgBox "读卡错误,请检查!", vbInformation
        Exit Function
    End If
    
    intSAddr = Val(GetSetting("ZLSOFT", "公共全局\ICCard\4", "起始地址", 32))
    If intSAddr <= 0 Then intSAddr = 1
    intLen = Val(GetSetting("ZLSOFT", "公共全局\ICCard\4", "长度", 10))
    If intLen > 16 Then intLen = 16
    If intLen <= 0 Then intLen = 1
 
    i = dv_beep(mintICTarget, 10)
Read:
    i = srd_4442(mintICTarget, intSAddr, intLen, strResult)
                
    If i = 0 Then
        clsCardDev_ReadCard = Trim(strResult)
    Else
       If MsgBox("卡内数据错误,是否重试?", vbYesNo) = vbYes Then
          GoTo Read
       End If
    End If
    
    Call WritLog("CardDev_ReadCard", "", clsCardDev_ReadCard)
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmICSet_Demo.ShowMe(4)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function
