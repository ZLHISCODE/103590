VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_MHCX_715K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev


Private mbytPort As Integer
Private mintStart As Integer
Private mintLength As Integer

'读卡
Private Declare Function readcard Lib "MHCX715.dll" (ByVal nPort As Long, ByVal nTrack As Long, ByVal sData1 As String, ByVal sData2 As String, ByVal nModel As Long, ByVal cmto As Long) As Long
'写卡
Private Declare Function writecard Lib "MHCX715.dll" (ByVal nPort As Long, ByVal nTrack As Long, ByVal sData1 As String, ByVal sData2 As String, ByVal nModel As Long, ByVal cmto As Long) As Long


Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean

End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_initCard = True
End Function

Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim data2 As String
    Dim data13 As String
    Dim n   As Long
    Dim err As Long
    
    '初始化
    mbytPort = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "端口", 0)) + 1
    mintStart = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "起始地址", 1))
    mintStart = IIf(mintStart = 0, 1, mintStart)
    mintLength = CInt(GetSetting("ZLSOFT", "公共全局\ICCard\11", "长度", 8))
    
'    s2 = Trim(sOneThree.Text) + Chr(0)
'    err = writecard(1, n, s1, s2, 2, cmto)

    data2 = String(255, Chr(0))
    data13 = String(255, Chr(0))
    err = readcard(mbytPort, 2, data2, data13, 1, 100000)
    
    
    Select Case err
    Case 0:
        clsCardDev_ReadCard = Trim(Mid(data2, mintStart, mintLength))
        Call WritLog("clsCardDev_initCard", "", "读卡成功！")
    Case 1
        clsCardDev_ReadCard = ""
        Call WritLog("clsCardDev_initCard", "", "读卡失败！")
    End Select
    
    

    
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(12)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean

End Function
