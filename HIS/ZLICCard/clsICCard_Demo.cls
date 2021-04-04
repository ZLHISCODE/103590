VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCardDev_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements clsCardDev

Private Function clsCardDev_CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_CloseCard = True
End Function

Private Function clsCardDev_GetSpare() As Currency

End Function

Private Function clsCardDev_initCard(Optional strInput As String, Optional strOut As String) As Boolean
    clsCardDev_initCard = True
End Function


Private Function clsCardDev_PaymentSwap(ByVal dblMoney As Double, dblBalance As Double, intCardType As Integer, ByVal lngHOSCode As Long, strCardNO As String, strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean

End Function

Private Function clsCardDev_ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    Dim strNo As String
    strNo = InputBox("请输入IC卡号，如果输入内容为空则采用默认值123456789", "DEMO", "123456789")
    If strNo = "" Then strNo = "123456789"
    clsCardDev_ReadCard = strNo
End Function

Private Function clsCardDev_ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean

End Function

Private Sub clsCardDev_SetCard(Optional strInput As String, Optional strOut As String)
    Call frmParameterSet.ShowMe(1)
End Sub

Private Function clsCardDev_UploadSwap(ByVal lng病人ID As Long, strOut As String) As Boolean
    
End Function
