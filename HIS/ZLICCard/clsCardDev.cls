VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量

'接口类
Public Sub SetCard(Optional strInput As String, Optional strOut As String)
    '函数模板，不实现具体功能
End Sub

Public Function ReadCard(Optional strInput As String, Optional blnCancel As Boolean = False) As String
    '函数模板，不实现具体功能
End Function

Public Function initCard(Optional strInput As String, Optional strOut As String) As Boolean
    '函数模板，不实现具体功能
End Function

Public Function CloseCard(Optional strInput As String, Optional strOut As String) As Boolean
    '函数模板，不实现具体功能
End Function

Public Function GetSpare() As Currency
    '函数模板，返回一卡通的卡当前余额
End Function

Public Function PaymentSwap(ByVal dblMoney As Double, ByRef dblBalance As Double, ByRef intCardType As Integer, _
    ByVal lngHOSCode As Long, ByRef strCardNO As String, ByRef strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    '函数模板，执行一卡通扣费交易
End Function

Public Function ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
    '函数模板，执行一卡通退费交易
End Function

Public Function UploadSwap(ByVal lng病人ID As Long, ByRef strOut As String) As Boolean
    '函数模板，完成与中心数据上传的工作，目前仅应用于宁波一卡通，完成病人信息及卡信息的上传
    
End Function

Private Sub Class_Terminate()
    Call CloseCard
End Sub
