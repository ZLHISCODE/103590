VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashPayEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event zlSendCash(dblTotal As Double, dblMony As Double, blnCancel As Boolean, strErrMsg As String)
Public Event zlSendCard(strCardNO As String, blnCancel As Boolean, strErrMsg As String)
'Public Event zlRetrunMoneyComplete(dbl缺币数量 As Double)

Public Sub zlRaiseSendCash(dblTotal As Double, dblMony As Double, blnCancel As Boolean, strErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:触发投入现金事件
    '入参:dblTotal-投放总额
    '       dblMoney-当前投放额
    '出参:blnCancel-true,操作失败;false-操作成功
    '       strErrMsg-返回的错误信息
    '编制:刘兴洪
    '日期:2012-04-11 14:29:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    RaiseEvent zlSendCash(dblTotal, dblMony, blnCancel, strErrMsg)
End Sub
Public Sub zlRaiseSendCard(strCardNO As String, blnCancel As Boolean, strErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:触发发卡事件
    '入参:strCardNo-当前发卡的卡号
    '出参:blnCancel-true,操作失败;false-操作成功
    '       strErrMsg-返回的错误信息
    '编制:刘兴洪
    '日期:2012-04-11 14:29:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    RaiseEvent zlSendCard(strCardNO, blnCancel, strErrMsg)
End Sub

'Public Sub zlRaiseRetrunMoneyComplete(dbl缺币数量 As Double)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:获取缺币数量
'    '入参:byt缺币数量
'    '编制:王吉
'    '日期:2012-11-19 14:29:02
'    '问题号:56055
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    RaiseEvent zlRetrunMoneyComplete(dbl缺币数量)
'End Sub
