VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ShowCardInfor(ByVal strCardType As String, ByVal strCardNO As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取卡号信息
'入参:strCardNO-卡号
'        strXmlCardInfor-卡号信息
'        strCardType-卡类别(可以为类别名称,可以为卡类别ID)
'出参:strExpended-暂未用,待以后扩展
'       blnCancel-true,表示取消本次读取;否则读取成功
'编制:刘兴洪
'日期:2012-04-23 10:49:54
'---------------------------------------------------------------------------------------------------------------------------------------------

Public Sub zlKeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Public Sub zlShowCardInfor(ByVal strCardType As String, ByVal strCardNO As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    RaiseEvent ShowCardInfor(strCardType, strCardNO, strXmlCardInfor, strExpended, blnCancel)
End Sub
