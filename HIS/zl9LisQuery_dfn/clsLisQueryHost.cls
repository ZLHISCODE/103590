VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisQueryHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get OwnerFormHandle() As Long
    '主程序的句柄
End Property

Public Sub RaiseFinished(objQuery As clsLisQuery)
    '在客户端触发此事件，后通知主程序，程序已完成
End Sub

Public Function GetRecordSet(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    '开放给客户端，用于执行取数据
End Function

Public Function ClientTrigger(ByVal Index As Long, ByVal strAction As String, strData As String) As String
    '开放给客户端触发事件
    '           预留功能
End Function


