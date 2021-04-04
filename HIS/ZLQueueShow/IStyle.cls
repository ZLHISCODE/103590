VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub Show(ByVal lngWindowNo As Long)
'功能：实现排队数据显示
'lngWindowNo:窗口编号，根据窗口编号读取配置信息，并进行显示

End Sub


Public Function ShowCfg(ByVal lngWindowNo As Long, objOwner As Object) As Boolean
'功能：打开对应的样式配置窗口
'lngWindowNo:窗口编号
'objOwner:窗口拥有者

End Function



Public Function MsgProcess(ByVal lngWindowNo As Long, _
    ByVal strMsgKey As String, ByVal strXmlContext As String, rsData As ADODB.Recordset) As Boolean
'消息接收处理
    
    
End Function


Public Function WindNo() As Long
'获取当前样式窗口编号

End Function


Public Sub RefreshQueueData(Optional ByVal lngQueueId As Long)
'刷新界面排队呼叫信息

End Sub
