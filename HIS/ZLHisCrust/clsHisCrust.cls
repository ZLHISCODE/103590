VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHisCrust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ReceiveCommand(ByVal strCommand As String, Optional ByVal strParas As String, Optional ByRef pbSend As PropertyBag, Optional ByRef pbReturn As PropertyBag) As Boolean
'功能：升级程执行命令，方便扩展,保持接口兼容
'strCommand:执行的命令
'strParas:执行命令的参数
'pbSend：附加传送数据
'pbReturn:附加返回数据
'返回：命令是否执行成功
End Function

Public Sub ExitApp()
'功能：结束升级
    End
End Sub
