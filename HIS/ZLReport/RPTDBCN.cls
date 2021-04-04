VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTDBCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'说明：ZLHIS产品之外的数据连接扩展对象

Private mintNO As Integer
Private mstrName As String
Private mcnConnection As ADODB.Connection

Public Property Get Connection() As ADODB.Connection
    Set Connection = mcnConnection
End Property

Public Property Set Connection(ByVal conValue As ADODB.Connection)
    Set mcnConnection = conValue
End Property

Public Property Get Name() As String
    Name = mstrName
End Property

Public Property Let Name(ByVal strValue As String)
    mstrName = strValue
End Property

Public Property Get NO() As Integer
    NO = mintNO
End Property

Public Property Let NO(ByVal intValue As Integer)
    mintNO = intValue
End Property

