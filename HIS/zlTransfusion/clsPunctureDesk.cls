VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPunctureDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mlngSeqNo As Long '序号
Private mStrBeeperCode As String '呼叫器编号
Private mintStat As Integer '是否在用 0－未用  1-在用

'序号
Public Property Let SeqNo(ByVal vData As Long)
    mlngSeqNo = vData
End Property

Public Property Get SeqNo() As Long
    SeqNo = mlngSeqNo
End Property

'呼叫器编号
Public Property Let BeeperCode(ByVal vData As String)
    mStrBeeperCode = vData
End Property

Public Property Get BeeperCode() As String
    BeeperCode = mStrBeeperCode
End Property

'是否在用
Public Property Let Stat(ByVal vData As Integer)
    Stat = vData
End Property

Public Property Get Stat() As Integer
    Stat = mintStat
End Property




