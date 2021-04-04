VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPROutLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarDBID As Long          '数据库ID
Private mvar文件ID As Long
Private mvar名称 As String
Private mvar说明 As String
Private mvar保留 As Boolean
Private mvar预制ID As Long
Private mvar复用 As Boolean
Private mvar层次 As Long
Private mvar文本 As String

Private mvar流水号 As Long      '对应于集合的Key。

Public Property Let 流水号(ByVal vData As Long)
    mvar流水号 = vData
End Property

Public Property Get 流水号() As Long
    流水号 = mvar流水号
End Property

Public Property Let 文本(ByVal vData As String)
    mvar文本 = vData
End Property

Public Property Get 文本() As String
    文本 = mvar文本
End Property

Public Property Let 层次(ByVal vData As Long)
    mvar层次 = vData
End Property

Public Property Get 层次() As Long
    层次 = mvar层次
End Property

Public Property Let 复用(ByVal vData As Boolean)
    mvar复用 = vData
End Property

Public Property Get 复用() As Boolean
    复用 = mvar复用
End Property

Public Property Let 预制ID(ByVal vData As Long)
    mvar预制ID = vData
End Property

Public Property Get 预制ID() As Long
    预制ID = mvar预制ID
End Property

Public Property Let 保留(ByVal vData As Boolean)
    mvar保留 = vData
End Property

Public Property Get 保留() As Boolean
    保留 = mvar保留
End Property

Public Property Let 说明(ByVal vData As String)
    mvar说明 = vData
End Property

Public Property Get 说明() As String
    说明 = mvar说明
End Property

Public Property Let 名称(ByVal vData As String)
    mvar名称 = vData
End Property

Public Property Get 名称() As String
    名称 = mvar名称
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let DBID(ByVal vData As Long)
    mvarDBID = vData
End Property

Public Property Get DBID() As Long
    DBID = mvarDBID
End Property




