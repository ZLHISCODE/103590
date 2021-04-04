VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreeSwapCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvar卡类别ID As Long
Private mvar含医保 As Boolean
Private mvar多结算方式 As Boolean
Private mvar允许退现 As Boolean
Private mvar缺省退现 As Boolean
Private mvar缺省退现方式 As String

Public Property Let 卡类别ID(ByVal vData As Long)
    mvar卡类别ID = vData
End Property

Public Property Get 卡类别ID() As Long
    卡类别ID = mvar卡类别ID
End Property

    
Public Property Let 含医保(ByVal vData As Boolean)
    mvar含医保 = vData
End Property

Public Property Get 含医保() As Boolean
    含医保 = mvar含医保
End Property

    
Public Property Let 多结算方式(ByVal vData As Boolean)
    mvar多结算方式 = vData
End Property

Public Property Get 多结算方式() As Boolean
    多结算方式 = mvar多结算方式
End Property

    
Public Property Let 允许退现(ByVal vData As Boolean)
    mvar允许退现 = vData
End Property

Public Property Get 允许退现() As Boolean
    允许退现 = mvar允许退现
End Property

    
Public Property Let 缺省退现(ByVal vData As Boolean)
    mvar缺省退现 = vData
End Property

Public Property Get 缺省退现() As Boolean
    缺省退现 = mvar缺省退现
End Property

    
Public Property Let 缺省退现方式(ByVal vData As String)
    mvar缺省退现方式 = vData
End Property

Public Property Get 缺省退现方式() As String
    缺省退现方式 = mvar缺省退现方式
End Property
