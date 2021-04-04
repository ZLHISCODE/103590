VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "合作单位控制"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"号序信息集"
Option Explicit
'保持属性值的局部变量
Private mvar合作单位名称 As String '局部复制
Private mvar预约控制方式 As Integer '局部复制
Private mvar号序信息集 As 号序信息集  '局部复制
Private mvar类型 As Byte '局部复制

Public Property Let 类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar类型 = vData
End Property


Public Property Get 类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.类型
    类型 = mvar类型
End Property


Public Property Set 号序信息集(ByVal vData As 号序信息集)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.obj号序信息集 = 5
    Set mvar号序信息集 = vData
End Property


Public Property Get 号序信息集() As 号序信息集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.obj号序信息集
    Set 号序信息集 = mvar号序信息集
End Property



Public Property Let 预约控制方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预约控制方式 = 5
    mvar预约控制方式 = vData
End Property


Public Property Get 预约控制方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预约控制方式
    预约控制方式 = mvar预约控制方式
End Property



Public Property Let 合作单位名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.合作单位名称 = 5
    mvar合作单位名称 = vData
End Property


Public Property Get 合作单位名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.合作单位名称
    合作单位名称 = mvar合作单位名称
End Property


Public Function Clone() As 合作单位控制
    '创建一个副本
    Dim obj合作单位控制  As New 合作单位控制
    
    With obj合作单位控制
        .类型 = mvar类型
        .合作单位名称 = mvar合作单位名称
        .预约控制方式 = mvar预约控制方式
        If Not mvar号序信息集 Is Nothing Then Set .号序信息集 = mvar号序信息集.Clone
    End With
    Set Clone = obj合作单位控制
End Function
