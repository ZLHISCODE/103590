VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "分诊诊室"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'保持属性值的局部变量
Private mvar诊室ID As Long '局部复制
Private mvar诊室名称 As String '局部复制
Public Property Let 诊室名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.诊室名称 = 5
    mvar诊室名称 = vData
End Property


Public Property Get 诊室名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.诊室名称
    诊室名称 = mvar诊室名称
End Property



Public Property Let 诊室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.诊室ID = 5
    mvar诊室ID = vData
End Property


Public Property Get 诊室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.诊室ID
    诊室ID = mvar诊室ID
End Property

Public Function Clone() As 分诊诊室
    '创建一个副本
    Dim obj分诊诊室  As New 分诊诊室
    
    With obj分诊诊室
        .诊室ID = mvar诊室ID
        .诊室名称 = mvar诊室名称
    End With
    Set Clone = obj分诊诊室
End Function


