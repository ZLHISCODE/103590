VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "停诊记录"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'保持属性值的局部变量
Private mvar开始时间 As String '局部复制
Private mvar终止时间 As String '局部复制
Private mvar停诊原因 As String '局部复制
'保持属性值的局部变量
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



Public Property Let 停诊原因(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停诊原因 = 5
    mvar停诊原因 = vData
End Property


Public Property Get 停诊原因() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停诊原因
    停诊原因 = mvar停诊原因
End Property



Public Property Let 终止时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.终止时间 = 5
    mvar终止时间 = vData
End Property


Public Property Get 终止时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.终止时间
    终止时间 = mvar终止时间
End Property



Public Property Let 开始时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.开始时间 = 5
    mvar开始时间 = vData
End Property


Public Property Get 开始时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.开始时间
    开始时间 = mvar开始时间
End Property

Public Function Clone() As 停诊记录
    '创建一个副本
    Dim obj停诊记录  As New 停诊记录
    
    With obj停诊记录
        .类型 = mvar类型
        .开始时间 = mvar开始时间
        .终止时间 = mvar终止时间
        .停诊原因 = mvar停诊原因
    End With
    Set Clone = obj停诊记录
End Function

