VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFactProperty"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvar共享批次ID As Long '局部复制
Private mvar使用类别 As String '局部复制
Private mvar打印格式 As Integer '局部复制
Private mvar打印方式 As Integer '局部复制
'保持属性值的局部变量
Private mvar票种 As Integer '局部复制
Private mvar严格控制 As Boolean '局部复制
Private mvar启用使用类别 As Boolean '局部复制
Private mvar票号长度 As Integer   '局部复制
Private mvarLastUseID As Long     '局部复制

Public Property Let LastUseID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.LastUseID = 5
    mvarLastUseID = vData
End Property
Public Property Get LastUseID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.LastUseID
    LastUseID = mvarLastUseID
End Property


Public Property Let 票号长度(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.票号长度 = 5
    mvar票号长度 = vData
End Property
Public Property Get 票号长度() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.票号长度
    票号长度 = mvar票号长度
End Property


Public Property Let 启用使用类别(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.启用使用类别 = 5
    mvar启用使用类别 = vData
End Property


Public Property Get 启用使用类别() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.启用使用类别
    启用使用类别 = mvar启用使用类别
End Property


Public Property Let 严格控制(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.严格控制 = 5
    mvar严格控制 = vData
End Property


Public Property Get 严格控制() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.严格控制
    严格控制 = mvar严格控制
End Property



Public Property Let 票种(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.票种 = 5
    mvar票种 = vData
End Property


Public Property Get 票种() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.票种
      票种 = mvar票种
End Property



Public Property Let 打印方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.打印方式 = 5
    mvar打印方式 = vData
End Property


Public Property Get 打印方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.打印方式
      打印方式 = mvar打印方式
End Property



Public Property Let 打印格式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.打印格式 = 5
    mvar打印格式 = vData
End Property


Public Property Get 打印格式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.打印格式
    打印格式 = mvar打印格式
End Property



Public Property Let 使用类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.使用类别 = 5
    mvar使用类别 = vData
End Property


Public Property Get 使用类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.使用类别
    使用类别 = mvar使用类别
End Property



Public Property Let 共享批次ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.共享批次ID = 5
    mvar共享批次ID = vData
End Property


Public Property Get 共享批次ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.共享批次ID
    共享批次ID = mvar共享批次ID
End Property



