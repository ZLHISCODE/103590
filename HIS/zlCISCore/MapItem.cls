VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar类型 As Byte '局部复制
Private mvar内容 As String '局部复制
Private mvar点集 As String '局部复制
Private mvarX1 As Long '局部复制
Private mvarY1 As Long '局部复制
Private mvarX2 As Long '局部复制
Private mvarY2 As Long '局部复制
Private mvar填充色 As Long '局部复制
Private mvar填充方式 As Integer '局部复制
Private mvar线条色 As Long '局部复制
Private mvar线型 As Byte  '局部复制
Private mvar线宽 As Byte  '局部复制
Private mvar字体 As String   '局部复制

Private mvarKey As String   '局部复制
Private mvarOper As Byte    '局部复制

Public Property Let Key(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    Key = mvarKey
End Property

Public Property Let Oper(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvarOper = vData
End Property

Public Property Get Oper() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    Oper = mvarOper
End Property

Public Property Let 字体(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvar字体 = vData
End Property

Public Property Get 字体() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    字体 = mvar字体
End Property

Public Property Let 线宽(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvar线宽 = vData
End Property

Public Property Get 线宽() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    线宽 = mvar线宽
End Property

Public Property Let 线型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvar线型 = vData
End Property

Public Property Get 线型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    线型 = mvar线型
End Property

Public Property Let 线条色(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvar线条色 = vData
End Property

Public Property Get 线条色() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    线条色 = mvar线条色
End Property

Public Property Let 填充方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充方式 = 5
    mvar填充方式 = vData
End Property

Public Property Get 填充方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充方式
    填充方式 = mvar填充方式
End Property


Public Property Let 填充色(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.填充色 = 5
    mvar填充色 = vData
End Property

Public Property Get 填充色() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.填充色
    填充色 = mvar填充色
End Property

Public Property Let Y2(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Y2 = 5
    mvarY2 = vData
End Property

Public Property Get Y2() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Y2
    Y2 = mvarY2
End Property

Public Property Let X2(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.X2 = 5
    mvarX2 = vData
End Property

Public Property Get X2() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.X2
    X2 = mvarX2
End Property

Public Property Let Y1(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Y1 = 5
    mvarY1 = vData
End Property

Public Property Get Y1() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Y1
    Y1 = mvarY1
End Property

Public Property Let X1(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.X1 = 5
    mvarX1 = vData
End Property

Public Property Get X1() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.X1
    X1 = mvarX1
End Property

Public Property Let 点集(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.点集 = 5
    mvar点集 = vData
End Property

Public Property Get 点集() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.点集
    点集 = mvar点集
End Property

Public Property Let 内容(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.内容 = 5
    mvar内容 = vData
End Property

Public Property Get 内容() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.内容
    内容 = mvar内容
End Property

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
