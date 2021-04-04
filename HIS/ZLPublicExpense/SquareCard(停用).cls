VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SquareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'保持属性值的局部变量
Private mvarobjSquareCard As Object '局部复制
Private mvarobjCurCard As Card
Private mvarobjDefaultCard  As Card
Private mvarbln缺省卡号密文 As Boolean '局部复制
Private mvarint缺省卡号长度 As Integer   '局部复制
Private mvarbln按缺省卡查找   As Boolean

Public Property Let bln按缺省卡查找(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln按缺省卡查找 = 5
    mvarbln按缺省卡查找 = vData
End Property
Public Property Get bln按缺省卡查找() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln按缺省卡查找
    bln按缺省卡查找 = mvarbln按缺省卡查找
End Property

Public Property Set objDefaultCard(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objDefaultCard = Form1
    Set mvarobjDefaultCard = vData
End Property


Public Property Get objDefaultCard() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objDefaultCard
    Set objDefaultCard = mvarobjDefaultCard
End Property

Public Property Set objCurCard(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objCurCard = Form1
    Set mvarobjCurCard = vData
End Property


Public Property Get objCurCard() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objCurCard
    Set objCurCard = mvarobjCurCard
End Property
 
Public Property Let int缺省卡号长度(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.int医疗卡长度 = 5
    mvarint缺省卡号长度 = vData
End Property


Public Property Get int缺省卡号长度() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.int医疗卡长度
    int缺省卡号长度 = mvarint缺省卡号长度
End Property
 
Public Property Let bln缺省卡号密文(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln缺省卡号密文 = 5
    mvarbln缺省卡号密文 = vData
End Property
Public Property Get bln缺省卡号密文() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln缺省卡号密文
    bln缺省卡号密文 = mvarbln缺省卡号密文
End Property
Public Property Set objSquareCard(ByVal vData As Object)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objSquareCard = Form1
    Set mvarobjSquareCard = vData
End Property


Public Property Get objSquareCard() As Object
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objSquareCard
    Set objSquareCard = mvarobjSquareCard
End Property





