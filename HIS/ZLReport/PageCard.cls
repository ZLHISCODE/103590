VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Key As String
'保持属性值的局部变量
Private mvarID As Integer '局部复制
Private mvarX As Long '局部复制
Private mvarY As Long '局部复制
Private mvarW As Long '局部复制
Private mvarH As Long '局部复制
Private mvarRowB As Long '局部复制
Private mvarRowE As Long '局部复制
Private mvarCol As Long
Private mvarRow As Long
Private mColItem As Collection

Public Property Let RowE(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowE = 5
    mvarRowE = vData
End Property

Public Property Get RowE() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowE
    RowE = mvarRowE
End Property

Public Property Let RowB(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowB = 5
    mvarRowB = vData
End Property

Public Property Get RowB() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowB
    RowB = mvarRowB
End Property

Public Property Let Row(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowE = 5
    mvarRow = vData
End Property

Public Property Get Row() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowE
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowB = 5
    mvarCol = vData
End Property

Public Property Get Col() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowB
    Col = mvarCol
End Property

Public Property Let H(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.H = 5
    mvarH = vData
End Property

Public Property Get H() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.H
    H = mvarH
End Property

Public Property Let W(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.W = 5
    mvarW = vData
End Property

Public Property Get W() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.W
    W = mvarW
End Property

Public Property Let Y(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Y = 5
    mvarY = vData
End Property

Public Property Get Y() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.X = 5
    mvarX = vData
End Property

Public Property Get X() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.X
    X = mvarX
End Property

Public Property Let id(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get id() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ID
    id = mvarID
End Property

Public Property Let item(ByVal vData As Collection)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    Set mColItem = vData
End Property

Public Property Get item() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ID
    Set item = mColItem
End Property
