VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "一个元素就是记帐单上的一个控件"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvarKey As String '局部复制
Private mvarControl As Object '局部复制
Private mvarTabIndex As Long '局部复制
Private mvarValue As String '局部复制
Private mvarVisible As Boolean '局部复制

Public Property Let Visible(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Visible = 5
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property

Public Property Let Value(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Value = 5
    mvarValue = vData
End Property

Public Property Get Value() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property

Public Property Let TabIndex(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.TabIndex = 5
    mvarTabIndex = vData
End Property

Public Property Get TabIndex() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.TabIndex
    TabIndex = mvarTabIndex
End Property

Public Property Set Control(ByVal vData As Object)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Control = Form1
    Set mvarControl = vData
End Property


Public Property Get Control() As Object
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Control
    Set Control = mvarControl
End Property

Public Property Let Key(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Key = 5
    mvarKey = vData
End Property

Public Property Get Key() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property
