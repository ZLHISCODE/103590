VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTFmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvar序号 As Byte '局部复制
Private mvar说明 As String '局部复制
Private mvarW As Long '局部复制
Private mvarH As Long '局部复制
Private mvar纸张 As Integer '局部复制
Private mvar纸向 As Byte '局部复制
Private mvar动态纸张 As Boolean '局部复制
Private mvar图样 As Long

Public Property Let 图样(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.图样 = 5
    mvar图样 = vData
End Property

Public Property Get 图样() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.图样
    图样 = mvar图样
End Property

Public Property Let 说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.说明 = 5
    mvar说明 = vData
End Property

Public Property Get 说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.说明
    说明 = mvar说明
End Property

Public Property Let 序号(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property

Public Property Get 序号() As Byte
Attribute 序号.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.序号
    序号 = mvar序号
End Property

Public Property Let 纸向(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.纸向 = 5
    mvar纸向 = vData
End Property

Public Property Get 纸向() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.纸向
    纸向 = mvar纸向
End Property

Public Property Let 纸张(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.纸张 = 5
    mvar纸张 = vData
End Property

Public Property Get 纸张() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.纸张
    纸张 = mvar纸张
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

Public Property Let 动态纸张(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.动态纸张 = 5
    mvar动态纸张 = vData
End Property

Public Property Get 动态纸张() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.动态纸张
    动态纸张 = mvar动态纸张
End Property
