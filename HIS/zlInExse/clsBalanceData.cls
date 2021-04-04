VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvarobjBalanceItems As clsBalanceItems '局部复制
Private mvar结算总额 As Double '局部复制
Public Property Let 结算总额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算总额 = 5
    mvar结算总额 = vData
End Property

Public Property Get 结算总额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算总额
    结算总额 = mvar结算总额
End Property
Public Property Set objBalanceItems(ByVal vData As clsBalanceItems)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objBalanceItems = Form1
    Set mvarobjBalanceItems = vData
End Property


Public Property Get objBalanceItems() As clsBalanceItems
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objBalanceItems
    Set objBalanceItems = mvarobjBalanceItems
End Property



