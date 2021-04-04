VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatiFeeInfor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar未结费用 As Double '局部复制
Private mvar预交余额 As Double '局部复制
Private mvar担保额 As Double '局部复制
Private mvar预结费用 As Double '局部复制
Private mvar剩余款 As Double '局部复制
Public Property Let 剩余款(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.剩余款 = 5
    mvar剩余款 = vData
End Property
Public Property Get 剩余款() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.剩余款
    剩余款 = mvar剩余款
End Property



Public Property Let 预结费用(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预结费用 = 5
    mvar预结费用 = vData
End Property


Public Property Get 预结费用() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预结费用
    预结费用 = mvar预结费用
End Property



Public Property Let 担保额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.担保额 = 5
    mvar担保额 = vData
End Property


Public Property Get 担保额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.担保额
    担保额 = mvar担保额
End Property



Public Property Let 预交余额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预交余额 = 5
    mvar预交余额 = vData
End Property


Public Property Get 预交余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预交余额
    预交余额 = mvar预交余额
End Property



Public Property Let 未结费用(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.未结费用 = 5
    mvar未结费用 = vData
End Property


Public Property Get 未结费用() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未结费用
    未结费用 = mvar未结费用
End Property



