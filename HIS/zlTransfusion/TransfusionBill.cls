VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "输液单"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar医嘱ID As Long '局部复制
Private mvar医嘱内容 As String '局部复制

'保持属性值的局部变量
Private mvar单量 As Currency  '局部复制
Private mvar单位 As String '局部复制
Private mvar容量 As Long '局部复制
Private mvar医生嘱托 As String '局部复制
Private mvar序号 As Integer '局部复制
Private mvar时间 As Long
Private mvar金额 As Currency '
Private mvar计费状态 As Integer

'保持属性值的局部变量

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property


Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.序号
    序号 = mvar序号
End Property


Public Property Let 医生嘱托(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医生嘱托 = 5
    mvar医生嘱托 = MidUni(vData, 1, 200)
End Property


Public Property Get 医生嘱托() As String
'检索属性值时使用，位于赋值语句的右边。
' X.医生嘱托
    医生嘱托 = mvar医生嘱托
End Property

Public Property Let 容量(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.容量 = 5
    mvar容量 = vData
End Property


Public Property Get 容量() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.容量
    容量 = mvar容量
End Property

Public Property Let 单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.单位 = 5
    mvar单位 = vData
End Property


Public Property Get 单位() As String
'检索属性值时使用，位于赋值语句的右边。
' X.单位
    单位 = mvar单位
End Property

Public Property Let 金额(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.单量 = 5
    mvar金额 = vData
End Property


Public Property Get 金额() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.单量
    金额 = mvar金额
End Property

Public Property Let 单量(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.单量 = 5
    mvar单量 = vData
End Property


Public Property Get 单量() As Currency
'检索属性值时使用，位于赋值语句的右边。
' X.单量
    单量 = mvar单量
End Property

Public Property Get 时间() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.单量
    时间 = mvar时间
End Property

Public Property Let 时间(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.单量 = 5
    mvar时间 = vData
End Property

Public Property Let 医嘱内容(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医嘱内容 = 5
    mvar医嘱内容 = vData
End Property


Public Property Get 医嘱内容() As String
'检索属性值时使用，位于赋值语句的右边。
' X.医嘱内容
    医嘱内容 = mvar医嘱内容
End Property



Public Property Let 医嘱ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医嘱ID = 5
    mvar医嘱ID = vData
End Property


Public Property Get 医嘱ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.医嘱ID
    医嘱ID = mvar医嘱ID
End Property



Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property


Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.病人ID
    病人ID = mvar病人ID
End Property

Public Property Let 明细计费状态(ByVal vData As Integer)
    mvar计费状态 = vData
End Property

Public Property Get 明细计费状态() As Integer
    明细计费状态 = mvar计费状态
End Property
