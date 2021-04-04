VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'保持属性值的局部变量
Private mvarCardObject As Object '局部复制
Private mvar自制卡 As Boolean '局部复制
Private mvar接口序号 As Long '局部复制
Private mvarCardPreporty As clsCard '局部复制
Private mvar消费卡 As Boolean '局部复制
Private mvarInitCompents As Boolean '局部复制
Public Property Let InitCompents(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡 = 5
    mvarInitCompents = vData
End Property
Public Property Get InitCompents() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡
    InitCompents = mvarInitCompents
End Property

Public Property Let 消费卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡 = 5
    mvar消费卡 = vData
End Property
Public Property Get 消费卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡
    消费卡 = mvar消费卡
End Property



Public Property Set CardPreporty(ByVal vData As clsCard)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.CardPreporty = Form1
    Set mvarCardPreporty = vData
End Property


Public Property Get CardPreporty() As clsCard
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CardPreporty
    Set CardPreporty = mvarCardPreporty
End Property



Public Property Let 接口序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接口序号 = 5
    mvar接口序号 = vData
End Property


Public Property Get 接口序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.接口序号
    接口序号 = mvar接口序号
End Property



Public Property Let 自制卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.自制卡 = 5
    mvar自制卡 = vData
End Property


Public Property Get 自制卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.自制卡
    自制卡 = mvar自制卡
End Property



Public Property Set CardObject(ByVal vData As Object)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.CardObject = Form1
    Set mvarCardObject = vData
End Property


Public Property Get CardObject() As Object
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CardObject
    Set CardObject = mvarCardObject
End Property



