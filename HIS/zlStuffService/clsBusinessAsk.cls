VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'保持属性值的局部变量
Private mvar发布ID As String '局部复制
Private mvar过程名 As String '局部复制
Private mvar功能 As String '局部复制
Private mvar返回信息 As clsBusinessParsing

Public Property Set 返回信息(ByVal vData As clsBusinessParsing)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.功能 = 5
    Set mvar返回信息 = vData
End Property


Public Property Get 返回信息() As clsBusinessParsing
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.功能
   Set 返回信息 = mvar返回信息
End Property


Public Property Let 功能(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.功能 = 5
    mvar功能 = vData
End Property


Public Property Get 功能() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.功能
    功能 = mvar功能
End Property



Public Property Let 过程名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.过程名 = 5
    mvar过程名 = vData
End Property


Public Property Get 过程名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.过程名
    过程名 = mvar过程名
End Property



Public Property Let 发布ID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发布ID = 5
    mvar发布ID = vData
End Property


Public Property Get 发布ID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发布ID
    发布ID = mvar发布ID
End Property



