VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsFunc"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvar功能 As String '局部复制
Private mvar排列 As Integer '局部复制
Private mvar说明 As String '局部复制
Private mvarChecked As Boolean '局部复制
Private mvar分组 As Integer '局部复制
Private mvar关系 As Integer '局部复制
Private mvar主功能 As Integer
Private mvar主功能关系 As Integer
Private mvarKey As String

Public Property Let 关系(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关系 = 5
    mvar关系 = vData
End Property


Public Property Get 关系() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关系
    关系 = mvar关系
End Property



Public Property Let 分组(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分组 = 5
    mvar分组 = vData
End Property


Public Property Get 分组() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分组
    分组 = mvar分组
End Property

Public Property Let 主功能(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分组 = 5
    mvar主功能 = vData
End Property


Public Property Get 主功能() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分组
    主功能 = mvar主功能
End Property

Public Property Let 主功能关系(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分组 = 5
    mvar主功能关系 = vData
End Property


Public Property Get 主功能关系() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分组
    主功能关系 = mvar主功能关系
End Property

Public Property Let Checked(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
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


Public Property Let Key(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.说明 = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.说明
   Key = mvarKey
End Property

Public Property Let 排列(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.排列 = 5
    mvar排列 = vData
End Property


Public Property Get 排列() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.排列
    排列 = mvar排列
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



