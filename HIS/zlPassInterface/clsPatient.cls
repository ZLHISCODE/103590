VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mlng病人ID As Long '局部复制
Private mlng主页ID As Long '局部复制
Private mlng挂号ID As Long '局部复制
Private mdbl标识号 As Double '局部复制
Private mstr姓名 As String '局部复制
Private mstr性别 As String '局部复制
Private mstr床号 As String '局部复制
Private mint婴儿 As Integer '局部复制
Private mstr挂号单 As String '局部复制
Private mdat出生日期 As Date '局部复制
Private mlngDeptID As Long  '局部复制

Public Property Let dat出生日期(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.dat出生日期 = 5
    mdat出生日期 = vData
End Property

Public Property Get dat出生日期() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.dat出生日期
     dat出生日期 = mdat出生日期
End Property

Public Property Let str挂号单(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.str挂号单 = 5
    mstr挂号单 = vData
End Property

Public Property Get str挂号单() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.str挂号单
     str挂号单 = mstr挂号单
End Property


Public Property Let int婴儿(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.int婴儿 = 5
    mint婴儿 = vData
End Property


Public Property Get int婴儿() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.int婴儿
     int婴儿 = mint婴儿
End Property

Public Property Let str床号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.str床号 = 5
    mstr床号 = vData
End Property

Public Property Get str床号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.str床号
     str床号 = mstr床号
End Property

Public Property Let str性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.str性别 = 5
    mstr性别 = vData
End Property

Public Property Get str性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.str性别
     str性别 = mstr性别
End Property

Public Property Let str姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.str姓名 = 5
    mstr姓名 = vData
End Property

Public Property Get str姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.str姓名
    str姓名 = mstr姓名
End Property

Public Property Let dbl标识号(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.dbl标识号 = 5
    mdbl标识号 = vData
End Property

Public Property Get dbl标识号() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.dbl标识号
    dbl标识号 = mdbl标识号
End Property

Public Property Let lng挂号ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.lng挂号ID = 5
    mlng挂号ID = vData
End Property

Public Property Get lng挂号ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.lng挂号ID
    lng挂号ID = mlng挂号ID
End Property

Public Property Let lng主页ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.lng主页ID = 5
    mlng主页ID = vData
End Property

Public Property Get lng主页ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.lng主页ID
    lng主页ID = mlng主页ID
End Property

Public Property Let lng病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.lng病人ID = 5
    mlng病人ID = vData
End Property

Public Property Get lng病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.lng病人ID
    lng病人ID = mlng病人ID
End Property

Public Property Let lngDeptID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.lngDeptID = 5
    mlngDeptID = vData
End Property

Public Property Get lngDeptID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.lngDeptID
    lngDeptID = mlngDeptID
End Property

