VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'消费卡信息
Private mvar卡类型 As String
Private mvar卡号 As String
Private mvar充值卡 As Boolean
Private mvar有效期 As String
Private mvar发卡原因 As String
Private mvar发卡人 As String
Private mvar领卡人 As String
Private mvar病人ID As Long
Private mvar发卡时间 As String
Private mvar领卡部门id As Long
Private mvar领卡部门 As String '编码-名称
Private mvar备注 As String
Private mvar卡余额 As Double
Private mvar卡面值 As Double
Private mvar实际销售 As Double
Private mvar充值折扣率 As Double
Private mvar回收人 As String
Private mvar回收时间 As String
Private mvar停用人 As String
Private mvar停用日期 As String
Private mvar当前状态 As String
Private mvar已充值 As Boolean '存在充值记录的就不允许修改“允许充值”属性了
Private mvar原密码 As String
Private mvar限制类别 As String

Public Property Let 卡类型(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡类型 = 5
    mvar卡类型 = vData
End Property

Public Property Get 卡类型() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡类型
    卡类型 = mvar卡类型
End Property


Public Property Let 卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡号 = 5
    mvar卡号 = vData
End Property

Public Property Get 卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡号
    卡号 = mvar卡号
End Property
    
    
Public Property Let 充值卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.充值卡 = 5
    mvar充值卡 = vData
End Property

Public Property Get 充值卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.充值卡
    充值卡 = mvar充值卡
End Property
    

Public Property Let 有效期(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.有效期 = 5
    mvar有效期 = vData
End Property

Public Property Get 有效期() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.有效期
    有效期 = mvar有效期
End Property
    

Public Property Let 发卡原因(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发卡原因 = 5
    mvar发卡原因 = vData
End Property

Public Property Get 发卡原因() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发卡原因
    发卡原因 = mvar发卡原因
End Property
    
 
Public Property Let 发卡人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发卡人 = 5
    mvar发卡人 = vData
End Property

Public Property Get 发卡人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发卡人
    发卡人 = mvar发卡人
End Property
    

Public Property Let 领卡人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.领卡人 = 5
    mvar领卡人 = vData
End Property

Public Property Get 领卡人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.领卡人
    领卡人 = mvar领卡人
End Property


Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mvar病人ID
End Property


Public Property Let 发卡时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发卡时间 = 5
    mvar发卡时间 = vData
End Property

Public Property Get 发卡时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发卡时间
    发卡时间 = mvar发卡时间
End Property


Public Property Let 领卡部门id(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.领卡部门id = 5
    mvar领卡部门id = vData
End Property

Public Property Get 领卡部门id() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.领卡部门id
    领卡部门id = mvar领卡部门id
End Property


Public Property Let 领卡部门(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.领卡部门 = 5
    mvar领卡部门 = vData
End Property

Public Property Get 领卡部门() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.领卡部门
    领卡部门 = mvar领卡部门
End Property


Public Property Let 备注(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.备注 = 5
    mvar备注 = vData
End Property

Public Property Get 备注() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.备注
    备注 = mvar备注
End Property

    
Public Property Let 卡余额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡余额 = 5
    mvar卡余额 = vData
End Property

Public Property Get 卡余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡余额
    卡余额 = mvar卡余额
End Property


Public Property Let 卡面值(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡面值 = 5
    mvar卡面值 = vData
End Property

Public Property Get 卡面值() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡面值
    卡面值 = mvar卡面值
End Property


Public Property Let 实际销售(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实际销售 = 5
    mvar实际销售 = vData
End Property

Public Property Get 实际销售() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.实际销售
    实际销售 = mvar实际销售
End Property


Public Property Let 充值折扣率(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.充值折扣率 = 5
    mvar充值折扣率 = vData
End Property

Public Property Get 充值折扣率() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.充值折扣率
    充值折扣率 = mvar充值折扣率
End Property


Public Property Let 回收人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.回收人 = 5
    mvar回收人 = vData
End Property

Public Property Get 回收人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.回收人
    回收人 = mvar回收人
End Property


Public Property Let 回收时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.回收时间 = 5
    mvar回收时间 = vData
End Property

Public Property Get 回收时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.回收时间
    回收时间 = mvar回收时间
End Property


Public Property Let 停用人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停用人 = 5
    mvar停用人 = vData
End Property

Public Property Get 停用人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停用人
    停用人 = mvar停用人
End Property


Public Property Let 停用日期(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停用日期 = 5
    mvar停用日期 = vData
End Property

Public Property Get 停用日期() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停用日期
    停用日期 = mvar停用日期
End Property


Public Property Let 当前状态(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.当前状态 = 5
    mvar当前状态 = vData
End Property

Public Property Get 当前状态() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.当前状态
    当前状态 = mvar当前状态
End Property


Public Property Let 已充值(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.已充值 = 5
    mvar已充值 = vData
End Property

Public Property Get 已充值() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.已充值
    已充值 = mvar已充值
End Property


Public Property Let 原密码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原密码 = 5
    mvar原密码 = vData
End Property

Public Property Get 原密码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原密码
    原密码 = mvar原密码
End Property


Public Property Let 限制类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.限制类别 = 5
    mvar限制类别 = vData
End Property

Public Property Get 限制类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.限制类别
    限制类别 = mvar限制类别
End Property
