VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClinicChargeInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvarShareUserID As Long '局部复制
Private mvarPatiUseType As String '局部复制
Private mvar费别 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar收费结算 As String '局部复制
Private mvar预结结算 As String '局部复制
Private mvar当前发票号 As String '局部复制
Private mvarintInsure As Integer '局部复制
Private mvar应收金额 As Double '局部复制
Private mvar实收金额 As Double '局部复制
Private mvar本次应缴 As Double  '局部复制
Private mvar医保预结金额 As Double '局部复制
Private mvar医保结算金额 As Double '局部复制
Private mvar消费合计 As Double '局部复制
Private mvar医保不走票号 As Boolean   '局部复制
Private mvar应缴累计 As Double '局部复制
Private mvar缴款 As Double '局部复制
Private mvar找补 As Double '局部复制
Private mvarNos As String '局部复制
Private mvar结帐ID As Long '局部复制
Private mvar冲销ID As Long  '局部复制
Private mvar结算序号 As Long '局部复制
Private mvarrsBalance As ADODB.Recordset   '局部复制
Private mvarrs结算方式 As ADODB.Recordset   '局部复制
Private mvar家属IDs As String '局部复制
Private mvar缺省结算方式 As String '局部复制
Private mvar费用来源 As String

Public Property Let 费用来源(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.费用来源 = 5
    mvar费用来源 = vData
End Property


Public Property Get 费用来源() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.费用来源
    费用来源 = mvar费用来源
End Property


Public Property Let 缺省结算方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缺省结算方式 = 5
    mvar缺省结算方式 = vData
End Property


Public Property Get 缺省结算方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缺省结算方式
    缺省结算方式 = mvar缺省结算方式
End Property

Public Property Let 家属IDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.家属IDs = 5
    mvar家属IDs = vData
End Property


Public Property Get 家属IDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.家属IDs
    家属IDs = mvar家属IDs
End Property


Public Property Let Nos(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Nos = 5
    mvarNos = vData
End Property


Public Property Get Nos() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Nos
    Nos = mvarNos
End Property



Public Property Let 找补(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.找补 = 5
    mvar找补 = vData
End Property


Public Property Get 找补() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.找补
    找补 = mvar找补
End Property

Public Property Let 缴款(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缴款 = 5
    mvar缴款 = vData
End Property


Public Property Get 缴款() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缴款
    缴款 = mvar缴款
End Property



Public Property Let 应缴累计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.应缴累计 = 5
    mvar应缴累计 = vData
End Property


Public Property Get 应缴累计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.应缴累计
    应缴累计 = mvar应缴累计
End Property

Public Property Let 本次应缴(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.本次应缴 = 5
    mvar本次应缴 = vData
End Property


Public Property Get 本次应缴() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.本次应缴
    本次应缴 = mvar本次应缴
End Property
 


Public Property Let 医保结算金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医保结算金额 = 5
    mvar医保结算金额 = vData
End Property


Public Property Get 医保结算金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保结算金额
    医保结算金额 = mvar医保结算金额
End Property




Public Property Let 医保预结金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医保预结金额 = 5
    mvar医保预结金额 = vData
End Property


Public Property Get 医保预结金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保预结金额
    医保预结金额 = mvar医保预结金额
End Property



Public Property Let 实收金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实收金额 = 5
    mvar实收金额 = vData
End Property


Public Property Get 实收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.实收金额
    实收金额 = mvar实收金额
End Property



Public Property Let 应收金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.应收金额 = 5
    mvar应收金额 = vData
End Property


Public Property Get 应收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.应收金额
    应收金额 = mvar应收金额
End Property

Public Property Let 消费合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费合计 = 5
    mvar消费合计 = vData
End Property


Public Property Get 消费合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费合计
    消费合计 = mvar消费合计
End Property


Public Property Let intInsure(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.intInsure = 5
    mvarintInsure = vData
End Property


Public Property Get intInsure() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.intInsure
    intInsure = mvarintInsure
End Property

Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mvar性别
End Property



Public Property Let 费别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.费别 = 5
    mvar费别 = vData
End Property


Public Property Get 费别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.费别
    费别 = mvar费别
End Property
Public Property Let 收费结算(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费结算 = 5
    mvar收费结算 = vData
End Property
Public Property Get 收费结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.收费结算
    收费结算 = mvar收费结算
End Property
Public Property Let 预结结算(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预结结算 = 5
    mvar预结结算 = vData
End Property
Public Property Get 预结结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预结结算
    预结结算 = mvar预结结算
End Property
 
 Public Property Let 当前发票号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.当前发票号 = 5
    mvar当前发票号 = vData
End Property
Public Property Get 当前发票号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.当前发票号
    当前发票号 = mvar当前发票号
End Property
 
Public Property Let PatiUseType(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PatiUseType = 5
    mvarPatiUseType = vData
End Property


Public Property Get PatiUseType() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PatiUseType
    PatiUseType = mvarPatiUseType
End Property



Public Property Let ShareUserID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ShareUserID = 5
    mvarShareUserID = vData
End Property


Public Property Get ShareUserID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ShareUserID
    ShareUserID = mvarShareUserID
End Property

Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property
Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    年龄 = mvar年龄
End Property



Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property


Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mvar姓名
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

Public Property Get 结帐ID() As Long
    结帐ID = mvar结帐ID
End Property

Public Property Let 结帐ID(ByVal vData As Long)
    mvar结帐ID = vData
End Property

Public Property Get 结算序号() As Long
    结算序号 = mvar结算序号
End Property

Public Property Let 结算序号(ByVal vData As Long)
    mvar结算序号 = vData
End Property

Public Property Let 医保不走票号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医保不走票号 = 5
    mvar医保不走票号 = vData
End Property

Public Property Get 医保不走票号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保不走票号
    医保不走票号 = mvar医保不走票号
End Property
Public Property Get 冲销ID() As Long
   冲销ID = mvar冲销ID
End Property

Public Property Let 冲销ID(ByVal vData As Long)
   mvar冲销ID = vData
End Property
 
Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.rsBalance = 5
    Set mvarrsBalance = vData
End Property
Public Property Get rsBalance() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsBalance
   Set rsBalance = mvarrsBalance
End Property
Public Property Set rs结算方式(ByVal vData As ADODB.Recordset)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.rs结算方式 = 5
    Set mvarrs结算方式 = vData
End Property
Public Property Get rs结算方式() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rs结算方式
   Set rs结算方式 = mvarrs结算方式
End Property

 
