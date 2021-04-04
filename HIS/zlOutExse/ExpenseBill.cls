VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpenseBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Pages"
Option Explicit '要求变量声明

Public Key As String
'保持属性值的局部变量
Private mvarNO As String '局部复制
Private mvar病人ID As Long '局部复制
Private mvar主页ID As Long '局部复制
Private mvar病区ID As Long '局部复制
Private mvar科室ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar标识号 As String  '局部复制
Private mvar床号 As String '局部复制
Private mvar费别 As String '局部复制
Private mvar门诊标志 As Byte '局部复制
Private mvar加班标志 As Byte '局部复制
Private mvar多病人单 As Boolean '局部复制
Private mvar婴儿费 As Byte '局部复制
Private mvar划价人 As String '局部复制
Private mvar操作员姓名 As String '局部复制
Private mvar操作员编号 As String '局部复制
Private mvar发生时间 As Date '局部复制
Private mvar登记时间 As Date '局部复制
Private mvarPages As BillPages '局部复制
Private mvar收费结算 As String '局部复制
Private mvar结帐ID As Long
Public Property Let 结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费结算 = 5
    mvar结帐ID = vData
End Property
Public Property Get 结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.收费结算
    结帐ID = mvar结帐ID
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




Private Sub Class_Initialize()
    Dim objPage As New BillPage
    
    Set Me.Pages = New BillPages
    
    '整个单据对象初始化时,缺省固定包含一张单据
    Me.Pages.Add objPage.Details
End Sub

Private Sub Class_Terminate()
    Set Me.Pages = Nothing
End Sub

Public Property Set Pages(ByVal vData As BillPages)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Pages = vData
    Set mvarPages = vData
End Property

Public Property Get Pages() As BillPages
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Pages
    Set Pages = mvarPages
End Property

Public Property Let 登记时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发生时间 = 5
    mvar登记时间 = vData
End Property

Public Property Get 登记时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.发生时间
    登记时间 = mvar登记时间
End Property

Public Property Let 发生时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发生时间 = 5
    mvar发生时间 = vData
End Property

Public Property Get 发生时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.发生时间
    发生时间 = mvar发生时间
End Property

Public Property Let 操作员编号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.操作员编号 = 5
    mvar操作员编号 = vData
End Property

Public Property Get 操作员编号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.操作员编号
    操作员编号 = mvar操作员编号
End Property

Public Property Let 操作员姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.操作员姓名 = 5
    mvar操作员姓名 = vData
End Property

Public Property Get 操作员姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.操作员姓名
    操作员姓名 = mvar操作员姓名
End Property

Public Property Let 划价人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.划价人 = 5
    mvar划价人 = vData
End Property

Public Property Get 划价人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.划价人
    划价人 = mvar划价人
End Property

Public Property Let 多病人单(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.多病人单 = 5
    mvar多病人单 = vData
End Property

Public Property Get 多病人单() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.多病人单
    多病人单 = mvar多病人单
End Property

Public Property Let 加班标志(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.加班标志 = 5
    mvar加班标志 = vData
End Property

Public Property Get 加班标志() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.加班标志
    加班标志 = mvar加班标志
End Property

Public Property Let 门诊标志(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊标志 = 5
    mvar门诊标志 = vData
End Property

Public Property Get 门诊标志() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.门诊标志
    门诊标志 = mvar门诊标志
End Property

Public Property Let 费别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.费别 = 5
    mvar费别 = vData
End Property

Public Property Get 费别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.费别
    费别 = mvar费别
End Property

Public Property Let 床号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.床号 = 5
    mvar床号 = vData
End Property

Public Property Get 床号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.床号
    床号 = mvar床号
End Property

Public Property Let 标识号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标识号 = 5
    mvar标识号 = vData
End Property

Public Property Get 标识号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标识号
    标识号 = mvar标识号
End Property

Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property

Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.年龄
    年龄 = mvar年龄
End Property

Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property

Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.性别
    性别 = mvar性别
End Property

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.姓名
    姓名 = mvar姓名
End Property

Public Property Let 主页ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.主页ID = 5
    mvar主页ID = vData
End Property

Public Property Get 主页ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.主页ID
    主页ID = mvar主页ID
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    病人ID = mvar病人ID
End Property

Public Property Let NO(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.NO = 5
    mvarNO = vData
End Property

Public Property Get NO() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.NO
    NO = mvarNO
End Property

Public Property Let 病区ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病区ID = vData
End Property

Public Property Get 病区ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    病区ID = mvar病区ID
End Property

Public Property Let 科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar科室ID = vData
End Property

Public Property Get 科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    科室ID = mvar科室ID
End Property

Public Property Let 婴儿费(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.婴儿费 = 5
    mvar婴儿费 = vData
End Property

Public Property Get 婴儿费() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.婴儿费
    婴儿费 = mvar婴儿费
End Property
