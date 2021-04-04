VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliniDelBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar病人类型 As String '局部复制
Private mvarCurDelNos As String '局部复制
Private mvarAllNos As String '局部复制
Private mvarPrintNOs As String '局部复制
Private mvar退费时间 As Date   '局部复制
Private mvar部分退费 As Boolean   '局部复制
Private mvarShareUserID As Long '局部复制
Private mvarPatiUseType As String '局部复制
Private mvar费别 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar退费结算 As String '局部复制
Private mvar当前发票号 As String '局部复制
Private mvar回收发票 As String '局部复制
Private mvarintInsure As Integer '局部复制
Private mvar医保结算金额 As Double '局部复制
Private mvar退费合计 As Double '局部复制
Private mvar医保不走票号 As Boolean   '局部复制
Private mvar缴款 As Double '局部复制
Private mvar找补 As Double '局部复制
Private mvar结帐ID As Long '局部复制
Private mvar冲销ID As Long  '局部复制
Private mvar原结帐ID As Long '局部复制
Private mvar结算序号 As Long '局部复制
Private mvarrsBalance As ADODB.Recordset   '局部复制
Private mvarrs结算方式 As ADODB.Recordset   '局部复制
Private mvar缺省结算方式 As String '局部复制
Private mvarSaveBilled As Boolean   '局部复制
Private mvar原样退 As Boolean   '局部复制
Private mvar家属IDs As String '局部复制
Private mvarblnOnePatiPrint As Boolean   '局部复制
Private mvarstrOnePatiPrintNos As String     '局部复制
Public Property Let strOnePatiPrintNos(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strOnePatiPrintNos = 5
    mvarstrOnePatiPrintNos = vData
End Property


Public Property Get strOnePatiPrintNos() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strOnePatiPrintNos
    strOnePatiPrintNos = mvarstrOnePatiPrintNos
End Property

Public Property Let blnOnePatiPrint(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.blnOnePatiPrint = 5
    mvarblnOnePatiPrint = vData
End Property


Public Property Get blnOnePatiPrint() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnOnePatiPrint
    blnOnePatiPrint = mvarblnOnePatiPrint
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


Public Property Let 原样退(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原样退 = 5
    mvar原样退 = vData
End Property


Public Property Get 原样退() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原样退
    原样退 = mvar原样退
End Property

Public Property Let SaveBilled(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SaveBilled = 5
    mvarSaveBilled = vData
End Property


Public Property Get SaveBilled() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SaveBilled
    SaveBilled = mvarSaveBilled
End Property
Public Property Let AllNos(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.AllNos = 5
    mvarAllNos = vData
End Property
Public Property Get AllNos() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.AllNos
    AllNos = mvarAllNos
End Property
Public Property Let CurDelNos(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CurDelNos = 5
    mvarCurDelNos = vData
End Property


Public Property Get CurDelNos() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CurDelNos
    CurDelNos = mvarCurDelNos
End Property
 
Public Property Let 回收发票(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.回收发票 = 5
    mvar回收发票 = vData
End Property


Public Property Get 回收发票() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.回收发票
    回收发票 = mvar回收发票
End Property
Public Property Let 退费时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退费时间 = 5
    mvar退费时间 = vData
End Property


Public Property Get 退费时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退费时间
    退费时间 = mvar退费时间
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


Public Property Let 退费合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退费合计 = 5
    mvar退费合计 = vData
End Property


Public Property Get 退费合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退费合计
    退费合计 = mvar退费合计
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
Public Property Let 退费结算(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退费结算 = 5
    mvar退费结算 = vData
End Property
Public Property Get 退费结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退费结算
    退费结算 = mvar退费结算
End Property
 
Public Property Let PrintNOs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PrintNOs = 5
    mvarPrintNOs = vData
End Property
Public Property Get PrintNOs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PrintNOs
    PrintNOs = mvarPrintNOs
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
 

Public Property Let 病人类型(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人类型 = 5
    mvar病人类型 = vData
End Property


Public Property Get 病人类型() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人类型
    病人类型 = mvar病人类型
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
Public Property Let 原结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原结帐ID = 5
    mvar原结帐ID = vData
End Property

Public Property Get 原结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原结帐ID
    原结帐ID = mvar原结帐ID
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

 

Public Property Get rsBalance() As ADODB.Recordset
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.rsBalance
   Set rsBalance = mvarrsBalance
End Property

Public Property Set rsBalance(ByVal vData As ADODB.Recordset)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.rs结算方式 = 5
    Set mvarrsBalance = vData
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
Public Property Let 部分退费(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.部分退费 = 5
      mvar部分退费 = vData
End Property

Public Property Get 部分退费() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.部分退费
    部分退费 = mvar部分退费
End Property

 
