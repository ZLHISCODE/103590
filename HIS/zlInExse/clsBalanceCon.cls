VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrTime As String  '病人结帐次数(初始="",可以为"1,2,3...")
Private mdtBeginDate As Date       '病人结帐的开始时间,初始为'1900-01-01'
Private mdtEndDate As Date         '病人结帐的结束时间,初始为'3000-01-01'
Private mstrDeptIDs As String      '病人结帐科室ID串(初始="",可以为"0,1,2,3...",0表示开单部门ID为空)
Private mstrClass As String       '费用类型=""-所有费用(含未设置),"'公费','比例',..."
Private mstrChargeType As String      '收费类别 '34260
Private mstrBaby As String      '是否仅结算婴儿费用(0-所有费用,1-病人费用,2及以上-第mbytbaby-1个婴儿费用)
Private mstrItem As String      '要结的收据费目
Private mbytKind As Byte       '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
Private mblnCurBalanceOwnerFee As Boolean      '当前是否正在结“自费费用”
Private mstrDiag As String
Private mlng病人ID As Long
'保持属性值的局部变量
Private mvarblnNormalFee As Boolean '局部复制
Private mvarblnHealthCheckFee As Boolean '局部复制
Public Property Let blnHealthCheckFee(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.blnHealthCheckFee = 5
    mvarblnHealthCheckFee = vData
End Property


Public Property Get blnHealthCheckFee() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnHealthCheckFee
    blnHealthCheckFee = mvarblnHealthCheckFee
End Property



Public Property Let blnNormalFee(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.blnNormalFee = 5
    mvarblnNormalFee = vData
End Property


Public Property Get blnNormalFee() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnNormalFee
    blnNormalFee = mvarblnNormalFee
End Property




Public Property Get lng病人ID() As Long
    lng病人ID = mlng病人ID
End Property

Public Property Let lng病人ID(ByVal vNewValue As Long)
    mlng病人ID = vNewValue
End Property

Public Property Get strTime() As String
    strTime = mstrTime
End Property

Public Property Let strTime(ByVal vNewValue As String)
    mstrTime = vNewValue
End Property

Public Property Get strDiag() As String
    strDiag = mstrDiag
End Property

Public Property Let strDiag(ByVal vNewValue As String)
    mstrDiag = vNewValue
End Property

Public Property Get strDeptIDs() As String
    strDeptIDs = mstrDeptIDs
End Property

Public Property Let strDeptIDs(ByVal vNewValue As String)
    mstrDeptIDs = vNewValue
End Property

Public Property Get strItem() As String
    strItem = mstrItem
End Property

Public Property Let strItem(ByVal vNewValue As String)
    mstrItem = vNewValue
End Property

Public Property Get strClass() As String
    strClass = mstrClass
End Property

Public Property Let strClass(ByVal vNewValue As String)
    mstrClass = vNewValue
End Property

Public Property Get strChargeType() As String
    strChargeType = mstrChargeType
End Property

Public Property Let strChargeType(ByVal vNewValue As String)
    mstrChargeType = vNewValue
End Property

Public Property Get strBaby() As String
    strBaby = mstrBaby
End Property

Public Property Let strBaby(ByVal vNewValue As String)
    mstrBaby = vNewValue
End Property

Public Property Get bytKind() As Byte
    bytKind = mbytKind
End Property

Public Property Let bytKind(ByVal vNewValue As Byte)
    mbytKind = vNewValue
End Property

Public Property Get blnCurBalanceOwnerFee() As Boolean
    blnCurBalanceOwnerFee = mblnCurBalanceOwnerFee
End Property

Public Property Let blnCurBalanceOwnerFee(ByVal vNewValue As Boolean)
    mblnCurBalanceOwnerFee = vNewValue
End Property

Public Property Get dtBeginDate() As Date
    dtBeginDate = mdtBeginDate
End Property

Public Property Let dtBeginDate(ByVal vNewValue As Date)
    mdtBeginDate = vNewValue
End Property

Public Property Get dtEndDate() As Date
    dtEndDate = mdtEndDate
End Property

Public Property Let dtEndDate(ByVal vNewValue As Date)
    mdtEndDate = vNewValue
End Property
