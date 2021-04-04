VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvar系统 As Long '局部复制

Private mvar编号 As String '局部复制
Private mvar名称 As String '局部复制
Private mvar说明 As String '局部复制
Private mvar打印机 As String '局部复制
Private mvar进纸 As Integer   '局部复制
Private mvar票据 As Boolean '局部复制
Private mvar打印方式 As Integer '局部复制
Private mvar修改时间 As Date '局部复制
Private mvarItems As RPTItems '局部复制
Private mvarDatas As RPTDatas '局部复制
Private mvarFmts As RPTFmts '局部复制
Private mvar禁止开始时间 As Date '局部复制
Private mvar禁止结束时间 As Date '局部复制

Private mvarbytFormat As Byte '局部复制
Private mvarblnLoad As Boolean  '局部复制
Private mvarintGridCount As Integer
Private mvarintGridID As Integer

Public Property Let 修改时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.修改时间 = 5
    mvar修改时间 = vData
End Property

Public Property Get 修改时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.修改时间
    修改时间 = mvar修改时间
End Property

Public Property Let 系统(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.intGridID = 5
    mvar系统 = vData
End Property

Public Property Get 系统() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.intGridID
    系统 = mvar系统
End Property

Public Property Let intGridID(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.intGridID = 5
    mvarintGridID = vData
End Property

Public Property Get intGridID() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.intGridID
    intGridID = mvarintGridID
End Property

Public Property Let intGridCount(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.intGridCount = 5
    mvarintGridCount = vData
End Property

Public Property Get intGridCount() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.intGridCount
    intGridCount = mvarintGridCount
End Property

Public Property Let bytFormat(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bytFormat = 5
    mvarbytFormat = vData
End Property

Public Property Get bytFormat() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bytFormat
    bytFormat = mvarbytFormat
End Property

Public Property Let blnLoad(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.blnLoad = 5
    mvarblnLoad = vData
End Property

Public Property Get blnLoad() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.blnLoad
    blnLoad = mvarblnLoad
End Property
'-----------------------------------------------------

Public Property Set Fmts(ByVal vData As RPTFmts)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Fmts = Form1
    Set mvarFmts = vData
End Property

Public Property Get Fmts() As RPTFmts
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Fmts
    Set Fmts = mvarFmts
End Property

Public Property Set Datas(ByVal vData As RPTDatas)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Datas = Form1
    Set mvarDatas = vData
End Property

Public Property Get Datas() As RPTDatas
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Datas
    Set Datas = mvarDatas
End Property

Public Property Set Items(ByVal vData As RPTItems)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Items = Form1
    Set mvarItems = vData
End Property

Public Property Get Items() As RPTItems
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Items
    On Error Resume Next
    Set Items = mvarItems
    On Error GoTo 0
End Property

Public Property Let 打印机(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.打印机 = 5
    mvar打印机 = vData
End Property

Public Property Get 打印机() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.打印机
    打印机 = mvar打印机
End Property

Public Property Let 进纸(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.进纸 = 5
    mvar进纸 = vData
End Property

Public Property Get 进纸() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.进纸
    进纸 = mvar进纸
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

Public Property Let 名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar名称 = vData
End Property

Public Property Get 名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    名称 = mvar名称
End Property

Public Property Let 编号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编号 = 5
    mvar编号 = vData
End Property

Public Property Get 编号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.编号
    编号 = mvar编号
End Property

Private Sub Class_Initialize()
    '自动初始化
    Set Datas = New RPTDatas
    Set Items = New RPTItems
    Set Fmts = New RPTFmts
End Sub

Public Property Let 票据(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.票据 = 5
    mvar票据 = vData
End Property

Public Property Get 票据() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.票据
    票据 = mvar票据
End Property

Public Property Let 打印方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.打印方式 = 5
    mvar打印方式 = vData
End Property

Public Property Get 打印方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.打印方式
    打印方式 = mvar打印方式
End Property

Public Property Let 禁止开始时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.修改时间 = 5
    mvar禁止开始时间 = vData
End Property

Public Property Get 禁止开始时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.修改时间
    禁止开始时间 = mvar禁止开始时间
End Property

Public Property Let 禁止结束时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.修改时间 = 5
    mvar禁止结束时间 = vData
End Property

Public Property Get 禁止结束时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.修改时间
    禁止结束时间 = mvar禁止结束时间
End Property
