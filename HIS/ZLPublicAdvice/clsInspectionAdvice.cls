VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInspectionAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明

'保持属性值的局部变量
Private mvar病人ID As Long '
Private mvar主页ID As Long
Private mvar挂号单 As String
Private mvar婴儿 As Long
Private mvar病人来源 As Long '：1-门诊，2-住院,3-外来,4-体检
Private mvar采集项目ID As Long '
Private mvar检验项目IDs As String '检验项目1:是否组合项目,检验项目2:是否组合项目,···
Private mvar标本 As String '
Private mvar嘱托 As String '
Private mvar执行科室ID As Long
Private mvar采集科室ID As Long
Private mvar开始时间 As Date '医嘱开始执行时间
Private mvar病人科室ID As Long
Private mvar开单科室ID As Long
Private mvar开嘱医生 As String
Private mvar开嘱时间 As Date
Private mvar紧急标志 As Long '0-普通;1-紧急;2-补录(对门诊无效)
Private mvar申请附项串 As String '格式="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>…
Private mvar申请序号 As Long  '可不传入
Private mvar关联诊断IDs As String
Private mvar修改医嘱ID As Long

'------------------------------------------------------------------------
Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    病人ID = mvar病人ID
End Property

Public Property Let 主页ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar主页ID = vData
End Property

Public Property Get 主页ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    主页ID = mvar主页ID
End Property


Public Property Let 挂号单(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar挂号单 = vData
End Property

Public Property Get 挂号单() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    挂号单 = mvar挂号单
End Property

Public Property Let 病人来源(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar病人来源 = vData
End Property

Public Property Get 病人来源() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    病人来源 = mvar病人来源
End Property

 
Public Property Let 采集项目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar采集项目ID = vData
End Property

Public Property Get 采集项目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    采集项目ID = mvar采集项目ID
End Property

Public Property Let 检验项目IDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar检验项目IDs = vData
End Property

Public Property Get 检验项目IDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    检验项目IDs = mvar检验项目IDs
End Property


Public Property Let 标本(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar标本 = vData
End Property

Public Property Get 标本() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    标本 = mvar标本
End Property

Public Property Let 嘱托(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar嘱托 = vData
End Property

Public Property Get 嘱托() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    嘱托 = mvar嘱托
End Property

Public Property Let 申请附项串(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar申请附项串 = vData
End Property

Public Property Get 申请附项串() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    申请附项串 = mvar申请附项串
End Property

Public Property Let 执行科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar执行科室ID = vData
End Property

Public Property Get 执行科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    执行科室ID = mvar执行科室ID
End Property

Public Property Let 采集科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar采集科室ID = vData
End Property

Public Property Get 采集科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    采集科室ID = mvar采集科室ID
End Property

Public Property Let 病人科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar病人科室ID = vData
End Property

Public Property Get 病人科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    病人科室ID = mvar病人科室ID
End Property

Public Property Let 开单科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar开单科室ID = vData
End Property

Public Property Get 开单科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    开单科室ID = mvar开单科室ID
End Property

Public Property Let 紧急标志(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar紧急标志 = vData
End Property

Public Property Get 紧急标志() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    紧急标志 = mvar紧急标志
End Property

Public Property Let 开始时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar开始时间 = vData
End Property

Public Property Get 开始时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    开始时间 = mvar开始时间
End Property


Public Property Let 开嘱时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar开嘱时间 = vData
End Property

Public Property Get 开嘱时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    开嘱时间 = mvar开嘱时间
End Property


Public Property Let 开嘱医生(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar开嘱医生 = vData
End Property

Public Property Get 开嘱医生() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    开嘱医生 = mvar开嘱医生
End Property

Public Property Let 申请序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar申请序号 = vData
End Property

Public Property Get 申请序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    申请序号 = mvar申请序号
End Property

Public Property Let 关联诊断IDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar关联诊断IDs = vData
End Property

Public Property Get 关联诊断IDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    关联诊断IDs = mvar关联诊断IDs
End Property

Public Property Let 修改医嘱ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar修改医嘱ID = vData
End Property

Public Property Get 修改医嘱ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    修改医嘱ID = mvar修改医嘱ID
End Property

Public Property Let 婴儿(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar婴儿 = vData
End Property

Public Property Get 婴儿() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    婴儿 = mvar婴儿
End Property

