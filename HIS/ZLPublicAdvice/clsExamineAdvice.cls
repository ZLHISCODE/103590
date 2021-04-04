VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExamineAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明

'保持属性值的局部变量
Private mvar病人ID As Long '
Private mvar主页ID As Long
Private mvar婴儿 As Long
Private mvar挂号单 As String
Private mvar病人来源 As Long '：1-门诊，2-住院,3-外来,4-体检
Private mvar检查项目ID As Long '局部复制
Private mvar部位方法 As String '部位名1;方法名1,方法名2|部位名2;方法名1,方法名2
Private mvar执行类型 As Long '0-常规,1-床旁,2-术中,3-院外执行
Private mvar执行科室ID As Long
Private mvar开始时间 As Date '医嘱开始执行时间
Private mvar病人科室ID As Long
Private mvar开单科室ID As Long
Private mvar开嘱医生 As String
Private mvar开嘱时间 As Date
Private mvar紧急标志 As Long '0-普通;1-紧急;2-补录(对门诊无效)
Private mvar申请附项串 As String '格式="项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>…
Private mvar申请序号 As Long  '可不传入
Private mvar执行科室性质 As Long '0-无执行的叮嘱；1-病人所在科室；2-病人所在病区；3-开单人所在科室；4-指定科室;5-院外执行；
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

 
Public Property Let 检查项目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar检查项目ID = vData
End Property

Public Property Get 检查项目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    检查项目ID = mvar检查项目ID
End Property

Public Property Let 部位方法(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar部位方法 = vData
End Property

Public Property Get 部位方法() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    部位方法 = mvar部位方法
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

Public Property Let 执行类型(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar执行类型 = vData
End Property

Public Property Get 执行类型() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    执行类型 = mvar执行类型
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

Public Property Let 执行科室性质(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar执行科室性质 = vData
End Property

Public Property Get 执行科室性质() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    执行科室性质 = mvar执行科室性质
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

