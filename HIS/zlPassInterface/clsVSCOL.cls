VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSCOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mintCOLID As Integer '局部复制
Private mintCOL相关ID As Integer '局部复制
Private mintCOL诊疗类别 As Integer '局部复制
Private mIntCOL诊疗项目ID As Integer '局部复制
Private mintCOL收费细目ID As Integer '局部复制
Private mintCOL医嘱内容 As Integer '局部复制
Private mintCOL期效 As Integer '局部复制
Private mintCOL单量 As Integer '局部复制
Private mintCOL单量单位 As Integer '局部复制
Private mintCOL用法 As Integer

Private mintCOL天数 As Integer
Private mintCOL婴儿 As Integer
Private mintCOL开嘱时间 As Integer
Private mintCOL开嘱医生 As Integer
Private mintCOL开始时间 As Integer
Private mintCOL开嘱科室ID As Integer
Private mintCOL终止时间 As Integer
Private mintCOL频率 As Integer
Private mintCOL频率次数 As Integer
Private mintCOL频率间隔 As Integer

Private mintCOL间隔单位 As Integer
Private mintCOL警示 As Integer
Private mintCOL序号 As Integer
Private mintCOL状态 As Integer
Private mintCOLEDIT As Integer
Private mintCOL病人ID As Integer
Private mintCOL主页ID As Integer
Private mintCOL选择 As Integer       'gbytUseType =1时传人

Private mintCOL执行性质 As Integer   '
Private mintCOL药品名称 As Integer   '中西成药名称
Private mintCOL总量 As Integer
Private mintCol总量单位 As Integer
Private mintCol医嘱嘱托 As Integer
Private mintCol用药目的 As Integer
Private mintCol操作类型 As Integer
Private mintCol禁忌药品说明 As Integer
Private mintCol用药理由  As Integer
Private mintCol标志  As Integer
Private mintCol处方号  As Integer
Private mintCol执行分类  As Integer
Private mintCol执行科室ID  As Integer

'为药品模块提供
Private mintCOL单据 As Integer
Private mintCOLNO As Integer
Private mintCOL住院号 As Integer
Private mintCOL门诊号 As Integer
Private mintCOL审查结果 As Integer

Public Property Let intCOL单量单位(ByVal intData As Integer)
    mintCOL单量单位 = intData
End Property

Public Property Get intCOL单量单位() As Integer
    intCOL单量单位 = mintCOL单量单位
End Property

Public Property Let intCOL单量(ByVal intData As Integer)
    mintCOL单量 = intData
End Property

Public Property Get intCOL单量() As Integer
    intCOL单量 = mintCOL单量
End Property

Public Property Let intCOL期效(ByVal intData As Integer)
    mintCOL期效 = intData
End Property

Public Property Get intCOL期效() As Integer
    intCOL期效 = mintCOL期效
End Property

Public Property Let intCOL医嘱内容(ByVal intData As Integer)
    mintCOL医嘱内容 = intData
End Property

Public Property Get intCOL医嘱内容() As Integer
    intCOL医嘱内容 = mintCOL医嘱内容
End Property

Public Property Let intCOL收费细目ID(ByVal intData As Integer)
    mintCOL收费细目ID = intData
End Property

Public Property Get intCOL收费细目ID() As Integer
    intCOL收费细目ID = mintCOL收费细目ID
End Property

Public Property Let intCOL诊疗项目ID(ByVal intData As Integer)
    mIntCOL诊疗项目ID = intData
End Property

Public Property Get intCOL诊疗项目ID() As Integer
    intCOL诊疗项目ID = mIntCOL诊疗项目ID
End Property

Public Property Let intCOL诊疗类别(ByVal intData As Integer)
    mintCOL诊疗类别 = intData
End Property

Public Property Get intCOL诊疗类别() As Integer
    intCOL诊疗类别 = mintCOL诊疗类别
End Property

Public Property Let intCOL相关ID(ByVal intData As Integer)
    mintCOL相关ID = intData
End Property

Public Property Get intCOL相关ID() As Integer
    intCOL相关ID = mintCOL相关ID
End Property

Public Property Let intCOLID(ByVal intData As Integer)
    mintCOLID = intData
End Property

Public Property Get intCOLID() As Integer
    intCOLID = mintCOLID
End Property

Public Property Let intCOL用法(ByVal intData As Integer)
    mintCOL用法 = intData
End Property


Public Property Get intCOL用法() As Integer
    intCOL用法 = mintCOL用法
End Property

Public Property Let intCOL天数(ByVal intData As Integer)
    mintCOL天数 = intData
End Property

Public Property Get intCOL天数() As Integer
    intCOL天数 = mintCOL天数
End Property

Public Property Let intCOL婴儿(ByVal intData As Integer)
    mintCOL婴儿 = intData
End Property

Public Property Get intCOL婴儿() As Integer
    intCOL婴儿 = mintCOL婴儿
End Property

Public Property Let intCOL开嘱时间(ByVal intData As Integer)
    mintCOL开嘱时间 = intData
End Property

Public Property Get intCOL开嘱时间() As Integer
    intCOL开嘱时间 = mintCOL开嘱时间
End Property

Public Property Let intCOL开嘱医生(ByVal intData As Integer)
    mintCOL开嘱医生 = intData
End Property

Public Property Get intCOL开嘱医生() As Integer
    intCOL开嘱医生 = mintCOL开嘱医生
End Property

Public Property Let intCOL开始时间(ByVal intData As Integer)
    mintCOL开始时间 = intData
End Property

Public Property Get intCOL开始时间() As Integer
    intCOL开始时间 = mintCOL开始时间
End Property

Public Property Let intCOL开嘱科室ID(ByVal intData As Integer)
    mintCOL开嘱科室ID = intData
End Property

Public Property Get intCOL开嘱科室ID() As Integer
    intCOL开嘱科室ID = mintCOL开嘱科室ID
End Property

Public Property Let intCOL终止时间(ByVal intData As Integer)
    mintCOL终止时间 = intData
End Property

Public Property Get intCOL终止时间() As Integer
    intCOL终止时间 = mintCOL终止时间
End Property

Public Property Let intCOL频率(ByVal intData As Integer)
    mintCOL频率 = intData
End Property

Public Property Get intCOL频率() As Integer
    intCOL频率 = mintCOL频率
End Property

Public Property Let intCOL频率次数(ByVal intData As Integer)
    mintCOL频率次数 = intData
End Property

Public Property Get intCOL频率次数() As Integer
    intCOL频率次数 = mintCOL频率次数
End Property

Public Property Let intCOL频率间隔(ByVal intData As Integer)
    mintCOL频率间隔 = intData
End Property

Public Property Get intCOL频率间隔() As Integer
    intCOL频率间隔 = mintCOL频率间隔
End Property

Public Property Let intCOL间隔单位(ByVal intData As Integer)
    mintCOL间隔单位 = intData
End Property

Public Property Get intCOL间隔单位() As Integer
    intCOL间隔单位 = mintCOL间隔单位
End Property

Public Property Let intCOL警示(ByVal intData As Integer)
    mintCOL警示 = intData
End Property

Public Property Get intCOL警示() As Integer
    intCOL警示 = mintCOL警示
End Property

Public Property Let intCOL序号(ByVal intData As Integer)
    mintCOL序号 = intData
End Property

Public Property Get intCOL序号() As Integer
    intCOL序号 = mintCOL序号
End Property

Public Property Let intCOL状态(ByVal intData As Integer)
    mintCOL状态 = intData
End Property

Public Property Get intCOL状态() As Integer
    intCOL状态 = mintCOL状态
End Property

Public Property Let intCOLEDIT(ByVal intData As Integer)
    mintCOLEDIT = intData
End Property

Public Property Get intCOLEDIT() As Integer
    intCOLEDIT = mintCOLEDIT
End Property

Public Property Let intCOL病人ID(ByVal intData As Integer)
    mintCOL病人ID = intData
End Property

Public Property Get intCOL病人ID() As Integer
    intCOL病人ID = mintCOL病人ID
End Property

Public Property Let intCOL主页ID(ByVal intData As Integer)
    mintCOL主页ID = intData
End Property

Public Property Get intCOL主页ID() As Integer
    intCOL主页ID = mintCOL主页ID
End Property

Public Property Let intCOL选择(ByVal intData As Integer)
    mintCOL选择 = intData
End Property

Public Property Get intCOL选择() As Integer
    intCOL选择 = mintCOL选择
End Property

Public Property Let intCOL执行性质(ByVal intData As Integer)
    mintCOL执行性质 = intData
End Property

Public Property Get intCOL执行性质() As Integer
    intCOL执行性质 = mintCOL执行性质
End Property

'中西成药名称
Public Property Let intCOL药品名称(ByVal intData As Integer)
    mintCOL药品名称 = intData
End Property

Public Property Get intCOL药品名称() As Integer
    intCOL药品名称 = mintCOL药品名称
End Property
'为药品模块提供
Public Property Let intCOL单据(ByVal intData As Integer)
    mintCOL单据 = intData
End Property

Public Property Get intCOL单据() As Integer
    intCOL单据 = mintCOL单据
End Property

Public Property Let intCOLNO(ByVal intData As Integer)
    mintCOLNO = intData
End Property

Public Property Get intCOLNO() As Integer
    intCOLNO = mintCOLNO
End Property

Public Property Let intCOL住院号(ByVal intData As Integer)
    mintCOL住院号 = intData
End Property

Public Property Get intCOL住院号() As Integer
    intCOL住院号 = mintCOL住院号
End Property

Public Property Let intCOL门诊号(ByVal intData As Integer)
    mintCOL门诊号 = intData
End Property

Public Property Get intCOL门诊号() As Integer
    intCOL门诊号 = mintCOL门诊号
End Property

Public Property Let intCOL审查结果(ByVal intData As Integer)
    mintCOL审查结果 = intData
End Property

Public Property Get intCOL审查结果() As Integer
    intCOL审查结果 = mintCOL审查结果
End Property

Public Property Let intCOL总量(ByVal intData As Integer)
    mintCOL总量 = intData
End Property

Public Property Get intCOL总量() As Integer
    intCOL总量 = mintCOL总量
End Property

Public Property Let intcol总量单位(ByVal intData As Integer)
    mintCol总量单位 = intData
End Property

Public Property Get intcol总量单位() As Integer
    intcol总量单位 = mintCol总量单位
End Property

Public Property Let intcol医嘱嘱托(ByVal intData As Integer)
    mintCol医嘱嘱托 = intData
End Property

Public Property Get intcol医嘱嘱托() As Integer
    intcol医嘱嘱托 = mintCol医嘱嘱托
End Property

Public Property Let intcol用药目的(ByVal intData As Integer)
    mintCol用药目的 = intData
End Property

Public Property Get intcol用药目的() As Integer
    intcol用药目的 = mintCol用药目的
End Property

Public Property Let intcol用药理由(ByVal intData As Integer)
    mintCol用药理由 = intData
End Property

Public Property Get intcol用药理由() As Integer
    intcol用药理由 = mintCol用药理由
End Property

'操作类型
Public Property Let intCol操作类型(ByVal intData As Integer)
    mintCol操作类型 = intData
End Property

Public Property Get intCol操作类型() As Integer
    intCol操作类型 = mintCol操作类型
End Property

Public Property Let intCol禁忌药品说明(ByVal intData As Integer)
    mintCol禁忌药品说明 = intData
End Property

Public Property Get intCol禁忌药品说明() As Integer
    intCol禁忌药品说明 = mintCol禁忌药品说明
End Property

Public Property Let intCol标志(ByVal intData As Integer)
    mintCol标志 = intData
End Property

Public Property Get intCol标志() As Integer
    intCol标志 = mintCol标志
End Property

Public Property Let intCol处方号(ByVal intData As Integer)
    mintCol处方号 = intData
End Property

Public Property Get intCol处方号() As Integer
    intCol处方号 = mintCol处方号
End Property

Public Property Let intCol执行分类(ByVal intData As Integer)
    mintCol执行分类 = intData
End Property

Public Property Get intCol执行分类() As Integer
    intCol执行分类 = mintCol执行分类
End Property

Public Property Let intCol执行科室ID(ByVal intData As Integer)
    mintCol执行科室ID = intData
End Property

Public Property Get intCol执行科室ID() As Integer
    intCol执行科室ID = mintCol执行科室ID
End Property
