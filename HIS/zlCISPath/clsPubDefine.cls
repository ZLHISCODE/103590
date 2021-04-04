VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'路径版本
Public Type TYPE_PATH_VERSION
    版本号          As Integer
    标准住院日      As String           '住院才有
    标准治疗时间      As String           '门诊才有
    标准费用        As String
    版本说明        As String
    创建人          As String
    创建时间        As Date
    审核人          As String
    审核时间        As Date
    停用人          As String
    停用时间        As Date
    药剂科审核人 As String
    药剂科审核时间 As Date
End Type

'分支路径
Public Type TYPE_PATH_BRANCH
    分支ID As Long
    分支名称 As String
    版本号 As Integer
    前一阶段ID As Long
    前一阶段名称 As String
    标准住院日 As String
    标准费用 As String
    说明 As String
    创建人 As String
    创建时间 As Date
End Type

'评估指标
Public Type TYPE_PATH_EvalMark
    ID As Long
    序号 As Integer
    评估指标 As String
    指标类型 As Byte
    指标结果 As String
End Type

'评估条件
Public Type TYPE_PATH_EvalCond
    指标ID As Long
    项目ID As Long
    关系式 As String
    条件值 As String
    条件组合 As Byte
End Type

'路径评估
Public Type TYPE_PATH_EVAL
    指标集 As Collection
    条件集 As Collection
End Type

'时间阶段
Public Type TYPE_PATH_STEP
    ID As Long
    父ID As Long
    序号 As Integer
    名称 As String
    开始天数 As Integer
    结束天数 As Integer
    标志 As String
    分类 As String
    说明 As String
    评估 As TYPE_PATH_EVAL
    Edit As Integer '0-原始,1-新增,2-修改
    存在分支 As Boolean
End Type

'路径项目
Public Type TYPE_PATH_ITEM
    ID As Long
    项目序号 As Integer
    项目内容 As String
    执行方式 As Integer
    执行者 As Integer
    生成者 As Integer
    项目结果 As String
    图标ID As Long
    病历IDs As String '(老版)ID1,ID2,...数字型
    新版病历IDs As String  ' ID1,ID2,....字符型号
    病历详情 As String  '文件ID1,原型ID1,名称1,序号1;文件ID2,原型ID2,名称2,序号2;...
    医嘱IDs As String 'ID1,ID2,...
    原医嘱IDs As String 'ID1,ID2,...  已审核版本的路径才传人该值
    待审核医嘱IDs As String 'ID1,ID2,...  已审核版本的路径才传人该值
    Edit As Integer '0-原始,1-新增,2-修改
    Tip As String '用于提示临时使用
    内容要求 As Integer '0-全部生成，1-选择生成
    变异原因 As String
    导入参考 As String
    导入结果 As Integer
    前一版本项目ID As Long '显示差异的时候赋值
    审核状态 As Byte       '1-药剂审核;2-医务审核;3-双审核
End Type


'病人路径
Public Type TYPE_PATH_Pati
    路径ID As Long
    版本号 As Long
    病人路径ID As Long
    病人路径状态 As Long    '-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
    当前阶段ID As Long
    阶段父ID As Long
    当前天数 As Long
    当前日期 As String
    未导入原因 As String
    原路径ID As Long        '路径跳转时，之前的路径ID
    当前阶段分支ID As Long
    结束路径控制 As Long
    合并路径个数 As Long
    导入时间  As Date
End Type

'病人信息
Public Type TYPE_Pati
    病人ID          As Long
    主页ID          As Long             '住院病人才有
    挂号ID          As Long             '门诊病人才有
    挂号NO          As String           '门诊病人才有
    病区ID          As Long
    科室ID          As Long
    病人状态        As Integer          '病案主页中的状态：0-正常住院，1-尚未入科，2-正在转科，3-已预出院；
                                  '病人挂号记录里面的执行状态：0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
End Type
