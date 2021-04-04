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

'------------------------
'审查接口病人相关信息
Public Type TYPE_PASSPati
    lng病人ID As Long
    lng主页ID As Long
    lng挂号ID As Long
    dbl标识号 As Double  '门诊传门诊号,住院 传住院号
    str姓名 As String
    str性别 As String
    str床号  As String
    int婴儿 As Integer
    str挂号单 As String '门诊
    Dat出生日期 As Date
End Type

'诊断信息
Public Type TYPE_PASSDiag
    objDiag As Object
    lng疾病编码 As Long
    lng诊断编码 As Long
    lng诊断描述 As Long
End Type

'映射vsAdvice列号,未映射的列传值为-1
Public Type TYPE_PASSCOL
    intCOLID As Integer
    intCOL相关ID As Integer
    intCOL诊疗类别 As Integer
    intCOL诊疗项目ID As Integer
    intCOL收费细目ID As Integer
    intCOL医嘱内容 As Integer
    intCOL期效 As Integer       '住院
    intCOL单量 As Integer
    intCOL单量单位 As Integer
    intCOL用法 As Integer
    intCOL天数 As Integer
    intCOL婴儿 As Integer
    intCOL开嘱时间 As Integer
    intCOL开嘱医生 As Integer
    intCOL开始时间 As Integer
    intCOL开嘱科室ID As Integer
    intCOL终止时间 As Integer
    intCOL频率 As Integer
    intCOL频率次数 As Integer
    intCOL频率间隔 As Integer
    intCOL间隔单位 As Integer
    intCOL警示 As Integer
    intCOL序号 As Integer
    intCOL状态 As Integer
    intCOLEDIT As Integer
    intCOL病人ID As Integer
    intCOL主页ID As Integer
    intCOL选择 As Integer       'gbytUseType =1时传人
    
    intCOL执行性质 As Integer   '
    intCOL药品名称 As Integer   '中西成药名称
    '为药品模块提供
    intCOL单据 As Integer
    intCOLNO As Integer
    intCOL住院号 As Integer
    intCOL门诊号 As Integer
    intCOL审查结果 As Integer
End Type

'传人接口的模块参数
Public Type TYPE_PASS_MainPara
    lngModel As Long                 '调用环境
    udtCOL As TYPE_PASSCOL          '医嘱列映射
    objAdvice As Object             '医嘱列表对象 vsAdvice
    objCommandBar As Object         '菜单对象
    objCmdAlley As Object           'PASS过敏史按钮
    bytUseType As Byte              'bytUseType =1 临床路径生成医嘱时调用住院医嘱编辑界面
    frmMain As Object               '主窗体
    '诊断信息  门诊传人
    udtDiag As TYPE_PASSDiag        '门诊诊断对象
End Type

'----------------------------
'调用模块编号
Public Enum Enu_Model
    m门诊编辑 = 0
    M住院编辑 = 1
    M住院医嘱清单 = 2
    M护士校对 = 3
    M门诊医嘱清单 = 4
    M住院首页 = 5
    M处方发药 = 1341        '1341    药品处方发药
    M部门发药 = 1342        '1342    药品部门发药
    MPIVA管理 = 1345        '1345    PIVA管理
End Enum

