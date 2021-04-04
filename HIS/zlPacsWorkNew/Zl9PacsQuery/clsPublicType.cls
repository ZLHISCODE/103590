VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'或用于PacsWork列表选择需要的数据  TPacsInfo
Public Type TStudyInfo
     lngPatId As Long                 '1 病人ID
     lngPageID As Long                  '2 主页ID
     lngAdviceID As Long                '3 医嘱ID
     lngSendNo As Long                  '4 发送号
     strPatientName As String           '5 病人姓名
     strPatientSex As String
     strPatientAge As String
     
     dtRefreshDate As Date
'Public strNO As String
'Public lngRecordKind As Long

'Public lngPatDept As Long                 '6 病人所属科室
'Public strRegNo As String                 '7 挂号单
'Public lngRegId As Long                   '8 挂号id
     intMoved As Integer                '9 是否转出
     intState As Integer                '10 检查状态
     intStep As Integer                 '11 检查过程
'Public lngUnit As Long                    '12 病区ID
'Public strStudyUID As String              '13 检查UID
'Public blnCanPrint As Boolean             '14 是否能够打印
'Public blnIsInsidePatient As Boolean      '15 是否门诊或住院病人
'Public lngExeDepartmentId As Long         '16 执行部门ID
'Public strDoDoctor As String              '17 检查技师
'Public strExeRoom As String               '18 执行间
     lngPatientFrom As Long             '19 病人来源
 strStudyNum As String              '20 检查号
'     strBedNum As String                '21 床号
     lngMarkNum As Double               '22 标志号
'Public lngBaby As Long                    '23 婴儿
'Public strPatientDepartment As String     '24 病人科室名称
'Public strReportDoctor As String          '25 报告人
'Public strReportOperation As String       '26 报告操作
     lngLinkId As Long                  '27 关联ID
'Public strImgType As String               '28 影像类别
'Public intImageLocation As Integer        '29 PACS影像所在的位置，0在中联PACS；1在新网PACS
'Public strPatientType As String           '30 病人类型
'Public intFilmGiveOut As Integer          '胶片发放
'Public intReportGiveOut As Integer        '报告发放
'Public intDangerState As Integer          '危急状态
'Public intEmergentTag As Integer          '紧急状态
'Public intGreenChannel As Integer         '绿色通道
'Public strAdviceContext As String             '医嘱内容
'Public strAdviceDepartAndMethod As String     '部位方法
     strStuStateDesc As String          '检查状态描述
'Public blnIsTechincalSure As Boolean      '是否技师确认
'Public strMoneyState As String            '费用状态描述
'Public blnIsReported As Boolean           '已经有报告
End Type


'控件类型
Public Enum TControlType
    ctText = 0  '文本框
    ctDate = 1  '日期框
    ctTime = 2  '时间框
    ctDateTime = 3  '日期时间框
    ctCombobox = 4  '下拉框
    ctList = 5      '多选框
    ctAgeCbx = 6    '年龄快选框
    ctMutxCbx = 7   '互斥框
    ctChk = 8       '可选框
    ctQueryWay = 9  '条件选择框
    ctFastDate = 10  '日期快选框
End Enum

'数据来源类型
Public Enum TDBFromType
    dbftText = 0
    dbftSql = 1
End Enum

''快速过滤选择类型
'Public Enum TFilterType
'    ftSelect = 0
'    ftFast = 1
'End Enum

'快速过滤选择类型
Public Enum TSelectWay
    swSingle = 0
    swMulti = 1
End Enum

Public Enum TInputType
    itPopup = 0
    itFast = 1
    itBoth = 2
End Enum

Public Enum TLikeWay
    lwNormal = 0    '正常
    lwLeft = 1      '左匹配
    lwRight = 2     '右匹配
    lwAll = 3       '全匹配
End Enum

'过滤窗口中录入条件基本信息结构
Public Type TInputTag
    ParName As String   '条件参数名称
    DisplayName As String   '界面显示名称，如[系统.开始日期]可能显示为[开始日期]
    FromType As TDBFromType    '数据来源类型0-文本，1-sql
    ExtProperty As String
    DataFrom As String  '数据来源配置内容
    Default As String   '默认值
    ControlType As Long '控件类型
    
    ParList() As String '所使用参数列表
    InputControl As Object  '对应的录入控件
    ReleationInputIndex() As Long    '关联录入配置索引位置
    Index As Long
    Tag As String
End Type

'查询方案基本信息结构
Public Type TSchemeBaseInfo
    SchemeId As Double    '方案ID
    Name As String      '方案名称
    Ver As Double       '方案版本
    Memo As String '方案说明
    IsDefault As Boolean '是否默认
    IsOften As Boolean  '是否常用
End Type

'方案结构
Public Type TSchemeItem
    IdKey As Double    '方案识别关键字
    BaseInfo As TSchemeBaseInfo '基本信息
    Index As Long           '索引位置
    IsLoaded As Boolean '是否载入完毕
    SqlScheme As clsSqlScheme   '方案配置对象
    SqlParse As clsSqlParse     'sql解析对象
    Startdate As Date
    EndDate As Date
    InitValues As clsSqlFilterValue
    FilterValues As clsSqlFilterValue
    FilterWindow As Object
End Type

'查询方案的过滤条件录入情况
Public Type FilterData
    Name As String
    Value As Variant
End Type

