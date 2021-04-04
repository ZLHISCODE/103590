VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------
'太元通接口结构
'--------------------------------------------------------------------------------------------------------------------------------------
'药嘱结构
Public Type PatDrug
    drugID As String                'his 系统的药品代码不为空
    DrugName As String              'his 系统的药品名称不为空
    recMainNo As String             'his 系统的医嘱组号，在一次就诊/住院中唯一，只能为数字编码，不为空
    recSubNo As String              'his 系统的医嘱序号，在一次就诊/住院中唯一，只能为数字编码,不为空
    dosage As String                'his 系统的医嘱药品使用剂量不为空
    doseUnits As String             'his 系统的医嘱药品剂量单位不为空
    administrationID As String      'his 系统的医嘱途径代码不为空
    performFreqDictID As String     'his 系统的医嘱频次代码不为空
    performFreqDictText As String   'his 系统的医嘱执行频率描述不为空
    startDateTime As String         'his 系统的医嘱开始时间,格式 YYYY-MM-DD HH: MM: SS 不为空
    stopDateTime As String          'his 系统的医嘱结束时间,格式 YYYY-MM-DD HH: MM: SS
    doctorDept As String            'his 系统的开医嘱医生所在科室代码
    DoctorID As String              'his 系统的开医嘱医生编码
    Doctor As String                'his 系统的开医嘱医生姓名,
    isNew As String                 '新增医嘱值为1；否则为0
End Type
'诊断信息
Public Type PatDiagnosis
    diagnosisID As String          'his 系统的诊断编码
    diagnosisName As String          'his 系统的诊断名称
    diagnosisType As String          'his 系统的诊断类型，如门诊诊断、入院诊断等
End Type
'过敏源
Public Type PatDrugSensitive
    patOrderDrugSensitiveID As String   '默认为0
    drugAllergenID As String        'his 系统的过敏编码
End Type
'症状数据结构
Public Type PatSymptom
    symptomID As String           'his 系统的症状编码
    symptomName As String           'his 系统的症状名称
End Type

'审查传人信息
Public Type PatientOrder
    PatientID As String        '病人标识号不为空
    Pname As String             '病人姓名不为空
    pSex As String              '病人性别 1-女, 0-男, 2-不详
    pdateOfBirth As String      '病人出生日期，格式 YYYY-MM-DD 不为空（必填）
    pWeight As String           ' 病人体重
    pHeight As String           ' 病人身高

    isLact As String                '是否哺乳，是为1，否为0 不为空
    isPregnant As String            '是否孕妇，是为1 ，否为0 不为空
    isLiverWhole As String          '是否肝功异常 1-异常，0-正常不为空
    isKidneyWhole As String         '是否肾功异常 1-异常，0-正常不为空
    PvisitID As String               '就诊序号，唯一标识患者的一次就诊(门诊号\住院号)

    PatDrugs As Variant              '病人药嘱详情
    PatDiagnoses As Variant   '诊断详情
    PatDrugSensitives As Variant   '过敏记录
    PatSymptoms As Variant       '症状详情

    DoctDeptID As String           '登陆医生部门代码不为空
    DoctDeptName As String         '登陆医生部门名称不为空
    DoctID As String               '登陆医生代码不为空
    DoctName As String             '登陆医生姓名不为空
    DoctTitleID As String         '登陆医生职称类别不为空'C --副教授；教授；副主任医师；主任医师；专家
    'B—主治医师；讲师  'A—除以上的其他职称
    DoctTitleName As String          '登陆医生职称名称不为空
    SysFlag As String              '1— 门诊医生工作站 2— 住院医生工作
End Type

'审查结果集<table></table>
Public Type AuditResult
    recMainNo As String         '医嘱组号
    recSubNo As String           '医嘱序号
    alertLevel As String         '警示级别 1— 禁忌（红色警示灯）2— 慎用（黄色警示灯示) 3— 提示（蓝色警示灯）
    strChecksum As String        '审查结果
End Type
