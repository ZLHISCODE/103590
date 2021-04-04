VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_HZYY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --接口定义 【杭州逸曜信息技术有限公司】
'--------------------------------------------------------------------------------------------------------------------------------------
'<-begin------------------------------------------------------------------------basexml----------------------------------------------------->

Public Type HZYYBASE
    strHospCode         As String
    strPatiID           As String
    strEventNO          As String
    strSource           As String
End Type
'<-begin------------------------------------------------------------------------1.2 optpatient门诊患者就诊标签格式如下:----------------------------------------------------->
Public Type OPTPATIENT
    strSex              As String
    strName             As String
    strIDType           As String
    strIDNO             As String
    strBirthWeight      As String           '出生体重
    strBirthDay         As String           '出生日期
    strEthnicGroup      As String           '民族
    strNativePlace      As String           '籍贯
    strRace             As String           '人种
    strMedCardNO        As String           '就诊卡号
    strEventTime       As String           '就诊时间
    strDeptID           As String           '挂号科室ID
    strDeptName         As String           '挂号科室名称
    strPayType          As String           '费用类型
    strPregnancy        As String           '是否怀孕
    strTimeOfPreg       As String           '孕期
    strBreastFeeding    As String          '是否哺乳
    strHeight           As String                 '身高
    strWeight           As String                 '体重
    strAddress          As String                '病人地址
    strPhoneNo          As String                '病人电话
    strDialysis         As String               '是否透析
    strmarital          As String                 '婚姻状况
    strOccupation       As String             '职业
    strSpecialConstitution As String    '特殊体质
    strVisitType        As String              '就诊类别
    strPatiCondition    As String          '患者状态
End Type

'<-begin------------------------------------------------------------------------1.3 iptpatient住院患者就诊标签:----------------------------------------------------->
Public Type IPTPATIENT
    strSex                      As String
    strName                     As String
    strIDType                   As String
    strIDNO                     As String
    strBirthWeight              As String           '出生体重
    strBirthDay                 As String           '出生日期
    strEthnicGroup              As String           '民族
    strNativePlace              As String           '籍贯
    strRace                     As String           '人种
    strMedCardNO                As String           '就诊卡号
    strPayType                  As String           '费用类型
    strPregnancy                As String           '是否怀孕
    strTimeOfPreg               As String           '孕期
    strBreastFeeding            As String           '是否哺乳
    strHeight                   As String           '身高
    strWeight                   As String           '体重
    strAddress                  As String           '病人地址
    strPhoneNo                  As String           '病人电话
    strDialysis                 As String           '是否透析
    strmarital                  As String           '婚姻状况
    strOccupation               As String           '职业
    strSpecialConstitution      As String           '特殊体质
    strINDeptId                 As String           '入院科室ID
    strINDeptName               As String           '入院科室名称
    strHospitalTime             As String           '入院时间
    strInWardID                 As String           '入院病区ID
    strInWardName               As String           '入院病区名称
    strInWardBedNo              As String           '入院病床号
    strInConditon               As String           '入院病情
    strWeightOfBaby             As String           '新生儿入院体重
    strPatientConditon          As String           '患者状态
End Type
'<-begin------------------------------------------------------------------------1.4 optprescriptions处方和处方明细信息标签:----------------------------------------------------->

Public Type OPTPRESCRIPTIONSINFO
    strRecipeId             As String                   '处方id
    strRecipeNo             As String                   '处方号
    strRecipeSource         As String                   '处方来源
    strRecipeCategory       As String                   '处方类别
    strRecipeType           As String                   '处方类型
    strDeptID               As String                   '处方科室ID
    strDeptName             As String                   '开方科室名称
    strRecipeDocTitle       As String                   '开方医生职称
    strRecipeDocId          As String                   '开方医生工号
    strRecipeDocName        As String                   '开方医生姓名
    strRecipeTime           As String                   '开方时间
    strHerbUnitPrice        As String                   '饮片单帖价格
    strHerbPacketCount     As String                    '饮片帖数
    strIsCream              As String                   '膏方
    strRecipeFeeTotal       As String                   '处方金额
    strOriginalRecipeId     As String                   '原始处方号
    strRecipeStatus         As String                   '处方状态
    strUrgentFlag           As String                   '紧急标志
    strReviewPharmId        As String                   '核对药师工号
    strReviewPharmName      As String                   '核对药师姓名
    strReviewPharmNTitle    As String                   '核对药师职称
    strPrepPharmId          As String                   '调配药师工号
    strPrepPharmName        As String                   '调配药师姓名
    strPrepPharmTitle       As String                   '调配药师职称
    strCheckPharmID         As String                   '审核药师工号
    strCheckPharmName       As String                   '审核药师姓名
    strCheckPharmTitle      As String                   '审核药师职称
    strDespensingPharmId    As String                   '发药药师工号
    strDespensingPharmName  As String                   '发药药师工号
    strDespensingPharmTitle As String                   '发药药师工号
End Type

Public Type OptPRESCRIPTIONSITEM
    strRecipeItemId             As String               '处方明细编号
    strRecipeId                 As String               '处方id
    strDrugPurpose              As String               '药嘱目的
    strGroupNO                  As String               '组号
    strDrugID                   As String               '药品ID
    strDrugName                 As String               '药品名称
    strCountUnit                As String               '包装规格数量
    strPackUnit                 As String               '包装规格单位
    strManufacturerID           As String               '生产厂家id
    strManufacturerName         As String               '生产厂家名称
    strDrugdose                 As String               '每次给药剂量
    strDrugadminRouteName       As String               '给药途径
    strDrugUsingFreq            As String               '给药频率
    strDrugUsingTimePoint       As String               '给药时机
    strDrugUsingAim             As String               '给药目的
    strDrugUsingArea            As String               '给药部位
    strDuration                 As String               '疗程
    strPreparation              As String               '剂型
    strSpecification            As String               '规格
    strUnitPrice                As String               '单价
    strDespensingNum            As String               '发药数量
    strFeeTotal                 As String               '总价
    strStartTime                As String               '开始使用时间
    strEndTime                  As String               '停止使用时间
    strSpecialPrompt            As String               '特殊要求
    strSkinTestFlag             As String               '皮试标志
    strSkinTestResult           As String               '皮试结果
    strSkinTestTime             As String               '皮试时间
    strDrugSource               As String               '药品来源
    strdrugReturnFlag           As String               '是否退药标志
    strOuvasFlag                As String               '门诊静配标志
    strDrippingSpeed            As String               '滴速
    strLimitTime                As String               '限用时间
    strTherapeuticRegimen       As String               '用药方案
    strDispensingWindow         As String               '发药窗口号
    strDrugstoreArea            As String               '商品货架号
End Type

'<-begin------------------------------------------------------------------------1.5 orders医嘱信息标签:----------------------------------------------------->
Public Type NonMedicalOrderItem          '标记一条非药品医嘱
    strOrderId              As String       '医嘱id
    strOrderTime            As String       '医嘱时间
    strOrderDeptID          As String       '医嘱科室id
    strOrderDeptName        As String       '医嘱科室名称
    strDocGroup             As String       '医嘱医疗组名称
    strOrderDocName         As String       '医嘱医生姓名
    strOrderDocID           As String       '医嘱医生工号
    strOrderDocTitle        As String       '医嘱医生职称
    strOrderType            As String       '医嘱类型
    strOrderCode            As String       '医嘱代码
    strOrderName            As String       '医嘱名称
    strOrderCategory        As String       '医嘱类别
    strOrderFreq            As String       '医嘱频率
    strOrderValidTime       As String       '医嘱生效时间
    strOrderInvalidTime     As String       '医嘱失效时间
    strDuration             As String       '疗程
    strCheckTime            As String       '复核时间
    strCheckNurseID         As String       '复核护士id
    strCheckNurseName       As String       '复核护士姓名
    strStopFlag             As String       '医嘱停止标志
End Type

Public Type MedicalOrderItem              '标签一条药物医嘱，多条药物医嘱多个
    strOrderId                     As String         '医嘱id
    strOrderTime                   As String         '医嘱时间
    strOrderDeptID                As String         '医嘱科室代码
    strOrderDeptName             As String         '医嘱科室名称
    strDocGroup                    As String         '医嘱医疗组名称
    strOrderDocID                 As String         '医嘱医生id
    strOrderDocName              As String         '医嘱医生姓名
    strOrderDocTitle              As String         '医嘱医生职称
    strOrderType                   As String         '医嘱类型
    strDrugPurpose                 As String         '药嘱目的
    strGroupNO                     As String         '组号
    strDrugID                      As String         '药品ID
    strDrugName                   As String         '药品通用名
    strCountUnit                   As String         '包装规格数量
    strPackUnit                    As String         '包装规格单位
    strManufacturerID              As String         '生产厂家id
    strManufacturerName           As String         '生产厂家名称
    strDrugdose                    As String         '每次给药剂量
    strDrugadminRouteName       As String         '给药途径
    strDrugUsingFreq              As String         '给药频率
    strDrugUsingTimePoint        As String         '给药时机
    strDrugUsingAim               As String         '给药目的
    strDrugUsingArea              As String         '给药部位
    strDrugSource                  As String         '药品来源
    strDuration                     As String         '疗程
    strPreparation                  As String         '剂型
    strSpecifications               As String         '规格
    strUnitPrice                   As String         '单价
    strDespensingNum               As String         '发药数量
    strFeeTotal                    As String         '总价
    strCheckTime                   As String         '复核时间
    strCheckNurseID               As String         '复核护士id
    strCheckNurseName            As String         '复核护士姓名
    strOrderValidTime             As String         '医嘱生效时间
    strOrderInvalidTime           As String         '医嘱失效时间
    strSpecialPrompt               As String         '特殊要求
    strSkinTestTime               As String         '皮试时间
    strSkinTestFlag               As String         '皮试标志
    strSkinTestResult             As String         '皮试结果
    strdrugReturnFlag             As String         '是否退药标志
    strStopFlag                    As String         '是否停药标志
    strPivasFlag                   As String         '住院静配标志
    strUrgentFlag                  As String         '紧急标志
    strDrippingSpeed               As String         '滴速
    strLimitTime                   As String         '限用时间
    strTherapeuticRegimen          As String         '用药方案
    strExeDeptID                  As String         '医嘱执行科室id
    strExeDeptName               As String         '医嘱执行科室名称
    strDispensingWindow            As String         '发药窗口号
    strDrugstoreArea              As String         '商品货架号
End Type



Public Type HerbMedicalOrderInfo
    strOrderId                      As String                  '医嘱id
    strOrderTime                    As String                  '医嘱时间
    strOrderDeptID                  As String                  '医嘱科室代码
    strOrderDeptName                As String                  '医嘱科室名称
    strDocGroup                     As String                  '医嘱医疗组名称
    strOrderDocID                   As String                  '医嘱医生id
    strOrderDocName                 As String                  '医嘱医生姓名
    strOrderDocTitle                As String                  '医嘱医生职称
    strOrderType                    As String                  '医嘱类型
    strHerbUnitPrice                As String                  '饮片单帖价格
    strHerbPacketCount              As String                  '饮片帖数
    strIsCream                      As String                  '膏方
    strCheckTime                    As String                  '复核时间
    strCheckNurseID                 As String                  '复核护士id
    strCheckNurseName               As String                  '复核护士姓名
    strOrderValidTime               As String                  '医嘱生效时间
    strOrderInvalidTime             As String                  '医嘱失效时间
    strdrugReturnFlag               As String                  '是否退药标志
    strStopFlag                     As String                  '是否停药标志
    strUrgentFlag                   As String                  '紧急标志
    strExeDeptID                    As String                  '医嘱执行科室id
    strExeDeptName                  As String                  '医嘱执行科室名称
End Type

Public Type HerbMedicalOrderItem                     '标签一条饮片医嘱明细， 多条饮片医嘱明细多个
    strOrderId                     As String                                                    '医嘱id
    strOrderitemID                As String                                                    '医嘱明细
    strGroupNO                     As String                                                    '组号
    strDrugID                      As String                                                    '药品ID
    strDrugName                   As String                                                    '药品通用名
    strManufacturerID              As String                                                    '生产厂家id
    strManufacturerName           As String                                                    '生产厂家名称
    strDrugdose                    As String                                                    '每次给药剂量
    strDrugadminRouteName       As String                                                    '给药途径
    strDrugUsingFreq              As String                                                    '给药频率
    strPreparation                  As String                                                    '剂型
    strSpecifications               As String                                                    '规格
    strUnitPrice                   As String                                                    '单价
    strDespensingNum               As String                                                    '发药数量
    strFeeTotal                    As String                                                    '总价
    strSpecialPrompt               As String                                                    '特殊要求
End Type
'<-begin------------------------------------------------------------------------1.6 Diagnoses诊断信息标签:----------------------------------------------------->

Public Type Diagnosis               '一条诊断信息，多条诊断信息多个
    strDiagID                      As String                 '诊断流水号
    strDiagDeptID                 As String                 '诊断科室id
    strDiagDeptName              As String                 '诊断科室名称
    strDiagDocID                  As String                 '诊断医生工号
    strDiagDocName               As String                 '诊断医生姓名
    strDiagDocTitle               As String                 '诊断医生职称
    strDiagDate                    As String                 '诊断时间
    strDiagCategory                As String                 '诊断类别
    strDiagType                    As String                 '诊断类型
    strDiagName                   As String                 '诊断名称
    strDiagCode                    As String                 '诊断编码
    strDiagCodeType               As String                 '诊断编码类型
    strDiseaseClassification       As String                 '疾病分型
    strDiseaseStaging              As String                 '疾病分期
    strDiseaseScore                As String                 '疾病评分
End Type

'<-begin------------------------------------------------------------------------1.7 allergies过敏信息标签:----------------------------------------------------->
Public Type Allergy
    strAllergyID           As String          '过敏信息
    strAllergyDrug         As String          '过敏药物
    strAnaphylaxis          As String          '过敏反应
    strRecordTime          As String          '记录时间
End Type

'<-begin------------------------------------------------------------------------1.8 Operations手术信息标签:----------------------------------------------------->

Public Type Operation
    strOperationID                 As String                 '手术id
    strOperationCode               As String                 '手术编码
    strOperationName              As String                 '手术名称
    strDeptID                      As String                 '科室id
    strDeptName                   As String                 '科室名称
    strOperationStartTime         As String                 '手术开始时间
    strOperationEndTime           As String                 '手术结束时间
    strOperationIncisionType      As String                 '切口类型
    strAnesthesiaCode              As String                 '麻醉方法
    strAsa                          As String                 '麻醉分级
    strAnesthesiaEndTime          As String                 '麻醉停止时间
    strAnesthesiaStartTime        As String                 '麻醉开始时间
    strIsImplant                   As String                 '是否有植入物
    strImplantNO                   As String                 '植入物id
    strImplantName                As String                 '植入物名称
    strIsReOperation               As String                 '是否二次手术
    strOperationDocID             As String                 '手术医生工号
    strOperationDocName          As String                 '手术医生姓名
    strOperationlevel              As String                 '手术等级
    strOperationSiteCode          As String                 '手术部位
    strhemorrhageVolume            As String                 '出血量
    strOperationSource             As String                 '手术来源
    strpreOPDiagCode             As String                 '术前诊断编码
    strpreOPDiagName            As String                 '术前诊断名称
    strpostOPDiagCode            As String                 '术后诊断编码
    strpostOPDiagName           As String                 '术后诊断名称
    strNnis                         As String                 '手术风险分级
    strisSelectiveOperation       As String                 '择期手术
End Type

'<-begin------------------------------------------------------------------------1.9  Exams检验标签:----------------------------------------------------->
Public Type ExamInfo
    strReportID                    As String                 '检验报告单号
    strExamItemCode               As String                 '检验项目代码
    strExamItemName              As String                 '检验项目名称
    strSampleCollectTime          As String                 '采样时间
    strSampleCode                  As String                 '标本代码
    strSampleName                 As String                 '标本名称
    strSampleCollectOpporunity    As String                 '采样时机
    strApplyNO                     As String                 '申请单编号
    strApplicantID                 As String                 '申请医生工号
    strApplicantName              As String                 '申请医生姓名
    strApplicantDeptID            As String                 '申请科室id
    strApplicantDeptName         As String                 '申请科室名称
    strReporterID                  As String                 '报告医生工号
    strReporterName               As String                 '报告医生姓名
    strReportTime                  As String                 '报告时间
    strExamAim                    As String                 '检验目的
End Type

Public Type ExamItem
    strReportID                    As String                 '检验报告单号
    strReportItemID               As String                 '检验报告明细编号
    strindicatorCode               As String                 '检验指标代码
    strindicatorName              As String                 '检验指标名称
    strindicatorename              As String                 '检验指标英文简称
    strExamResult                  As String                 '检测结果
    strExamResultUnit             As String                 '检验计量单位
    strreferenceResult             As String                 '参考结果
    strupperlimit                  As String                 '参考值上限
    strlowerlimit                  As String                 '参考值下限
    strcriticalFlag                As String                 '危急值标志
End Type
'<-begin------------------------------------------------------------------------1.10 Imageinfos影像标签:----------------------------------------------------->
Public Type ImageInfo
    strImageID                 As String                 '影像编号
    strImageCode               As String                 '检查项目代码
    strImageName              As String                 '检查项目名称
    strperformMethod           As String                 '检查方法
    strperformSite             As String                 '检查部位
    strimagingPosition         As String                 '体位
    strimagingDiagnosis        As String                 '影像诊断
    strApplyNO                 As String                 '申请单编号
    strApplicantID             As String                 '申请医生工号
    strApplicantName          As String                 '申请医生姓名
    strApplicantDeptID        As String                 '申请科室id
    strApplicantDeptName     As String                 '申请科室名称
    strPerformTime             As String                 '检查时间
    strReporterID              As String                 '报告医生工号
    strReporterName           As String                 '报告医生姓名
    strReportTime              As String                 '报告时间
    strperformAim             As String                 '检查目的
End Type

'<-begin------------------------------------------------------------------------1.11 SpecialExams特殊检查项目标签:----------------------------------------------------->
Public Type SpecialExam
    strReportID                As String                 '特殊检查报告编号
    strExamItemCode           As String                 '检查项目代码
    strExamItemName          As String                 '检查项目名称
    strExamConclusion          As String                 '检查结论
    strApplyNO                 As String                 '申请单编号
    strApplicantID             As String                 '申请医生工号
    strApplicantName          As String                 '申请医生姓名
    strApplicantDeptID        As String                 '申请科室id
    strApplicantDeptName     As String                 '申请科室名称
    strPerformTime             As String                 '检查时间
    strReporterID              As String                 '报告医生工号
    strReporterName           As String                 '报告医生姓名
    strReportTime              As String                 '报告时间
    strperformAim             As String                 '检查目的
End Type

'<-begin------------------------------------------------------------------------1.12    ElectronicMedical 门诊电子病历标签:----------------------------------------------------->
Public Type ElectronicMedical
    strElectronicMedicalID        As String                 '电子病历编号
    strChiefComplaint              As String                 '主诉
    strMedicalHistory              As String                 '现病史
    strPastHistory                 As String                 '既往史
    strPersonalHistory             As String                 '个人史
    strFamilyDiseaseHistory       As String                 '家族史
    strMenstrualHistory            As String                 '月经史
    strObstericalHistory           As String                 '婚育史
    strRecordDocID                As String                 '记录医生工号
    strRecordDocName             As String                 '记录医生姓名
    strRecordTime                  As String                 '记录时间
End Type
'<-begin------------------------------------------------------------------------1.13    AdmissionRecord入院记录标签:----------------------------------------------------->
Public Type AdmissionRecord
    strAdmissionRecordID              As String                 '入院记录ID
    strAdmissionRecordType            As String                 '入院记录类型
    strChiefComplaint                  As String                 '主诉
    strMedicalHistory                  As String                 '现病史
    strPastHistory                     As String                 '既往史
    strPersonalHistory                 As String                 '个人史
    strFamilyDiseaseHistory           As String                 '家族史
    strMenstrualHistory                As String                 '月经史
    strObstericalHistory               As String                 '婚育史
    strOperationHistory                As String                 '手术史
    strTransfusionHistory              As String                 '输血史
    strInfectionHistory                As String                 '传染病史
    strVaccinationHistory              As String                 '预防接种史
    strPhysicalExam                    As String                 '体格检查
    strSpecialExam                     As String                 '专科检查
    strAuxiliaryExam                   As String                 '辅助检查结果
    strRecordDocID                    As String                 '记录医生工号
    strRecordDocName                 As String                 '记录医生姓名
    strRecordTime                      As String                 '记录时间
End Type


'<-begin------------------------------------------------------------------------1.14    Progressnotes病程录标签:----------------------------------------------------->
Public Type HZYYProgressNote
    strProgressNoteID         As String                 '病程录ID
    strProgressNoteType       As String                 '病程录类型
    strProgressNoteContent    As String                 '病程录内容
    strRecordDocID            As String                 '记录医生工号
    strRecordDocName         As String                 '记录医生姓名
    strRecordTime              As String                 '记录时间
End Type


'<-begin------------------------------------------------------------------------1.15    vitalsigns生命体征标签:----------------------------------------------------->
Public Type VitalSign
    strVitalSignNO        As String                 '体征记录编号
    strTemperature          As String                 '体温
    strSbp                  As String                 '收缩压
    strDbp                  As String                 '舒张压
    strBreathingRate       As String                 '呼吸频率
    strPulseRate           As String                 '脉率
    strHeartRate           As String                 '心率
    strPainScore           As String                 '疼痛评分
    strHour24Amountin     As String                 '24小时入量
    strHour24Amountout    As String                 '24小时出量
    strTestTime            As String                 '测定时间
End Type

'<-begin------------------------------------------------------------------------1.16    pathologicalExams病理信息标签:----------------------------------------------------->
Public Type PathologicalExam
    strReportID                As String                 '检验报告单号
    strExamItemCode           As String                 '检查项目代码
    strExamItemName          As String                 '检查项目名称
    strSampleName             As String                 '标本名称
    strPathologicDiagnosis     As String                 '病理诊断
    strApplyNO                 As String                 '申请单编号
    strApplicantID             As String                 '申请医生工号
    strApplicantName          As String                 '申请医生姓名
    strApplicantDeptID        As String                 '申请科室id
    strApplicantDeptName     As String                 '申请科室名称
    strPerformTime             As String                 '送检时间
    strReporterID              As String                 '报告医生工号
    strReporterName           As String                 '报告医生姓名
    strReportTime              As String                 '报告时间
    strExamAim                As String                 '检查目的
End Type


'<-begin------------------------------------------------------------------------1.17    bacterialreports细菌培养报告标签:----------------------------------------------------->
Public Type BacterialReportInfo
      strReportID                      As String                 '检验报告单号
      strExamItemCode                 As String                 '检验项目代码
      strExamItemName                As String                 '检验项目名称
      strSampleCollectTime            As String                 '采样时间
      strSampleCode                    As String                 '标本代码
      strSampleName                   As String                 '标本名称
      strSampleCollectOpporunity      As String                 '采样时机
      strApplyNO                       As String                 '申请单编号
      strApplicantID                   As String                 '申请医生工号
      strApplicantName                As String                 '申请医生姓名
      strApplicantDeptID      As String                 '申请科室id
      strApplicantDeptName   As String                 '申请科室名称
      strReporterID            As String                 '报告医生工号
      strReporterName         As String                 '报告医生姓名
      strReportTime            As String                 '报告时间
End Type

Public Type BacterialReportItem
    strReportID            As String                 '检验报告单号
    strReportItemID       As String                 '细菌培养明细编号
    strExamItemResult     As String                 '检验结果
End Type

'<-begin------------------------------------------------------------------------1.18    drugsensitives药物敏感试验信息标签:----------------------------------------------------->
Public Type DrugSensitiveinfo
    strReportID                    As String                 '检验报告单号
    strExamItemCode               As String                 '检验项目代码
    strExamItemName              As String                 '检验项目名称
    strSampleCollectTime          As String                 '采样时间
    strSampleCode                  As String                 '标本代码
    strSampleName                 As String                 '标本名称
    strSampleCollectOpporunity    As String                 '采样时机
    strgermID                      As String                 '细菌id
    strgermName                   As String                 '细菌名称
    strApplyNO                     As String                 '申请单编号
    strApplicantID                 As String                 '申请医生工号
    strApplicantName              As String                 '申请医生姓名
    strApplicantDeptID            As String                 '申请科室id
    strApplicantDeptName         As String                 '申请科室名称
    strReporterID                  As String                 '报告医生工号
    strReporterName               As String                 '报告医生姓名
    strReportTime                  As String                 '报告时间
End Type

Public Type DrugSensitiveItem
    strReportID                    As String                 '检验报告单号
    strReportItemID               As String                 '药敏明细编号
    strantibioticID                As String                 '抗菌药物id
    strantibioticName             As String                 '抗菌药物
    strsensitivity                  As String                 '敏感程度
    strmic                          As String                 '最低抑菌浓度
End Type

'<-end------------------------------------------------------------------------xml----------------------------------------------------->

Public Type OptPrescription
    udtOptPresInfo                  As OPTPRESCRIPTIONSINFO        '处方信息
    colPresItem                     As Collection                    '处方明细
End Type

'单个饮片医嘱信息结构
Public Type HerbMedicalOrder
    udtHerbInfo         As HerbMedicalOrderInfo
    colItemHerb  As Collection                       'HerbMedicalOrderItem的集合
End Type
'医嘱信息XML标签
Public Type Order
    colNonMedical                   As Collection           '多条非药物医嘱
    colMedical                      As Collection           '多条药物医嘱多个
    colHerbMedical                  As Collection           '多条饮片医嘱信息HerbMedicalOrder
End Type
'
Public Type HZYYExam
    udtInfo As ExamInfo
    colExamItem As Collection                                      '多条检验明细 ExamItem
End Type

'细菌培养报告标签
Public Type BacterialReport
    udtInfo As BacterialReportInfo
    colItem As Collection                       'BacterialReportItem的集合
End Type


Public Type DrugSensitive
    udtInfo As DrugSensitiveinfo
    colItem As Collection                       'DrugSensitiveItem的集合
End Type

