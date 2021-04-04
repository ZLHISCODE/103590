Attribute VB_Name = "mdlMedRecEvent"
Option Explicit
'----------------------------------------------------------
'功能    ：首页控件事件进行封装，以及界面控件的处理
'编制人  ：刘硕
'编制日期：2013/10/30
'过程函数：
'
'
'修改记录:
'
'----------------------------------------------------------
'----------------------------------------------------------
'控件枚举
'说明：1、住院首页（标准版、四川版、云南版、湖南版）、门诊首页
'       与病案首页（标准版、四川版、云南版、湖南版）
'       的相同信息的编辑控件的名称与Index要保持相同，
'       此处控件枚举将9个窗体的枚举整合在一起
'      2、由于门诊首页控件较少，因此门诊首页先行枚举
'      3、缩写：OM=门诊首页，PM=病案首页，IM=住院首页，DS=诊断选择器
'               ST=标准版，YN=云南版，SC=四川版，HN=湖南版
'      4、特殊符号：!=只有，&=与,/=或者
'      5、控件数组符号说明：+：可以有多个，至少1个   -：可能没有，至多1个
'                          无后缀：一定有一个
'      6、相同信息的表格控件的列Index保持相同，不相同时通过增加隐藏列来实现相同
'----------------------------------------------------------
   
Public Enum ErrCol
    ERR_ID = 0
    ERR_类型 = 1
    ERR_信息 = 2
End Enum

Public Enum Pane_ID
    Pane_导航 = 1
    Pane_首页 = 2
    Pane_检查 = 3
End Enum

'页面枚举(OM /IM / PM)
Public Enum PIC菜单
    PIC_住院首页 = 0
    PIC_基本信息 = 1
    PIC_西医诊断 = 2
    PIC_西医诊断情况 = 3
    PIC_中医诊断 = 4
    PIC_中医诊断情况 = 5
    PIC_药物过敏 = 6
    PIC_输血信息 = 7
    PIC_签名信息 = 8
    PIC_手术记录 = 9
    PIC_住院费用 = 10
    PIC_住院情况 = 11
    PIC_化疗信息 = 12
    PIC_放疗记录 = 13
    PIC_抗菌药物 = 14
    PIC_抗精神病 = 15
    PIC_重症监护 = 16
    PIC_病案附加 = 17
    PIC_附页1 = 18
    PIC_附页2 = 19
End Enum

'参数相关控件(OM /IM / PM)
Public Enum ParaCtrl
    '西医诊断(IM / PM): optDiag
    PC_XY按诊断输入 = 0
    PC_XY按疾病编码输入 = 1
    '中医诊断(IM /PM): optDiag
    PC_ZY按诊断输入 = 2
    PC_ZY按疾病编码输入 = 3
    '就诊信息(OM):optDiag
    PC_按诊断输入 = 0
    PC_按疾病编码输入 = 1
    '过敏与手术(IM)/就诊信息(OM):optAller
    PC_按药品目录输入 = 0
    PC_按过敏源输入 = 1
    '过敏与手术(IM / PM):OptParaOPSInfo，chkParaOPSInfo
    PC_按诊疗项目输入 = 0
    PC_按ICDCM9编码输入 = 1
    PC_未找到时自由录入 = 0 'chkParaOPSInfo
End Enum

'人员展示相关控件(IM / PM):lblManInfo,cboManInfo
Public Enum ManCtrl
    '住院情况(IM)/医生与手术(PM)
    MC_门诊医师 = 0
    MC_科主任 = 1
    MC_主任或副主任 = 2
    MC_进修医师 = 3
    MC_主治医师 = 4
    MC_住院医师 = 5
    MC_研究生医师 = 6
    MC_实习医师 = 7
    MC_质控医师 = 8
    MC_质控护士 = 9
    MC_责任护士 = 10
    '住院情况(PM)
    MC_编目员 = 11
    '住院情况(IM)/医生与手术(PM) SC
    MC_主诊医师 = 12
End Enum

'时间输入相关控件(OM/IM/PM)：lblDateInfo，mskDateInfo(+),cmdDateInfo(-)
Public Enum DateCtrl
    '基本信息(OM/PM/IM)
    DC_出生日期 = 0
    '基本信息(IM/PM)
    DC_入院时间 = 2
    DC_出院时间 = 3
    '西医诊断(IM/PM)
    DC_确诊日期 = 4
    DC_死亡时间 = 5
    '住院情况(IM/PM)/就诊情况(OM)
    DC_发病日期 = 6
    DC_发病时间 = 7
    '住院情况(IM)/医生与手术(PM)
    DC_质控日期 = 8
    '住院情况(PM)
    DC_编目日期 = 9
    DC_收回日期 = 10
End Enum

'地址相关控件(OM/IM/PM):lblAdressInfo,txtAdressInfo,padrInfo(-),cmdAdressInfo(-)
Public Enum AdressCtrl
    '基本信息(OM/IM/PM)
    ADRC_出生地点 = 0
    ADRC_籍贯 = 1
    ADRC_现住址 = 2
    ADRC_户口地址 = 3
    '基本信息(IM/PM)
    ADRC_联系人地址 = 4
    '基本信息(OM/IM/PM)
    ADRC_病人区域 = 5
    ADRC_单位地址 = 6
End Enum

'基础字典或固定数据 下拉列表型控件(OM/IM/PM)：lblBaseInfo,cboBaseInfo
Public Enum BaseCodeCtrl
    '基本信息(OM/IM/PM)
    BCC_付款方式 = 0
    BCC_性别 = 1
    BCC_婚姻 = 2
    BCC_职业 = 3
    BCC_国籍 = 4
    BCC_民族 = 5
    '基本信息(IM/PM)
    BCC_关系 = 6
    BCC_入院途径 = 7
    '基本信息(OM)
    BCC_文化程度 = 8
    '就诊信息(OM)
    BCC_去向 = 9
    '西医诊断(IM/PM)
    BCC_感染与死亡关系 = 10
    BCC_入院情况 = 11
    BCC_分化程度 = 12
    BCC_最高诊断依据 = 13
    BCC_门诊与出院XY = 14
    BCC_入院与出院XY = 15
    BCC_门诊与入院 = 16
    BCC_术前与术后 = 17
    BCC_放射与病理 = 18
    BCC_临床与病理 = 19
    BCC_死亡期间 = 20 '!PM
    BCC_临床与尸检 = 21
    '中医诊断(IM/PM)
    BCC_门诊与出院ZY = 22
    BCC_入院与出院ZY = 23
    BCC_辩证 = 24
    BCC_治法 = 25
    BCC_方药 = 26
    BCC_治疗类别 = 27
    BCC_中医诊疗设备 = 28
    BCC_抢救方法 = 29
    BCC_中医诊疗技术 = 30
    BCC_自制中药制剂 = 31
    BCC_辨证施护 = 32
    '医生与手术(PM)/住院情况(IM)
    BCC_病案质量 = 33
    '住院情况(IM/PM)
    BCC_病例分型 = 34
    BCC_HBsAg = 35
    BCC_血型 = 36 '基本信息(OM)
    BCC_HCVAb = 37
    BCC_RH = 38 '基本信息(OM)
    BCC_HIVAb = 39
    BCC_输液反应 = 40
    BCC_输血反应 = 41
    BCC_输血前9项检查 = 42
    BCC_生育状况 = 43 '基本信息(OM)
    BCC_出院方式 = 44
    BCC_再入院计划天数 = 49
    '其他(IM/PM)
    BCC_压疮发生期间 = 45
    BCC_压疮分期 = 46
    BCC_跌倒或坠床伤害 = 47
    BCC_跌倒或坠床原因 = 48
    'YN:附页1（PM)
    BCC_距上次住院时间 = 50
    'YN:附页2（IM/PM)
    BCC_重返间隔时间 = 51
    BCC_约束方式 = 52
    BCC_约束工具 = 53
    BCC_约束原因 = 54
    BCC_新生儿离院方式 = 55
    'HN:其他（IM / PM),SC:住院情况（IM / PM)
    BCC_肿瘤分期 = 56
    'HN:其他（IM / PM)
    BCC_临床路径管理 = 57
    BCC_法定传染病 = 58
    BCC_实施DGRS管理 = 59
    BCC_死亡患者尸检 = 60
    BCC_身份证 = 61
    BCC_变异原因 = 62
    BCC_健康卡号 = 63
End Enum

'CheckBox控件(IM/PM/OM):chkInfo
Public Enum CheckCtrl
    '基本信息(IM/PM)
    CHK_再入院 = 0
    CHK_入院前外院治疗 = 1
    '西医诊断(IM/PM)
    CHK_是否确诊 = 2
    CHK_病原学检查 = 3
    'CHK_死亡患者尸检 = 4
    CHK_新发肿瘤 = 5
    '中医诊断(IM/PM)
    CHK_危重 = 6
    CHK_急症 = 7
    CHK_疑难 = 8
    '住院情况(IM/PM)
    CHK_示教病案 = 9
    CHK_科研病案 = 10
    CHK_疑难病例 = 11
    CHK_随诊 = 12
    '其他(IM/PM)
    CHK_CT = 13
    CHK_MRI = 14
    CHK_彩色多普勒 = 15
    '就诊信息(OM)
    CHK_传染病上传 = 16
    '过敏与手术(YN:IM/PM)
    CHK_围术期死亡 = 17
    CHK_术后猝死 = 18
    'YN:附页1（IM/PM)
    CHK_进入路径 = 19
    CHK_完成路径 = 20
    CHK_变异 = 21
    CHK_住院出现危重 = 22
    'YN:附页1(PM)，SC:附页2 (IM / PM)
    CHK_是否同一疾病 = 23
    'YN:    附页2 (IM / PM)
    CHK_人工气道脱出 = 24
    CHK_重返重症医学科 = 25
    CHK_住院物理约束 = 26
    'HN:其他（IM / PM)
    CHK_单病种管理 = 27
    CHK_细菌标本送检 = 28
    'SC_住院情况（IM / PM)
    CHK_会诊情况 = 29
    '就诊信息（OM)
    CHK_无过敏记录 = 30
End Enum

'限定输入控件(IM/PM/OM):lblSpecificInfo(+),txtSpecificInfo(+,-)，cmdSpecificInfo(-),cboSpecificInfo(-)
Public Enum SpecificLimitCtrl
    '基本情况(IM/PM/OM)
    SLC_单位电话 = 1
    SLC_单位邮编 = 2
    SLC_家庭电话 = 3
    SLC_家庭邮编 = 4
    SLC_户口邮编 = 5
    SLC_身高 = 6
    SLC_身高单位 = 7
    SLC_体重 = 8
    SLC_体重单位 = 9
    '基本情况(OM)
    SLC_体温 = 10
    '基本情况(IM/PM)
    SLC_入院次数 = 11
    '就诊情况(OM)
    SLC_收缩压 = 12
    SLC_舒张压 = 13
    '基本情况(IM/PM)
    SLC_联系人电话 = 14
    SLC_年龄 = 15
    SLC_婴幼儿年龄 = 16
    SLC_新生儿出生体重 = 17
    SLC_新生儿入院体重 = 18
    SLC_住院天数 = 19
    SLC_住院号 = 20
    '西医诊断(IM/PM)
    SLC_抢救次数 = 21
    SLC_成功次数 = 22
    '医生与手术(PM)
    SLC_特护 = 23
    SLC_一级护理 = 24
    SLC_二级护理 = 25
    SLC_三级护理 = 26
    SLC_ICU = 27
    SLC_CCU = 28
    '住院情况(IM/PM)
    SLC_输红细胞 = 29
    SLC_输血小板 = 30
    SLC_输血浆 = 31
    SLC_输全血 = 32
    SLC_自体回收 = 33
    SLC_呼吸机使用 = 34
    SLC_昏迷时间入院前_天 = 35
    SLC_昏迷时间入院前_小时 = 36
    SLC_昏迷时间入院前_分钟 = 37
    SLC_昏迷时间入院后_天 = 38
    SLC_昏迷时间入院后_小时 = 39
    SLC_昏迷时间入院后_分钟 = 40
    SLC_随诊期限 = 41
    '住院情况(PM)
    SLC_费用和 = 42
    'YN:附页2（IM/PM）
    SLC_约束总时间 = 43
    'HN:其他（IM/PM）
    SLC_重症监护天 = 44
    SLC_重症监护小时 = 45
    SLC_Apgar = 46
    'SC:基本信息（IM/PM）
    SLC_QQ = 47
    'SC:住院情况（IM/PM）
    SLC_输白蛋白 = 48
    SLC_院内会诊 = 49
    SLC_外院会诊 = 50
    'SC:其他
    SLC_距上次住院时间 = 51 'YN:附页1（PM)BCC_距上次住院时间 = 50
    SLC_婴幼儿年龄_DAY = 52 '婴儿年龄单位为月时按分数表现形式 数据存储格式:2月15天 分母固定为30
End Enum

'普通控件信息(IM/PM/OM)：lblInfo(+),txtInfo(+)
Public Enum GeneralCtrl
    '基本信息(PM)
    GC_病案号 = 0
    GC_档案号 = 1
    GC_X线号 = 2
    '基本信息(IM/PM/OM)
    GC_姓名 = 3
    GC_其他证件 = 4
    'GC_单位地址 = 5  更名为 ADRC_单位地址 = 6
    '基本信息(IM/PM)
    GC_联系人姓名 = 6
    GC_入院科室 = 7
    GC_入院病房 = 8
    GC_出院科室 = 9
    GC_出院病房 = 10
    '基本信息(PM)
    GC_医保号 = 11
    '就诊信息(OM)
    GC_摘要 = 12
    '基本信息(OM)
    GC_门诊号 = 14
    GC_监护人 = 15
    '就诊信息(OM)
    GC_发病地址 = 16
    '住院情况(IM/PM)/就诊信息(OM)
    GC_医学警示 = 17
    GC_其他医学警示 = 18
    '西医诊断(IM/PM)
    GC_病理号 = 19
    GC_死亡原因 = 20
    GC_病原学诊断 = 21
    GC_抢救病因 = 22
    '住院情况(IM/PM)
    GC_输其他 = 23
    GC_转入医疗机构 = 24
    GC_31天内再住院 = 25
    '基本信息(IM/PM)
    GC_转科1 = 27
    GC_转科2 = 28
    GC_转科3 = 29
    'YN:附页1（IM/PM）
    GC_退出原因 = 30
    GC_变异原因 = 31
    'YN:附页2（IM/PM）
    GC_重症监护室名称 = 32
    'HN:其他（IM/PM）
    GC_肿瘤T = 33
    GC_肿瘤N = 34
    GC_肿瘤M = 35
    'SC:基本信息（IM/PM）
    GC_Email = 36
    'SC:住院情况（IM/PM）
    GC_其他会诊 = 37
    'SC:附页2（IM/PM）
    GC_引发药物 = 38
    GC_临床表现 = 39
    GC_透析尿素氮值 = 40
    '基本信息（IM/PM）
    GC_其他关系 = 41
    GC_入院转入 = 42
    GC_监护人身份证号 = 64
End Enum
'OptionButton
Public Enum OPCtrl
    '住院情况(IM/PM)  optInput
    OP_再住院无 = 0
    OP_再住院有 = 1
    'OM optState
    OP_初诊 = 0
    OP_复诊 = 1
    'HN:其他(IM/PM)  optInput
    OP_ICU无 = 2
    OP_ICU有 = 3
End Enum
Public Enum DeptRow
    DR_转科科室 = 0
    DR_转科时间 = 1
End Enum

'自动提取控件索引cmdAutoLoad
Public Enum AuoLoadCtrl
    ALC_抗生素 = 0
    ALC_手术 = 1
    ALC_过敏记录 = 2
    ALC_临床路径 = 3
End Enum

'表格控件列枚举(IM/PM/OM)：过敏信息，vsAller
Public Enum AllerColsIndex
    AI_过敏药物 = 0
    AI_过敏反应 = 1
    AI_过敏时间 = 2
    AI_过敏源编码 = 3
    AI_药物ID = 4
    AI_过敏来源 = 5
End Enum

'表格控件列枚举(IM/PM/OM/DS)：西医诊断，vsDiagXY，中医诊断，vsDiagZY
Public Enum DiagColsIndex
    DI_诊断类型 = 0
    DI_关联 = 1
    DI_诊断编码 = 2
    DI_诊断描述 = 3
    DI_中医证候 = 4
    DI_发病时间 = 5
    DI_备注 = 6
    DI_入院病情 = 7
    DI_出院情况 = 8
    DI_ICD附码 = 9
    DI_是否未治 = 10
    DI_是否疑诊 = 11
    DI_增加 = 12
    DI_Del = 13
    DI_诊断ID = 14
    DI_疾病ID = 15
    DI_证候ID = 16
    DI_医嘱IDs = 17 '与当前诊断关联的医嘱ID组成的字符串，医嘱ID间以逗号分割
    DI_诊断分类 = 18
    DI_固定附码 = 19
    DI_是否病人 = 20
    DI_疗效限制 = 21
    DI_分娩信息 = 22
    DI_附码ID = 23
    DI_诊断来源 = 24
    DI_疾病编码 = 25
    DI_疾病类别 = 26
    DI_证候编码 = 27
    DI_记录日期 = 28
    DI_记录人员 = 29
End Enum

'特殊检查枚举
Public Enum TSJCRow
    'HN，YN,ST
    TR_特殊检查4 = 0
    TR_特殊检查5 = 1
    TR_特殊检查6 = 2
    'SC
    TR_CT = 0
    TR_PETCT = 1
    TR_双源CT = 2
    TR_X片 = 3
    TR_B超 = 4
    TR_超声心动图 = 5
    TR_MRI = 6
    TR_同位素检查 = 7
End Enum
'手术列枚举
Public Enum OPSColsIndex
    PI_Copy = 0
    PI_手术日期 = 1
    PI_结束日期 = 2
    PI_抗菌用药时间 = 3
    PI_手术情况 = 4
    PI_准备天数 = 5
    PI_手术编码 = 6
    PI_手术名称 = 7
    PI_再次手术 = 8
    PI_主刀医师 = 9
    PI_助产护士 = 10
    PI_助手1 = 11
    PI_助手2 = 12
    PI_麻醉开始时间 = 13
    PI_麻醉类型 = 14 '界面名称麻醉方式
    PI_ASA分级 = 15
    PI_NNIS分级 = 16
    PI_手术级别 = 17
    PI_麻醉医师 = 18
    PI_切口愈合 = 19
    PI_切口部位 = 20
    PI_重返手术室计划 = 21
    PI_重返手术室目的 = 22
    PI_切口感染 = 23
    PI_并发症 = 24
    PI_预防用抗菌药 = 25
    PI_抗菌药天数 = 26
    PI_非预期的二次手术 = 27
    PI_麻醉并发症 = 28
    PI_术中异物遗留 = 29
    PI_手术并发症 = 30
    PI_术后出血或血肿 = 31
    PI_手术伤口裂开 = 32
    PI_术后深静脉血栓 = 33
    PI_术后生理代谢紊乱 = 34
    PI_术后呼吸衰竭 = 35
    PI_术后肺栓塞 = 36
    PI_术后败血症 = 37
    PI_术后髋关节骨折 = 38
    PI_手术操作ID = 39
    PI_诊疗项目ID = 40
    PI_麻醉ID = 41
    PI_麻醉方式 = 42 '界面名称麻醉类型
    PI_手麻来源 = 43
End Enum
'化疗列枚举
Public Enum ChemothColsIndex
    CI_化学治疗编码 = 0
    CI_开始日期 = 1
    CI_结束日期 = 2
    CI_疗程数 = 3
    CI_化疗方案 = 4
    CI_总量 = 5
    CI_化疗效果 = 6
    CI_疾病ID = 7
End Enum
'放疗列枚举
Public Enum RadiothColsIndex
    RI_放射治疗编码 = 0
    RI_开始日期 = 1
    RI_结束日期 = 2
    RI_设野部位 = 3
    RI_放射剂量 = 4
    RI_累计量 = 5
    RI_放疗效果 = 6
    RI_疾病ID = 7
End Enum

'精神药品列枚举
Public Enum SpiritColsIndex
    SI_药物名称 = 0
    SI_疗程 = 1
    SI_最高日量 = 2
    SI_特殊反应 = 3
    SI_疗效 = 4
    SI_药品ID = 5
End Enum
'重症监护枚举
Public Enum ICUColsIndex
    UI_序号 = 0
    UI_监护室名称 = 1
    UI_进入时间 = 2
    UI_退出时间 = 3
    UI_再入住计划 = 4
    UI_再入住原因 = 5
End Enum
'重症监护器械枚举
Public Enum ICUInstruColsIndex
    TI_ICU类型 = 0
    TI_器械及导管 = 1
    TI_开始时间 = 2
    TI_结束时间 = 3
    TI_感染累计小时 = 4
End Enum
'医院感染枚举
Public Enum InfectColsIndex
    FI_确诊日期 = 0
    FI_感染部位 = 1
    FI_医院感染名称 = 2
    FI_医院感染编码 = 3
End Enum
'标本来源枚举
Public Enum SampleColsIndex
    MI_标本 = 0
    MI_病原学代码及名称 = 1
    MI_送检日期 = 2
End Enum
'抗菌药枚举
Public Enum KSSColsIndex
    KI_序号 = 0
    KI_抗菌药物名 = 1
    KI_用药目的 = 2
    KI_使用阶段 = 3
    KI_使用天数 = 4
    KI_一类切口预防用 = 5
    KI_DDD数 = 6
    KI_联合用药 = 7
End Enum

Private mblnChk  As Boolean  '是否执行chk点击事件
Private mobjDiag As Object   '诊断表格对象
Private mblnReturn As Boolean

'--------------------------------------------------------------------------
'控件事件封装
'函数命名方法：控件名+事件名
'--------------------------------------------------------------------------
'Form事件
Public Sub FormActivate()
'Form_Activate事件
'    If gclsPros.IsLoad Then
'        Call ChangePage(, 0)
'    End If
'    gclsPros.IsLoad = False
End Sub

Public Sub FormKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'Form_KeyDown事件
    Dim lngIndex As Long, i As Long, lngCount As Long
    Dim monTmp As MonthView

    With gclsPros.CurrentForm
        Select Case intKeyCode
            '向下翻
            Case vbKeyPageDown
                Call ChangePage(True)
                Exit Sub
            '向上翻
            Case vbKeyPageUp
                Call ChangePage(False)
                Exit Sub
            '翻到最前
            Case vbKeyHome
                .vsbMain.Value = .vsbMain.Min
                Exit Sub
            '翻到最后
            Case vbKeyEnd
                .vsbMain.Value = .vsbMain.Max
                Exit Sub
            Case vbKeyUp
                If intShift = 2 Then
                    i = .vsbMain.Value
                    If i - 10 < .vsbMain.Min Then
                        i = .vsbMain.Min
                    Else
                        i = i - 10
                    End If
                    .vsbMain.Value = i
                    Exit Sub
                End If
            Case vbKeyDown
                If intShift = 2 Then
                    i = .vsbMain.Value
                    If i + 10 > .vsbMain.Max Then
                        i = .vsbMain.Max
                    Else
                        i = i + 10
                    End If
                    .vsbMain.Value = i
                    Exit Sub
                End If
            Case vbKeyEscape
                '错误屏蔽，主要是由于某些窗体没有monInfo控件，如门诊首页
                On Error Resume Next
                Set monTmp = .monInfo
                If Err.Number = 0 Then
                    monTmp.Visible = False
                    Err.Clear: On Error GoTo 0
                    Call ShowInfectInfo(False)
                End If
            Case vbKeyF5
                Call ChangePage(True, 1)
            Case vbKeyF6
                Call ChangePage(True, 2)
            Case vbKeyF7
                Call ChangePage(True, 6)
            Case vbKeyF8
                Call ChangePage(True, 8)
            Case vbKeyF9
                Call ChangePage(True, 9)
            Case vbKeyF10
                Call ChangePage(True, 12)
            Case vbKeyF11
                Call ChangePage(True, 14)
            Case vbKeyF12
                Call ChangePage(True, 17)
        End Select
        If gclsPros.FuncType = f病案首页 Then
            If gclsPros.OpenMode = EM_查阅 Or gclsPros.OpenMode = EM_编辑 Then
                If intShift = 2 And intKeyCode = vbKeyU Then
                    CmdUPClick
                ElseIf intShift = 2 And intKeyCode = vbKeyD Then
                    CmdDownClick
                End If
            End If
        End If
        If intKeyCode = vbKeyS And intShift = 2 Then
            If gclsPros.FuncType = f医生首页 Then
                If gclsPros.InfosChange Then
                    Call menuPageOperate(MOP_确定)
                End If
             ElseIf gclsPros.FuncType = f病案首页 And gclsPros.OpenMode <> EM_查阅 Then
                If gclsPros.InfosChange Then
                    Call menuPageOperate(MOP_确定)
                End If
            End If
        End If
    End With
End Sub

Public Sub FormKeyPress(ByRef intKeyAscii As Integer)
'Form_KeyPress事件
    If intKeyAscii = Asc("'") Then intKeyAscii = 0
End Sub

Public Function FormLoad(Optional ByVal blnChange As Boolean) As Boolean
    Dim i As Integer
'功能：首页窗体Form_Load事件
'参数：blnChange=是否是获取上一份或者下一份病案调用
'返回：Ture-成功，False-失败
    gclsPros.IsOpen = True
    gclsPros.IsLoad = True
    On Error GoTo errH
    With gclsPros.CurrentForm
        '界面初始化以及数据加载
        On Error GoTo errH:
        '病案可能需要新增主页，或者获取上一次住院或下一次住院的主页ID
        If gclsPros.FuncType = f病案首页 Then
            If Not ValiAndGet主页ID Then Exit Function
            Call OpenExtraData
            Select Case gclsPros.OpenMode
                Case EM_编辑
                    gclsPros.NoType = IT_Old
                    gclsPros.IsExistPati = True
                Case EM_新增病案
                    gclsPros.NoType = IT_New
                Case EM_新增首页
                    gclsPros.NoType = IT_Old
                    gclsPros.IsExistPati = True
            End Select
        End If
        
        '页面切换时,不重复加载
        If Not blnChange Then
            If gclsPros.FuncType = f医生首页 Or gclsPros.FuncType = f病案首页 Then
                '检查是否开启外挂部件加载首页附页
                Call CreatePlugInOK(gclsPros.Module)
                If Not gobjPlugIn Is Nothing Then
                    Err.Clear: On Error Resume Next
                    If gobjPlugIn.gblnfrmMec = True Then
                        '调用病案加载自定义附页接口
                        If Err.Number = 0 Then
                            Set gfrmMecCol = gobjPlugIn.GetMeRecFormCol(gclsPros.SysNo, gclsPros.Module, gclsPros.病人ID, gclsPros.主页ID, gclsPros.PatiType)
                        End If
                        If Err.Number = 0 Then
                            gBlnNew = True
                        Else
                            gBlnNew = False
                        End If
                        Call zlPlugInErrH(Err, "GetMeRecFormCol")
                        Err.Clear: On Error GoTo 0
                    End If
                End If
            End If
            
            If gBlnNew = True And (Not gfrmMecCol Is Nothing) Then
                gIntPic = gclsPros.CurrentForm.PicPage.Count - 1
                For i = 1 To gfrmMecCol.Count
                    gPic外挂附页 = gclsPros.CurrentForm.PicPage.Count
                    Load gclsPros.CurrentForm.PicPage(gPic外挂附页)
                    gclsPros.CurrentForm.PicPage(gPic外挂附页).Height = gfrmMecCol(i).Height + 50
                    SetParent gfrmMecCol(i).hwnd, gclsPros.CurrentForm.PicPage(gPic外挂附页).hwnd
                    gfrmMecCol(i).Top = IIf(gclsPros.FuncType = f病案首页, IIf(gclsPros.MedPageSandard = ST_卫生部标准, -300, -1900), -300): gfrmMecCol(i).Left = 0: gfrmMecCol(i).Tag = gPic外挂附页
                    gfrmMecCol(i).Show
                    Set gclsPros.CurrentForm.PicPage(gPic外挂附页).Container = gclsPros.CurrentForm.picMain
                Next
            End If
        End If
        Call SetAllObject
        If Not InitMedRecEnv Then Exit Function
        If gclsPros.OpenMode <> EM_新增病案 Then   '新增病案不用加载数据
            If Not LoadMedPageData(gclsPros.病人ID, gclsPros.主页ID, gclsPros.RegistNo, gclsPros.PatiType = PF_门诊) Then Exit Function
        End If
        Call SetPageVisible '设置默认页面以及页面可见性
        Call gclsPros.InitFacePara '设置界面参数控件状态
        If Not InitMedRecEnv(True) Then Exit Function '数据加载后的界面调整
        If gclsPros.PatiType = PF_住院 And gclsPros.FuncType = f医生首页 Then
            gclsPros.IsSigned = SetSignature
        End If
        Call SetFaceInit                        '将界面回复到初始状态
        Call SetFaceEditable(gclsPros.IsSigned) '设置界面控件可用性
        .subcMain.hwnd = .hwnd
        .subcMain.Messages(WM_MOUSEWHEEL) = True
        Call SetAllVSF
        Call SetPicPosition(True, True)
        
        Call SetComboBoxProperty(True) '屏蔽掉ComboBox的鼠标滚轮事件
    End With
    FormLoad = True
    gclsPros.LoadFinish = True
    '默认婴幼儿年龄
    Call CboSpecificInfoClick(SLC_婴幼儿年龄)
    gclsPros.InfosChange = False
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FormUnLoad(ByRef blnCancel As Integer)
    Dim i As Integer
    If (gclsPros.InfosChange And Not gclsPros.IsOK And gclsPros.FuncType = f诊断选择) Or (gclsPros.FuncType <> f诊断选择 And gclsPros.InfosChange And gclsPros.OpenMode <> EM_查阅) Then
        If MsgBox("如果退出，刚才所修改的内容将不会被保存。确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnCancel = True: Exit Sub
        End If
    End If
    If gclsPros.FuncType = f诊断选择 Or gclsPros.FuncType = f医生首页 Then
        If gclsPros.PatiType = PF_门诊 Then
            Call zlDatabase.SetPara("门诊诊断输入", gclsPros.DiagInputXY, gclsPros.SysNo, p门诊医生站, InStr(gclsPros.Privs, "参数设置") > 0)
            If gclsPros.FuncType = f医生首页 Then
                Call zlDatabase.SetPara("过敏输入来源", gclsPros.AllerInput, gclsPros.SysNo, p门诊医生站, gclsPros.AllerSource = 0 And gclsPros.PassType = 3 And InStr(gclsPros.Privs, "参数设置") > 0)
            End If
        Else
            Call zlDatabase.SetPara("西医诊断输入", gclsPros.DiagInputXY, gclsPros.SysNo, p住院医生站, InStr(gclsPros.Privs, "参数设置") > 0)
            Call zlDatabase.SetPara("中医诊断输入", gclsPros.DiagInputZY, gclsPros.SysNo, p住院医生站, InStr(gclsPros.Privs, "参数设置") > 0)
            If gclsPros.FuncType = f医生首页 Then
                Call zlDatabase.SetPara("过敏输入来源", gclsPros.AllerInput, gclsPros.SysNo, p门诊医生站, gclsPros.AllerSource = 0 And gclsPros.PassType = 3 And InStr(gclsPros.Privs, "参数设置") > 0)
                Call zlDatabase.SetPara("手术情况输入", gclsPros.OPSInput & IIf(gclsPros.OPSFree, 1, 0), gclsPros.SysNo, p住院医生站, InStr(gclsPros.Privs, "参数设置") > 0)
            End If
        End If
    End If
    Call SaveWinState(gclsPros.CurrentForm, App.ProductName)
    gclsPros.IsOpen = False
    If gclsPros.FuncType <> f病案首页 Then
        Call gclsMain.Closed(Not gclsPros.IsOK, gclsPros.DiseaseIDs, gclsPros.DiagIDs, gclsPros.PictureFile)
    End If
    If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
        With gclsPros.CurrentForm
            .subcMain.Messages(WM_MOUSEWHEEL) = False
        End With
        Call SetComboBoxProperty(False)
    End If
    '卸载外挂附页
    If gBlnNew = True And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Unload gfrmMecCol(i)
        Next
        gBlnNew = False
        Set gfrmMecCol = Nothing
    End If
End Sub

'cmdCancel事件
Public Sub CmdCancelClick()
    Unload gclsPros.CurrentForm
End Sub

Public Sub CmdCancelGotFocus()
'cmdCancel_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'CmdDown事件
Public Sub CmdDownClick()
'CmdDown_Click事件
    If gclsPros.OpenMode <> EM_查阅 And gclsPros.InfosChange Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认查看另一份病案？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
           Exit Sub
        End If
    End If
    Call ClearPageContent
    Call gclsPros.RefreshPara
    Call gclsPros.InitCacheRecInfo
    gclsPros.主页ID = Get主页IDByCur(gclsPros.主页ID, True)
    gclsPros.Is编目 = True
    Call FormLoad(True)
End Sub

Public Sub CmdDownGotFocus()
'CmdDown_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdHelp事件
Public Sub CmdHelpClick()
'cmdHelp_Click事件
    ShowHelp App.ProductName, gclsPros.CurrentForm.hwnd, gclsPros.CurrentForm.Name, gclsPros.SysNo \ 100
End Sub

'cmdDiagMove事件
Public Sub CmdDiagMoveClick(ByRef intIndex As Integer)
'cmdDiagMove_Click事件
    Call MoveDiagRows(IIf(intIndex \ 2 = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY), IIf(intIndex Mod 2 = 0, -1, 1))
End Sub

Public Sub CmdDiagMoveGotFocus(ByRef intIndex As Integer)
'cmdDiagMove_GotFocus事件
    '西医需要隐藏
    If intIndex \ 2 = 0 Then Call ShowInfectInfo(False)
End Sub

Public Sub CmdHelpGotFocus()
'cmdHelp_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdOPSMove事件
Public Sub cmdOPSMoveClick(ByRef intIndex As Integer)
'cmdOPSMove_Click事件
    Call MoveOPSRows(gclsPros.CurrentForm.vsOPS, IIf(intIndex Mod 2 = 0, -1, 1))
End Sub

'CmdUp事件
Public Sub CmdUPClick()
'CmdUP_Click事件
    If gclsPros.OpenMode <> EM_查阅 And gclsPros.InfosChange Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认查看另一份病案？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
           Exit Sub
        End If
    End If
    Call ClearPageContent
    Call gclsPros.RefreshPara
        Call gclsPros.InitCacheRecInfo
    gclsPros.主页ID = Get主页IDByCur(gclsPros.主页ID, False)
    gclsPros.Is编目 = True
    Call FormLoad(True)
End Sub

Public Sub CmdUPGotFocus()
'CmdUP_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

Public Sub cmdDoctorDiagClick(ByVal Index As Integer, ByVal frmParent As Form)
    Dim vPoint As POINTAPI
    If Index = 1 Then
        vPoint = GetCoordPos(gclsPros.CurrentForm.vsDiagXY.hwnd, gclsPros.CurrentForm.vsDiagXY.Left + 15, gclsPros.CurrentForm.vsDiagXY.CellTop)
        frmPublicTable.ShowMe Index, gclsPros.病人ID, gclsPros.主页ID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsDiagXY.Height
    Else
        vPoint = GetCoordPos(gclsPros.CurrentForm.vsDiagZY.hwnd, gclsPros.CurrentForm.vsDiagZY.Left + 15, gclsPros.CurrentForm.vsDiagZY.CellTop)
        frmPublicTable.ShowMe Index, gclsPros.病人ID, gclsPros.主页ID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsDiagZY.Height
    End If
End Sub

Public Sub cmdDoctorOPSClick(ByVal frmParent As Form)
    Dim vPoint As POINTAPI
    vPoint = GetCoordPos(gclsPros.CurrentForm.vsOPS.hwnd, gclsPros.CurrentForm.vsOPS.Left + 15, gclsPros.CurrentForm.vsOPS.CellTop)
    frmPublicTable.ShowMe 3, gclsPros.病人ID, gclsPros.主页ID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsOPS.Height
End Sub

'cmdDeliceryInfo事件
Public Sub CmdDeliceryInfoClick(Optional ByVal bytFunc As Byte = 0, Optional ByRef objFrmMain As Object, Optional ByVal lngPatiID As Long, Optional ByVal lngMainID As Long)
'cmdDeliceryInfo_Click事件
'参数:
'   bytFunc=0 病案系统调用
'   bytFunc=1 新生儿登记调用
    Dim LngRow As Long
    Dim str诊断 As String, dat入院日期 As Date, dat出院日期 As Date
    Dim strTmp As String, blnOK As Boolean
    
    If bytFunc = 0 Then
        With gclsPros.CurrentForm
            If .txtInfo(GC_病案号).Text = "" Or .txtSpecificInfo(SLC_住院号).Text = "" Or .txtInfo(GC_姓名).Text = "" Then
               MsgBox "请先录入首页中的病案号,住院号,姓名等基本信息!", vbInformation, gstrSysName
               Exit Sub
            End If
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("病案号", .txtInfo(GC_病案号).Text, 1))
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("住院号", .txtSpecificInfo(SLC_住院号).Text, 1))
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("姓名", .txtInfo(GC_姓名).Text, 1))
    
            strTmp = .mskDateInfo(DC_入院时间).Text
            If Not IsDate(strTmp) Then strTmp = ""
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("入院日期", strTmp, 1))
            strTmp = .mskDateInfo(DC_出院时间).Text
            If Not IsDate(strTmp) Then strTmp = ""
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("出院日期", strTmp, 1))
            LngRow = FindDiagRow(DT_出院诊断XY)
            str诊断 = .vsDiagXY.TextMatrix(LngRow, DI_诊断编码) & Space(2) & .vsDiagXY.TextMatrix(LngRow, DI_诊断描述)
            Call grsDeliceryInfo.AddNew(Array("信息名", "信息值", "类型"), Array("主要诊断", str诊断, 1))
            Call frmDeliceryInfo.EditDelivery(gclsPros.CurrentForm, gclsPros.病人ID, gclsPros.主页ID, Val(.vsDiagXY.TextMatrix(LngRow, DI_疾病ID)), gclsPros.OpenMode <> EM_查阅, grsDeliceryInfo, grsBabyInfo, grsBabyDiag, blnOK)
            If blnOK Then
                Call CheckValueChange
            End If
        End With
    ElseIf bytFunc = 1 Then
        If gclsPros Is Nothing Then
            Set gclsPros = New clsProperty
        End If
        
        Set gclsPros.CurrentForm = objFrmMain
        gclsPros.病人ID = lngPatiID
        gclsPros.主页ID = lngMainID
        Set grsDeliceryInfo = zlDatabase.CopyNewRec(GetPatiAuxiInfoData(lngPatiID, lngMainID, , 2), , "信息名,信息值,信息值 信息现值", Array("类型", adInteger, 1, 0, "记录性质", adInteger, 1, Empty))
        Do While Not grsDeliceryInfo.EOF
            grsDeliceryInfo!类型 = 0
            grsDeliceryInfo.MoveNext
        Loop
        Set grsBabyDiag = zlDatabase.CopyNewRec(GetBabyDiagData(lngPatiID, lngMainID), , , Array("记录性质", adInteger, 1, Empty))
        Set grsBabyInfo = zlDatabase.CopyNewRec(GetBabyInfoData(lngPatiID, lngMainID), , , Array("记录性质", adInteger, 1, Empty))
        Call frmDeliceryInfo.EditDelivery(objFrmMain, lngPatiID, lngMainID, 0, True, grsDeliceryInfo, grsBabyInfo, grsBabyDiag, , 2)
    End If
End Sub

Public Sub CmdDeliceryInfoGotFocus()
'cmdDeliceryInfo_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdPrint事件
Public Sub CmdPrintClick()
'cmdPrint_Click事件
    Call PageOperate(MOP_打印)
End Sub

Public Sub CmdPrintGotFocus()
'cmdPrint_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdPrintdown事件
Public Sub CmdPrintdownGotFocus()
'cmdPrintdown_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdPriviewDown
Public Sub CmdPriviewDownGotFocus()
'cmdPriviewDown_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdSign事件
Public Sub CmdSignClick(ByRef intIndex As Integer)
'功能：cmdSign_Click事件，签名
    If gclsPros.CurrentForm.cmdSign(intIndex).Caption = "签名" Then
        Call SetSign(intIndex)              '签名
    Else
        Call SetSign(intIndex, True)        '取消签名
    End If
End Sub

'ManInfo事件封装
'ManInfo事件封装
Public Sub ManInfoClick(ByRef intIndex As Integer)
'功能：cboManInfo_Click事件封装
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim blnRestore As Boolean
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If gclsPros.CurrentForm.Visible Then
        If cboTmp.ListIndex <> -1 Then
            If cboTmp.ItemData(cboTmp.ListIndex) = -1 Then
                If gclsPros.FuncType = f医生首页 Then Set gclsPros.ManInfo = Nothing   '清空缓存，重新读取数据
                Set rsInput = zlDatabase.CopyNewRec(GetManData(intIndex), , "ID,编码,简码 拼音简码,五笔简码,名称 姓名,缺省")
                If rsInput.RecordCount <> 0 Then
                    '如果下拉列表展开则关闭下拉列表
                    If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
                        SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
                    End If
                    If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                        If cboTmp.ListCount = 0 Then Call SetCboFromRec(intIndex, 1)
                        intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                        If intIdx = -1 Then
                            cboTmp.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!姓名, cboTmp.ListCount - 1
                            cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                            intIdx = cboTmp.NewIndex
                        End If
                        Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
                        cboTmp.Tag = cboTmp.ListIndex
                    Else
                        blnRestore = True
                    End If
                Else
                    MsgBox "没有住院医生或护士的数据，请先到部门/人员管理中设置。", vbInformation, gstrSysName
                    blnRestore = True
                End If
            Else
                cboTmp.Tag = cboTmp.ListIndex
            End If
            '恢复成现有的人员(不引发Click)
            If blnRestore Then
                If Val(cboTmp.Tag) <> -1 Then
                    Call zlControl.CboSetIndex(cboTmp.hwnd, Val(cboTmp.Tag))
                End If
            Else
                '医师更改,刷新签名状态
                If gclsPros.FuncType = f医生首页 Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
        End If
    End If
    
    Call CheckValueChange(cboTmp)
End Sub

Public Sub ManInfoDropDown(ByRef intIndex As Integer)
'功能：cboManInfo_DropDown事件封装
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If cboTmp.Tag = "" Then
       Call ManInfoGotFocus(intIndex)
    End If
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = f医生首页, "[其他...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            Call SetCboFromName(strTmp, cboTmp, "人员")
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    Call TxtGotFocus(cboTmp, True, True)
End Sub

Public Function ManInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
    If intIndex = MC_科主任 Then
        If intKeyAscii = vbKeyPageDown Or intKeyAscii = vbKeyPageUp Then intKeyAscii = 0
    End If
End Function

Public Sub ManInfoGotFocus(ByRef intIndex As Integer)
'功能：cboManInfo_GotFocus事件封装
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim blnAdd As Boolean
    
    Call ChangeCtl

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        blnAdd = True
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = f医生首页, "[其他...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            blnAdd = True
            Call SetCboFromName(strTmp, cboTmp, "人员", blnAdd)
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    Call TxtGotFocus(cboTmp, True, True)
End Sub

Public Sub ManInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'功能：cboManInfo_KeyPress事件封装
    Dim cboTmp As ComboBox
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim lngIdx As Integer
    Dim blnRestore As Boolean
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If intKeyAscii = vbKeyReturn Then
        strTmp = cboTmp.Text
        If cboTmp.ListCount = 0 Then
            Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = f医生首页, "[其他...]", "NULL"))
        End If
        If strTmp <> "" Then
            lngIdx = Cbo.FindIndex(cboTmp, strTmp)
            If lngIdx = -1 Then
                Call SetCboFromName(strTmp, cboTmp, "人员")
            Else
                Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
                Call CheckValueChange(cboTmp)
            End If
        End If
        '查找匹配项
        strInput = Trim(cboTmp.Text)
        If strInput = "" Then
            '如果之前有内容，现在删除之后，保存按钮也应该变得可用，签名按钮也应该发生变化
            If cboTmp.Tag >= 0 Then
                cboTmp.Tag = cboTmp.ListIndex
                Call CheckValueChange(cboTmp)
                 '医师更改,刷新签名状态
                If gclsPros.FuncType = f医生首页 Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Exit Sub
        End If
        '相同的项目则不进行处理
        If cboTmp.ListIndex <> -1 Then
            If cboTmp.Tag <> cboTmp.ListIndex Then
                cboTmp.Tag = cboTmp.ListIndex
                '医师更改,刷新签名状态
                If gclsPros.FuncType = f医生首页 Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
            
            If zlStr.NeedName(strInput) = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
                Exit Sub
            End If
        End If

        strInput = UCase(strInput)
        strFilter = "编码 Like '" & strInput & "*' Or 名称 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or " & IIf(gclsPros.BriefCode = 0, "简码", "五笔简码") & " Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
        Set rsInput = Rec.FilterNew(GetManData(intIndex), strFilter, "ID,编码,简码 拼音简码,五笔简码,名称 姓名,缺省")
        If rsInput.RecordCount = 0 Then
            MsgBox "未找到对应的医生或护士。", vbInformation, gstrSysName
            blnRestore = True
        Else
            '如果下拉列表展开则关闭下拉列表
            If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
                SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
            End If
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                If cboTmp.ListCount = 0 Then Call SetCboFromRec(intIndex, 1)
                lngIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                If lngIdx = -1 Then
                    cboTmp.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!姓名, cboTmp.ListCount - 1
                    cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                    lngIdx = cboTmp.NewIndex
                End If

                Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
                cboTmp.Tag = cboTmp.ListIndex
                cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                blnRestore = True
            End If
        End If
        '恢复成现有的人员(不引发Click)
        If blnRestore Then
            If Val(cboTmp.Tag) <> -1 Then
                Call zlControl.CboSetIndex(cboTmp.hwnd, Val(cboTmp.Tag))
            End If
            cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
        Else
            Call CheckValueChange(cboTmp)
            '医师更改,刷新签名状态
            If gclsPros.FuncType = f医生首页 Then
                gclsPros.IsSigned = SetSignature()
                Call SetFaceEditable(gclsPros.IsSigned)
            End If
        End If
    End If
End Sub

Public Sub ManInfoLostFocus(ByRef intIndex As Integer)
'功能：cboManInfo_LostFocus事件封装
    Dim intIdx As Integer, i As Long
    Dim blnHave As Boolean
    Dim cboTmp As ComboBox

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If cboTmp.ListIndex >= 0 Then
       If cboTmp.Text <> zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
           cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
       End If
    End If
End Sub

Public Sub ManInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'功能：cboManInfo_Validate事件封装
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = f医生首页, "[其他...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            Call SetCboFromName(strTmp, cboTmp, "人员")
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
        cboTmp.Tag = cboTmp.ListIndex
    Else
        cboTmp.Tag = cboTmp.ListIndex
        Exit Sub '无输入
    End If
    If cboTmp.ListIndex <> -1 Then cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)): Exit Sub '已选中

    strInput = UCase(zlStr.NeedName(cboTmp.Text))
    strFilter = "编码 Like '" & strInput & "*' Or 名称 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or " & IIf(gclsPros.BriefCode = 0, "简码", "五笔简码") & " Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
    Set rsInput = Rec.FilterNew(GetManData(intIndex), strFilter, "ID,编码,简码 拼音简码,五笔简码,名称 姓名,缺省")
    If rsInput.RecordCount <> 0 Then
        '如果下拉列表展开则关闭下拉列表
        If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
            SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
        End If
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
            intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
            If intIdx = -1 Then
                cboTmp.AddItem rsTmp!编码 & "-" & Chr(13) & rsTmp!姓名, cboTmp.ListCount - 1
                cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                intIdx = cboTmp.NewIndex
            End If
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
            cboTmp.Tag = cboTmp.ListIndex
            cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
        Else
            blnCancel = True: cboTmp.Text = ""
        End If
    Else
        MsgBox "未找到对应的医生或护士。", vbInformation, gstrSysName
        blnCancel = True: cboTmp.Text = ""
    End If
End Sub

'DateInfo事件封装
Public Sub DateInfoChange(ByRef intIndex As Integer)
'功能：MskDateInfo_Change
    Select Case intIndex
        Case DC_发病日期
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病时间), Not IsDate(gclsPros.CurrentForm.mskDateInfo(intIndex).Text))
    End Select
    Call CheckValueChange(gclsPros.CurrentForm.mskDateInfo(intIndex))
End Sub

Public Sub DateInfoClick(ByRef intIndex As Integer)
'功能：cmdDateInfo_Click
    Dim objmonInfo As MonthView  '方便调用控件属性
    Dim objCmd As CommandButton
    Dim objMSK As MaskEdBox
    Dim datStart As Date
    Dim dateEnd As Date
    Dim datTmp As Date
    On Error GoTo errH
    gclsPros.DateIndex = intIndex
    With gclsPros.CurrentForm
        Set objmonInfo = .monInfo
        Set objCmd = .cmdDateInfo(intIndex)
        Set objMSK = .mskDateInfo(intIndex)
        If IsDate(gclsPros.InTime) Then
            datStart = CDate(gclsPros.InTime)
        End If
        If IsDate(gclsPros.OutTime) Then
            dateEnd = CDate(gclsPros.OutTime)
        Else
            dateEnd = zlDatabase.Currentdate
        End If
        objmonInfo.MinDate = 0
        objmonInfo.MaxDate = zlDatabase.Currentdate
        Select Case intIndex
            Case DC_出生日期
                objmonInfo.MaxDate = datStart
            Case DC_入院时间
                objmonInfo.MaxDate = dateEnd
            Case DC_出院时间
                objmonInfo.MinDate = datStart
            Case DC_确诊日期
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = dateEnd
            Case DC_死亡时间
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = zlDatabase.Currentdate
            Case DC_发病日期
                objmonInfo.MaxDate = dateEnd
            Case DC_编目日期, DC_收回日期
                objmonInfo.MinDate = dateEnd
                objmonInfo.MaxDate = zlDatabase.Currentdate
            Case DC_质控日期
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = CDate("3000-01-01")
'                objmonInfo.Value = zlDatabase.Currentdate
        End Select
        If IsDate(objMSK.Text) Then
            datTmp = CDate(objMSK.Text)
            If datTmp > objmonInfo.MaxDate Then
                datTmp = objmonInfo.MaxDate
            ElseIf datTmp < objmonInfo.MinDate Then
                datTmp = objmonInfo.MinDate
            End If
            objmonInfo.Value = datTmp
        End If
        objmonInfo.Left = objCmd.Left + objCmd.Width - objmonInfo.Width + objMSK.Container.Left + .PicPage(0).Left
        If intIndex = DC_出生日期 Then
            objmonInfo.Top = objCmd.Top + objCmd.Height + 20 + objMSK.Container.Top
        Else
            objmonInfo.Top = objCmd.Top - objmonInfo.Height - 20 + objMSK.Container.Top
        End If
        objmonInfo.ZOrder
        objmonInfo.Visible = True
        objmonInfo.SetFocus
    End With
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Public Sub DateInfoGotFocus(ByRef intIndex As Integer)
'功能：MskDateInfo_GotFocus
    Dim objMSK As MaskEdBox
    '日期输入不用输入中文
    Call ChangeCtl
    zlCommFun.OpenIme False
End Sub

Public Sub DateInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'功能：MskDateInfo_KeyDown
    If intKeyCode = vbKeyF4 Or (intKeyCode = vbKeyDown And intShift = vbAltMask) Then
        Call DateInfoClick(gclsPros.CurrentForm.cmdDateInfo(intIndex))
    End If
End Sub

Public Sub DateInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'功能：MskDateInfo_KeyPress
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub DateInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'功能：MskDateInfo_Validate
    Dim objMSK As MaskEdBox
    Dim str年龄 As String
    Dim str入院时间 As String

    With gclsPros.CurrentForm
        Set objMSK = .mskDateInfo(intIndex)
        If Not IsDate(objMSK.Text) And objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
            Call ShowMessage(objMSK, "你输入的时间不是有效的时间，请重新输入。")
            blnCancel = True
            Exit Sub
        End If
        Select Case intIndex
            Case DC_出生日期

            Case DC_发病日期, DC_发病时间, DC_死亡时间, DC_确诊日期
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If intIndex = DC_确诊日期 Then
                        If objMSK.Text <> Replace(objMSK.Tag, "#", "_") Then
                            If Not CheckDateRange(objMSK.Text, True) Then
                                Call ShowMessage(objMSK, "你输入的时间不在病人入出院时间范围内，请重新输入。")
                                blnCancel = True
                            End If
                        End If
                    End If
                End If
            Case DC_出院时间, DC_入院时间
                If gclsPros.InTime = "" And intIndex = DC_出院时间 Then
                    If IsDate(.mskDateInfo(DC_入院时间).Text) Then
                        gclsPros.InTime = .mskDateInfo(DC_入院时间).Text
                    End If
                ElseIf intIndex = DC_入院时间 And gclsPros.OutTime = "" Then
                    If IsDate(.mskDateInfo(DC_出院时间).Text) Then
                        gclsPros.InTime = .mskDateInfo(DC_出院时间).Text
                    End If
                End If
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If intIndex = DC_出院时间 And gclsPros.InTime <> "" Then
                        If CDate(gclsPros.InTime) > CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "你输入的出院时间小于入院时间，请重新输入。")
                            Call DateInfoClick(intIndex)
                        ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                            Call ShowMessage(objMSK, "你输入的出院时间大于当前时间，请重新输入。")
                            Call DateInfoClick(intIndex)
                        Else
                            gclsPros.OutTime = objMSK.Text
                        End If
                    ElseIf intIndex = DC_入院时间 And gclsPros.OutTime <> "" Then
                        If CDate(gclsPros.OutTime) < CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "你输入的入院时间大于出院时间，请重新输入。")
                            Call DateInfoClick(intIndex)
                        ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                            Call ShowMessage(objMSK, "你输入的入院时间大于当前时间，请重新输入。")
                            Call DateInfoClick(intIndex)
                        Else
                            gclsPros.InTime = objMSK.Text
                        End If
                    End If
                End If
            Case DC_编目日期
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If CDate(Format(gclsPros.OutTime, "yyyy-mm-dd")) > CDate(objMSK.Text) And gclsPros.OutTime <> "" Then
                        Call ShowMessage(objMSK, "你输入的编目日期小于出院时间，请重新输入。")
                        Call DateInfoClick(intIndex)
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        Call ShowMessage(objMSK, "你输入的编目日期大于当前时间，请重新输入。")
                        Call DateInfoClick(intIndex)
                    End If
                End If
            Case DC_质控日期
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If IsDate(gclsPros.InTime) Then
                        If CDate(Format(gclsPros.InTime, "yyyy-mm-dd")) > CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "你输入的质控日期小于入院时间，请重新输入。")
                            Call DateInfoClick(intIndex)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

'chkInfo事件
Public Sub chkInfoClick(ByRef intIndex As Integer)
'功能：chkInfo_Click
    Dim blnCheck As Boolean

    With gclsPros.CurrentForm
         blnCheck = .chkInfo(intIndex).Value = 1
        Select Case intIndex
            Case CHK_是否确诊
                Call SetCtrlLocked(.mskDateInfo(DC_确诊日期), Not blnCheck, True)
                Call SetCtrlLocked(.cmdDateInfo(DC_确诊日期), Not blnCheck, True)
            Case CHK_随诊
                Call SetCtrlLocked(.txtSpecificInfo(SLC_随诊期限), Not blnCheck, True)
                Call SetCtrlLocked(.cboSpecificInfo(SLC_随诊期限), Not blnCheck, True)
                If blnCheck Then
                    Call CboSpecificInfoClick(SLC_随诊期限)
                End If
            Case CHK_病原学检查
                If Not blnCheck Then
                    .txtInfo(GC_病原学诊断).Tag = ""
                    .cmdInfo(GC_病原学诊断).Tag = ""
                End If
                Call SetCtrlLocked(.txtInfo(GC_病原学诊断), Not blnCheck, True)
                Call SetCtrlLocked(.cmdInfo(GC_病原学诊断), Not blnCheck)
            Case CHK_变异
                If gclsPros.PathVCauses Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_变异原因), Not blnCheck, True)
                Else
                    Call SetCtrlLocked(.txtInfo(GC_变异原因), Not blnCheck, True)
                End If
            Case CHK_进入路径
                Call SetCtrlLocked(.chkInfo(CHK_变异), Not blnCheck, True)
                Call SetCtrlLocked(.chkInfo(CHK_完成路径), Not blnCheck, True)
                Call SetCtrlLocked(.txtInfo(GC_退出原因), Not blnCheck, True)
                If Not blnCheck Then
                    If gclsPros.PathVCauses Then
                        Call SetCtrlLocked(.cboBaseInfo(BCC_变异原因), Not blnCheck, True)
                    Else
                        Call SetCtrlLocked(.txtInfo(GC_变异原因), Not blnCheck, True)
                    End If
                End If
            Case CHK_完成路径
                Call SetCtrlLocked(.txtInfo(GC_退出原因), blnCheck, True)
            Case CHK_住院物理约束
                If gclsPros.MedPageSandard = ST_云南省标准 Then
                    Call SetCtrlLocked(.txtSpecificInfo(SLC_约束总时间), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_约束方式), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_约束工具), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_约束原因), Not blnCheck, True)
                End If
            Case CHK_会诊情况
                Call SetCtrlLocked(.txtSpecificInfo(SLC_院内会诊), Not blnCheck, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_外院会诊), Not blnCheck, True)
                Call SetCtrlLocked(.txtInfo(GC_其他会诊), Not blnCheck, True)
            Case CHK_无过敏记录
                If mblnChk = False Then
                    If .vsAller.TextMatrix(.vsAller.FixedRows, AI_过敏药物) <> "" And .vsAller.TextMatrix(.vsAller.FixedRows, AI_过敏药物) <> "—" Then
                        If blnCheck Then
                            MsgBox "已经有过敏药物，不能标记为无。", vbInformation, gstrSysName
                            mblnChk = True
                            .chkInfo(intIndex).Value = 0
                            Exit Sub
                        End If
                    End If
                    Call SetCtrlLocked(.vsAller, blnCheck)
                    .vsAller.TextMatrix(.vsAller.FixedRows, AI_过敏药物) = IIf(blnCheck, "—", "")
                End If
                mblnChk = False
        End Select
        Call CheckValueChange(.chkInfo(intIndex))
    End With
End Sub

Public Sub ChkInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'功能：ChkInfo_KeyPress
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub ChkInfoGotFocus(ByRef intIndex As Integer)
'功能：ChkInfo_GotFocus
    '医院感染相关控件可见性以及位置
    Call ChangeCtl
    Call ShowInfectInfo(False)
End Sub

'chkFeeEdit事件
Public Sub ChkFeeEditClick()
'功能：ChkFeeEdit_Click
    Call SetCtrlLocked(gclsPros.CurrentForm.vsFees, gclsPros.CurrentForm.chkFeeEdit.Value = 0)
End Sub

Public Sub ChkFeeEditKeyPress(ByRef intKeyAscii As Integer)
'功能：ChkFeeEdit_KeyPress

    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        If gclsPros.CurrentForm.chkFeeEdit.Value = 1 Then
            gclsPros.CurrentForm.vsFees.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End If
End Sub

'optInput事件
Public Sub OptInputClick(ByRef intIndex As Integer)
'optInput_Click事件
    With gclsPros.CurrentForm
        Select Case intIndex
            Case OP_再住院无, OP_再住院有
                Call SetCtrlLocked(.txtInfo(GC_31天内再住院), intIndex = OP_再住院无, True)
            Case OP_ICU无, OP_ICU有
                Call SetCtrlLocked(.txtSpecificInfo(SLC_重症监护天), intIndex = OP_ICU无, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_重症监护小时), intIndex = OP_ICU无, True)
        End Select
    End With
    Call CheckValueChange
End Sub

Public Sub OptInputKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optInput_KeyPress事件
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optDiag事件
Public Sub optDiagClick(ByRef intIndex As Integer)
    If gclsPros.PatiType = PF_门诊 Then
        gclsPros.DiagInputXY = intIndex Mod 2
        gclsPros.DiagInputZY = gclsPros.DiagInputXY
    Else
        If intIndex < 2 Then
            gclsPros.DiagInputXY = intIndex Mod 2
        Else
            gclsPros.DiagInputZY = intIndex Mod 2
        End If
    End If
    Call CheckValueChange
End Sub

Public Sub optDiagGotFocus(ByRef intIndex As Integer)
'optDiag_GotFocus事件
    Call ChangeCtl
    Call ShowInfectInfo(False)
End Sub

Public Sub optDiagKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optDiag_KeyPress事件
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optState事件，初诊、复诊切换，门诊首页
Public Sub optStateClick(ByRef intIndex As Integer)
    Dim blnDo As Boolean
    Dim rsTmp As ADODB.Recordset

    '复诊：在诊断尚未录入的情况下则自动提取上次诊断
    If intIndex = OP_复诊 Then
        With gclsPros.CurrentForm
            If .chkInfo(CHK_传染病上传).Value = 1 Then
                blnDo = .vsDiagXY.Rows = .vsDiagXY.FixedRows + 1 And .vsDiagZY.Rows = .vsDiagZY.FixedRows + 1
                If blnDo Then blnDo = blnDo And .vsDiagXY.TextMatrix(.vsDiagXY.FixedRows, DI_诊断描述) = "" And .vsDiagZY.TextMatrix(.vsDiagZY.FixedRows, DI_诊断描述) = ""
                If blnDo Then
                    Set rsTmp = GetPatiDiagData(gclsPros.病人ID, gclsPros.主页ID, 0, True, , gclsPros.Moved)
                    If rsTmp.RecordCount <> 0 Then gclsPros.Is复诊 = True: gclsPros.IsLastDiag = True
                    Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, DT_门诊诊断XY, , -1)
                    If gclsPros.Have中医 Then
                        Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, DT_门诊诊断ZY, , -1)
                    End If
                End If
            End If
        End With
    End If
End Sub

'optAller事件
Public Sub OptAllerClick(ByRef intIndex As Integer)
'optAller_Click事件
    If intIndex = PC_按药品目录输入 Then
        gclsPros.AllerInput = 0
        gclsPros.UseTYT = False
    Else
        If Not gobjPass Is Nothing Then
            gclsPros.AllerInput = 1
            gclsPros.UseTYT = True
        Else
            gclsPros.AllerInput = 1
        End If
    End If
    Call CheckValueChange
End Sub

Public Sub OptAllerKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optAller_KeyPress事件
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optParaOPSInfo事件
Public Sub OptParaOPSInfoClick(ByRef intIndex As Integer)
'optParaOPSInfo_Click事件
    If intIndex = PC_按诊疗项目输入 Then
        gclsPros.OPSInput = 0
    Else
        gclsPros.OPSInput = 1
    End If
End Sub

Public Sub OptParaOPSInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'OptParaOPSInfo_KeyPress事件
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'chkParaOPSInfo事件
Public Sub ChkParaOPSInfoClick(ByRef intIndex As Integer)
'chkParaOPSInfo_Click事件
    If gclsPros.CurrentForm.chkParaOPSInfo(PC_未找到时自由录入).Value = 1 Then
        gclsPros.OPSFree = True
    Else
        gclsPros.OPSFree = False
    End If
    Call CheckValueChange
End Sub

Public Sub ChkParaOPSInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'chkParaOPSInfo_KeyPress事件
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'SpecificInfo事件
Public Sub SpecificInfoChange(ByRef intIndex As Integer)
'功能：txtSpecificInfo_Change
    Dim objTextBox As TextBox
    Dim objComboBox As ComboBox

    With gclsPros.CurrentForm
        Select Case intIndex
            Case SLC_年龄, SLC_婴幼儿年龄
            '数字年龄才带标准年龄单位
                Set objTextBox = .txtSpecificInfo(intIndex)
                Set objComboBox = .cboSpecificInfo(intIndex)
                If IsNumeric(objTextBox.Text) Or objTextBox.Text = "" Then
                    objComboBox.Visible = True
                    objComboBox.Tag = ""
                    If objComboBox.Container.Name = "fraCbo" Then
                        objComboBox.Container.Visible = True
                    End If
                
                    If intIndex = SLC_年龄 Then
                        If gclsPros.FuncType = f病案首页 Then
                            objTextBox.Width = 450
                        Else
                            objTextBox.Width = 360
                        End If
                    ElseIf intIndex = SLC_婴幼儿年龄 Then
                        DrawLineCTL objTextBox, 1
                        objTextBox.Width = 360
                        DrawLineCTL objTextBox
                    End If
                    If objComboBox.ListIndex = -1 Then objComboBox.ListIndex = 0
                Else
                    objComboBox.Visible = False
                    objComboBox.Tag = "年龄"
                    objComboBox.ListIndex = -1
                    If objComboBox.Container.Name = "fraCbo" Then
                        objComboBox.Container.Visible = False
                    End If
                
                    If intIndex = SLC_年龄 Then
                        If gclsPros.FuncType = f病案首页 Then
                            objTextBox.Width = 1250
                        Else
                            objTextBox.Width = 1150
                        End If
                    ElseIf intIndex = SLC_婴幼儿年龄 Then
                        DrawLineCTL objTextBox, 1
                        objTextBox.Width = 1250
                        DrawLineCTL objTextBox
                    End If
                End If
            Case SLC_抢救次数
                Set objTextBox = .txtSpecificInfo(intIndex)
                Call SetCtrlLocked(.txtInfo(GC_抢救病因), Val(objTextBox.Text) = 0, True)
                Call SetCtrlLocked(.cmdInfo(GC_抢救病因), Val(objTextBox.Text) = 0, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_成功次数), Val(objTextBox.Text) = 0, True)
                If Val(objTextBox.Text) > 0 Then
                    '主要诊断的出院情况不为死亡时,缺省：成功次数=抢救次数
                    If .Visible Then
                        If .vsDiagXY.TextMatrix(FindDiagRow(DT_出院诊断XY), DI_出院情况) <> "死亡" Then
                            .txtSpecificInfo(SLC_成功次数).Text = objTextBox.Text
                        ElseIf Val(objTextBox.Text) > 1 Then
                            .txtSpecificInfo(SLC_成功次数).Text = Val(objTextBox.Text) - 1
                        End If
                    End If
                End If
        End Select
        Call CheckValueChange(.txtSpecificInfo(intIndex))
    End With
End Sub

Public Sub SpecificInfoClick(ByRef intIndex As Integer, Optional ByVal blnCmdButton As Boolean)
'CmdSpecificInfo_Click事件
'参数：blnCmdButton=True-cmdButton控件，False-非cmdButton控件
    Dim blnALLPati As Boolean
    Dim arrDate() As String
    Dim str提取病人 As String, strIfdate As String
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim int主页id As Integer ''记录主页id
    Dim strKEY As String
    Dim objText住院号 As TextBox
    Dim blnEditInfo As Boolean
    Dim blnSign As Boolean

    On Error GoTo errH

    If blnCmdButton Then
        '病案系统才有住院号选择
        If intIndex = SLC_住院号 Then
            ReDim arrDate(2)
            arrDate(0) = zlDatabase.GetPara("开始日期", gclsPros.SysNo, gclsPros.Module)
            arrDate(1) = zlDatabase.GetPara("结束日期", gclsPros.SysNo, gclsPros.Module)
            blnSign = zlDatabase.GetPara("已签名的出院病人", gclsPros.SysNo, gclsPros.Module)
            If arrDate(0) = "" Or arrDate(1) = "" Then
                arrDate(0) = "": arrDate(1) = ""
                blnALLPati = True
            Else
                arrDate(0) = Format(arrDate(0), "yyyy-mm-dd")
                arrDate(1) = Format(arrDate(1), "yyyy-mm-dd")
            End If
            If Not blnALLPati Then blnALLPati = Val(zlDatabase.GetPara("提取所有出院病人", gclsPros.SysNo, gclsPros.Module)) = 1

            If Val(zlDatabase.GetPara("提取24小时内出院病人", gclsPros.SysNo, gclsPros.Module)) <> 1 Then
                If gclsPros.OutFile = "" Then
                    str提取病人 = " And (B.出院日期-B.入院日期)*24>=24"
                Else
                    str提取病人 = "住院时间>=24"
                End If
            End If

            If gclsPros.EditUnrecive = False And gclsPros.OutFile = "" Then
                str提取病人 = str提取病人 & " And E.接收时间 IS NOT NULL"
            End If
            If Not blnALLPati Then
                If gclsPros.OutFile = "" Then
                    strIfdate = " And B.出院日期 Between Trunc(To_Date('" & arrDate(0) & "','yyyy-mm-dd')) And Trunc(To_Date('" & arrDate(1) & "','yyyy-mm-dd'))+1-1/24/60/60"
                Else
                    strIfdate = " 出院日期 >= #" & Format(arrDate(0), "yyyy-mm-dd 00:00:00") & "# and 出院日期 <= #" & Format(arrDate(1), "yyyy-mm-dd 23:59:59") & "#"
                End If
            End If
            Set objText住院号 = gclsPros.CurrentForm.txtSpecificInfo(SLC_住院号)
            vRect = zlControl.GetControlRect(objText住院号.hwnd)
            '39906:刘鹏飞,2013-05-07,添加病案接收标志
            If gclsPros.OutFile = "" Then
                '问题26488 by lesfeng 2010-03-18 结清
                strSql = "" & _
                    " Select    to_char(Id) Id ,上级id,0 as 主页ID,末级,编码,名称,性别,入院日期,出院日期,住院次数,结清,接收,级数 as 序号id " & _
                    "   From (  Select id as id,上级id,0 as 末级,编码,名称 ,'' as 性别,'' as 入院日期,'' as 出院日期,'' as 住院次数,'' As 结清,'' as 接收,max(Level) as 级数" & _
                    "           From 部门表 " & _
                    "           Start with id in (  Select distinct(b.出院科室id)  as 上级id " & _
                    "                               From 病案主页 b,病人信息 a,病案接收记录 E" & _
                    "                               Where a.病人id=b.病人id And B.病人ID=E.病人ID(+) And B.主页ID=E.主页ID(+) and b.住院号 is not null and b.编目日期 is null and b.出院日期 is not null and nvl(b.病人性质,0)=0  " & _
                                                            strIfdate & str提取病人 & _
                    "                             )  Connect by prior 上级id = id group by id,上级id,末级,编码,名称 " & _
                    "       )"
                If blnSign Then
                    strSql = strSql & vbNewLine & _
                    "Union All" & _
                    " Select a.病人id || '-' || b.主页id ,c.id,b.主页ID,1 as 末级, to_char(b.住院号) as 编码,a.姓名 as 名称,a.性别,to_char(b.入院日期,'yyyy-mm-dd'),to_char(b.出院日期,'yyyy-mm-dd')," & _
                    "         to_char(Zl_获取住院次数或主页id(a.病人id,b.主页id,0)) ,decode(D.费用余额,null,'是',0,'是','否') As 结清,Decode(E.接收时间,Null,'否','是') As 接收,-9999 as 序号id" & _
                    " From 病人信息 a,病案主页 b,部门表 c,病人余额 D,病案接收记录 E " & _
                    " Where a.病人ID = b.病人ID and B.病人id = D.病人id(+) And D.类型(+)=2 And B.病人ID=E.病人ID(+) And B.主页ID=E.主页ID(+) " & _
                    "     and b.编目日期 is null " & _
                    "     and b.出院日期 is not null and b.住院号 is not null and nvl(b.病人性质,0)=0 " & _
                    "     and b.出院科室id =c.id " & strIfdate & str提取病人 & _
                    "And Exists" & _
                    " (Select *" & vbNewLine & _
                    "       From 病案主页从表" & vbNewLine & _
                    "       Where 病人id = b.病人id And 主页id = b.主页id And 信息名 In ('科主任签名', '主任医师签名', '住院医师签名', '住院医师签名'))" & vbNewLine & _
                    " order by 序号ID desc "
                Else
                    strSql = strSql & vbNewLine & _
                    "Union All" & _
                     " Select a.病人id || '-' || b.主页id ,c.id,b.主页ID,1 as 末级, to_char(b.住院号) as 编码,a.姓名 as 名称,a.性别,to_char(b.入院日期,'yyyy-mm-dd'),to_char(b.出院日期,'yyyy-mm-dd')," & _
                    "         to_char(Zl_获取住院次数或主页id(a.病人id,b.主页id,0)) ,decode(D.费用余额,null,'是',0,'是','否') As 结清,Decode(E.接收时间,Null,'否','是') As 接收,-9999 as 序号id" & _
                    " From 病人信息 a,病案主页 b,部门表 c,病人余额 D,病案接收记录 E " & _
                    " Where a.病人ID = b.病人ID and B.病人id = D.病人id(+) And D.类型(+)=2 And B.病人ID=E.病人ID(+) And B.主页ID=E.主页ID(+) " & _
                    "     and b.编目日期 is null " & _
                    "     and b.出院日期 is not null and b.住院号 is not null and nvl(b.病人性质,0)=0 " & _
                    "     and b.出院科室id =c.id " & strIfdate & str提取病人 & _
                    " order by 序号ID desc "
                End If

                    '刘兴宏:留观病人不能建病案
                    '39906:刘鹏飞,2013-05-07,需要显示编目总人数和接收人数
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "病人住院号", False, "", "待编目总人数[Count]人，其中已接收的病案人数[接收='是']人", False, False, True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
                    If rsTmp Is Nothing Then Exit Sub
                    If rsTmp.State <> 1 Or rsTmp.EOF Then Exit Sub
                    objText住院号.Text = rsTmp!编码 & ""
                    int主页id = Val(rsTmp!主页ID & "")
                    If gclsPros.OpenMode <> EM_编辑 Then
                        '如果得到住院号有问题就不继续
                        '78747:这段本来就多余，因为LoadPatiByInNo中已经有类似检查
                        'If Get住院次数Or主页id(Val(Split(rsTmp!ID & "", "-")(0)), Val(rsTmp!主页ID & ""), False, True) = False Then: Exit Sub
                        gclsPros.IsSelPati = True
                        '在住院号改变后清空病案录入信息
                        If Val(objText住院号.Text) <> Val(gclsPros.InNo) Then
                            If Not CheckMedPageChange Then
                                gclsPros.InfosChange = False
                            End If
                            If gclsPros.InfosChange = True And Val(gclsPros.InNo) <> 0 Then
                                gclsPros.InfosChange = False
                                If MsgBox("信息已发生变化，是否确认更换录入病人？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    Call gclsPros.InitCacheRecInfo
                                ElseIf Val(gclsPros.InNo) <> 0 Then
                                    objText住院号.Text = gclsPros.InNo
                                    Exit Sub
                                End If
                            Else
                                gclsPros.InfosChange = False
                            End If
                        End If
                        Call LoadPatiByInNo(objText住院号.Text, int主页id)
                        gclsPros.IsSelPati = False
                        Call AfterLoadPatiByNo
                    End If
            Else
                If str提取病人 <> "" Then
                    strIfdate = IIf(strIfdate = "", "", strIfdate & " and ") & str提取病人
                End If
                With frmPageMedRecNOSel
                    .Top = vRect.Top + 300
                    .Left = vRect.Left
                    strKEY = .ShowMe(gclsPros.CurrentForm, gclsPros.PatiOut, strIfdate)
                    If strKEY = "" Then Exit Sub
                    objText住院号.Text = Split(strKEY, "_")(0)
                    If Val(objText住院号.Text) = 0 Then
                        objText住院号.Text = ""
                        Exit Sub
                    End If
                    int主页id = Split(strKEY, "_")(1)
                    If gclsPros.OpenMode <> EM_编辑 Then
                        '在住院号改变后清空病案录入信息
                        gclsPros.IsSelPati = True
                        '在住院号改变后清空病案录入信息
                        If Val(objText住院号.Text) <> Val(gclsPros.InNo) Then
                            If Not CheckMedPageChange Then
                                gclsPros.InfosChange = False
                            End If
                            If gclsPros.InfosChange = True And Val(gclsPros.InNo) <> 0 Then
                                gclsPros.InfosChange = False
                                If MsgBox("信息已发生变化，是否确认更换录入病人？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    Call gclsPros.InitCacheRecInfo
                                ElseIf Val(gclsPros.InNo) <> 0 Then
                                    objText住院号.Text = gclsPros.InNo
                                    Exit Sub
                                End If
                            Else
                                gclsPros.InfosChange = False
                            End If
                        End If
                        '如果得到住院号有问题就不继续
                        If LoadPatiByInNo(objText住院号.Text, int主页id) = False Then gclsPros.IsSelPati = False
                        Call AfterLoadPatiByNo
                    End If
                End With
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SpecificInfoGotFocus(ByRef intIndex As Integer)
'功能：txtSpecificInfo_GotFocus
    '使用不到中文关闭中文输入法
    Call ChangeCtl
    zlCommFun.OpenIme False
    Call TxtGotFocus(gclsPros.CurrentForm.txtSpecificInfo(intIndex), True, True)
End Sub
    
Public Sub SpecificInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'功能：txtSpecificInfo_KeyDown
    Dim objTextBox As TextBox
    
    Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If intKeyCode = vbKeyReturn Then
        If intIndex = SLC_住院号 Then
            gclsPros.IsReturn = True: Exit Sub
        ElseIf (intIndex = SLC_成功次数 Or intIndex = SLC_抢救次数) Then
             Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(SLC_抢救次数)
        ElseIf intIndex = SLC_单位邮编 Or intIndex = SLC_家庭邮编 Or intIndex = SLC_户口邮编 Then
            If ((Not IsNumeric(objTextBox.Text)) Or Len(objTextBox.Text) > 6 Or InStr(objTextBox.Text, ".") > 0) And objTextBox.Text <> "" Then
                Call SelectYouBian(objTextBox)
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    Else
        gclsPros.IsReturn = False
    End If
End Sub


Public Sub SelectYouBian(objTextBox As TextBox)
    '功能：邮编选择器
    Dim strInput As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    strInput = objTextBox.Text
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(strInput) Then
            strSql = strSql & " And A.名称 Like [1] "
        Else
            strSql = strSql & " And A.简码 Like [1] "
        End If
    Else
        Exit Sub
    End If
    strSql = "Select Rownum as ID,名称,简码,邮编  From 区域 A " & _
             "Where 邮编 is not null " & strSql & " Order by 编码"
    vPoint = GetCoordPos(objTextBox.hwnd, 0, 0)
    Set rsTmp = zlDatabase.ShowSQLSelect(objTextBox.Parent, strSql, 0, "邮编", False, "", "", False, _
        False, True, vPoint.X, vPoint.Y, objTextBox.Height, False, False, False, UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        objTextBox.Text = rsTmp!邮编 & ""
    End If
End Sub


Public Sub SpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'功能：txtSpecificInfo_KeyPress
    Dim objTextBox As TextBox
    Dim objCboTmp As ComboBox
    Dim blnCBO As Boolean
    Dim strMask As String
    Dim strTmp As String
    Dim blnEditInfo As Boolean

    On Error Resume Next
    Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If Err.Number <> 0 Then
        Set objCboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex): blnCBO = True
        Err.Clear: On Error GoTo 0
    Else
        On Error GoTo 0
    End If
    
    Select Case intIndex
        Case SLC_住院号
            If gclsPros.OpenMode <> EM_编辑 And gclsPros.IsReturn Then
            '检查信息是否变化
                If Not CheckMedPageChange Then
                    gclsPros.InfosChange = False
                End If
                If Val(objTextBox.Text) <> Val(gclsPros.InNo) And Trim(objTextBox.Text) <> "" And IsHavePageNos(CT_住院号, False, Val(objTextBox.Text)) Then
                    If gclsPros.InfosChange And Val(gclsPros.InNo) <> 0 Then
                        gclsPros.InfosChange = False
                        If MsgBox("信息已发生变化，是否确认更换录入病人？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYesNo Then
                            objTextBox.Text = gclsPros.InNo
                            Exit Sub '不更换病人则不做处理
                        Else
                            Call gclsPros.InitCacheRecInfo
                        End If
                    Else
                        gclsPros.InfosChange = False
                    End If
                End If
                If Val(objTextBox.Text) <> Val(gclsPros.InNo) Or objTextBox.Text = "" Then
                    If LoadPatiByInNo(objTextBox.Text) Then
                        Call AfterLoadPatiByNo
                    Else
                        gclsPros.InNo = ""
                    End If
                    gclsPros.IsReturn = False
                ElseIf Val(objTextBox.Text) = Val(gclsPros.InNo) Then
                    intKeyAscii = 0
                    Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                End If
            End If
        Case SLC_单位邮编, SLC_户口邮编, SLC_家庭邮编
            If Chr(intKeyAscii) = "." Then
                intKeyAscii = 0
            End If
    End Select

    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
        '限制输入长度
        If Not blnCBO Then
            If objTextBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If

        Select Case intIndex
            Case SLC_家庭电话, SLC_单位电话, SLC_联系人电话
                strMask = "1234567890-()"
            Case SLC_住院号, SLC_抢救次数, SLC_成功次数, _
                    SLC_随诊期限, SLC_昏迷时间入院前_小时, SLC_昏迷时间入院前_分钟, SLC_昏迷时间入院后_分钟, SLC_昏迷时间入院后_小时, _
                    SLC_昏迷时间入院前_天, SLC_昏迷时间入院后_天, SLC_呼吸机使用, SLC_重症监护天, SLC_重症监护小时, SLC_Apgar, SLC_QQ, SLC_外院会诊, SLC_院内会诊
                strMask = "1234567890"
            Case SLC_输红细胞, SLC_输血小板, SLC_输血浆, SLC_输全血, SLC_输白蛋白, SLC_自体回收, SLC_ICU, _
                    SLC_CCU, SLC_一级护理, SLC_二级护理, SLC_三级护理, SLC_特护, SLC_约束总时间, SLC_身高, SLC_体重, SLC_距上次住院时间
                strMask = "1234567890."
            Case SLC_新生儿出生体重, SLC_新生儿入院体重
                strMask = "1234567890.;"
            Case SLC_婴幼儿年龄_DAY
                strMask = "0123456789"
        End Select

        If strMask <> "" Then
            If InStr(strMask, Chr(intKeyAscii)) = 0 Then
                intKeyAscii = 0: Exit Sub
            ElseIf intIndex = SLC_Apgar Then
                    '保证txtApgar输入值在0-10之间
                    If objTextBox.Text <> "" And objTextBox.Text <> "1" Or _
                        objTextBox.Text <> "" And objTextBox.Text = "1" And Chr(intKeyAscii) <> "0" Then
                        intKeyAscii = 0: Exit Sub
                    End If
            End If
        End If
    End If
End Sub

Public Sub SpecificInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：txtSpecificInfo_MouseDown
    Call TxtMouseDown(gclsPros.CurrentForm.txtSpecificInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub SpecificInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：txtSpecificInfo_MouseUp
    Call TxtMouseUp(gclsPros.CurrentForm.txtSpecificInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub SpecificInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'功能：txtSpecificInfo_Validate
    Dim objText As TextBox
    Dim objTextDate As MaskEdBox

    Set objText = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    Select Case intIndex
        Case SLC_年龄
            '没有年龄有出生日期时计算一下年龄
            Set objTextDate = gclsPros.CurrentForm.mskDateInfo(DC_出生日期)
            If objText.Text = "" And IsDate(objTextDate.Text) Then
                objTextDate.Tag = ""
'                Call txt出生日期_Validate(False)
            End If
        Case SLC_抢救次数, SLC_成功次数, SLC_随诊期限, SLC_输红细胞, SLC_输血小板, SLC_输血浆, SLC_输全血, SLC_自体回收
            If objText.Text <> "" Then
                If Not IsNumeric(objText.Text) Then
                    objText.Text = ""
                ElseIf Val(objText.Text) <= 0 And intIndex <> SLC_成功次数 Then
                    objText.Text = ""
                ElseIf intIndex = SLC_抢救次数 Or intIndex = SLC_成功次数 Or intIndex = SLC_随诊期限 Then
                    If IsNumeric(objText.Text) Then
                        objText.Text = Int(Val(objText.Text))
                    End If
                End If
            End If
    End Select
End Sub

'CboBaseInfo事件
Public Sub CboBaseInfoChange(ByRef intIndex As Integer)
'CboBaseInfo_Change事件
    Dim cboTmp As ComboBox
    Dim lngPos As Long, lnglen As Long

    If gclsPros.IsReturn Then Exit Sub
    Select Case intIndex
        Case BCC_身份证
            Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
            gclsPros.IsReturn = True
            If Cbo.FindIndex(cboTmp, cboTmp.Text, True) = -1 Then
                '不规则的输入
                If Not zlStr.CheckCharScope(cboTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                    cboTmp.Text = ""
                Else
                    If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                        If zlCommFun.ActualLen(cboTmp.Text) > 18 Then
                            cboTmp.Text = Mid(cboTmp.Text, 1, 18)
                        End If
                    End If
                End If
            End If
            '门诊首页密文现实
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                If cboTmp.Tag <> "不触发Change事件" Then
                    lngPos = InStr(cboTmp.Text, "*")
                    lnglen = Len(Mid(cboTmp.Text, 13, 2))
                    Select Case lngPos
                        Case 0
                            cboTmp.Tag = cboTmp.Text
                        Case Else 'Is <= 12
                            cboTmp.Tag = Mid(cboTmp.Text, 1, lngPos - 1)
                            cboTmp.Text = cboTmp.Tag
                            cboTmp.SelStart = Len(cboTmp.Text)
                    End Select
                End If '
            Else
                cboTmp.Tag = cboTmp.Text
            End If
            gclsPros.IsReturn = False
    End Select
    Call CheckValueChange
End Sub


Public Sub CboSpecificInfoClick(ByRef intIndex As Integer)
'cboSpecificInfo_Click事件
    Dim objPic As Object
    Dim objFra As Object
    Dim lngNum As Long
    
    With gclsPros.CurrentForm
        Select Case intIndex
            Case SLC_随诊期限
                Call SetCtrlLocked(.txtSpecificInfo(intIndex), .cboSpecificInfo(intIndex).Text = "终身", True)
                If .cboSpecificInfo(intIndex).Text <> "终身" Then
                    If .Visible Then zlControl.ControlSetFocus (.txtSpecificInfo(intIndex))
                End If
            Case SLC_婴幼儿年龄
                If gclsPros.LoadFinish Then
                    Set objFra = .cboSpecificInfo(intIndex).Container
                    If .cboSpecificInfo(intIndex).Text = "月" Then
                        .txtSpecificInfo(SLC_婴幼儿年龄_DAY).Visible = True
                        .lblSpecificInfo(SLC_婴幼儿年龄_DAY).Visible = True
                        DrawLineCTL .txtSpecificInfo(SLC_婴幼儿年龄_DAY), 1
                        lngNum = .txtSpecificInfo(SLC_婴幼儿年龄).Left + .txtSpecificInfo(SLC_婴幼儿年龄).Width + 120
                        .txtSpecificInfo(SLC_婴幼儿年龄_DAY).Left = lngNum
                        .lblSpecificInfo(SLC_婴幼儿年龄_DAY).Left = lngNum
                        DrawLineCTL .txtSpecificInfo(SLC_婴幼儿年龄_DAY)
                        DrawLineCTL objFra, 1
                        objFra.Left = lngNum + .txtSpecificInfo(SLC_婴幼儿年龄_DAY).Width + 120
                        DrawLineCTL objFra
                    Else
                        .txtSpecificInfo(SLC_婴幼儿年龄_DAY).Text = ""
                        .txtSpecificInfo(SLC_婴幼儿年龄_DAY).Visible = False
                        .lblSpecificInfo(SLC_婴幼儿年龄_DAY).Visible = False
                        
                        If .cboSpecificInfo(intIndex).Tag <> "年龄" Then
                            DrawLineCTL objFra, 1
                            objFra.Left = .txtSpecificInfo(SLC_婴幼儿年龄).Left + .txtSpecificInfo(SLC_婴幼儿年龄).Width + 120
                            DrawLineCTL objFra
                        Else
                            DrawLineCTL objFra, 1  '清除线条
                            DrawLineCTL .txtSpecificInfo(SLC_婴幼儿年龄) '重绘线条避免单位线条清除时将婴幼儿年龄的线条也清除掉部分
                        End If
                    End If
                End If
        End Select
        Call CheckValueChange(.txtSpecificInfo(intIndex))
    End With
End Sub

Public Sub CboSpecificInfoGotFocus(ByRef intIndex As Integer)
'CboSpecificInfo_GotFocus事件
    Call ChangeCtl
    With gclsPros.CurrentForm
        '限定输入的项目一般不会输入汉字
        zlCommFun.OpenIme False
        Call ShowInfectInfo(False)
    End With
End Sub

Public Sub CboSpecificInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboSpecificInfo_KeyDown事件
    With gclsPros.CurrentForm
        If intKeyCode = vbKeyDelete Then
            If .cboSpecificInfo(intIndex).Style = 2 And .cboSpecificInfo(intIndex).ListIndex <> -1 Then
                .cboSpecificInfo(intIndex).ListIndex = -1
            End If
        End If
    End With
End Sub

Public Sub cboSpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'cboSpecificInfo_KeyPress事件
    Dim lngIdx As Long
    Dim cboTmp As ComboBox
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        zlCommFun.PressKey vbKeyTab: mblnReturn = True
    Else
        Set cboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
        If intIndex = SLC_年龄 And gclsPros.FuncType = f病案首页 Or intIndex = SLC_婴幼儿年龄 Then
          If cboTmp.ListCount + vbKey1 >= intKeyAscii And intKeyAscii >= vbKey1 Then
                If intKeyAscii - vbKey1 <= cboTmp.ListCount Then
                    cboTmp.ListIndex = intKeyAscii - vbKey1
                End If
            End If
        Else
            lngIdx = zlControl.CboMatchIndex(cboTmp.hwnd, intKeyAscii)
            If lngIdx = -1 And cboTmp.ListCount > 0 Then lngIdx = 0
            cboTmp.ListIndex = lngIdx
        End If
    End If
End Sub

Public Sub cboSpecificInfoLostFocus(ByRef intIndex As Integer)
'cboSpecificInfo_LostFocus事件
    Dim lngIdx As Long
    Dim cboTmp As ComboBox, txtTmp As TextBox

    Set cboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
    Set txtTmp = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If intIndex = SLC_年龄 And gclsPros.FuncType = f病案首页 Or intIndex = SLC_婴幼儿年龄 Then
        If Not ValidateAge(txtTmp, cboTmp, IIf(intIndex = SLC_婴幼儿年龄, 1, 0)) Then Exit Sub
    End If
End Sub

Public Sub txtDateInfoGotFocus(Index As Integer)
    Call ChangeCtl
    Call TxtGotFocus(gclsPros.CurrentForm.txtDateInfo(Index), True, True)
End Sub

'cmdAutoLoad事件
Public Sub CmdAutoLoadClick(ByRef intIndex As Integer)
'cmdAutoLoad_Click事件
    Dim strSql As String, rsTmp As Recordset
    Dim DateSs As Date          '该病人最早的手术时间
    Dim rsTime As ADODB.Recordset
    Dim vsTmp As VSFlexGrid
    Dim i As Long, j As Long, LngRow As Long
    Dim blnClear As Boolean
    Dim strPrivs As String
    Dim strUseStage As String

    On Error GoTo errH
    Select Case intIndex
        Case ALC_抗生素 '抗菌药自动提取
            strSql = "Select Min(NVL(to_date(c.标本部位,'yyyy-mm-dd hh24:mi:ss'),c.开始执行时间)) as 使用时间" & vbNewLine & _
                    " From 诊疗项目目录 A, 病人医嘱记录 C" & vbNewLine & _
                    " Where  a.Id = c.诊疗项目id and a.类别='F' And c.病人id = [1] And c.主页id = [2] And c.医嘱状态=8"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID)
            If rsTmp.RecordCount > 0 Then DateSs = CDate(Format(NVL(rsTmp!使用时间, 0), "yyyy-MM-dd"))

            strSql = "Select distinct ID, 医嘱id, 上级id, 编码, 名称, 单位, 执行时间方案, 频率间隔, 间隔单位, 频率次数, 上次执行时间, 开始执行时间, 结束时间," & vbNewLine & _
                    "       Sum(Ddd数) Over(Partition By ID,用药目的) As Ddd数, Count(1) Over(Partition By 相关id) As 联合用药,药名ID,decode(用药目的,1,'预防',2,'治疗',' ') as 用药目的" & vbNewLine & _
                    "From   (Select Distinct ID, 医嘱id, 上级id, 编码, 名称, 单位, 执行时间方案, 频率间隔, 间隔单位, 频率次数, 上次执行时间, 开始执行时间, 结束时间," & vbNewLine & _
                    "                Sum(数次) Over(Partition By ID, 医嘱id, 相关id,用药目的) * 剂量系数 / Decode(Ddd值, 0, Null, Ddd值) As Ddd数, 相关id,药名ID,用药目的" & vbNewLine & _
                    "         From   (Select z.Id, a.Id As 医嘱id, z.分类id As 上级id, z.编码, z.名称, z.计算单位 As 单位, a.执行时间方案, a.频率间隔, a.间隔单位, a.频率次数," & vbNewLine & _
                    "                         a.上次执行时间, a.开始执行时间, Nvl(a.上次执行时间, Nvl(a.执行终止时间, a.开始执行时间)) As 结束时间, a.相关id, f.数次, h.剂量系数," & vbNewLine & _
                    "                         Nvl((Select e.Ddd值 From 诊疗用法用量 E Where e.项目id = a.诊疗项目id And e.用法id = r.诊疗项目id), h.Ddd值) As Ddd值,A.诊疗项目ID as 药名ID,A.用药目的" & vbNewLine & _
                    "                  From   病人医嘱记录 A, 病人医嘱记录 R, 住院费用记录 F, 药品规格 H, 药品特性 B, 诊疗项目目录 Z" & vbNewLine & _
                    "                  Where  a.诊疗项目id = b.药名id And a.诊疗类别 In ('5', '6') And" & vbNewLine & _
                    "                         (a.医嘱期效 = 0 And a.上次执行时间 Is Not Null Or a.医嘱期效 = 1 And a.医嘱状态 = 8) And Nvl(b.抗生素, 0) <> 0 And" & vbNewLine & _
                    "                         a.相关id = r.Id And a.Id = f.医嘱序号 And f.记录状态 <> 0 And f.收费细目id = h.药品id And b.药名id = z.Id And" & vbNewLine & _
                    "                         f.记录性质 <> 12 And a.病人id = [1] And a.主页id = [2]))" & vbNewLine & _
                    "Order  By Ddd数 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID)

            If rsTmp.RecordCount = 0 Then
                MsgBox "没有找到该病人的抗菌药物使用记录。", vbInformation, gstrSysName
                Exit Sub
            End If
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                Do While Not rsTmp.EOF
                    LngRow = 0
                    strUseStage = GetKSSUseStage(CDate(Format(rsTmp!开始执行时间 & "", "yyyy-MM-dd")), CDate(Format(rsTmp!结束时间 & "", "yyyy-MM-dd")), DateSs)
                    For i = .FixedRows To .Rows - 1
                        '定位行
                        LngRow = 0
                        If Val(rsTmp!ID & "") = 0 Then
                            Exit For '加载的时候不会加载，因此退出进入下一次循环
                        ElseIf .TextMatrix(i, KI_抗菌药物名) & "" = "" Then
                            LngRow = i: Exit For
                        ElseIf Val(rsTmp!ID & "") = Val(.RowData(i) & "") And .TextMatrix(i, KI_用药目的) = rsTmp!用药目的 & "" And (.TextMatrix(i, KI_使用阶段) = "" Or .TextMatrix(i, KI_使用阶段) = strUseStage) Then
                            LngRow = -1 * i: Exit For
                        ElseIf i = .Rows - 1 Then
                            .AddItem ""
                            LngRow = .Rows - 1
                            Exit For
                        End If
                    Next

                    If LngRow > 0 Then
                        .RowData(LngRow) = Val(rsTmp!药名id & "")
                        .TextMatrix(LngRow, KI_抗菌药物名) = rsTmp!名称 & ""
                        .Cell(flexcpData, LngRow, KI_抗菌药物名) = .TextMatrix(LngRow, KI_抗菌药物名)
                        .TextMatrix(LngRow, KI_用药目的) = rsTmp!用药目的 & ""
                        .TextMatrix(LngRow, KI_DDD数) = FormatEx(Val(rsTmp!DDD数 & ""), 2)
                        .TextMatrix(LngRow, KI_联合用药) = decode(Val(rsTmp!联合用药 & ""), 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                        .TextMatrix(LngRow, KI_使用阶段) = strUseStage
                    Else '相同记录
                        LngRow = Abs(LngRow)
                        If .TextMatrix(LngRow, KI_DDD数) = "" Then .TextMatrix(LngRow, KI_DDD数) = FormatEx(Val(rsTmp!DDD数 & ""), 2)
                        If decode(.TextMatrix(i, KI_联合用药), "Ⅰ种", 1, "Ⅱ联", 2, "Ⅲ联", 3, "Ⅳ联", 4, ">Ⅳ联", 999, 0) < Val(rsTmp!联合用药 & "") Then
                            .TextMatrix(i, KI_联合用药) = decode(Val(rsTmp!联合用药 & ""), 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                        End If
                    End If
                    If LngRow <> 0 Then '获取使用天数
                        .TextMatrix(LngRow, KI_使用天数) = GetKSSUseDay(Val(rsTmp!医嘱ID), Val(.RowData(LngRow)), NVL(rsTmp!执行时间方案) & "", CDate(rsTmp!开始执行时间), CDate(rsTmp!结束时间), _
                                NVL(rsTmp!频率次数, 0), NVL(rsTmp!频率间隔, 0), NVL(rsTmp!间隔单位), NVL(rsTmp!用药目的), rsTime) & ""
                    End If
                    rsTmp.MoveNext
                Loop
                Call ChangeVSFHeight(vsTmp, True)
            End With
        Case ALC_手术 '手术自动读取
            strPrivs = GetInsidePrivs(p手麻接口, , 2400)
            If InStr(strPrivs, "内部接口") > 0 Then
                gclsPros.CurrentForm.lblAutoInfo.Visible = True
                gclsPros.CurrentForm.lblAutoInfo = "数据来源：手麻管理系统(默认)"
                Set rsTmp = AutoGetOPSInfo(True, gclsPros.病人ID, gclsPros.主页ID)
            Else
                If gblnHaveOPS Then
                    gclsPros.CurrentForm.lblAutoInfo.Visible = True
                    gclsPros.CurrentForm.lblAutoInfo = "数据来源：病人医嘱相关(没有【手麻管理系统-手麻接口管理-内部接口】权限)"
                Else
                    gclsPros.CurrentForm.lblAutoInfo.Visible = True
                    gclsPros.CurrentForm.lblAutoInfo = "数据来源：病人医嘱相关(未安装手麻管理系统)"
                End If
                Set rsTmp = AutoGetOPSInfo(False, gclsPros.病人ID, gclsPros.主页ID)
            End If

            If Not rsTmp.EOF Then
                Set vsTmp = gclsPros.CurrentForm.vsOPS
                '检查界面手术表格中是否有手术信息
                For i = vsTmp.FixedRows To vsTmp.Rows - 1
                    If vsTmp.TextMatrix(i, PI_手术日期) <> "" Or vsTmp.TextMatrix(i, PI_手术名称) <> "" Then
                        If MsgBox("是否清空原有的手术信息？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            blnClear = True
                        End If
                        Exit For
                    End If
                Next
                rsTmp.MoveFirst
                With vsTmp
                    If blnClear Then .Rows = .FixedRows
                    LngRow = IIf(.TextMatrix(.Rows - 1, PI_手术名称) <> "", .Rows, .Rows - 1)
                    .Rows = .Rows + rsTmp.RecordCount + IIf(.TextMatrix(.Rows - 1, PI_手术名称) <> "", 1, 0)
                    Call ChangeVSFHeight(vsTmp, True)
                    For i = LngRow To LngRow + rsTmp.RecordCount - 1
                        .TextMatrix(i, PI_手术日期) = Format(NVL(rsTmp!手术开始时间, rsTmp!手术日期) & "", "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, PI_结束日期) = Format(NVL(rsTmp!手术结束时间, rsTmp!手术日期) & "", "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, PI_手术编码) = rsTmp!手术编码 & ""
                        .TextMatrix(i, PI_手术名称) = rsTmp!已行手术 & ""
                        .TextMatrix(i, PI_主刀医师) = rsTmp!主刀医师 & ""
                        .TextMatrix(i, PI_助产护士) = rsTmp!助产护士 & ""
                        .TextMatrix(i, PI_助手1) = rsTmp!第一助手 & ""
                        .TextMatrix(i, PI_助手2) = rsTmp!第二助手 & ""
                        .TextMatrix(i, PI_麻醉方式) = rsTmp!麻醉方式 & ""
                        .TextMatrix(i, PI_麻醉医师) = rsTmp!麻醉医师 & ""
                        If rsTmp!切口 & rsTmp!愈合 & "" <> "" Then
                            .TextMatrix(i, PI_切口愈合) = rsTmp!切口 & "/" & rsTmp!愈合
                        End If
                        .TextMatrix(i, PI_手术操作ID) = Val(rsTmp!手术操作ID & "")
                        .TextMatrix(i, PI_诊疗项目ID) = Val(rsTmp!诊疗项目id & "")
                        .TextMatrix(i, PI_麻醉ID) = Val(rsTmp!ID & "")
                        .TextMatrix(i, PI_麻醉类型) = rsTmp!麻醉类型 & ""
                        .TextMatrix(i, PI_手术情况) = rsTmp!手术情况 & ""
                        .TextMatrix(i, PI_ASA分级) = rsTmp!asa分级 & ""
                        .TextMatrix(i, PI_NNIS分级) = rsTmp!NNIS分级 & ""
                        .TextMatrix(i, PI_手术级别) = rsTmp!手术级别 & ""
                        .TextMatrix(i, PI_再次手术) = IIf(Val(rsTmp!再次手术 & "") = 1, -1, 0)
                        .TextMatrix(i, PI_麻醉开始时间) = Format(rsTmp!麻醉开始时间 & "", "yyyy-MM-dd HH:mm")
                        .Cell(flexcpData, i, PI_手术名称) = rsTmp!手术原名 & ""
                        '记录用于编辑恢复
                        For j = 0 To .Cols - 1
                            If j = PI_手术名称 And .TextMatrix(i, PI_手术编码) <> "" Then
                                If .Cell(flexcpData, i, j) = "" Then
                                    .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                                End If
                            Else
                                .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                            End If
                        Next
                        .Cell(flexcpData, i, PI_手术级别) = IIf(rsTmp!手术级别 & "" = "", 0, 1)
                        rsTmp.MoveNext
                    Next
                End With
            End If
        Case ALC_过敏记录
             If gclsPros.CurrentForm.chkInfo(CHK_无过敏记录).Value = 1 Then gclsPros.CurrentForm.chkInfo(CHK_无过敏记录).Value = 0
             strSql = " Select Distinct a.Id, a.记录来源, a.过敏时间, a.药物id, a.药物名, a.过敏反应, a.过敏源编码, a.记录时间" & vbNewLine & _
                      " From (Select a.Id, a.记录来源, a.过敏时间, a.药物id, a.药物名, a.过敏反应, a.过敏源编码, a.记录时间" & vbNewLine & _
                      "       From 病人过敏记录 A," & vbNewLine & _
                      "            (Select c.病人id, c.主页id, c.药物id, Max(c.记录时间) As 记录时间" & vbNewLine & _
                      "              From 病人过敏记录 C" & vbNewLine & _
                      "              Where c.记录来源 = 2 And c.病人id = [1] And c.主页id = [2]" & vbNewLine & _
                      "              Group By c.病人id, c.主页id, c.药物id) B" & vbNewLine & _
                      "       Where a.结果 = 1 And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
                      "             ((a.记录来源 = 2 And a.记录时间 = b.记录时间 And a.药物id = b.药物id) Or a.记录来源 in (1,3))" & vbNewLine & _
                      "       Union" & vbNewLine & _
                      "       Select a.Id, a.记录来源, a.过敏时间, a.药物id, a.药物名, a.过敏反应, a.过敏源编码, a.记录时间" & vbNewLine & _
                      "       From 病人过敏记录 A" & vbNewLine & _
                      "       Where a.结果 = 1 And a.病人id = [1] And a.主页id = [2] And a.记录来源 in (1,3) ) A" & vbNewLine & _
                      " Order By Nvl(Trunc(a.过敏时间), a.记录时间) Desc,a.记录来源 Desc,a.药物名"
             
            On Error GoTo errH
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "首页获取过敏信息", gclsPros.病人ID, gclsPros.主页ID)
            
            With gclsPros.CurrentForm.vsAller
                If rsTmp.EOF Then
                    MsgBox "没有提取到任何的过敏记录信息。", vbInformation, gstrSysName
                Else
                    rsTmp.MoveFirst

                    .Rows = .FixedRows
                    For i = 1 To rsTmp.RecordCount
                        LngRow = -1
                        If Not IsNull(rsTmp!药物ID) Then
                            LngRow = .FindRow(rsTmp!药物ID & "", , AI_药物ID, , True)
                        ElseIf Not IsNull(rsTmp!药物名) Then
                            LngRow = .FindRow(rsTmp!药物名 & "", , AI_过敏药物, , True)
                        End If
                        If LngRow = -1 Then
                            For j = .FixedRows To .Rows - 1
                                If .TextMatrix(j, AI_过敏药物) = "" Then
                                    LngRow = j
                                End If
                            Next
                            
                            If LngRow = -1 Then
                                .Rows = .Rows + 1
                                LngRow = .Rows - 1
                            End If
                            
                            .TextMatrix(LngRow, AI_过敏时间) = Format(rsTmp!过敏时间, "yyyy-MM-dd")
                            .TextMatrix(LngRow, AI_过敏药物) = NVL(rsTmp!药物名)
                            .TextMatrix(LngRow, AI_过敏反应) = NVL(rsTmp!过敏反应)
                            .TextMatrix(LngRow, AI_过敏源编码) = NVL(rsTmp!过敏源编码)
                            .TextMatrix(LngRow, AI_药物ID) = rsTmp!药物ID & ""
                            .TextMatrix(LngRow, AI_过敏来源) = rsTmp!记录来源 & ""
                            '数据备份存储
                            .Cell(flexcpData, LngRow, AI_过敏时间) = .TextMatrix(LngRow, AI_过敏时间)
                            .Cell(flexcpData, LngRow, AI_过敏药物) = .TextMatrix(LngRow, AI_过敏药物)
                            .Cell(flexcpData, LngRow, AI_过敏反应) = .TextMatrix(LngRow, AI_过敏反应)
                            .Cell(flexcpData, LngRow, AI_过敏源编码) = .TextMatrix(LngRow, AI_过敏源编码)
                            .Cell(flexcpData, LngRow, AI_药物ID) = .TextMatrix(LngRow, AI_药物ID)
                            .RowData(LngRow) = Val(rsTmp!ID & "")
                        End If
                        rsTmp.MoveNext
                    Next
                    .Rows = .Rows + 1   '增加一行空行
                    .Row = .FixedRows
                    .Col = AI_过敏药物
                    Call ChangeVSFHeight(gclsPros.CurrentForm.vsAller, True, 300, 3)
                End If
            End With
        Case ALC_临床路径
            strSql = "Select Decode(c.性质, 2, c.名称, '') As 名称,b.状态" & vbNewLine & _
                "From 病人路径评估 A, 病人临床路径 B, 变异常见原因 C" & vbNewLine & _
                "Where a.路径记录id(+) = b.Id And b.当前天数 = a.天数(+) And Nvl(b.当前阶段id, b.前一阶段id) = a.阶段id(+) And b.状态 <> 0 And a.变异原因 = c.编码(+) And b.病人id = [1] And b.主页id = [2]"

            On Error GoTo errH
            With gclsPros.CurrentForm
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, gclsPros.病人ID, gclsPros.主页ID)
                If rsTmp.RecordCount > 0 Then
                    .chkInfo(CHK_进入路径).Value = 1
                    If Val(rsTmp!状态 & "") = 3 Then
                        .chkInfo(CHK_完成路径).Value = 0
                        .txtInfo(GC_退出原因).Text = rsTmp!名称 & ""
                    ElseIf Val(rsTmp!状态 & "") = 2 Then
                        .chkInfo(CHK_完成路径).Value = 1
                    End If
                Else
                    .chkInfo(CHK_进入路径).Value = 0
                End If
                '提取变异情况
                strSql = "Select Count(1) Over(Partition By b.病人id, b.主页id) As 变异数, c.名称 As 变异原因" & vbNewLine & _
                        "From 病人路径评估 A, 病人临床路径 B, 变异常见原因 C" & vbNewLine & _
                        "Where a.路径记录id = b.Id And c.编码(+) = a.变异原因 And a.评估结果 = -1 And b.病人id = [1] And b.主页id = [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, gclsPros.病人ID, gclsPros.主页ID)
                If rsTmp.RecordCount > 0 Then
                    .chkInfo(CHK_变异).Value = 1
                    If Val(rsTmp!变异数 & "") = 1 And Not gclsPros.PathVCauses Then
                        .txtInfo(GC_变异原因).Text = rsTmp!变异原因 & ""
                    End If
                Else
                    .chkInfo(CHK_变异).Value = 0
                End If
            End With
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'cboBaseInfo事件
Public Sub CboBaseInfoClick(ByRef intIndex As Integer)
'cboBaseInfo_Click事件
    Dim rsTmp As ADODB.Recordset
    Dim objTextBox As TextBox
    Dim strTmp As String
    Dim blnLocked As Boolean

    With gclsPros.CurrentForm
        Select Case intIndex
            Case BCC_出院方式
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                blnLocked = Not (strTmp Like "*转院*" Or strTmp Like "*转社区*")
                Call SetCtrlLocked(.txtInfo(GC_转入医疗机构), blnLocked, True)
                Call SetCtrlLocked(.cmdInfo(GC_转入医疗机构), blnLocked, True)
                If gblnSet Then Exit Sub
                Call ChangeOutInfo(strTmp, True) '调整诊断的出院情况
            Case BCC_关系
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                If strTmp Like "*其他*" Then
'                    .txtInfo(GC_其他关系).Visible = True
                    .picRelation.Visible = True
                Else
'                    .txtInfo(GC_其他关系).Visible = False
                    .picRelation.Visible = False
                    .txtInfo(GC_其他关系).Text = ""
                End If
            Case BCC_入院途径
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                blnLocked = Not (strTmp Like "*转入*" And Not strTmp Like "*非转入*")
                Call SetCtrlLocked(.txtInfo(GC_入院转入), blnLocked, True)
                Call SetCtrlLocked(.cmdInfo(GC_入院转入), blnLocked, True)
            Case BCC_输液反应
                On Error Resume Next
                Set objTextBox = .txtInfo(GC_引发药物) '仅四川版有
                strTmp = objTextBox.Text
                If Err.Number = 0 Then
                    blnLocked = zlStr.NeedName(.cboBaseInfo(intIndex).Text) <> "有"
                    Call SetCtrlLocked(objTextBox, blnLocked, True)
                    Call SetCtrlLocked(.txtInfo(GC_临床表现), blnLocked, True)
                Else
                    Err.Clear
                End If
                On Error GoTo 0
            Case BCC_死亡患者尸检
                If .cboBaseInfo(BCC_死亡患者尸检).ListIndex = 1 Then
                    .cboBaseInfo(BCC_临床与尸检).Clear
                    .cboBaseInfo(BCC_临床与尸检).AddItem "0-未做"
                    .cboBaseInfo(BCC_临床与尸检).AddItem "1-符合"
                    .cboBaseInfo(BCC_临床与尸检).AddItem "2-不符合"
                    .cboBaseInfo(BCC_临床与尸检).AddItem "3-不肯定"
                Else
                    .cboBaseInfo(BCC_临床与尸检).Clear
                    .cboBaseInfo(BCC_临床与尸检).AddItem "-"
                End If
                Call SetDiagMatchInfo(BCC_临床与尸检)
            Case BCC_变异原因
                .txtInfo(GC_变异原因).Text = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
        End Select
        Call CheckValueChange(.cboBaseInfo(intIndex))
    End With
End Sub

Public Sub cboBaseInfoDropDown(ByRef intIndex As Integer)
'功能：cboBaseInfo_DropDown事件封装
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
    strTmp = cboTmp.Text
    If (intIndex = BCC_民族 Or intIndex = BCC_职业 Or intIndex = BCC_国籍) And cboTmp.ListCount = 0 Then
        Call SetCboFromRec(BCC_民族, 0)
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx <> -1 Then
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    '使用不到中文关闭中文输入法
    zlCommFun.OpenIme False
    Call ShowInfectInfo(intIndex = BCC_感染与死亡关系, cboTmp)
    If cboTmp.Style = 0 Then
        Call zlControl.TxtSelAll(cboTmp)
    End If
End Sub

Public Sub CboBaseInfoGotFocus(ByRef intIndex As Integer)
'CboBaseInfo_GotFocus事件
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Call ChangeCtl
    With gclsPros.CurrentForm
        Set cboTmp = .cboBaseInfo(intIndex)
        If (intIndex = BCC_民族 Or intIndex = BCC_职业 Or intIndex = BCC_国籍) And cboTmp.ListCount = 0 Then
            Call SetCboFromRec(BCC_民族, 0)
        End If
        strTmp = cboTmp.Text
        If strTmp <> "" Then
            intIdx = Cbo.FindIndex(cboTmp, strTmp)
            If intIdx <> -1 Then
                Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
            End If
        End If
         If intIndex <> BCC_身份证 Then cboTmp.Tag = cboTmp.ListIndex
        '使用不到中文关闭中文输入法
        zlCommFun.OpenIme False
        Call ShowInfectInfo(intIndex = BCC_感染与死亡关系, cboTmp)
    End With
End Sub

Public Sub CboBaseInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboBaseInfo_KeyDown事件
    With gclsPros.CurrentForm
        If intKeyCode = vbKeyDelete Then
            If .cboBaseInfo(intIndex).Style = 0 Then
                .cboBaseInfo(intIndex).ListIndex = -1
                .cboBaseInfo(intIndex).Text = ""
            Else
                If .cboBaseInfo(intIndex).ListIndex <> -1 Then
                    .cboBaseInfo(intIndex).ListIndex = -1
                End If
            End If
        ElseIf intKeyCode = vbKeyEscape And intIndex = BCC_感染与死亡关系 Then
            Call ShowInfectInfo(False)
        End If
    End With
End Sub

Public Sub CboBaseInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'CboBaseInfo_KeyPress事件
    Dim lngIdx As Long, cboTmp As ComboBox
    Dim strInput As String
    Dim strFilter As String
    Dim rsInput As ADODB.Recordset

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)

    If intKeyAscii = vbKeyReturn And (intIndex = BCC_民族 Or intIndex = BCC_职业 Or intIndex = BCC_国籍) And cboTmp.Style = 0 Then
        strInput = Trim(cboTmp.Text)
        If strInput = "" Then zlCommFun.PressKey vbKeyTab: mblnReturn = True: Exit Sub
        '相同的项目则不进行处理
        If cboTmp.ListIndex <> -1 Then
            If zlStr.NeedName(strInput) = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
                Exit Sub
            End If
        End If
        strInput = UCase(strInput)
        'ADO的通配符有*与%,只能做开头匹配或结尾匹配，或者双向匹配，不能在字符串中间匹配
        If zlCommFun.IsCharChinese(strInput) Then
            strFilter = "名称 Like '*" & strInput & "*'"
        Else
            strFilter = "简码 like '*" & strInput & "*' or 编码 like '*" & strInput & "*'"
        End If
        Set rsInput = Rec.FilterNew(GetBaseCode(intIndex), strFilter)
        If rsInput.RecordCount = 0 Then Exit Sub
        '如果下拉列表展开则关闭下拉列表
        If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
            SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
        End If
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsInput) Then
            lngIdx = Cbo.FindIndex(cboTmp, rsInput!ID)
        Else
            lngIdx = Val(cboTmp.Tag)
        End If
        Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
        cboTmp.Tag = cboTmp.ListIndex
    ElseIf intIndex = BCC_身份证 Then
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Else
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                If zlCommFun.ActualLen(cboTmp.Text) >= 18 And intKeyAscii <> vbKeyBack Then
                    intKeyAscii = 0 '最多只能输入18个长度
                Else
                    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                        intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                            intKeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                            cboTmp.Text = "": cboTmp.Tag = ""
                        End If
                        If gclsPros.FuncType = f医生首页 And intKeyAscii <> 0 Then
                            '门诊首页密文现实
                            Select Case zlCommFun.ActualLen(cboTmp.Text)
                                Case 12
                                    cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                                Case 13
                                    cboTmp.Tag = cboTmp.Tag & Chr(intKeyAscii)
                            End Select
                        End If
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If gclsPros.FuncType = f医生首页 And intKeyAscii <> 0 Then
                        cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    ElseIf intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub CboBaseInfoKeyUp(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboBaseInfo_KeyUp事件
    If intKeyCode = vbKeyDelete Then
        If intIndex = BCC_术前与术后 Then
            gclsPros.CurrentForm.cboBaseInfo(intIndex).ListIndex = -1
        End If
    End If
End Sub

Public Sub cboBaseInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'CboBaseInfo_Validate事件
    '由于将民族修改可以输入匹配，需要验证是否选择的民族
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
    If (intIndex = BCC_民族 Or intIndex = BCC_职业 Or intIndex = BCC_国籍) And cboTmp.ListCount = 0 Then
        Call SetCboFromRec(BCC_民族, 0)
    End If
    strTmp = cboTmp.Text
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx <> -1 Then
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    If intIndex <> BCC_身份证 Then cboTmp.Tag = cboTmp.ListIndex

    If intIndex = BCC_民族 Or intIndex = BCC_职业 Or intIndex = BCC_国籍 Then
        If cboTmp.ListIndex <> -1 Then Exit Sub '已选中
        If cboTmp.Text = "" Then
            MsgBox "请输入" & decode(intIndex, BCC_民族, "民族", BCC_职业, "职业", BCC_国籍, "国籍") & "。", vbInformation, gstrSysName
            blnCancel = True: Exit Sub '无输入
        End If
        strInput = UCase(zlStr.NeedName(cboTmp.Text))
        'ADO的通配符有*与%,只能做开头匹配或结尾匹配，或者双向匹配，不能在字符串中间匹配
        If zlCommFun.IsCharChinese(strInput) Then
            strFilter = "名称 Like '*" & strInput & "*'"
        Else
            strFilter = "简码 like '*" & strInput & "*' or 编码 like '*" & strInput & "*'"
        End If

        blnCancel = True: cboTmp.Text = ""
        Set rsInput = Rec.FilterNew(GetBaseCode(intIndex), strFilter, "ID,编码,简码,名称,缺省")
        If rsInput.RecordCount <> 0 Then
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                If intIdx <> -1 Then
                    cboTmp.ListIndex = intIdx: blnCancel = False
                End If
            End If
        End If
        If blnCancel Then
            MsgBox "请输入" & decode(intIndex, BCC_民族, "民族", BCC_职业, "职业", BCC_国籍, "国籍") & "。", vbInformation, gstrSysName
        End If
    End If
End Sub

'monInfo事件封装
Public Sub monInfoDateClick(ByVal datDateClicked As Date)
'功能：monInfo_DateClick
    Dim strDate As String, strFMT As String
    Dim objMSK As MaskEdBox
    Dim datCurrent As Date

    Set objMSK = gclsPros.CurrentForm.mskDateInfo(gclsPros.DateIndex)
    '获取时分秒数据
    If objMSK.MaxLength >= Len("####-##-## ##:##") Then
        'yyyy-MM-dd HH:mm:ss 格式时间
        If objMSK.MaxLength > Len("####-##-## ##:##") Then
            strFMT = "HH:mm:ss"
        Else
            'yyyy-MM-dd HH:mm 格式时间
            strFMT = "HH:mm"
        End If
        '原时间是时间类型，这取该时间的时分秒数据，否则取当前时间的时分秒
        If IsDate(objMSK.Text) Then
            strDate = " " & Format(objMSK.Text, strFMT)
        Else
            strDate = " " & Format(zlDatabase.Currentdate, strFMT)
        End If
    End If
    '获取时间
    strDate = Format(datDateClicked, "yyyy-MM-dd") & strDate
    objMSK.Text = strDate
    Select Case gclsPros.DateIndex
        Case DC_确诊日期
            If Not CheckDateRange(strDate, True) Then
                MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                Exit Sub
            End If
        Case DC_入院时间, DC_出院时间
            If gclsPros.InTime = "" And gclsPros.DateIndex = DC_出院时间 Then
                If IsDate(gclsPros.CurrentForm.mskDateInfo(DC_入院时间).Text) Then
                    gclsPros.InTime = gclsPros.CurrentForm.mskDateInfo(DC_入院时间).Text
                End If
            ElseIf gclsPros.DateIndex = DC_入院时间 And gclsPros.OutTime = "" Then
                If IsDate(gclsPros.CurrentForm.mskDateInfo(DC_出院时间).Text) Then
                    gclsPros.InTime = gclsPros.CurrentForm.mskDateInfo(DC_出院时间).Text
                End If
            End If
            If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                If gclsPros.DateIndex = DC_出院时间 And gclsPros.InTime <> "" Then
                    If CDate(gclsPros.InTime) > CDate(objMSK.Text) Then
                        MsgBox "你输入的出院时间小于入院时间，请重新输入。", vbInformation, gstrSysName
                        Exit Sub
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        MsgBox "你输入的出院时间大于当前时间，请重新输入。", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        gclsPros.OutTime = objMSK.Text
                    End If
                ElseIf gclsPros.DateIndex = DC_入院时间 And gclsPros.OutTime <> "" Then
                    If CDate(gclsPros.OutTime) < CDate(objMSK.Text) Then
                        MsgBox "你输入的入院时间大于出院时间，请重新输入。", vbInformation, gstrSysName
                        Exit Sub
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        MsgBox "你输入的入院时间大于当前时间，请重新输入。", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        gclsPros.InTime = objMSK.Text
                    End If
                End If
            End If
    End Select
    gclsPros.CurrentForm.txtDateInfo(objMSK.Index).Text = objMSK.Text
    gclsPros.CurrentForm.monInfo.Visible = False
    zlControl.ControlSetFocus objMSK
End Sub

Public Sub monInfoKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'功能：monInfo_KeyDown
    If intKeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
End Sub

Public Sub monInfoKeyPress(ByRef intKeyAscii As Integer)
'功能：monInfo_KeyPress
    If intKeyAscii = 13 Then
        intKeyAscii = 0
        Call monInfoDateClick(gclsPros.CurrentForm.monInfo.Value)
    End If
End Sub

Public Sub monInfoValidate(ByRef blnCancel As Boolean)
'功能：monInfo_Validate
    gclsPros.CurrentForm.monInfo.Visible = False
End Sub
'lstInfectParts事件，lstAdvEvent事件，lstInfection事件
Public Sub LstGotFocus(ByRef lstInput As ListBox)
'lstInfectParts_GotFocus事件，lstAdvEvent_GotFocus事件，lstInfection_GotFocus事件
    Call ChangeCtl
    lstInput.ListIndex = 0
End Sub

Public Sub lstLostFocus(ByRef lstInput As ListBox)
'lstInfectParts_LostFocus事件，lstAdvEvent_LostFocus事件，lstInfection_LostFocus事件
    lstInput.ListIndex = -1
End Sub

Public Sub LstItemCheck(ByRef lstInput As ListBox, ByRef intItem As Integer)
'lstInfectParts_ItemCheck事件，lstAdvEvent_ItemCheck事件，lstInfection_ItemCheck事件
    Dim cboTmp As ComboBox
    With gclsPros.CurrentForm
        If lstInput.Name = "lstAdvEvent" Then
            If lstInput.List(intItem) = "压疮" Then
                Call SetCtrlLocked(.cboBaseInfo(BCC_压疮发生期间), Not lstInput.Selected(intItem), True)
                Call SetCtrlLocked(.cboBaseInfo(BCC_压疮分期), Not lstInput.Selected(intItem), True)
            ElseIf lstInput.List(intItem) = "医院内跌倒/坠床" Then
                Call SetCtrlLocked(.cboBaseInfo(BCC_跌倒或坠床伤害), Not lstInput.Selected(intItem), True)
                Call SetCtrlLocked(.cboBaseInfo(BCC_跌倒或坠床原因), Not lstInput.Selected(intItem), True)
            End If
        End If
        Call CheckValueChange
    End With
End Sub

Public Sub LstKeyDown(ByRef lstInput As ListBox, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'功能：lvwInfectParts_KeyDown
    If intKeyCode = vbKeyEscape And lstInput.Name = "lstInfectParts" Then
        Call ShowInfectInfo(False)
    End If
End Sub

Public Sub LstKeyPress(ByRef lstInput As ListBox, ByRef intKeyAscii As Integer)
'功能：lvwInfectParts_KeyPress
    With gclsPros.CurrentForm
        If intKeyAscii = vbKeyReturn Then
             intKeyAscii = 0
            If lstInput.ListIndex = lstInput.ListCount - 1 Then
                If lstInput.Name = "lstAdvEvent" Then
                    If Not .cboBaseInfo(BCC_压疮发生期间).Locked Or Not .cboBaseInfo(BCC_跌倒或坠床伤害).Locked Then
                        If Not .cboBaseInfo(BCC_压疮发生期间).Locked Then
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_压疮发生期间)
                        Else
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_跌倒或坠床伤害)
                        End If
                    Else
                       Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                    End If
                ElseIf lstInput.Name = "lstInfection" Then
                    Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                ElseIf lstInput.Name = "lstInfectParts" Then
                    Call ShowInfectInfo(False)
                    zlControl.ControlSetFocus .vsDiagXY
                End If
            Else
                lstInput.ListIndex = lstInput.ListIndex + 1
            End If
        End If
    End With
End Sub

'lvwFee事件
Public Sub lvwFeeItemCheck(ByVal Item As MSComctlLib.ListItem)
'lvwFee_ItemCheck事件
    Call AddOrDelFreeCols(gclsPros.CurrentForm.vsFees, Item.Text, Item.SubItems(1), Item.Checked)
End Sub

'cmdFeeEdit事件
Public Sub cmdFeeEditClick()
'cmdFeeEdit_Click事件
    Dim intSeq As Integer
    Dim i As Integer, LngRow As Long, LngCol As Long
    Dim lstTmp As ListItem

    '1.如果列表不可见
    If Not gclsPros.CurrentForm.lvwFee.Visible Then
    ''1.检查此病人是否在外部文件中存在
        gclsPros.FeesOut.Filter = "住院号 = " & IIf(gclsPros.InNo = "", 0, gclsPros.InNo)

        If gclsPros.FeesOut.State = adStateClosed Then Exit Sub
        If gclsPros.FeesOut.RecordCount = 0 Then
            MsgBox "住院号为" & gclsPros.InNo & "的病人其费用数据在外部文件中没找到。", vbInformation, gstrSysName
            Exit Sub
        Else
        ''2.显示列表
            gclsPros.CurrentForm.lvwFee.ListItems.Clear
            With gclsPros.FeesOut
                Do While Not .EOF
                    Set lstTmp = gclsPros.CurrentForm.lvwFee.ListItems.Add(, "K" & intSeq, IIf(IsNull(!费用名), "", !费用名))
                    lstTmp.SubItems(1) = Format(!金额, gclsPros.FreeFormat)
                    lstTmp.SubItems(2) = !住院次数
                    'lstTmp.Checked = !住院次数 = mlng主页ID
                    For i = 3 To gclsPros.CurrentForm.vsFees.Rows * 3
                        LngRow = i \ 3: LngCol = (i Mod 3) * 2
                        If GetTextByDot(gclsPros.CurrentForm.vsFees.TextMatrix(LngRow, LngCol)) = lstTmp.Text And _
                                gclsPros.CurrentForm.vsFees.TextMatrix(LngRow, LngCol + 1) = lstTmp.SubItems(1) And gclsPros.主页ID = lstTmp.SubItems(2) Then
                            lstTmp.Checked = True
                            Exit For
                        End If
                    Next
                    intSeq = intSeq + 1
                    .MoveNext
                Loop
            End With
            gclsPros.CurrentForm.lvwFee.Visible = True
            gclsPros.CurrentForm.lvwFee.Top = gclsPros.CurrentForm.cmdFeeEdit.Top + gclsPros.CurrentForm.cmdFeeEdit.Height
            gclsPros.CurrentForm.lvwFee.Left = gclsPros.CurrentForm.cmdFeeEdit.Left
        End If
    Else
    '2.如果可见,隐藏列表
        gclsPros.CurrentForm.lvwFee.Visible = False
    End If
End Sub

Public Sub ModifyPatiInfo()
'功能：修改病人基本信息
    On Error GoTo errH
    '初始化病人信息接口
    If gobjPatient Is Nothing Then
        On Error Resume Next
        Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo errH
        Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
    End If
    If gobjPatient Is Nothing Then
        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
    '修改病人基本信息后,刷新界面数据
    If gobjPatient.ModiPatiBaseInfo(gclsPros.CurrentForm, "住院首页", gclsPros.病人ID, gclsPros.主页ID, gclsPros.PatiType, False) Then
        Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.病人ID, gclsPros.主页ID) '病案主页以及病人信息
        Call SetCtrlValues("姓名", gclsPros.PatiInfo!姓名 & "")
        Call SetCtrlValues("性别", gclsPros.PatiInfo!性别 & "")
        Call SetCtrlValues("年龄", gclsPros.PatiInfo!年龄 & "")
        Call SetCtrlValues("出生日期", gclsPros.PatiInfo!出生日期 & "")
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub cmdModifyDownGotFocus()
'cmdPriviewDown_GotFocus事件
    Call ShowInfectInfo(False)
End Sub

'cmdMakeLog事件
Public Sub cmdMakeLogClick()
'功能：cmdMakeLog_Click
    Dim strLog As String, i As Long
    Dim vsTmp As VSFlexGrid

    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, DI_诊断描述) & IIf(.TextMatrix(i, DI_是否疑诊) <> "", "(？)", "")
            End If
        Next
    End With

    Set vsTmp = gclsPros.CurrentForm.vsDiagZY
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, DI_诊断描述) & IIf(.TextMatrix(i, DI_是否疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    With gclsPros.CurrentForm.txtInfo(GC_摘要)
        If strLog <> "" Then
            If .SelStart = 0 And .SelLength = Len(.Text) Then
                .SelStart = Len(.Text)
            End If
            i = .SelStart
            .SelText = Mid(strLog, 2)
            .SelStart = i
            .SelLength = Len(Mid(strLog, 2))
        End If
        .SetFocus
    End With
End Sub

'TxtInfo事件
Public Sub CmdInfoClick(ByRef intIndex As Integer)
'功能：cmdInfo_KeyPress
    Dim strSql As String, strCaption As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnMultiSel As Boolean
    Dim strMsg As String, strResult As String
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte

    Select Case intIndex
        Case GC_转科1, GC_转科2, GC_转科3, GC_出院病房, GC_出院科室, GC_入院科室, GC_出院病房
            '选择转科科室
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            grsDeptInfo.Filter = "": Set rsTmp = Rec.Distinct(Rec.FilterNew(grsDeptInfo, "工作性质='临床' OR 工作性质='手术'", "ID,编码,名称,简码"))
            strCaption = "临床科室": strMsg = "部门管理"
        Case GC_重症监护室名称
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            grsDeptInfo.Filter = "": Set rsTmp = Rec.FilterNew(grsDeptInfo, "工作性质='ICU'", "ID,编码,名称,简码")
            strCaption = "ICU重症监护室": strMsg = "部门管理"
        Case GC_病原学诊断

        Case GC_死亡原因
            Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("住院死亡原因"), , "ID,编码,名称,简码")
            strCaption = "死亡原因": strMsg = "字典管理工具"
        Case GC_抢救病因
            Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("抢救病因分类"), , "ID,编码,名称,简码")
            strCaption = "抢救原因": strMsg = "字典管理工具"
        Case GC_医学警示
            '选择医学警示
            strSql = "Select Rownum ID,编码,名称,简码 From 医学警示 Order by 编码"
            strCaption = "医学警示": strMsg = "字典管理工具": blnMultiSel = True
        Case GC_转入医疗机构
            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码,末级 From 出院转入 Order By 编码"
            strCaption = "出院转入": strMsg = "字典管理工具": blnMultiSel = False: bytStyle = 2
        Case GC_入院转入
            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码,末级 From 医疗机构 Order By 编码"
            strCaption = "医疗机构": strMsg = "字典管理工具": blnMultiSel = False: bytStyle = 2
    End Select
    '数据处理
    On Error GoTo errH
    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    If intIndex <> GC_病原学诊断 Then
        If strSql <> "" Then
            vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
            If blnMultiSel Then
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(gclsPros.CurrentForm, strSql, 0, strCaption, True, "", "", True, True, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, True, True)
            Else
                Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, bytStyle, strCaption, , , , , True, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)
            End If
        Else
            blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
        End If
    Else
        'D-ICD-10疾病编码
        Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "D", gclsPros.出院科室ID, gclsPros.CurrentForm.cboBaseInfo(BCC_性别).Text, False)
    End If

    If rsTmp Is Nothing Then
        If intIndex = GC_病原学诊断 Then Exit Sub
        If Not blnCancel Then
            MsgBox "没有设置""" & strCaption & """数据，请先到" & strMsg & "中设置。", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        If blnMultiSel Then '多选以逗号分割
            While Not rsTmp.EOF
                strResult = strResult & "," & rsTmp!名称
                rsTmp.MoveNext
            Wend
            objTXTBox.Text = Mid(strResult, 2)
        Else
            If intIndex = GC_病原学诊断 Then
                objTXTBox.Text = IIf(Not IsNull(rsTmp!编码), "(" & rsTmp!编码 & ")", "") & NVL(rsTmp!名称)
                objTXTBox.Tag = objTXTBox.Text
                gclsPros.CurrentForm.cmdInfo(intIndex).Tag = rsTmp!项目ID
            Else
                objTXTBox.Text = rsTmp!名称
                If gclsPros.FuncType = f病案首页 Then
                    If intIndex = GC_出院科室 Then
                        '53638:刘鹏飞,2013-05-10,档案号编码规则
                        If gclsPros.UseFileRules = True And gclsPros.出院科室ID <> Val(rsTmp!ID & "") And Val(gclsPros.InNo) <> 0 Then
                            If IsPageNosCodeRule(CT_档案号) = True Then
                                gclsPros.CurrentForm.txtInfo(GC_档案号).Text = NVL(GetNextNo(5, , rsTmp!编码 & ""))
                            End If
                        End If
                        gclsPros.出院科室ID = Val(rsTmp!ID & "")
                    ElseIf intIndex = GC_入院科室 Then
                        gclsPros.入院科室ID = Val(rsTmp!ID & "")
                    End If
                    Call SetFaceInit(True)
                    Call SetPageVisible
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
        End If
        zlControl.ControlSetFocus objTXTBox
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
    If intIndex = GC_重症监护室名称 Then
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_人工气道脱出), objTXTBox.Text = "", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_重返重症医学科), objTXTBox.Text = "", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_重返间隔时间), objTXTBox.Text = "", True)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'cmdLastDiag事件
Public Sub cmdLastDiagClick()
'cmdLastDiag_Click事件
    Dim rsTmp As Recordset
    Dim strSql As String

    If gclsPros.主页ID = 1 Then Exit Sub
    On Error GoTo errH
    strSql = "Select 诊断类型, 诊断描述 || '(' || 编码 || ')' As 诊断内容" & vbNewLine & _
                "From 病人诊断记录 a, 疾病编码目录 b" & vbNewLine & _
                "Where a.疾病id = b.Id(+) And 病人id = [1] And 主页id = [2] And 诊断类型 = 3 And 诊断次序 = 1 And" & vbNewLine & _
                "      记录来源 = (Select Max(Nvl(记录来源, 0)) From 病人诊断记录 Where 病人id = [1] And 主页id = [2] And 记录来源 <= 4)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "", gclsPros.病人ID, gclsPros.主页ID - 1)
    gclsPros.CurrentForm.lblDiagInfo.BorderStyle = 1
    gclsPros.CurrentForm.lblDiagInfo.Visible = True
    If Not rsTmp.EOF Then
        gclsPros.CurrentForm.lblDiagInfo.Caption = NVL(rsTmp!诊断内容)
    Else
        gclsPros.CurrentForm.lblDiagInfo.Caption = "未找到上次住院的主要诊断信息"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub TxtInfoChange(ByRef intIndex As Integer)
'功能：txtInfo_Change
    Dim objTXTBox As TextBox
    Dim lngPos As Long, lnglen As Long

    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    Select Case intIndex
        Case GC_转科1
            If objTXTBox.Text = "" Then
                gclsPros.CurrentForm.txtInfo(GC_转科2).Text = ""
                gclsPros.CurrentForm.txtInfo(GC_转科3).Text = ""
            End If
        Case GC_转科2
            If objTXTBox.Text = "" Then
                gclsPros.CurrentForm.txtInfo(GC_转科3).Text = ""
            End If
        Case GC_监护人身份证号
            If gclsPros.IsReturn Then Exit Sub
            gclsPros.IsReturn = True
            If Not zlStr.CheckCharScope(objTXTBox.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                objTXTBox.Text = ""
            Else
                If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                    If zlCommFun.ActualLen(objTXTBox.Text) > 18 Then
                        objTXTBox.Text = Mid(objTXTBox.Text, 1, 18)
                    End If
                End If
            End If
            '门诊首页密文现实
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                If objTXTBox.Tag <> "不触发Change事件" Then
                    lngPos = InStr(objTXTBox.Text, "*")
                    lnglen = Len(Mid(objTXTBox.Text, 13, 2))
                    Select Case lngPos
                        Case 0
                            objTXTBox.Tag = objTXTBox.Text
                        Case Else 'Is <= 12
                            objTXTBox.Tag = Mid(objTXTBox.Text, 1, lngPos - 1)
                            objTXTBox.Text = objTXTBox.Tag
                            objTXTBox.SelStart = Len(objTXTBox.Text)
                    End Select
                End If '
            Else
                objTXTBox.Tag = objTXTBox.Text
            End If
            gclsPros.IsReturn = False
    End Select
    Call CheckValueChange(objTXTBox)
End Sub

Public Sub TxtInfoGotFocus(ByRef intIndex As Integer)
'功能：txtInfo_GotFocus
    Call ChangeCtl
    If Not (intIndex = GC_摘要 And gclsPros.CurrentForm.txtInfo(intIndex).SelLength <> 0) Then
        Call TxtGotFocus(gclsPros.CurrentForm.txtInfo(intIndex), True, True)
    End If
End Sub

Public Sub TxtInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'功能：txtInfo_KeyDown
    If intKeyCode = vbKeyDelete Then
        If intIndex = GC_医学警示 Then
            gclsPros.CurrentForm.txtInfo(intIndex).Text = ""
        End If
    End If
End Sub

Public Sub TxtInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'功能：txtInfo_KeyPress
    Dim objTXTBox As TextBox
    Dim strSql As String, strFilter As String, strInput As String
    Dim strCaption As String, strSeek As String, strNote As String
    Dim bln末级 As Boolean, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    On Error GoTo errH

    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    If intIndex = GC_监护人身份证号 Then
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Else
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_国籍).Text)) = "中国" Then
                If zlCommFun.ActualLen(objTXTBox.Text) >= 18 And intKeyAscii <> vbKeyBack Then
                    intKeyAscii = 0 '最多只能输入18个长度
                Else
                    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                        intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                            intKeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(objTXTBox.Text) Then
                            objTXTBox.Text = "": objTXTBox.Tag = ""
                        End If
                        If gclsPros.FuncType = f医生首页 And intKeyAscii <> 0 Then
                            '门诊首页密文现实
                            Select Case zlCommFun.ActualLen(objTXTBox.Text)
                                Case 12
                                    objTXTBox.Tag = objTXTBox.Text & Chr(intKeyAscii)
                                Case 13
                                    objTXTBox.Tag = objTXTBox.Tag & Chr(intKeyAscii)
                            End Select
                        End If
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(objTXTBox.Text) Then
                        objTXTBox.Text = "": objTXTBox.Tag = ""
                    End If
                    If gclsPros.FuncType = f医生首页 And intKeyAscii <> 0 Then
                        objTXTBox.Tag = objTXTBox.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    Else
        If intKeyAscii = vbKeyReturn Then
            If objTXTBox.Text <> "" Then
                strInput = UCase(objTXTBox.Text)
                Select Case intIndex
                    Case GC_转科1, GC_转科2, GC_转科3, GC_出院科室, GC_入院科室, GC_重症监护室名称
                        If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                        If intIndex = GC_转科1 Or intIndex = GC_转科2 Or intIndex = GC_转科3 Then
                            grsDeptInfo.Filter = "": Set rsTmp = grsDeptInfo: strCaption = "转科科室"
                        Else
                            grsDeptInfo.Filter = "工作性质 = '" & IIf(intIndex = GC_重症监护室名称, "ICU", "临床") & "'": Set rsTmp = grsDeptInfo: strCaption = "临床科室"
                        End If
                        strFilter = "编码 Like '" & strInput & "*' Or 简码 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or 名称 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
                        Set rsTmp = Rec.Distinct(Rec.FilterNew(rsTmp, strFilter, "Id,编码,名称,简码,位置"), "Id,编码,名称,简码,位置")
                    Case GC_抢救病因
                        strFilter = "编码 Like '" & strInput & "*' Or 简码 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or 名称 Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
                        Set rsTmp = Rec.FilterNew(GetBaseCode("抢救病因分类"), strFilter, "ID,编码,名称,简码")
                        strCaption = "抢救原因"
                    Case GC_转入医疗机构
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where 名称 Like [1]"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where zlWbCode(名称) Like [1]"
                            Else
                                strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where 简码 Like [1]"
                            End If
                        End If
                        strCaption = "出院转入"
                    Case GC_入院转入
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 名称 Like [1]"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where zlWbCode(名称) Like [1]"
                            Else
                                strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 简码 Like [1]"
                            End If
                        End If
                        strCaption = " 医疗机构"
                End Select
                If strSql <> "" Or strFilter <> "" Then
                    If strSql <> "" Then
                        vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
                        If intIndex = GC_入院转入 Or intIndex = GC_转入医疗机构 Then
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, bln末级, strSeek, strNote, False, _
                                False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                                gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                        Else
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, bln末级, strSeek, strNote, False, _
                                False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                                UCase(objTXTBox.Text) & "%", gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                        End If
                    Else
                        Call zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
                    End If
                    '可以任意输入,不一定要匹配
                    If Not rsTmp Is Nothing Then
                        objTXTBox.Text = rsTmp!名称
                        If gclsPros.FuncType = f病案首页 Then
                            If intIndex = GC_出院科室 Then
                                '53638:刘鹏飞,2013-05-10,档案号编码规则
                                If gclsPros.UseFileRules = True And gclsPros.出院科室ID <> Val(rsTmp!ID & "") And Val(gclsPros.InNo) <> 0 Then
                                    If IsPageNosCodeRule(CT_档案号) = True Then
                                        gclsPros.CurrentForm.txtInfo(GC_档案号).Text = NVL(GetNextNo(5, , rsTmp!编码 & ""))
                                    End If
                                End If
                                gclsPros.出院科室ID = Val(rsTmp!ID & "")
                            ElseIf intIndex = GC_入院科室 Then
                                gclsPros.入院科室ID = Val(rsTmp!ID & "")
                            End If
                            Call SetFaceInit(True)
                            Call SetPageVisible
                            Call SetFaceEditable(gclsPros.IsSigned)
                        End If
                    Else
                        objTXTBox.Tag = ""
                        If gclsPros.GetMedical Then
                            MsgBox "在字典表中未找到该数据,请重新录入！", vbInformation, gstrSysName
                            objTXTBox.Text = ""
                            objTXTBox.SetFocus
                        End If
                    End If
                End If
            End If
            intKeyAscii = 0
            '医保号病案独有，不用加功能条件
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        ElseIf Not (intKeyAscii >= 0 And intKeyAscii < vbKeySpace) Then
            '选择快捷键
            If intKeyAscii = Asc("*") Then
                '注意界面上要求CMD和对应TXT的Index相同
                On Error Resume Next
                strSql = ""
                strSql = gclsPros.CurrentForm.cmdInfo(intIndex).Name
                Err.Clear: On Error GoTo errH
                If strSql <> "" Then
                    intKeyAscii = 0
                    Call CmdInfoClick(intIndex)
                    Exit Sub
                End If
            End If
    
            '限制输入长度
            If objTXTBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objTXTBox.Text) > objTXTBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub TxtInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：txtInfo_MouseDown
    Call TxtMouseDown(gclsPros.CurrentForm.txtInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub TxtInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：txtInfo_MouseUp
    Call TxtMouseUp(gclsPros.CurrentForm.txtInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub TxtInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'功能：txtInfo_Validate
    Dim objTXTBox As TextBox, objCmdBtn As CommandButton
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim strInput As String, str性别 As String
    Dim blnCancelSel As Boolean
    Dim strMsg As String
    
    On Error GoTo errH
    
    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    On Error Resume Next
    Set objCmdBtn = gclsPros.CurrentForm.cmdInfo(intIndex) '可能没有对应的按钮
    Err.Clear: On Error GoTo errH
    
    Select Case intIndex
        Case GC_病原学诊断
            If objTXTBox.Text = "" Then
                objTXTBox.Tag = ""
                objCmdBtn.Tag = ""
            ElseIf objTXTBox.Text = objTXTBox.Tag Then
                'Nothing
            Else
                strInput = UCase(objTXTBox.Text)
                If gclsPros.CurrentForm.cboBaseInfo(BCC_性别).Text Like "*男*" Then
                    str性别 = "男"
                ElseIf gclsPros.CurrentForm.cboBaseInfo(BCC_性别).Text Like "*女*" Then
                    str性别 = "女"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSql = "编码 Like [1] Or 名称 Like [2] Or " & IIf(gclsPros.BriefCode = 0, "简码", "五笔码") & " Like [2]"
                End If
                strSql = _
                    " Select ID,ID as 项目ID,编码,附码,名称," & IIf(gclsPros.BriefCode = 0, "简码", "五笔码 as 简码") & ",说明" & _
                    " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSql & ")" & _
                    IIf(str性别 <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                    " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by 编码"

                If gclsPros.DiagSourceZY = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", "'D'", str性别)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                Else
                    vPoint = GetCoordPos(objTXTBox.hwnd, 0, 0)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "病原学诊断", False, "", "", False, False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancelSel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", "'D'", str性别, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnCancelSel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (gclsPros.DiagSourceZY = 2 Or gclsPros.DiagSourceZY = 3 And gclsPros.InsureType <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                        End If
                    End If
                End If
                
                If Not blnCancel Then
                    If rsTmp Is Nothing Then
                        objCmdBtn.Tag = ""
                    Else
                        objTXTBox.Text = IIf(Not IsNull(rsTmp!编码), "(" & rsTmp!编码 & ")", "") & NVL(rsTmp!名称)
                        objTXTBox.Tag = objTXTBox.Text
                        objCmdBtn.Tag = rsTmp!项目ID
                    End If
                End If
            End If
        Case GC_重症监护室名称
            strInput = Trim(objTXTBox.Text)
            If strInput <> "" Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                grsDeptInfo.Filter = "": Set rsTmp = Rec.FilterNew(grsDeptInfo, "工作性质='ICU'", "ID,编码,名称,简码")
                rsTmp.Filter = "名称='" & strInput & "'"
                If rsTmp.EOF Then
                    rsTmp.Filter = "编码 Like '" & strInput & "*' OR 编码 Like '" & strInput & "*' OR 名称 Like '" & IIf(gclsPros.LikeString <> "", "*", "") & strInput & "*' "
                    blnCancelSel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                        End If
                    Else
                        objTXTBox.Text = rsTmp!名称
                    End If
                End If
            End If
            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_人工气道脱出), objTXTBox.Text = "" Or blnCancel, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_重返重症医学科), objTXTBox.Text = "" Or blnCancel, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_重返间隔时间), objTXTBox.Text = "" Or blnCancel, True)
    Case GC_转入医疗机构, GC_入院转入
        strInput = UCase(objTXTBox.Text)
        If strInput = "" Then
            objTXTBox.Tag = ""
        Else
            
            If zlCommFun.IsCharChinese(strInput) Then
                If intIndex = GC_转入医疗机构 Then
                    strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where 名称 Like [1]"
                ElseIf intIndex = GC_入院转入 Then
                    strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 名称 Like [1]"
                End If
            Else
                If gclsPros.BriefCode = 1 Then
                    If intIndex = GC_转入医疗机构 Then
                        strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where zlWbCode(名称) Like [1]"
                    ElseIf intIndex = GC_入院转入 Then
                        strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where zlWbCode(名称) Like [1]"
                    End If
                Else
                    If intIndex = GC_转入医疗机构 Then
                        strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 出院转入 Where 简码 Like [1]"
                    ElseIf intIndex = GC_入院转入 Then
                        strSql = "Select 编码 As ID, 编码, 上级 As 上级id, 名称, 简码, 末级 From 医疗机构 Where 简码 Like [1]"
                    End If
                End If
            End If
            If strSql <> "" Then
                vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "医疗机构", True, False, True, False, _
                    False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                      gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                If Not rsTmp Is Nothing Then
                    objTXTBox.Text = rsTmp!名称
                Else
                    objTXTBox.Tag = ""
                    If gclsPros.GetMedical Then
                        MsgBox "在字典表中未找到该数据,请重新录入", vbInformation, gstrSysName
                        objTXTBox.Text = ""
                        objTXTBox.SetFocus
                    End If
                End If
            End If
        End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'AdressInfo事件
Public Sub CmdAdressInfoClick(ByRef intIndex As Integer)
'功能：cmdAdressInfo_Click
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte, strCaption As String, strMsg As String, blnRoot As Boolean, blnNonWin As Boolean

    On Error GoTo errH
    Select Case intIndex
        Case ADRC_单位地址
            '选择单位信息
            strSql = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strCaption = "合约单位": strMsg = "合约单位管理": bytStyle = 2: blnRoot = True: blnNonWin = True
        Case ADRC_出生地点, ADRC_现住址, ADRC_联系人地址, ADRC_户口地址
            '选择地区数据
            strSql = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            strCaption = "区域": strMsg = "字典管理工具": bytStyle = 0: blnRoot = False: blnNonWin = True
        Case ADRC_病人区域, ADRC_籍贯
            '选择区域数据
            strSql = "Select 1  From 区域 Where Nvl(级数,0)<>0 And RowNum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption)
            If rsTmp.RecordCount > 0 Then bytStyle = 2
            If bytStyle = 2 Then
                strSql = _
                        "Select Id, 上级id, Id 编码, 名称, 简码, 末级" & vbNewLine & _
                        "From (Select Rpad(编码, 15, '0') As Id, Rpad(Substr(编码, 1, Decode(Nvl(级数, 0), 0, 0, 1, 2, 4)), 15, '0') As 上级id, 名称, 简码," & vbNewLine & _
                        "              Decode(Nvl(级数, 0), 2, 1, 3, 1, 0) As 末级" & vbNewLine & _
                        "       From 区域" & vbNewLine & _
                        "       Where Nvl(级数, 0) < 3" & vbNewLine & _
                        "       Order By 编码)" & vbNewLine & _
                        "Start With 上级id Is Null" & vbNewLine & _
                        "Connect By Prior Id = 上级id"
            Else
                strSql = "Select Rownum as ID,编码,名称,简码 From 区域 Order by 编码"
            End If
            strCaption = "区域": strMsg = "字典管理工具": blnRoot = False: blnNonWin = IIf(bytStyle = 0, True, False)
    End Select

    '数据处理
    On Error GoTo errH
    '数据处理
    Set objTXTBox = gclsPros.CurrentForm.txtAdressInfo(intIndex)
    vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
    Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, bytStyle, strCaption, , , , , blnRoot, blnNonWin, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有设置""" & IIf(strCaption = "区域", "地区", strCaption) & """数据，请先到" & strMsg & "中设置。", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        If intIndex = ADRC_单位地址 Then
            objTXTBox.Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
            If gclsPros.PatiType = PF_门诊 Then
                If InStr(gclsPros.Privs, "合约病人登记") > 0 Then objTXTBox.Tag = Val(rsTmp!ID)
            Else
                objTXTBox.Tag = Val(rsTmp!ID)
            End If
            If gclsPros.CurrentForm.txtSpecificInfo(SLC_单位电话).Text = "" Then
                gclsPros.CurrentForm.txtSpecificInfo(SLC_单位电话).Text = NVL(rsTmp!电话)
            End If
            objTXTBox.SetFocus
        Else
            objTXTBox.Text = rsTmp!名称
            objTXTBox.SetFocus
            If intIndex = ADRC_出生地点 And gclsPros.FuncType = f病案首页 And gclsPros.DefautADD Then
                Call SetPatiAddress(ADRC_联系人地址, "联系人地址", rsTmp!名称, True)
                Call SetPatiAddress(ADRC_现住址, "家庭地址", rsTmp!名称, True)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_家庭邮编).Text = rsTmp!编码 & ""
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub txtAdressInfoGotFocus(ByRef intIndex As Integer)
'功能：txtAdressInfo_GotFocus
    Call ChangeCtl
    Call TxtGotFocus(gclsPros.CurrentForm.txtAdressInfo(intIndex), True, True)
End Sub

Public Sub txtAdressInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'txtAdressInfo_KeyPress事件
    Dim objBox As TextBox
    Dim strSql As String, strCaption As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI

    Set objBox = gclsPros.CurrentForm.txtAdressInfo(intIndex)

    If intKeyAscii = vbKeyReturn Then
        If objBox.Text <> "" Then
            Select Case intIndex
                Case ADRC_单位地址
                    '选择单位信息
                    strSql = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                        " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                        " And 上级id Is not Null and (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                        " Order by 编码"
                    strCaption = "工作单位"
                Case ADRC_出生地点, ADRC_现住址, ADRC_联系人地址, ADRC_户口地址
                    '输入地区数据
                    strSql = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                        " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                        " Order by 编码"
                    strCaption = "地区"
                Case ADRC_病人区域, ADRC_籍贯
                    '输入区域数据
                    strSql = "Select Rownum as ID,编码,名称,简码 From 区域 " & _
                        " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2]) And Nvl(级数, 0) < 3" & _
                        " Order by 编码"
                    strCaption = IIf(intIndex = ADRC_病人区域, "区域", "籍贯")
            End Select

            If strSql <> "" Then
                vPoint = GetCoordPos(objBox.Container.hwnd, objBox.Left, objBox.Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, objBox.Height, blnCancel, False, False, _
                    UCase(objBox.Text) & "%", gclsPros.LikeString & UCase(objBox.Text) & "%")
                '可以任意输入,不一定要匹配
                If Not rsTmp Is Nothing Then
                    If intIndex = ADRC_单位地址 Then
                        objBox.Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                        If gclsPros.PatiType = PF_门诊 Then
                            If InStr(gclsPros.Privs, "合约病人登记") > 0 Then objBox.Tag = Val(rsTmp!ID)
                        Else
                            objBox.Tag = Val(rsTmp!ID)
                        End If
                        If gclsPros.CurrentForm.txtSpecificInfo(SLC_单位电话).Text = "" Then
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_单位电话).Text = NVL(rsTmp!电话)
                        End If
                    Else
                        objBox.Text = rsTmp!名称
                        If intIndex = ADRC_出生地点 And gclsPros.FuncType = f病案首页 And gclsPros.DefautADD Then
                            Call SetPatiAddress(ADRC_联系人地址, "联系人地址", rsTmp!名称, True)
                            Call SetPatiAddress(ADRC_现住址, "家庭地址", rsTmp!名称, True)
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_家庭邮编).Text = rsTmp!编码 & ""
                        End If
                    End If
                Else
                    objBox.Tag = ""
                End If
                objBox.SetFocus
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
        Else
            intKeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    Else
        If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
            '选择快捷键
            If intKeyAscii = Asc("*") Then
                '注意界面上要求CMD和对应TXT的Index相同
                On Error Resume Next
                strSql = ""
                strSql = gclsPros.CurrentForm.cmdAdressInfo(intIndex).Name
                Err.Clear: On Error GoTo 0
                If strSql <> "" Then
                    intKeyAscii = 0
                    Call CmdAdressInfoClick(intIndex)
                    Exit Sub
                End If
            End If

            '限制输入长度
            If objBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objBox.Text) > objBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
End Sub

Public Sub txtAdressInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'txtAdressInfo_MouseDown事件
    Call TxtMouseDown(gclsPros.CurrentForm.txtAdressInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub txtAdressInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'txtAdressInfo_MouseUp事件
    Call TxtMouseUp(gclsPros.CurrentForm.txtAdressInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

'vsChemoth事件
Public Sub ChemothAfterEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsChemoth_AfterEdit事件
    Dim strInput As String
    With vsChemoth
        If LngCol = CI_化学治疗编码 Then
            If .ComboIndex < 0 Then Exit Sub
           .TextMatrix(LngRow, CI_疾病ID) = .ComboData(.ComboIndex)
        ElseIf LngCol = CI_结束日期 Or LngCol = CI_开始日期 Then
            strInput = zlStr.FullDate(.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
            If Not IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
            Else
                .TextMatrix(LngRow, LngCol) = strInput
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            End If
        End If
    End With
End Sub

Public Sub ChemothAfterRowColChange(ByRef vsChemoth As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsChemoth_AfterRowColChange 事件
'    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
'    Call zlVsGridRowChange(vsChemoth, lngOldRow, lngNewRow, lngOldCol, lngNewCol)
End Sub

Public Sub ChemothGotFocus(ByRef vsChemoth As VSFlexGrid)
'vsChemoth_GotFocus事件
'    Call zlVsGridGotFocus(vsChemoth)
End Sub

Public Sub ChemothKeyDown(ByRef vsChemoth As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsChemoth_KeyDown事件
   Dim LngCol As Long

    With vsChemoth
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If MsgBox("你是否真的要删除该行的化疗信息吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                 Call ChangeVSFHeight(vsChemoth, True)
            End If
            zlControl.ControlSetFocus vsChemoth, True
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, CI_化学治疗编码) = "" And .Col = CI_化疗效果 Then
                zlControl.ControlSetFocus gclsPros.CurrentForm.vsRadioth, True
                Exit Sub
            End If
            
            Select Case .Col
                Case .Cols - 1, CI_化疗效果
                    If Not .Row >= .Rows - 1 Then
                        .Col = 0
                        .Row = .Row + 1
                    Else
                        Call ChemothKeyDownEdit(vsChemoth, .Row, .Col, intKeyCode, intShift)
                    End If
                    .SetFocus
                Case Else
                    zlCommFun.PressKey vbKeyRight
            End Select
        End If
    End With
End Sub

Public Sub ChemothKeyDownEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsChemoth_KeyDownEdit事件
    Dim lngCurRow As Long
    Dim strKEY As String

    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsChemoth
        Call zlVsMoveGridCell(vsChemoth, CI_化学治疗编码, .Cols - 1, True, lngCurRow)
        If lngCurRow > 0 Then
            '表示新增加了一行,需要设置相关的缺省值
            strKEY = .ColData(CI_化学治疗编码)
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngCurRow, CI_化学治疗编码) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngCurRow, CI_化学治疗编码) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngCurRow, CI_疾病ID) = .Cell(flexcpData, lngCurRow, CI_化学治疗编码)
                .TextMatrix(lngCurRow, CI_疗程数) = 1
                .Col = CI_开始日期
            End If
        End If
    End With
End Sub

Public Sub ChemothKeyPress(ByRef vsChemoth As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsChemoth_KeyPress事件
    If vsChemoth.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
    End If
End Sub

Public Sub ChemothKeyPressEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsChemoth_KeyPressEdit 事件
    If intKeyAscii = Asc("'") Then intKeyAscii = 0: Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Exit Sub
    End If
    With gclsPros.CurrentForm.vsChemoth
        Select Case LngCol
            Case CI_化学治疗编码, CI_化疗方案
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m文本式)
            Case CI_开始日期, CI_结束日期
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m文本式)
            Case CI_疗程数, CI_总量
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m数字式)
        End Select
    End With
End Sub

Public Sub ChemothLostFocus(ByRef vsChemoth As VSFlexGrid)
'vsChemoth_LostFocus事件
'    Call zlVsGridLostFocus(vsChemoth)
End Sub

Public Sub ChemothValidateEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsChemoth_ValidateEdit事件
    Dim strInput As String

    With vsChemoth
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        If strInput = "" Then Exit Sub

        Select Case LngCol
            Case CI_化学治疗编码

            Case CI_开始日期, CI_结束日期
                strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
                If IsDate(strInput) Then
                    If Not CheckDateRange(strInput) Then
                        MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                    If IsDate(Trim(.TextMatrix(LngRow, IIf(LngCol = CI_结束日期, CI_开始日期, CI_结束日期)))) Then
                        If LngCol = CI_结束日期 And CDate(strInput) < CDate(Trim(.TextMatrix(LngRow, CI_开始日期))) Or _
                            LngCol = CI_开始日期 And CDate(Trim(.TextMatrix(LngRow, CI_结束日期))) < CDate(strInput) Then
                            MsgBox "结束日期不能小于开始日期,请检查!", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        End If
                    End If
'                    Call zlVsMoveGridCell(vsChemoth, CI_开始日期, .Cols - 1, True, LngCol)
                Else
                    MsgBox IIf(LngCol = CI_结束日期, "结束日期", "开始日期") & "必须为日期型,请检查！", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            Case CI_化疗方案
                blnCancel = Not zlCommFun.StrIsValid(strInput, 50, 0, "化疗方案")
            Case CI_疗程数
                If DblIsValid(strInput, 3, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
            Case CI_总量
                If DblIsValid(strInput, 10, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
        End Select
    End With
End Sub

'vsRadioth事件
Public Sub RadiothAfterEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsRadioth_AfterEdit事件
    Dim strInput As String
    With vsRadioth
        If LngCol = RI_放射治疗编码 Then
            If .ComboIndex < 0 Then Exit Sub
            .TextMatrix(LngRow, RI_疾病ID) = .ComboData(.ComboIndex)
        ElseIf LngCol = RI_结束日期 Or LngCol = RI_开始日期 Then
            strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
            If Not IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
            Else
                .TextMatrix(LngRow, LngCol) = strInput
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            End If
        End If
    End With
End Sub

Public Sub RadiothAfterRowColChange(ByRef vsRadioth As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsRadioth_AfterRowColChange事件
'    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
'    Call zlVsGridRowChange(vsRadioth, lngOldRow, lngNewRow, lngOldCol, lngNewCol)
End Sub

Public Sub RadiothGotFocus(ByRef vsRadioth As VSFlexGrid)
'vsRadioth_GotFocus事件
'    Call zlVsGridGotFocus(vsRadioth)
End Sub

Public Sub RadiothKeyDown(ByRef vsRadioth As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsRadioth_KeyDown事件
    Dim LngCol As Long

    With vsRadioth
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If MsgBox("你是否真的要删除该行的放疗信息吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsRadioth, True)
            End If
            zlControl.ControlSetFocus vsRadioth, True
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, RI_放射治疗编码) = "" And .Col = RI_放疗效果 Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                Exit Sub
            End If
            
            Select Case .Col
                Case .Cols - 1, RI_放疗效果
                    If Not .Row >= .Rows - 1 Then
                        .Col = RI_放射治疗编码
                        .Row = .Row + 1
                    Else
                        Call RadiothKeyDownEdit(vsRadioth, .Row, .Col, intKeyCode, intShift)
                    End If
                    .SetFocus
                Case Else
                    zlCommFun.PressKey vbKeyRight
            End Select
        End If
    End With
End Sub

Public Sub RadiothKeyDownEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsRadioth_KeyDownEdit事件
    Dim lngCurRow As Long, strKEY As String

    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsRadioth
        Call zlVsMoveGridCell(vsRadioth, RI_放射治疗编码, .Cols - 1, True, lngCurRow)
        If lngCurRow > 0 Then
'            表示新增加了一行 , 需要设置相关的缺省值
            strKEY = .ColData(RI_放射治疗编码)
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngCurRow, RI_放射治疗编码) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngCurRow, RI_放射治疗编码) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngCurRow, RI_疾病ID) = .Cell(flexcpData, lngCurRow, RI_放射治疗编码)
                .Col = RI_开始日期
            End If
        End If
    End With
End Sub

Public Sub RadiothKeyPress(ByRef vsRadioth As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsRadioth_KeyPress事件
    If vsRadioth.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
    End If
End Sub

Public Sub RadiothKeyPressEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsRadioth_KeyPressEdit事件
    Dim strInput As String
    With vsRadioth
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            If .Row = .Rows - 1 And .TextMatrix(.Row, RI_放射治疗编码) = "" Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
            Exit Sub
        End If
        Select Case LngCol
            Case RI_放射治疗编码, RI_设野部位
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m文本式)
            Case RI_开始日期, RI_结束日期
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m文本式)
            Case RI_放射剂量, RI_累计量
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m数字式)
            Case RI_放疗效果
        End Select
    End With
End Sub

Public Sub RadiothLostFocus(ByRef vsRadioth As VSFlexGrid)
'vsRadioth_LostFocus事件
'    Call zlVsGridLostFocus(vsRadioth)
End Sub

Public Sub RadiothValidateEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsRadioth_ValidateEdit事件
    Dim strInput As String

    With vsRadioth
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        If strInput = "" Then Exit Sub
        Select Case LngCol
            Case RI_放射治疗编码
            Case RI_开始日期, RI_结束日期
                strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
                If IsDate(strInput) Then
                    If Not CheckDateRange(strInput) Then
                        MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                    If IsDate(Trim(.TextMatrix(LngRow, IIf(LngCol = RI_结束日期, RI_开始日期, RI_结束日期)))) Then
                        If LngCol = RI_结束日期 And CDate(strInput) < CDate(Trim(.TextMatrix(LngRow, RI_开始日期))) Or _
                            LngCol = RI_开始日期 And CDate(Trim(.TextMatrix(LngRow, RI_结束日期))) < CDate(strInput) Then
                            MsgBox "结束日期不能小于开始日期,请检查!", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        End If
                    End If
                    Call zlVsMoveGridCell(vsRadioth, RI_开始日期, .Cols - 1, True, LngCol)
                Else
                    MsgBox IIf(LngCol = RI_结束日期, "结束日期", "开始日期") & "必须为日期型,请检查！", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            Case RI_设野部位
                blnCancel = Not zlCommFun.StrIsValid(strInput, 50, 0, "设野部位")
            Case RI_放射剂量, RI_累计量
                If DblIsValid(strInput, 10, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
            Case RI_放疗效果
        End Select
    End With
End Sub

'vsFlxAddICU事件
Public Sub FlxAddICUAfterEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsFlxAddICU_AfterEdit事件
    Dim strList As String, i As Long
    Dim strInput As String

    With vsFlxAddICU
        If LngCol = UI_监护室名称 Then
            If gclsPros.MedPageSandard = ST_卫生部标准 Then
                .ColComboList(0) = "..."
            ElseIf gclsPros.MedPageSandard = ST_四川省标准 Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, UI_序号) = i
                Next
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, UI_监护室名称) <> "" Then
                        strList = strList & "|" & .TextMatrix(i, UI_序号) & "-" & .TextMatrix(i, UI_监护室名称)
                    End If
                Next
                strList = Mid(strList, 2)
                gclsPros.CurrentForm.vsICUInstruments.ColComboList(TI_ICU类型) = strList
                gclsPros.CurrentForm.vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
            End If
        ElseIf LngCol = UI_进入时间 Or LngCol = UI_退出时间 Then
            strInput = zlStr.FullDate(.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime)
            If IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = strInput
            End If
        End If
    End With
End Sub

Public Sub FlxAddICUCellButtonClick(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsFlxAddICU_CellButtonClick事件
    Dim strSql As String, rsTmp As Recordset, vPoint As POINTAPI, blnCancel As Boolean

    With vsFlxAddICU
        Select Case LngCol
            Case UI_监护室名称
                strSql = " Select Distinct A.ID,A.编码,A.名称" & _
                        " From 部门表 A,部门性质说明 B" & _
                        " Where B.部门ID=A.ID And B.工作性质='ICU'" & _
                        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "重症监护室", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)

                If rsTmp Is Nothing Then
                    If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        MsgBox "没有设置ICU重症监护室。", vbInformation, gstrSysName
                    End If
                Else
                    .TextMatrix(LngRow, LngCol) = rsTmp!名称 & ""
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                End If
            Case Else
        End Select
    End With
End Sub

Public Sub FlxAddICUEnterCell(ByRef vsFlxAddICU As VSFlexGrid)
'vsFlxAddICU_EnterCell事件

    Dim datTemp As Date

    With vsFlxAddICU
        If vsFlxAddICU.Cols <> 0 Then
            On Error Resume Next
            If .TextMatrix(.Row, UI_监护室名称) <> "" And Trim(.TextMatrix(.Row, UI_进入时间)) = "" Then
                '看是否第一行
                If .Row > 1 Then
                    If .TextMatrix(.Row - 1, UI_监护室名称) <> "" And IsDate(.TextMatrix(.Row - 1, UI_退出时间)) Then
                        datTemp = CDate(.TextMatrix(.Row - 1, UI_退出时间))
                        If Format(datTemp, "yyyy-mm-dd HH:MM") < Format(gclsPros.InTime, "yyyy-mm-dd HH:MM") Then
                            .TextMatrix(.Row, UI_进入时间) = ""
                        Else
                            '以上行为准
                            .TextMatrix(.Row, UI_进入时间) = Trim(.TextMatrix(.Row - 1, UI_退出时间))
                        End If
                        datTemp = DateAdd("d", 1, datTemp)
    
                        If Format(datTemp, "yyyy-mm-dd HH:MM") < Format(gclsPros.InTime, "yyyy-mm-dd HH:MM") Or Format(datTemp, "yyyy-mm-dd HH:MM") > Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM") Then
                            .TextMatrix(.Row, UI_退出时间) = ""
                        Else
                            .TextMatrix(.Row, UI_退出时间) = Format(datTemp, "yyyy-mm-dd HH:MM")
                        End If
                    Else
                        '以入出院为准
                        .TextMatrix(.Row, UI_进入时间) = Format(gclsPros.InTime, "yyyy-mm-dd HH:MM")
                        .TextMatrix(.Row, UI_退出时间) = Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM")
                    End If
                Else
                    If Trim(.TextMatrix(.Row, UI_退出时间)) = "" Then
                        '以入出院为准
                        .TextMatrix(.Row, UI_进入时间) = Format(gclsPros.InTime, "yyyy-mm-dd HH:MM")
                        .TextMatrix(.Row, UI_退出时间) = Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM")
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub VSFlxGotFocus(ByRef vsFlex As VSFlexGrid)
'vsFlex_GotFocus事件
    Call ChangeCtl
    Select Case vsFlex.Name
        Case "vsChemoth"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, CI_化学治疗编码, CI_化疗效果, 1, CI_化学治疗编码)
        Case "vsRadioth"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, RI_放射治疗编码, RI_放疗效果, 1, RI_放射治疗编码)
        Case "vsSpirit"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, SI_药物名称, SI_疗效, 1, SI_药物名称)
        Case "vsKSS"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, KI_抗菌药物名, KI_联合用药, 1, KI_抗菌药物名)
        Case "vsFlxAddICU"
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, UI_监护室名称, UI_再入住原因, 1, UI_监护室名称)
            Else
                Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, UI_监护室名称, UI_退出时间, 1, UI_监护室名称)
            End If
        Case "vsfMain"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, 0, vsFlex.Cols - 1, 1, 1)
            If vsFlex.TextMatrix(0, vsFlex.Col) = "项目" Then vsFlex.Col = vsFlex.Col + 1
        Case "vsInfect"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, FI_确诊日期, FI_医院感染名称, 1, FI_确诊日期)
        Case "vsSample"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, MI_标本, MI_送检日期, 1, MI_标本)
        Case "vsTSJC"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, 1, vsFlex.Cols - 1, 0, 1)
    End Select
End Sub

Public Sub FlxAddICUKeyDown(ByRef vsFlxAddICU As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsFlxAddICU_KeyDown事件
    Dim i As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid
    Dim int序号 As Integer
    Dim lngRevRow As Long
    Dim strType As String

    If vsFlxAddICU.Editable = flexEDNone Then Exit Sub
    With vsFlxAddICU
        If intKeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行数据吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                '四川版删除重症监护时同步修改重症监护序号与重症监护器械的ICU类型序号
                Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
                int序号 = Val(.TextMatrix(.Row, UI_序号))
                For i = 1 To vsTmp.Rows - 1
                    If i < vsTmp.Rows Then
                        If Val(vsTmp.Cell(flexcpData, i, TI_ICU类型)) = int序号 Then
                            If strType = "" Then
                                lngRevRow = i
                                If lngRevRow <> 0 Then vsTmp.RemoveItem lngRevRow
                                i = i - 1
                            ElseIf Split(vsTmp.TextMatrix(i, TI_ICU类型), "-")(1) = .TextMatrix(.Row, UI_监护室名称) Then
                                lngRevRow = i
                                If lngRevRow <> 0 Then vsTmp.RemoveItem lngRevRow
                                i = i - 1
                            End If
                        ElseIf Val(vsTmp.Cell(flexcpData, i, TI_ICU类型)) > int序号 Then
                            strType = vsTmp.Cell(flexcpData, i, TI_ICU类型)
                            vsTmp.Cell(flexcpData, i, TI_ICU类型) = Val(vsTmp.Cell(flexcpData, i, TI_ICU类型)) - 1
                            vsTmp.TextMatrix(i, TI_ICU类型) = Val(vsTmp.Cell(flexcpData, i, TI_ICU类型)) & "-" & Split(vsTmp.TextMatrix(i, TI_ICU类型), "-")(1)
                        End If
                    End If
                Next
                '删除已经删除的重症监护器械类型序号的行
                If vsTmp.Rows = 1 Then vsTmp.Rows = vsTmp.Rows + 1
                Call ChangeVSFHeight(vsFlxAddICU, True)
            End If

            If .Row >= .FixedRows Then
                .RemoveItem .Row
                Call ChangeVSFHeight(vsFlxAddICU, True)
            End If
            If gclsPros.MedPageSandard = ST_四川省标准 Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, UI_序号) = i
                Next
            End If
        ElseIf intKeyCode = vbKeyReturn Then
            intKeyCode = 0
            If .TextMatrix(.Row, UI_监护室名称) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsFlxAddICU, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = UI_监护室名称
                    Call ChangeVSFHeight(vsFlxAddICU, True)
                Else
                    .Col = LngCol
                End If
            End If
        End If
    End With
End Sub

Public Sub FlxAddICUKeyDownEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsFlxAddICU_KeyDownEdit事件

    Dim strKEY As String
    If vsFlxAddICU.Editable = flexEDNone Then Exit Sub
    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsFlxAddICU
        Select Case LngCol
            Case UI_监护室名称
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                If strKEY = "" Then Exit Sub
            Case UI_进入时间, UI_退出时间
                If .TextMatrix(.Row, UI_监护室名称) = "" Then Exit Sub
        End Select
        '移动光标,只有标准版、四川版才有重症监护记录
        If LngCol = IIf(gclsPros.MedPageSandard = ST_四川省标准, UI_再入住原因, UI_退出时间) Then
            If .Row = .Rows - 1 Then .Rows = .Rows + 1: Call ChangeVSFHeight(vsFlxAddICU, True)
            .ShowCell .Row + 1, UI_监护室名称
        Else
            .ShowCell .Row, .Col + 1
        End If
    End With
End Sub

Public Sub FlxAddICUStartEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, blnCancel As Boolean)
'vsFlxAddICU_StartEdit
    With vsFlxAddICU
        Select Case LngCol
            Case UI_进入时间
                blnCancel = .TextMatrix(LngRow, UI_监护室名称) = ""
            Case UI_退出时间
                blnCancel = Not IsDate(.TextMatrix(LngRow, UI_进入时间))
            Case UI_再入住计划, UI_再入住原因
                blnCancel = Not IsDate(.TextMatrix(LngRow, UI_退出时间))
        End Select
    End With
End Sub

Public Sub FlxAddICUValidateEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
    Dim strSql As String, rsTmp As Recordset, vPoint As POINTAPI, blnCancelTmp As Boolean
    Dim strInput As String
    Dim strKEY As String

    With vsFlxAddICU
        Select Case LngCol
            Case UI_监护室名称
                If gclsPros.MedPageSandard = ST_卫生部标准 Then
                    strInput = UCase(.EditText)
                    If strInput = "" Then Exit Sub

                    strSql = " Select Distinct A.ID,A.编码,A.名称" & _
                            " From 部门表 A,部门性质说明 B" & _
                            " Where B.部门ID=A.ID And B.工作性质='ICU'" & _
                            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                            " And (A.编码 Like [1] Or A.简码 Like [2] Or A.名称 Like [2])" & _
                            " Order by A.编码"
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "重症监护室", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancelTmp, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%")
                    If rsTmp Is Nothing Then
                        If Not blnCancelTmp Then '无匹配输入时,按任意输入处理,取消不同
                            MsgBox "没有设置ICU重症监护室。", vbInformation, gstrSysName
                        End If
                    Else
                        .TextMatrix(LngRow, LngCol) = rsTmp!名称 & ""
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                End If
            Case UI_进入时间, UI_退出时间
                If .TextMatrix(.Row, UI_监护室名称) = "" Then Exit Sub
                If CheckInPutIsDate(vsFlxAddICU, LngRow, LngCol) = False Then
                    blnCancel = True
                    Exit Sub
                End If
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
                If strKEY <> "" And strKEY <> "-  -     :" And strKEY <> "____-__-__ __:__" And InStr("0123456789", Mid(strKEY, 1, 1)) > 0 Then
                    If Not CheckDateRange(strKEY) Then
                        MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                        blnCancel = True
                        Exit Sub
                    End If
                End If
            Case UI_再入住原因
                If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                    MsgBox "不能超过50个汉字或100个字符的长度。", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
        End Select
    End With
End Sub

'vsICUInstruments事件
Public Sub vsICUInstrumentsAfterEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
    Dim strInput As String

    If LngCol = TI_开始时间 Or LngCol = TI_结束时间 Then
        strInput = zlStr.FullDate(vsICUInstruments.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsICUInstruments.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
End Sub

Public Sub vsICUInstrumentsAfterRowColChange(ByRef vsICUInstruments As VSFlexGrid, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'vsICUInstruments_AfterRowColChange
    Dim vsFlxAddICU As VSFlexGrid
    Dim i As Long, strList As String
    Set vsFlxAddICU = gclsPros.CurrentForm.vsFlxAddICU
    With vsFlxAddICU
        If NewCol = TI_ICU类型 Then
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, UI_序号) = i
            Next
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, UI_监护室名称) <> "" Then
                    strList = strList & "|" & .TextMatrix(i, UI_序号) & "-" & .TextMatrix(i, UI_监护室名称)
                End If
            Next
            strList = Mid(strList, 2)
            vsICUInstruments.ColComboList(TI_ICU类型) = strList
            vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
        End If
    End With
End Sub

Public Sub vsICUInstrumentsKeyDown(ByRef vsICUInstruments As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsICUInstruments_KeyDown
    Dim LngCol As Long, i As Long
    If vsICUInstruments.Editable = flexEDNone Then
        If intKeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Exit Sub
    End If
    If intKeyCode = vbKeyDelete Then
        If MsgBox("你是否真的要删除该行数据吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        With vsICUInstruments
            If .Row = .Rows - 1 Then
                .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsICUInstruments, True)
            End If
        End With
    ElseIf intKeyCode = vbKeyReturn Then
        intKeyCode = 0
        With vsICUInstruments
            If .TextMatrix(.Row, TI_ICU类型) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsICUInstruments, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = TI_ICU类型
                Else
                    .Col = LngCol
                End If
            End If
        End With
    End If
End Sub

Public Sub vsICUInstrumentsKeyDownEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsICUInstruments_KeyDownEdit事件
    Dim strKEY As String
    If vsICUInstruments.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyReturn Then Exit Sub
    With vsICUInstruments
        Select Case LngCol
            Case TI_ICU类型
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                If strKEY = "" Then Exit Sub
            Case TI_开始时间, TI_结束时间
                If .TextMatrix(.Row, TI_器械及导管) = "" Then Exit Sub
                If CheckInPutIsDate(gclsPros.CurrentForm.vsFlxAddICU, LngRow, LngCol) = False Then
                    intKeyCode = 0
                    zlCommFun.PressKey vbKeySpace
                    .EditSelStart = 1
                    .EditSelLength = 1000
                    Exit Sub
                End If
        End Select
        '移动光标,四川版才有重症监护器械记录
        If LngCol = TI_感染累计小时 Then
            If .Row = .Rows - 1 Then .Rows = .Rows + 1: Call ChangeVSFHeight(vsICUInstruments, True)
            .ShowCell .Row + 1, TI_ICU类型
        Else
            .ShowCell .Row, .Col + 1
        End If
    End With
End Sub

Public Sub vsICUInstrumentsStartEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsICUInstruments_StartEdit事件
    With vsICUInstruments
        Select Case LngCol
            Case TI_器械及导管
                blnCancel = .TextMatrix(LngRow, TI_ICU类型) = ""
            Case TI_开始时间
                blnCancel = .TextMatrix(LngRow, TI_器械及导管) = ""
            Case TI_结束时间
                blnCancel = Not IsDate(.TextMatrix(LngRow, TI_开始时间))
            Case TI_感染累计小时
                blnCancel = Not IsDate(.TextMatrix(LngRow, TI_结束时间))
        End Select
    End With
End Sub

Public Sub vsICUInstrumentsValidateEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsICUInstruments_ValidateEdit
    Dim strKEY As String
    Dim str入住时间 As String, str转出时间 As String
    Dim strTmp As String, i As Long
    Dim lngDif As Long

    With vsICUInstruments
        strKEY = Trim(.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        If strKEY = "" Then strKEY = .TextMatrix(LngRow, LngCol)

        Select Case LngCol
            Case TI_ICU类型
                If .EditText <> "" Then
                    .Cell(flexcpData, LngRow, TI_ICU类型) = Mid(.EditText, 1, InStr(.EditText, "-") - 1)
                    .RowData(LngRow) = Val(.Cell(flexcpData, LngRow, TI_ICU类型))
                End If
            Case TI_器械及导管
                 i = Val(.Cell(flexcpData, LngRow, TI_ICU类型))
                str入住时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_进入时间))
                str转出时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_退出时间))
                '没有开始时间，则为进入时间+1分
                If .TextMatrix(LngRow, TI_开始时间) = "" And str入住时间 <> "" Then .TextMatrix(LngRow, TI_开始时间) = Format(CDate(str入住时间) + 1 / 24 / 60, "yyyy-mm-dd hh:mm")
                '没有开始时间，则为退出时间-1分
                If .TextMatrix(LngRow, TI_结束时间) = "" And str转出时间 <> "" Then .TextMatrix(LngRow, TI_结束时间) = Format(CDate(str转出时间) - 1 / 24 / 60, "yyyy-mm-dd hh:mm")
                 If .TextMatrix(LngRow, TI_结束时间) <> "" And .TextMatrix(LngRow, TI_开始时间) <> "" Then
                    lngDif = DateDiff("n", CDate(.TextMatrix(LngRow, TI_开始时间)), CDate(.TextMatrix(LngRow, TI_结束时间)))
                    .TextMatrix(LngRow, TI_感染累计小时) = Format(lngDif \ 60, "00") & ":" & Format(lngDif Mod 60, "00")
                 End If
            Case TI_开始时间, TI_结束时间
                i = Val(.Cell(flexcpData, LngRow, TI_ICU类型))
                strTmp = IIf(LngCol = TI_开始时间, "开始使用时间", "结束使用时间")
                str入住时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_进入时间))
                str转出时间 = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_退出时间))
                strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
                If strKEY = "" Then Exit Sub
                If Not IsDate(strKEY) Then
                    MsgBox strTmp & "必须为日期型,请重新输入！", vbInformation + vbDefaultButton1, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
                If IsDate(str入住时间) Then
                    If CDate(strKEY) < CDate(str入住时间) Then
                        .EditText = str入住时间
                        ShowMessage vsICUInstruments, "注:" & vbCrLf & "  " & strTmp & "小于了入住时间,请检查！"
                        blnCancel = True
                        Exit Sub
                    End If
                End If
                If IsDate(str转出时间) Then
                    If CDate(strKEY) > CDate(str转出时间) Then
                        .EditText = str入住时间
                        ShowMessage vsICUInstruments, "注:" & vbCrLf & "    " & strTmp & "大于了转出时间,请检查！"
                        blnCancel = True
                        Exit Sub
                    End If
                End If
                strTmp = .TextMatrix(LngRow, IIf(LngCol = TI_开始时间, TI_结束时间, TI_开始时间))
                If IsDate(strTmp) Then
                    If CDate(strKEY) >= CDate(strTmp) And LngCol = TI_开始时间 Then
                        ShowMessage vsICUInstruments, "您输入的开始使用时间大于结束使用时间，请检查。"
                        blnCancel = True
                        Exit Sub
                    ElseIf CDate(strKEY) <= CDate(strTmp) And LngCol = TI_结束时间 Then
                        ShowMessage vsICUInstruments, "您输入的结束使用时间小于开始使用时间，请检查。"
                        blnCancel = True
                        Exit Sub
                    End If
                    If .TextMatrix(LngRow, TI_结束时间) <> "" And .TextMatrix(LngRow, TI_开始时间) <> "" Then
                        lngDif = DateDiff("n", CDate(.TextMatrix(LngRow, TI_开始时间)), CDate(.TextMatrix(LngRow, TI_结束时间)))
                        .TextMatrix(LngRow, TI_感染累计小时) = Format(lngDif \ 60, "00") & ":" & Format(lngDif Mod 60, "00")
                    End If
                End If
                If Not CheckDateRange(strKEY) Then
                    ShowMessage vsICUInstruments, "您输入的时间必须在病人的住院期间。"
                    blnCancel = True
                    Exit Sub
                End If
        Case TI_感染累计小时
            If InStr(strKEY, ":") > 0 Then
                If Val(Mid(strKEY, InStr(strKEY, ":") + 1)) >= 60 Then
                    ShowMessage vsICUInstruments, "输入的分钟数不能超过59分钟。"
                    blnCancel = True
                    Exit Sub
                End If
            End If
        End Select
    End With
End Sub

'vsInfect事件
Public Sub vsInfectAfterEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsInfect_AfterEdit事件
    Dim strInput As String
    If LngCol = FI_确诊日期 Then
        strInput = zlStr.FullDate(vsInfect.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsInfect.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
    Call vsInfectAfterRowColChange(vsInfect, -1, -1, vsInfect.Row, vsInfect.Col)
End Sub

Public Sub vsInfectAfterRowColChange(ByRef vsInfect As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsInfect_AfterRowColChange事件
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsInfect.ColComboList(FI_医院感染名称) = "..."
End Sub

Public Sub vsInfectCellButtonClick(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsInfect_CellButtonClick事件
    Dim blnCancle As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim textTmp As TextBox
    Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("医院感染目录"), , "ID,编码,简码,名称")
    If rsTmp.RecordCount = 0 Then
        MsgBox "没有感染项目可以选择,请到字典管理工具中设置感染项目。", vbInformation, gstrSysName
    Else
        Set textTmp = GetReplaceObject(vsInfect)
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , , rsTmp) Then
            With vsInfect
                .TextMatrix(LngRow, FI_医院感染编码) = rsTmp!编码 & ""
                .TextMatrix(LngRow, LngCol) = rsTmp!名称 & ""
                If LngRow = .Rows - 1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsInfect, True)
                End If
                .ShowCell .Row + 1, FI_确诊日期
            End With
        End If
    End If
End Sub

Public Sub vsInfectKeyDown(ByRef vsInfect As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsInfect_KeyDown事件
    If vsInfect.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行数据吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            With vsInfect
                If .Rows = .FixedRows Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    .RemoveItem .Row
                     Call ChangeVSFHeight(vsInfect, True)
                End If
            End With
    ElseIf intKeyCode = vbKeyReturn And Trim(vsInfect.TextMatrix(vsInfect.Row, FI_确诊日期)) = "" Then
        zlCommFun.PressKey vbKeyTab: mblnReturn = True
    ElseIf intKeyCode = Asc("*") Then
        Call vsInfectCellButtonClick(vsInfect, vsInfect.Row, vsInfect.Col)
    Else
         vsInfect.ColComboList(FI_医院感染名称) = ""  '使按钮状态进入输入状态
    End If
    Call VsGriedFocuesMove(vsInfect, vsInfect.Row, vsInfect.Col, intKeyCode)
End Sub

Public Sub vsInfectKeyPress(ByRef vsInfect As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsInfect_KeyPress事件
    If vsInfect.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then intKeyAscii = 0
End Sub

Public Sub vsInfectKeyPressEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsInfect_KeyPressEdit事件

    If intKeyAscii = Asc("'") Then intKeyAscii = 0: Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
        Exit Sub
    End If
    Select Case LngCol
        Case FI_确诊日期, FI_感染部位, FI_医院感染名称
            Call VsFlxGridCheckKeyPress(vsInfect, LngRow, LngCol, intKeyAscii, m文本式)
        Case Else
    End Select
End Sub

Public Sub vsInfectStartEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsInfect_StartEdit事件
    Select Case LngCol
        Case FI_感染部位
            blnCancel = Not IsDate(vsInfect.TextMatrix(LngRow, FI_确诊日期))
        Case FI_医院感染名称
            blnCancel = vsInfect.TextMatrix(LngRow, FI_感染部位) = ""
    End Select
End Sub

Public Sub vsInfectValidateEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsInfect_ValidateEdit事件
    Dim strKEY As String
    Dim strFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim blnInputCancel As Boolean
    Dim textTmp As TextBox
    With vsInfect
        strKEY = Trim(vsInfect.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        Select Case LngCol
            Case FI_确诊日期
                If strKEY <> "" Then
                    strKEY = zlStr.FullDate(strKEY, False, gclsPros.InTime, gclsPros.OutTime)
                    If Not IsDate(strKEY) Then
                        MsgBox "确诊日期必须为日期型,请重新输入！", vbInformation + vbDefaultButton1, gstrSysName
                        blnCancel = True
                        zlCommFun.PressKey vbKeySpace
                        .EditSelStart = 0
                        .EditSelLength = 1000
                        Exit Sub
                    End If
                    If Not CheckDateRange(strKEY, True) Then
                        MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                        blnCancel = True
                        Exit Sub
                    End If
                End If
            Case FI_感染部位
                 Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
            Case FI_医院感染名称
                If strKEY = "" Then Exit Sub
                Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("医院感染目录"), , "ID,编码,简码,名称")
                If rsTmp.RecordCount = 0 Then
                    MsgBox "没有感染项目可以选择。", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                Else
                    strKEY = UCase$(strKEY)
                    strFilter = "简码 Like '" & strKEY & "*' OR 编码  Like '" & strKEY & "*'  OR 名称 like '" & IIf(gclsPros.LikeString <> "", "*", "") & strKEY & "*' "
                    Set textTmp = GetReplaceObject(vsInfect)
                    blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, Rec.FilterNew(rsTmp, strFilter), True, , , rsTmp)
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "未找到匹配的感染项目。", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        Else
                            blnCancel = True
                            Exit Sub
                        End If
                    Else
                        .TextMatrix(LngRow, FI_医院感染编码) = rsTmp!编码 & ""
                        .EditText = rsTmp!名称 & ""
                        vsInfect.SetFocus
                        Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
                    End If
                End If
        End Select
    End With
End Sub

'vsSample事件
Public Sub vsSampleAfterEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsSample_AfterEdit事件
    Dim strInput As String

    If LngCol = MI_送检日期 Then
        strInput = zlStr.FullDate(vsSample.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsSample.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
End Sub

Public Sub vsSampleKeyDown(ByRef vsSample As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsSample_KeyDown事件
    Dim LngCol As Long, i As Long
    With vsSample
        If vsSample.Editable = flexEDNone Then Exit Sub
        If intKeyCode = vbKeyDelete Then
                If MsgBox("你是否真的要删除该行数据吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                If .Rows = .FixedRows Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    .RemoveItem .Row
                     Call ChangeVSFHeight(vsSample, True)
                End If
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(vsSample.Row, MI_标本) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                     Call ChangeVSFHeight(vsSample, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = MI_标本
                     Call ChangeVSFHeight(vsSample, True)
                Else
                    .Col = LngCol
                End If
            End If
        End If
    End With
End Sub

Public Sub vsSampleStartEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSample_StartEdit事件
    Select Case LngCol
        Case MI_病原学代码及名称
            blnCancel = vsSample.TextMatrix(LngRow, MI_标本) = ""
        Case MI_送检日期
            blnCancel = vsSample.TextMatrix(LngRow, MI_病原学代码及名称) = ""
    End Select
End Sub

Public Sub vsSampleValidateEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSample_ValidateEdit事件
    Dim strKEY As String
    Dim strFilter As String
    Dim rsTmp As ADODB.Recordset

    With vsSample
        strKEY = Trim(vsSample.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        Select Case LngCol
            Case MI_送检日期
                strKEY = zlStr.FullDate(strKEY, False, gclsPros.InTime, gclsPros.OutTime)
                If Not IsDate(strKEY) Then
                    MsgBox "确诊日期必须为日期型,请重新输入！", vbInformation + vbDefaultButton1, gstrSysName
                    blnCancel = True
                    zlCommFun.PressKey vbKeySpace
                    .EditSelStart = 0
                    .EditSelLength = 1000
                    Exit Sub
                End If
                If Not CheckDateRange(strKEY, True) Then
                    MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
        End Select
    End With
End Sub

'vsTSJC事件
Public Sub TSJCAfterEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsTSJC_AfterEdit事件
    Call TSJCAfterRowColChange(vsTSJC, -1, -1, vsTSJC.Row, vsTSJC.Col)
End Sub

Public Sub TSJCAfterRowColChange(ByRef vsTSJC As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsTSJC_AfterRowColChange事件
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsTSJC.ComboList = "..."
End Sub

Public Sub TSJCCellButtonClick(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsTSJC_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String

    With vsTSJC
        strSQLItem = _
            " From 诊疗项目目录 A" & _
            " Where A.类别='D' And A.服务对象 IN(2,3) And A.单独应用=1" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
        strSql = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
            " Group by ID,上级ID,编码,名称"
        strSql = strSql & " Union ALL" & _
            " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位" & _
            strSQLItem & " Order By 末级,级ID Desc,编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "特殊检查", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有检查项目数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call TSJCSetDiagInput(LngRow, rsTmp)
            Call TSJCEnterNextCell
        End If
    End With
End Sub

Public Sub TSJCKeyDown(ByRef vsTSJC As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsTSJC_KeyDown事件
    'If mbln护士站 Or mblnReadOnly Then Exit Sub
    With vsTSJC
        If intKeyCode = vbKeyF4 Then
            Call zlCommFun.PressKey(vbKeySpace)
        ElseIf intKeyCode = vbKeyDelete Then
            If MsgBox("确实要删除该行内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .TextMatrix(.Row, 1) = ""
            End If
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call TSJCKeyPress(vsTSJC, intKeyCode)
        End If
    End With
End Sub

Public Sub TSJCKeyPress(ByRef vsTSJC As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsTSJC_KeyPress事件
    If vsTSJC.Editable = flexEDNone Then Exit Sub
    With vsTSJC
        If intKeyAscii = 13 Then
            intKeyAscii = 0
            Call TSJCEnterNextCell
        ElseIf gclsPros.MedPageSandard <> ST_四川省标准 Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call TSJCCellButtonClick(vsTSJC, .Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Public Sub TSJCKeyPressEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsTSJC_KeyPressEdit事件
    If intKeyAscii = vbKeyReturn Then
        gclsPros.IsReturn = True
    Else
        gclsPros.IsReturn = False
    End If
End Sub

Public Sub TSJCSetupEditWindow(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef lngEditWindow As Long, ByRef blnIsCombo As Boolean)
'vsTSJC_SetupEditWindow事件
    With vsTSJC
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub TSJCValidateEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTSJC_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsTSJC
        If .EditText = "" Then
            .EditText = .Cell(flexcpData, LngRow, LngCol)
            If gclsPros.IsReturn Then Call TSJCEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
            If gclsPros.IsReturn Then Call TSJCEnterNextCell
        Else
            strInput = UCase(.EditText)
            If LenB(StrConv(strInput, vbFromUnicode)) > 100 Then
                MsgBox "您输入的内容不能超过50个汉字。", vbInformation, gstrSysName
                blnCancel = True
                Exit Sub
            End If
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
            strSql = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='D' And A.服务对象 IN(2,3)" & _
                " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And A.单独应用=1 And B.码类=[3] And (" & strSql & ")" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " Order by A.编码"
            If zlCommFun.IsCharChinese(strInput) Then
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                If rsTmp.EOF Then
                    Set rsTmp = Nothing
                ElseIf rsTmp.RecordCount > 1 Then
                    Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                End If
                Call TSJCSetDiagInput(LngRow, rsTmp)
                .EditText = .Text
                If gclsPros.IsReturn Then Call TSJCEnterNextCell
            Else
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "特殊检查", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                    strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                    blnCancel = True
                Else
                    Call TSJCSetDiagInput(LngRow, rsTmp)
                    .EditText = .Text
                    If gclsPros.IsReturn Then Call TSJCEnterNextCell
                End If
            End If
        End If
        gclsPros.IsReturn = False
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsfMain事件
Public Sub vsfMainEnterCell(ByRef vsfMain As VSFlexGrid)
'vsfMain_EnterCell事件
    With vsfMain
        Select Case .Col
            Case 1, 4, 7
                If InStr(.TextMatrix(.Row, .Col + 1), ",") > 0 Then
                    .ColComboList(.Col) = Replace(.TextMatrix(.Row, .Col + 1), ",", "|")
                Else
                    .ColComboList(.Col) = ""
                End If
        End Select
    End With
End Sub

Public Sub vsfMainKeyPress(ByRef vsfMain As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsfMain_KeyPress事件
    With vsfMain
        If .Editable = flexEDNone Or .Rows <= 1 Then zlCommFun.PressKey (vbKeyTab): Exit Sub
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Select Case .Col
                Case 0, 3, 6
                    .Col = .Col + 1
                Case 1, 4, 7
                    If .Col = .Cols - 2 Then
                        If .Row <> .Rows - 1 Then
                            .Col = 1
                            .Row = .Row + 1
                        Else
                            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                        End If
                    Else
                        .Col = .Col + 3
                    End If
            End Select
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Public Sub vsfMainStartEdit(ByRef vsfMain As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsfMain_StartEdit事件
    Select Case LngCol Mod 3
        Case 0
            blnCancel = True
        Case 1
            blnCancel = vsfMain.TextMatrix(LngRow, LngCol - 1) = ""
    End Select
End Sub

Public Sub vsfMainValidateEdit(ByRef vsfMain As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsfMain_ValidateEdit事件
    Dim sngNum1, sngNum2 As Single

    With vsfMain
        If InStr(.TextMatrix(LngRow, LngCol + 1), "...") > 0 Then
            sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 1, InStr(.TextMatrix(LngRow, LngCol + 1), "...") - 1)
            sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(.TextMatrix(LngRow, LngCol + 1), "...") + 3)
            If Not IsNumeric(.EditText) Then
                blnCancel = True
            ElseIf CSng(.EditText) < sngNum1 Or CSng(.EditText) > sngNum2 Then
                MsgBox "数据应该在" & .TextMatrix(LngRow, LngCol + 1) & "的范围以内!", vbInformation, gstrSysName
                blnCancel = True
            End If
        ElseIf InStr(.TextMatrix(LngRow, LngCol + 1), "-") > 0 Then
            If InStr(.TextMatrix(LngRow, LngCol + 1), "-") = 1 Then
                sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 2, InStr(2, .TextMatrix(LngRow, LngCol + 1), "-") - 1)
                sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(2, .TextMatrix(LngRow, LngCol + 1), "-") + 1)
            Else
                sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 1, InStr(1, .TextMatrix(LngRow, LngCol + 1), "-") - 1)
                sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(1, .TextMatrix(LngRow, LngCol + 1), "-") + 1)
            End If
            If Not IsNumeric(.EditText) Then
                blnCancel = True
            ElseIf CSng(.EditText) < sngNum1 Or CSng(.EditText) > sngNum2 Then
                MsgBox "数据应该在" & .TextMatrix(LngRow, LngCol + 1) & "的范围以内!", vbInformation, gstrSysName
                blnCancel = True
            End If
        ElseIf .TextMatrix(LngRow, LngCol + 1) = "" Then
            If zlCommFun.ActualLen(.EditText) > gclsPros.ValueLen Then
                MsgBox "输入长度不能大于" & "[" & gclsPros.ValueLen & "]", vbInformation, gstrSysName
                blnCancel = True
            End If
        End If
    End With
End Sub

'vsFrees事件
Public Sub vsFeesComboDropDown(ByVal LngRow As Long, ByVal LngCol As Long)
'vsFees_ComboDropDown事件
    Dim vsFees As VSFlexGrid
    Dim i As Long

    Set vsFees = gclsPros.CurrentForm.vsFees
    With vsFees
        If LngCol Mod 2 = 0 Then
            '定位到匹配项
            If .TextMatrix(LngRow, LngCol) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Sub vsFeesKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsFees_KeyDown事件
    Dim vsFree As VSFlexGrid

    Set vsFree = gclsPros.CurrentForm.vsFees
    With vsFree
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If Not FreeHaveLowLevel(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)) Then
                If MsgBox("是否删除该费用？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    Call AddOrDelFreeCols(vsFree, .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)), .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col + 1, .Col)), False)
                End If
            End If

        ElseIf intKeyCode = vbKeyReturn Then
            intKeyCode = 0
            If .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)) = "" Or .Editable = flexEDNone Then
                If gclsPros.CurrentForm.cboBaseInfo(BCC_病例分型).Enabled And Not gclsPros.CurrentForm.cboBaseInfo(BCC_病例分型).Locked Then
                    Call gclsPros.CurrentForm.cboBaseInfo(BCC_病例分型).SetFocus
                End If
            Else
                If IIf(.Col Mod 2 = 0, .Col, .Col - 1) = 4 Then
                    .Col = 0: .Row = .Row + 1
                Else
                    .Col = IIf(.Col Mod 2 = 0, .Col, .Col - 1) + 2
                End If
            End If
        End If
    End With
End Sub

Public Sub vsFeesKeyPressEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsFees_KeyPressEdit事件
    Dim vsFees As VSFlexGrid

    Set vsFees = gclsPros.CurrentForm.vsFees
    If vsFees.Editable = flexEDNone Then Exit Sub
    With vsFees
         If intKeyAscii = vbKeyReturn Then
            gclsPros.IsReturn = True
            If LngCol Mod 2 = 0 Then
                intKeyAscii = 0
                If .ComboIndex <> -1 Then
                    '此时.TextMatrix尚未更新,所以取ComboItem
                    .TextMatrix(LngRow, LngCol) = .ComboItem(.ComboIndex)
                    Call EnterNextCellFees(vsFees)
                End If
            End If
         Else
             If LngCol Mod 2 = 1 Then
                 If .EditSelLength <> 0 Then Exit Sub
                 If Len(.EditText) > 17 Then intKeyAscii = 0: Exit Sub
                 Call VsFlxGridCheckKeyPress(vsFees, LngRow, LngCol, intKeyAscii, m金额式)
             End If
             gclsPros.IsReturn = False
         End If
    End With
End Sub

Public Sub vsFeesStartEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsFees_StartEdit事件
    '存在子级费用，则不允许编辑
    If FreeHaveLowLevel(LngRow, IIf(LngCol Mod 2 = 0, LngCol, LngCol - 1)) Then blnCancel = True
End Sub

Public Sub vsFeesValidateEdit(ByVal LngRow As Long, ByVal LngCol As Long, byrefCancel As Boolean)
'vsFees_ValidateEdit事件
    Dim vsFree As VSFlexGrid
    Dim i As Long, lngTmpRow As Long, lngTmpCol As Long

    Set vsFree = gclsPros.CurrentForm.vsFees
    With vsFree
        If LngCol Mod 2 = 0 Then
            For i = .FixedRows * 3 To (.Rows - 1) * 3
                lngTmpRow = i \ 3: lngTmpCol = (i Mod 3) * 2
                If .TextMatrix(lngTmpRow, lngTmpCol) = .EditText And lngTmpRow <> LngRow And lngTmpCol <> LngCol Then
                    If gclsPros.SameName Then
                        Call AddOrDelFreeCols(vsFree, .TextMatrix(LngRow, LngCol), "", True)
                        Exit Sub
                    End If
                End If
            Next
        Else
            .TextMatrix(LngRow, LngCol) = Format(.EditText, gclsPros.FreeFormat)
            Call SumAndSetFrees
        End If
    End With
End Sub

'vsTransfer事件
Public Sub vsTransferAfterRowColChange(ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsTransfer_AfterRowColChange事件
    Dim vsTransfer As VSFlexGrid
    Dim blnEdit As Boolean
    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    With vsTransfer
        If lngNewCol >= .FixedCols Then
            If lngNewRow = DR_转科科室 Then
                blnEdit = .TextMatrix(DR_转科科室, lngNewCol - 1) <> ""
            Else
                blnEdit = .TextMatrix(DR_转科科室, lngNewCol) <> ""
            End If
            If lngNewRow = DR_转科科室 Then
                .FocusRect = IIf(blnEdit, flexFocusSolid, flexFocusLight)
                .ComboList = IIf(blnEdit, "...", "")
            Else
                .ComboList = ""
                .FocusRect = IIf(blnEdit, flexFocusSolid, flexFocusLight)
            End If
        End If
    End With
End Sub

Public Sub vsTransferCellButtonClick(ByVal LngRow As Long, ByVal LngCol As Long)
'vsTransfer_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim vsTransfer As VSFlexGrid
    Dim textTmp As TextBox

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer

    With vsTransfer
        If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
        grsDeptInfo.Filter = "工作性质='临床'"
        If grsDeptInfo.RecordCount = 0 Then
            MsgBox "未找到临床部门数据，请在基础数据管理中设置部门性质为临床！", vbInformation, gstrSysName
        Else
            grsDeptInfo.Filter = "工作性质='临床'"
            Set textTmp = GetReplaceObject(vsTransfer)
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, grsDeptInfo, True, , , rsTmp) Then
                If rsTmp.RecordCount > 0 Then
                    .TextMatrix(LngRow, LngCol) = rsTmp!名称
                End If
            End If
        End If
    End With
End Sub

Public Sub vsTransferKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsTransfer_KeyDown事件
    Dim vsTransfer As VSFlexGrid
    Dim i As Long

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    If vsTransfer.Editable = flexEDNone Then Exit Sub
    With vsTransfer
        If intKeyCode = vbKeyDelete Then
            For i = .Col To .Cols - 2
                .TextMatrix(DR_转科时间, i) = .TextMatrix(DR_转科时间, i + 1)
                .TextMatrix(DR_转科科室, i) = .TextMatrix(DR_转科科室, i + 1)
            Next
            .TextMatrix(DR_转科时间, .Cols - 1) = ""
            .TextMatrix(DR_转科科室, .Cols - 1) = ""
        ElseIf intKeyCode = vbKeyInsert Then
            If .TextMatrix(0, .Col) <> "" Then
                For i = .Cols - 1 To .Col + 1 Step -1
                    .TextMatrix(DR_转科时间, i) = .TextMatrix(DR_转科时间, i - 1)
                    .TextMatrix(DR_转科科室, i) = .TextMatrix(DR_转科科室, i - 1)
                Next
                .TextMatrix(DR_转科时间, .Col) = ""
                .TextMatrix(DR_转科科室, .Col) = ""
            End If
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsTransferKeyPress(intKeyCode)
        End If
    End With
End Sub

Public Sub vsTransferStartEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTransfer_StartEdit事件
    With gclsPros.CurrentForm.vsTransfer
        If LngRow = DR_转科时间 And .TextMatrix(DR_转科科室, LngCol) = "" Then blnCancel = True
        If LngRow = DR_转科科室 And .TextMatrix(DR_转科科室, LngCol - 1) = "" Then blnCancel = True
    End With
End Sub

Public Sub vsTransferKeyPress(ByRef intKeyAscii As Integer)
'vsTransfer_KeyPress事件
    Dim vsTransfer As VSFlexGrid
    Dim i As Long

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    If vsTransfer.Editable = flexEDNone Then Exit Sub
    With vsTransfer
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            If .Col = .Cols - 1 And .Row = DR_转科时间 Then
                If ControlIsLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间)) Then
                    Call gclsPros.CurrentForm.txtInfo(GC_出院病房).SetFocus
                Else
                    Call gclsPros.CurrentForm.mskDateInfo(DC_出院时间).SetFocus
                End If
            ElseIf .TextMatrix(DR_转科科室, .Col) = "" Then
                If ControlIsLocked(gclsPros.CurrentForm.mskDateInfo(DC_出院时间)) Then
                    Call gclsPros.CurrentForm.txtInfo(GC_出院病房).SetFocus
                Else
                    Call gclsPros.CurrentForm.mskDateInfo(DC_出院时间).SetFocus
                End If
            ElseIf .Row = DR_转科时间 Then
                .Col = .Col + 1: .Row = DR_转科科室
            ElseIf .Row = DR_转科科室 Then
                .Row = DR_转科时间
            End If
        Else
            If .Row = DR_转科科室 Then
                If intKeyAscii = Asc("*") Then
                    intKeyAscii = 0
                    Call vsTransferCellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Public Sub vsTransferValidateEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTransfer_ValidateEdit事件
    Dim vsTransfer As VSFlexGrid
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String
    Dim textTmp As TextBox

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer

    With vsTransfer
        If .EditText = "" And .TextMatrix(.Row, .Col) <> "" Then
            If MsgBox("是否删除该列转科信息？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                .EditText = .TextMatrix(.Row, .Col + 1)
                For i = .Col To .Cols - 2
                    .TextMatrix(DR_转科时间, i) = .TextMatrix(DR_转科时间, i + 1)
                    .TextMatrix(DR_转科科室, i) = .TextMatrix(DR_转科科室, i + 1)
                Next
                .TextMatrix(DR_转科时间, .Cols - 1) = ""
                .TextMatrix(DR_转科科室, .Cols - 1) = ""
            End If
        Else
            If .Row = DR_转科科室 Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                grsDeptInfo.Filter = "工作性质='临床' "
                strInput = UCase(Trim(.EditText))
                If strInput = "" Then Exit Sub
                Set rsTmp = Rec.FilterNew(grsDeptInfo, "名称 Like '*" & strInput & "*' OR 编码 Like '" & strInput & "*' OR 简码 Like '" & strInput & "*'")
                If rsTmp.EOF Then
                    blnCancel = True
                Else
                    If rsTmp.RecordCount = 1 Then
                        .EditText = rsTmp!名称
                    Else
                        Set textTmp = GetReplaceObject(vsTransfer)
                        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , , rsTmp) Then
                            If rsTmp.RecordCount <> 0 Then
                                .EditText = rsTmp!名称
                            Else
                                blnCancel = True
                            End If
                        Else
                            blnCancel = True
                        End If
                    End If
                End If
            Else
                strInput = zlStr.FullDate(.EditText, , gclsPros.InTime, gclsPros.OutTime)
                If strInput <> "" Then
                    If IsDate(strInput) Then
                        .EditText = strInput
                    Else
                        MsgBox "请输入正确的转科时间，例如：""2012-12-21""或""20121221""。", vbInformation, gstrSysName
                        blnCancel = True
                    End If
                ElseIf .EditText <> "" Then
                    blnCancel = True
                End If
            End If
        End If
    End With
End Sub

'vsAller事件
Public Sub AllerAfterEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsAller_AfterEdit事件
    Dim strDate As String

    With vsAller
        Select Case LngCol
            Case AI_过敏时间
                strDate = zlStr.FullDate(.TextMatrix(LngRow, LngCol), False)
                If IsDate(strDate) Then
                    .TextMatrix(LngRow, LngCol) = strDate
                End If
        End Select
        Call AllerAfterRowColChange(vsAller, -1, -1, .Row, .Col)
    End With
End Sub

Public Sub AllerAfterRowColChange(ByRef vsAller As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsAller_AfterRowColChange事件
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    With vsAller
        If lngNewCol = AI_过敏药物 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(.TextMatrix(lngNewRow, AI_过敏药物)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Public Sub AllerCellButtonClick(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsAller_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int性别 As Integer
    Dim vPoint As POINTAPI

    With vsAller
        If gclsPros.UseTYT Then
            If Not gobjPass Is Nothing Then
                strSql = gobjPass.zlPassInputAllergy()
            End If
            If InStr(strSql, ";") > 0 Then
                Call SetAllerInput(LngRow, , strSql)
                Call AllerEnterNextCell
            End If
        Else
            If gclsPros.Sex Like "*男*" Then
                int性别 = 1
            ElseIf gclsPros.Sex Like "*女*" Then
                int性别 = 2
            End If
            If gclsPros.FuncType <> f病案首页 Then
                If gclsPros.CurrentForm.optAller(PC_按药品目录输入).Value = True Then
                    strSql = _
                        " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                        " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                        " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                        " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                        " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                        " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                        " Union All" & _
                        " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                        " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                        " From 诊疗项目目录 A,药品特性 B" & _
                        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                        IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
                Else
                    strSql = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Order By 编码"
                    vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "过敏源", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
                End If
            Else
                strSql = _
                    " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                    " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                    " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                    " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                    IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    If gclsPros.CurrentForm.optAller(PC_按药品目录输入).Value = True Then
                        MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
                    Else
                        MsgBox "没有过敏源数据可以选择", vbInformation, gstrSysName
                    End If
                End If
            Else
                Call SetAllerInput(LngRow, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Public Sub AllerKeyDown(ByRef vsAller As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsAller_KeyDown 事件
    Dim i As Long
    If vsAller.Editable = flexEDNone Then Exit Sub
    'If gbln护士站 Or gblnReadOnly Then Exit Sub

    With vsAller
        If intKeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AI_过敏药物) <> "" Then
                If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call ChangeVSFHeight(vsAller, True, 0)
                End If
            End If
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call AllerKeyPress(vsAller, intKeyCode)
        End If
    End With
End Sub

Public Sub AllerKeyPress(ByRef vsAller As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsAller_KeyPress事件

    If vsAller.Editable = flexEDNone Then Exit Sub
    With vsAller
        If intKeyAscii = vbKeySpace Then  'Space
            If .Col = AI_过敏药物 And gclsPros.UseTYT Then intKeyAscii = 0: Exit Sub
        End If
        If intKeyAscii = 13 Then
            intKeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AI_过敏药物 Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call AllerCellButtonClick(vsAller, .Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Public Sub AllerKeyPressEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsAller_KeyPressEdit 事件
    Dim blnIsNextchr As Boolean
    Dim strChr As String

    If intKeyAscii = 13 Then
        gclsPros.IsReturn = True
    Else
        gclsPros.IsReturn = False
    End If
    With vsAller
        If LngCol = AI_过敏反应 Then
            If intKeyAscii = 13 Then .Col = .Col + 1: .ShowCell LngRow, LngCol: Exit Sub
        ElseIf LngCol = AI_过敏药物 Then
            If intKeyAscii <> 13 Then
                If gclsPros.UseTYT Then intKeyAscii = 0
            End If
        End If
    End With
End Sub

Public Sub AllerSetupEditWindow(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsAller_SetupEditWindow 事件
    With vsAller
        If LngCol = AI_过敏药物 Or LngCol = AI_过敏时间 Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        End If
    End With
End Sub

Public Sub AllerStartEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsAller_StartEdit事件
    If LngCol = AI_过敏反应 And Trim(vsAller.TextMatrix(LngRow, AI_过敏药物)) = "" Then blnCancel = True
End Sub

Public Sub AllerValidateEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, blnCancel As Boolean)
'vsAller_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    Dim curDate As Date
    Dim strDate As String

    With vsAller
        If LngCol = AI_过敏药物 Then
            If .EditText = "" Then
                If .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem .Row
                        If .Rows = .FixedRows Then
                            .Rows = .FixedRows + 1
                            Call ChangeVSFHeight(vsAller, True)
                        End If
                    Else
                        .EditText = .Cell(flexcpData, LngRow, LngCol)
                    End If
                End If
                If gclsPros.IsReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                If gclsPros.IsReturn Then Call AllerEnterNextCell
            Else
                strInput = UCase(.EditText)
                If gclsPros.Sex Like "*男*" Then
                    int性别 = 1
                ElseIf gclsPros.Sex Like "*女*" Then
                    int性别 = 2
                End If
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                If gclsPros.FuncType <> f病案首页 Then
                    If gclsPros.CurrentForm.optAller(PC_按药品目录输入).Value = True Then
                        strSql = _
                            " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                            " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                            " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                            " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                            " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                            IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                            decode(gclsPros.BriefCode, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                            " Order by A.编码"
    
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "过敏药物", False, "", "", False, _
                            False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", int性别, gclsPros.BriefCode + 1)
                    Else
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where 名称 Like [1] Order By 编码"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where zlWbCode(名称) Like [1] Order By 编码"
                            Else
                                strSql = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where 简码 Like [1] Order By 编码"
                            End If
                        End If
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "过敏源", False, "", "", False, _
                            False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            gclsPros.LikeString & UCase(strInput) & "%")
                    End If
                Else
                    strSql = _
                        " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                        " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                        " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                        " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                        IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                        decode(gclsPros.BriefCode, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"
    
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "过敏药物", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", int性别, gclsPros.BriefCode + 1)
                End If
                If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                    blnCancel = True
                Else
                    Call SetAllerInput(LngRow, rsTmp): .EditText = .Text
                    If gclsPros.IsReturn Then Call AllerEnterNextCell
                End If
            End If
            gclsPros.IsReturn = False
        ElseIf LngCol = AI_过敏时间 Then
            If .EditText <> "" Then
                strDate = zlStr.FullDate(.EditText, False)
                If IsDate(strDate) Then
                    curDate = zlDatabase.Currentdate
                    If CDate(strDate) > curDate Then
                        MsgBox "您输入的日期不能大于当前时间。当前时间：" & Format(curDate, "yyyy-mm-dd") & "。", vbInformation, gstrSysName
                        blnCancel = True
                        .EditText = .TextMatrix(LngRow, LngCol)
                    End If
                    .EditText = Format(strDate, "yyyy-MM-dd")
                    If .Cell(flexcpData, LngRow, LngCol) <> .EditText Then
                        .Cell(flexcpData, LngRow, LngCol) = .EditText
                    End If
                Else
                    MsgBox "请输入正确的过敏时间，例如：""2012-12-21""或""121221""。", vbInformation, gstrSysName
                    blnCancel = True
                End If
            End If
        End If
    End With
End Sub

'vsDiagXY事件,vsDiagZY事件
Public Sub DiagAfterEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagXY_AfterEdit事件,vsDiagZY_AfterEdit事件
    Dim bln西医 As Boolean
    Dim i As Long, lngStart As Long, lngEnd As Long

    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        If LngCol = DI_出院情况 Then
            '主要处理非回车离开:不用ComboIndex,取消编辑时不对
            .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
            If bln西医 Then
                If Not DiagCellEditable(vsDiag, LngRow, DI_是否未治) Then
                    .TextMatrix(LngRow, DI_是否未治) = ""
                End If
                If .TextMatrix(LngRow, DI_出院情况) = "死亡" Then
                    lngEnd = FindDiagRow(DT_病理诊断)
                    lngStart = FindDiagRow(DT_出院诊断XY)
                    For i = lngStart To lngEnd - 1
                        If .TextMatrix(i, DI_诊断描述) <> "" Then .TextMatrix(LngRow, DI_是否未治) = ""
                    Next
                End If
            End If
            Call ChangeOutInfo(zlStr.NeedName(.TextMatrix(LngRow, DI_出院情况)))
        ElseIf LngCol = DI_诊断描述 Or gclsPros.FuncType = f病案首页 And gclsPros.CNIndent And LngCol = DI_诊断编码 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If Not (gclsPros.CNIndent And (LngCol = DI_诊断描述 And .TextMatrix(LngRow, DI_诊断编码) <> "" Or LngCol = DI_诊断编码 And .TextMatrix(LngRow, DI_诊断描述) <> "")) And .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
            End If
        End If
        Call DiagAfterRowColChange(vsDiag, -1, -1, .Row, .Col)
         If LngCol = DI_出院情况 Then
              .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
        End If
    End With
End Sub

Public Sub DiagAfterRowColChange(ByRef vsDiag As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsDiagZY_AfterRowColChange事件，vsDiagXY_AfterRowColChange事件
    Dim i As Long
    Dim bln西医 As Boolean
    Dim vPoint As POINTAPI
    Dim blnEdit As Boolean
    Dim j As Long, arrTmp As Variant

    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        '清除图片
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, DI_增加) = Nothing
            Set .Cell(flexcpPicture, i, DI_Del) = Nothing
        Next
        If bln西医 And gclsPros.FuncType <> f诊断选择 Then Call ShowInfectInfo(False)
        If Not DiagCellEditable(vsDiag, lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            blnEdit = True
            Set .CellButtonPicture = Nothing
            If bln西医 And gclsPros.FuncType <> f诊断选择 Then
                If .TextMatrix(lngNewRow, 0) = "院内感染" Then
                    If .TextMatrix(lngNewRow, DI_诊断描述) <> "" Then
                        If lngNewCol = DI_诊断描述 Or lngNewCol = DI_备注 Then
                             vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                             Call ShowInfectInfo(True, , vPoint.X, vPoint.Y)
                        End If
                    End If
                End If
            End If
            Select Case lngNewCol
                Case DI_诊断描述
                    If Not (.TextMatrix(lngNewRow, DI_诊断编码) <> "" And gclsPros.CNIndent And gclsPros.FuncType = f病案首页) Then
                        .ComboList = "..."
                    End If
                Case DI_诊断编码
                    If gclsPros.FuncType = f病案首页 And gclsPros.CNIndent Then .ComboList = "..."
                Case DI_增加, DI_Del
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = IIf(lngNewCol = DI_增加, gclsPros.CurrentForm.imgButtonNew.Picture, gclsPros.CurrentForm.imgButtonDel.Picture)
                Case DI_入院病情
                    If blnEdit Then
                        .ComboList = "有|临床未确定|情况不明|无"
                         If Not gclsPros.IsCheckData Then OS.PressKey vbKeySpace
                    Else
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    End If
                Case DI_出院情况
                    .ComboList = .ColData(lngNewCol)
                    If Trim(.TextMatrix(lngNewRow, lngNewCol)) <> "" Then
                        arrTmp = Split(.ColData(lngNewCol) & "", "|")
                        For j = LBound(arrTmp) To UBound(arrTmp)
                            If zlStr.NeedName(arrTmp(j) & "") = .TextMatrix(lngNewRow, lngNewCol) Then
                                .TextMatrix(lngNewRow, lngNewCol) = arrTmp(j)
                                Exit For
                            End If
                        Next
                    End If
                Case DI_中医证候
                    If .TextMatrix(lngNewRow, DI_诊断描述) = "" Then
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    Else
                        .ComboList = "..."
                    End If
                Case DI_ICD附码
                    .ComboList = "..."
                Case Else
                    .ComboList = ""
            End Select
        End If
        If lngNewRow >= .FixedRows Then
            '显示图片
            If lngNewCol <> DI_增加 And .TextMatrix(lngNewRow, DI_诊断描述) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '下一行诊断为空则不能新增行
                    If Not (.TextMatrix(lngNewRow, DI_诊断分类) = .TextMatrix(lngNewRow + 1, DI_诊断分类) And .TextMatrix(lngNewRow + 1, DI_诊断描述) = "") Then
                         Set .Cell(flexcpPicture, lngNewRow, DI_增加) = gclsPros.CurrentForm.imgButtonNew.Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, DI_增加) = gclsPros.CurrentForm.imgButtonNew.Picture
                End If
            End If
            '显示图片
            If lngNewCol <> DI_Del Then Set .Cell(flexcpPicture, lngNewRow, DI_Del) = gclsPros.CurrentForm.imgButtonDel.Picture
        End If
    End With
End Sub

Public Sub DiagAfterUserResize(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagZY_BeforeUserResize事件，vsDiagXY_BeforeUserResize事件
    If LngCol = DI_诊断描述 And gclsPros.PatiType = PF_门诊 And gclsPros.FuncType = f医生首页 Then
        If gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_中医证候) < gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol) Then
            gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_中医证候) = False
            gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol) - gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_中医证候)
        Else
            gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_中医证候) = True
            gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol)
        End If
    Else
        gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol)
    End If
End Sub

Public Sub DiagBeforeUserResize(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagZY_BeforeUserResize事件，vsDiagXY_BeforeUserResize事件
    If LngCol = DI_增加 Or LngCol = DI_Del Or LngCol < DI_诊断编码 Then blnCancel = True
End Sub

Public Sub DiagCellButtonClick(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagZY_CellButtonClick事件，vsDiagXY_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngCurRow As Long
    Dim bln西医 As Boolean
    
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        Select Case LngCol
            Case DI_诊断描述, DI_诊断编码
                If IIf(bln西医, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0 And gclsPros.FuncType <> f病案首页 Then
                    '按诊断输入:中医部份，一个诊断可能属于多个分类
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(bln西医, "1", "2"), gclsPros.出院科室ID, , True, False)
                Else
                    'B-中医疾病编码，7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(bln西医, decode(Val(.TextMatrix(LngRow, DI_诊断分类)), DT_损伤中毒码, "Y", DT_病理诊断, IIf(gclsPros.M病理, "M", "M,D"), "D"), "B"), gclsPros.出院科室ID, gclsPros.Sex, True, True, , gclsPros.SysNo)
                End If
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, LngRow, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_ICD附码 '只有病案首页可见，可见才会触发事件
                'B-中医疾病编码，7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(.TextMatrix(LngRow, DI_诊断编码) Like "S*" Or .TextMatrix(LngRow, DI_诊断编码) Like "T*", "Y", IIf(.TextMatrix(LngRow, DI_诊断编码) Like "C*" Or (.TextMatrix(LngRow, DI_诊断编码) Like "D*" And Val(Mid(.TextMatrix(LngRow, DI_诊断编码), 2, 2)) <= 48), "M", "D")), gclsPros.出院科室ID, gclsPros.Sex, False, gclsPros.SysNo)
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, LngRow, rsTmp, True)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_中医证候
                If gclsPros.DiagInputZY = 0 Then
                    '按诊断输入:先查是否有对应
                    If Set中医证候(LngRow, Val(.TextMatrix(LngRow, DI_诊断ID))) Then Exit Sub
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "Z", gclsPros.出院科室ID, gclsPros.Sex, True, , , gclsPros.SysNo)
                Else
                    'Z-中医疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "Z", gclsPros.出院科室ID, gclsPros.Sex, True, , , gclsPros.SysNo)
                End If
                If Not rsTmp Is Nothing Then
                    Call Set中医证候(LngRow, 0, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_增加
                If Not .Cell(flexcpPicture, LngRow, DI_增加) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyInsert, 0)
                    Set .CellButtonPicture = Nothing
                End If
            Case DI_Del
                If Not .Cell(flexcpPicture, LngRow, DI_Del) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
                End If
        End Select
    End With
End Sub

Public Sub DiagClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_Click事件，vsDiagZY_Click事件
    Dim bln西医 As Boolean

    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        If (.MouseCol = DI_增加 Or .MouseCol = DI_Del) And .MouseRow >= .FixedRows Then
            If .MouseCol = DI_增加 Then
                If .TextMatrix(.MouseRow, DI_诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = IIf(bln西医, "出院诊断", "主要诊断") Then Exit Sub
            End If
            .Select .MouseRow, .MouseCol
            Call DiagCellButtonClick(vsDiag, .MouseRow, .MouseCol)
        End If
    End With
End Sub

Public Sub DiagComboDropDown(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagXY_ComboDropDown事件，vsDiagZY_ComboDropDown事件
    Dim i As Long

    With vsDiag
        If LngCol = DI_出院情况 Or LngCol = DI_入院病情 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Public Sub DiagDblClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_DblClick事件，vsDiagZY_DblClick事件
    Call DiagKeyPress(vsDiag, vbKeySpace)
End Sub

Public Sub DiagGotFocus(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_GotFocus事件，vsDiagZY_GotFocus事件
    If vsDiag.Row >= vsDiag.FixedRows And vsDiag.Col >= vsDiag.FixedCols Then
        Call DiagAfterRowColChange(vsDiag, -1, -1, vsDiag.Row, vsDiag.Col)
    End If
End Sub

Public Sub DiagKeyDown(ByRef vsDiag As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsDiagXY_KeyDown事件，vsDiagZY_KeyDown事件
    Dim i As Long, j As Long
    Dim dtCurRow As DiagType, LngRow As Long
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long, strMsg As String
    Dim blnDel As Boolean

    On Error GoTo errH
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyCode = vbKeyF4 Then
            If .Col = DI_诊断描述 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If (.TextMatrix(.Row, DI_诊断描述) <> "" Or gclsPros.PatiType = PF_门诊 And .Rows = .FixedRows + 1) And Not (gclsPros.FuncType = f病案首页 And gclsPros.CNIndent) Or gclsPros.FuncType = f病案首页 And gclsPros.CNIndent Then
                If .TextMatrix(.Row, DI_医嘱IDs) <> "" Then
                    strMsg = "该条诊断已经关联了医嘱，不能删除。"
                    lng医嘱ID = Val(Mid(.TextMatrix(.Row, DI_医嘱IDs), 1, InStr(.TextMatrix(.Row, DI_医嘱IDs) & ",", ",") - 1))
                    If lng医嘱ID > 0 Then
                        strSql = "Select 医嘱内容 from 病人医嘱记录 where 病人ID = [1] and 主页ID = [2] and id =[3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "关联医嘱内容查询", gclsPros.病人ID, gclsPros.主页ID, lng医嘱ID)
                        If rsTmp.RecordCount > 0 Then
                            strMsg = "该条诊断已经关联了医嘱:" & rsTmp!医嘱内容 & "，不能删除。"
                        End If
                    End If
                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
                If Not DiagCellEditable(vsDiag, .Row, DI_诊断描述) Then Exit Sub
                '病案首页若定位到附码列，且附码列不为空，则清空附码列，否则删除行,  出院情况同理
                If gclsPros.FuncType = f病案首页 Then
                    Select Case .Col
                        Case DI_ICD附码
                            If .TextMatrix(.Row, DI_ICD附码) <> "" Then
                                .TextMatrix(.Row, DI_ICD附码) = ""
                                .TextMatrix(.Row, DI_附码ID) = ""
                                .Cell(flexcpData, .Row, DI_ICD附码) = ""
                                Exit Sub
                            End If
                        Case DI_出院情况
                            If .TextMatrix(.Row, DI_出院情况) <> "" Then
                                .TextMatrix(.Row, DI_出院情况) = ""
                                Call ChangeOutInfo
                                Exit Sub
                            End If
                        Case DI_中医证候
                            If .TextMatrix(.Row, DI_中医证候) <> "" Then
                                .TextMatrix(.Row, DI_中医证候) = ""
                                .TextMatrix(.Row, DI_证候ID) = ""
                                .Cell(flexcpData, .Row, DI_中医证候) = ""
                                Exit Sub
                            End If
                    End Select
                End If

                blnDel = True
                If gclsPros.FuncType <> f病案首页 Then
                    If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        blnDel = False
                    End If
                End If

                If blnDel Then
                    '删除主/次要诊断后调用外挂接口
                    If gclsPros.FuncType <> f病案首页 Then
                        If CreatePlugInOK(IIf(gclsPros.PatiType = PF_门诊, p门诊医生站, p住院医生站)) Then
                            If Not gobjPlugIn Is Nothing Then
                                On Error Resume Next
                                Call gobjPlugIn.DiagnosisDeleted(gclsPros.SysNo, IIf(gclsPros.PatiType = PF_门诊, p门诊医生站, p住院医生站), gclsPros.病人ID, gclsPros.主页ID, IIf(IIf(vsDiag.Name = "vsDiagXY", gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, Val(.TextMatrix(.Row, DI_诊断ID)), Val(.TextMatrix(.Row, DI_疾病ID))), .TextMatrix(.Row, DI_诊断描述))
                                Call zlPlugInErrH(Err, "DiagnosisDeleted")
                                Err.Clear: On Error GoTo 0
                            End If
                        End If
                    End If
                    dtCurRow = Val(.TextMatrix(.Row, DI_诊断分类))
                     '院内感染，感染部位框体隐藏,西医才有
                    If dtCurRow = DT_院内感染 Then Call ShowInfectInfo(False)
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, DI_诊断分类) = dtCurRow
                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, DI_诊断类型) = "" Or gclsPros.PatiType = PF_门诊 And .Rows <> .FixedRows + 1 Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, DI_诊断类型) = "" Then
                                For j = .FixedCols To .Cols - 1
                                    .TextMatrix(.Row, j) = .TextMatrix(.Row + 1, j)
                                    .Cell(flexcpData, .Row, j) = .Cell(flexcpData, .Row + 1, j)
                                Next
                                .RowData(.Row) = .RowData(.Row + 1)
                                .RemoveItem .Row + 1
                            End If
                        End If
                    End If
                    Call ChangeVSFHeight(vsDiag, True)
                End If
            ElseIf .TextMatrix(.Row, DI_诊断类型) = "" Or gclsPros.PatiType = PF_门诊 And .Rows <> .FixedRows + 1 Then
                .RemoveItem .Row
                Call ChangeVSFHeight(vsDiag, True)
            End If
            '设置诊断相关信息
            If Not (gclsPros.FuncType = f诊断选择 And gclsPros.PatiType = PF_门诊) Then
                Call SetDiagReletedInfo(vsDiag)
                If gclsPros.PatiType <> PF_门诊 Then Call ChangeOutInfo
            End If
        ElseIf intKeyCode = vbKeyInsert Then '新增行
            LngRow = .Row + 1: .AddItem "", LngRow
            Call ChangeVSFHeight(vsDiag, True)
            .TextMatrix(LngRow, DI_诊断分类) = .TextMatrix(LngRow - 1, DI_诊断分类)
            If gclsPros.PatiType = PF_门诊 Then .TextMatrix(LngRow, DI_诊断类型) = .TextMatrix(LngRow - 1, DI_诊断类型)
            .Cell(flexcpData, LngRow, DI_诊断类型) = IIf(.TextMatrix(LngRow - 1, DI_诊断类型) = "", .Cell(flexcpData, LngRow - 1, DI_诊断类型), .TextMatrix(LngRow - 1, DI_诊断类型))
            .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Row = LngRow: .Col = DI_诊断编码
            .ShowCell .Row, .Col
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call DiagKeyPress(vsDiag, intKeyCode)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DiagKeyPress(ByRef vsDiag As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPress事件，vsDiagZY_KeyPress事件
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellDiag(vsDiag)
        Else
            If Not DiagCellEditable(vsDiag, .Row, .Col) Then Exit Sub
            Select Case .Col
                Case DI_是否未治, DI_是否疑诊 '中医这两列隐藏
                    If intKeyAscii <> vbKeySpace Then Exit Sub
                    intKeyAscii = 0
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", IIf(.Col = DI_是否疑诊, "？", "√"), "")
                Case DI_诊断编码, DI_诊断描述, DI_中医证候, DI_ICD附码 '西医中医证候隐藏,中医无ICD附码隐藏
                    If intKeyAscii = Asc("*") Then
                        intKeyAscii = 0
                        Call DiagCellButtonClick(vsDiag, .Row, .Col)
                    Else
                        .ComboList = "" '使按钮状态进入输入状态
                    End If
            End Select
        End If
    End With
End Sub

Public Sub DiagKeyPressEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPressEdit事件，vsDiagZY_KeyPressEdit事件
    Dim bln西医 As Boolean

    If intKeyAscii = 13 Then
        gclsPros.IsReturn = True
        With vsDiag
            bln西医 = .Name = "vsDiagXY"
            If LngCol = DI_出院情况 Or LngCol = DI_入院病情 Then
                intKeyAscii = 0
                If .ComboIndex <> -1 Then
                    '此时.TextMatrix尚未更新,所以取ComboItem
                    .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.ComboItem(.ComboIndex))
                    If bln西医 And LngCol = DI_出院情况 Then
                        If Not DiagCellEditable(vsDiag, LngRow, DI_是否未治) Then .TextMatrix(LngRow, DI_是否未治) = ""
                    End If
                    Call EnterNextCellDiag(vsDiag)
                 End If
            End If
        End With
    Else
        gclsPros.IsReturn = False
    End If
End Sub


Public Sub DiagSetupEditWindow(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsDiagXY_SetupEditWindow事件，vsDiagZY_SetupEditWindow事件
    With vsDiag
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub DiagStartEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_StartEdit事件，vsDiagZY_StartEdit事件
    If gclsPros.FuncType = f诊断选择 And gclsPros.IsSigned And LngCol <> DI_关联 Then
        blnCancel = True
        MsgBox "该病人的首页已经签名不能修改诊断。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not DiagCellEditable(vsDiag, LngRow, LngCol) Then
        blnCancel = True
    ElseIf LngCol = DI_是否未治 Or LngCol = DI_是否疑诊 Then '西医才可能进入该分支
        blnCancel = True '不直接编辑
    End If
End Sub

Public Sub DiagValidateEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_ValidateEdit事件，vsDiagZY_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim int诊断输入 As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim bln西医 As Boolean
    Dim str性别 As String

    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        Select Case LngCol
            Case DI_诊断描述, DI_诊断编码
                If bln西医 Then
                    strDiagType = decode(Val(.TextMatrix(LngRow, DI_诊断分类)), 7, "'Y'", 6, IIf(gclsPros.M病理, "'M'", "'M,D'"), "'D'")
                Else
                    strDiagType = IIf(gclsPros.DiagInputZY = 0, "", "B")
                End If
                If gclsPros.FuncType = f病案首页 Then
                     If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        If LngCol = DI_诊断描述 Or LngCol = DI_诊断编码 And (.TextMatrix(LngRow, DI_诊断描述) = "" Or gclsPros.DaigFree And Not bln西医) Then
                            .EditText = ""
                        Else
                            .EditText = .Cell(flexcpData, LngRow, LngCol)
                        End If
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                    '编码名称独立时，若编辑诊断名称时，若编码不为空，则当作自由录入，其他当作匹配
                    ElseIf Not (LngCol = DI_诊断描述 And .TextMatrix(LngRow, DI_诊断编码) <> "" And gclsPros.CNIndent) Then
                        If Val(.TextMatrix(LngRow, DI_诊断分类)) = IIf(bln西医, DT_门诊诊断XY, DT_门诊诊断ZY) Then
                            int诊断输入 = gclsPros.DiagSourceMZ
                        Else
                            int诊断输入 = gclsPros.DiagSourceZY
                        End If
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别, strDiagType)
                        vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln西医, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "疾病诊断", "疾病编码"), _
                            False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                        If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                            If Not (gclsPros.DaigFree And Not bln西医 And LngCol = DI_诊断描述) Then
                                blnCancel = True
                            End If
                        Else
                            '检查诊断输入方式
                            If rsTmp Is Nothing Then
                                If Not (gclsPros.DaigFree And Not bln西医 And LngCol = DI_诊断描述) Then
                                    MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                    blnCancel = True
                                End If
                            Else
                                Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call XYEnterNextCell    '暂不跳到下一行，因为可能还要改描述内容
                            End If
                        End If
                    End If
                Else
                    If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        .EditText = ""
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                    ElseIf .TextMatrix(LngRow, DI_诊断编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, LngRow, LngCol) & "*" Then
                        '判断加了前缀后的名称是否存在其他的诊断编码
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别, strDiagType)
                        On Error GoTo errH
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput, strInput, strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID)
                        If rsTmp.RecordCount = 1 Then
                            Call SetDiagInput(vsDiag, LngRow, rsTmp)
                            .EditText = .Text
                        Else
                            '允许在标准的名称前后输入附加信息
                            '不处理.Cell(flexcpData, lngRow, lngCol)，以便修改内容时再次使用like判断
                            .TextMatrix(LngRow, DI_诊断描述) = .EditText
                        End If
                    ElseIf .TextMatrix(LngRow, DI_诊断编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别, strDiagType)
                        On Error GoTo errH
                        vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln西医, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "疾病诊断", "疾病编码"), _
                            False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                        If blnInputCancel Then
                            blnCancel = True
                        Else
                            If rsTmp Is Nothing Then
                                .TextMatrix(LngRow, DI_诊断描述) = .EditText
                            Else
                                 Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                            End If
                        End If
                    Else
                        If Val(.TextMatrix(LngRow, DI_诊断分类)) = IIf(bln西医, DT_门诊诊断XY, DT_门诊诊断ZY) Then
                            int诊断输入 = gclsPros.DiagSourceMZ
                        Else
                            int诊断输入 = gclsPros.DiagSourceZY
                        End If
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别, strDiagType)
                        If False And int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                            '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID)
                            If rsTmp.EOF Then
                                Set rsTmp = Nothing
                            ElseIf rsTmp.RecordCount > 1 Then
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput, strInput, strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID)
                                If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                            End If
                            Call SetDiagInput(vsDiag, LngRow, rsTmp)
                            .EditText = .Text
                            If gclsPros.IsReturn And rsTmp Is Nothing Then Call EnterNextCellDiag(vsDiag) '不是自由录入时，暂不跳到下一行，因为可能还要改描述内容
                        Else
                            vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln西医, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "疾病诊断", "疾病编码"), _
                                False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                                strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                            If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                                blnCancel = True
                            Else
                                '检查诊断输入方式
                                If rsTmp Is Nothing And ((int诊断输入 = 2 Or int诊断输入 = 3 And gclsPros.InsureType <> 0)) Then
                                    MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                    blnCancel = True
                                Else
                                    Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                                    'If mblnReturn Then Call XYEnterNextCell    '暂不跳到下一行，因为可能还要改描述内容
                                End If
                            End If
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_中医证候
                If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    .EditText = ""
                    .Cell(flexcpData, LngRow, LngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                ElseIf .TextMatrix(LngRow, DI_诊断编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                    strDiagType = "Z"
                    strInput = UCase(.EditText)
                    strSql = GetMedInputSQL(1, strInput, str性别, strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then      '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(LngRow, DI_中医证候) = .EditText
                        Else
                            Call Set中医证候(LngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                Else
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSql = GetMedInputSQL(1, strInput, str性别, strDiagType)

                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        If Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_门诊诊断ZY Then
                            int诊断输入 = gclsPros.DiagSourceMZ
                        Else
                            int诊断输入 = gclsPros.DiagSourceZY
                        End If
                        '检查诊断输入方式
                         If rsTmp Is Nothing And (int诊断输入 = 2 Or (int诊断输入 = 3 And gclsPros.InsureType <> 0)) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                         Else
                            Call Set中医证候(LngRow, 0, rsTmp, rsTmp Is Nothing)
                         End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_ICD附码
                If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    .EditText = ""
                    .TextMatrix(LngRow, DI_附码ID) = ""
                    .Cell(flexcpData, LngRow, LngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                Else
                    strInput = UCase(.EditText)
                    'B-中医疾病编码，7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    strDiagType = IIf(.TextMatrix(LngRow, DI_诊断编码) Like "S*" Or .TextMatrix(LngRow, DI_诊断编码) Like "T*", "Y", IIf(.TextMatrix(LngRow, DI_诊断编码) Like "C*" Or (.TextMatrix(LngRow, DI_诊断编码) Like "D*" And Val(Mid(.TextMatrix(LngRow, DI_诊断编码), 2, 2)) <= 48), "M", "D"))
                    strSql = GetMedInputSQL(0, strInput, str性别)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln西医, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str性别, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.出院科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                        Else
                            Call SetDiagInput(vsDiag, LngRow, rsTmp, True): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_出院情况
                If .EditText <> "" Then
                    If .TextMatrix(.Row, DI_疗效限制) <> "" And InStr(.EditText, .TextMatrix(.Row, DI_疗效限制)) > 0 Then
                        MsgBox "请注意，该疾病通常不能达到这种疗效的。", vbInformation, gstrSysName
                        .EditText = "": blnCancel = True: Exit Sub
                    End If
                End If
            Case DI_发病时间
                If .EditText <> "" Then
                    strInput = zlStr.FullDate(.EditText)
                    If IsDate(strInput) Then
                        .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    Else
                        MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。", vbInformation, gstrSysName
                        blnCancel = True
                    End If
                End If
                If LngRow = .FixedRows And gclsPros.FuncType <> f诊断选择 Then
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病时间), IsDate(.EditText), True)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病日期), IsDate(.EditText), True)
                End If
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsOPS事件
Public Sub OPSAfterEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_AfterEdit事件
    Dim strInput As String

    With vsOPS
        Select Case LngCol
            Case PI_手术日期, PI_结束日期, PI_抗菌用药时间, PI_麻醉开始时间
                If LngCol <> PI_抗菌用药时间 Then
                    strInput = Format(zlStr.FullDate(.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime), "yyyy-mm-dd hh:mm")
                Else
                    strInput = Format(zlStr.FullDate(.TextMatrix(LngRow, LngCol)), "yyyy-mm-dd hh:mm")
                End If
                If Not IsDate(strInput) Then
                    .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                Else
                    .TextMatrix(LngRow, LngCol) = strInput
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                End If
            Case PI_切口愈合, PI_麻醉类型
                .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
        End Select
'        Call OPSAfterRowColChange(vsOPS, -1, -1, LngRow, LngCol)
        '设置诊断符合情况
        Call SetDiagMatchInfo(BCC_术前与术后)
    End With
End Sub

Public Sub OPSAfterRowColChange(ByRef vsOPS As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsOPS_AfterRowColChange事件
    Dim blnEdit As Boolean

    With vsOPS
        If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
        If vsOPS.Editable <> flexEDNone Then Call SetCopyImage(vsOPS)
        If Not OPSCellEditable(lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Select Case lngNewCol
                Case PI_手术编码, PI_主刀医师, PI_助产护士, PI_助手1, PI_助手2, PI_麻醉方式, PI_麻醉医师, PI_切口部位 'PI_麻醉方式为隐藏列
                    .ComboList = "..."
                Case PI_手术名称
                    If gclsPros.FuncType <> f病案首页 Then
                        '手术名称不能输入
                        blnEdit = gclsPros.CurrentForm.chkParaOPSInfo(PC_未找到时自由录入).Value
                    Else
                        blnEdit = gclsPros.CNIndent
                    End If
                    If blnEdit Then
                        .ComboList = "..."
                    Else
                        .ComboList = ""
                    End If
                Case PI_手术情况
                    If Not gclsPros.IsCheckData Then OS.PressKey vbKeySpace
                Case PI_切口愈合, PI_麻醉类型
                    .ComboList = .ColData(lngNewCol)
                Case Else
                    .ComboList = ""
            End Select
        End If
    End With
End Sub

Public Sub OPSBeforeUserResize(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_BeforeUserResize事件
    blnCancel = LngCol = PI_Copy
End Sub

Public Sub OPSCellButtonClick(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int性别 As Integer, int手术输入方式 As Integer
    Dim vPoint As POINTAPI, strInfoName As String
    Dim textTmp As TextBox

    With vsOPS
        Select Case LngCol
            Case PI_手术编码, PI_手术名称
                If gclsPros.Sex Like "*男*" Then
                    int性别 = 1
                ElseIf gclsPros.Sex Like "*女*" Then
                    int性别 = 2
                End If
                int手术输入方式 = Val(gclsPros.OPSInput)
                If int手术输入方式 = 0 And gclsPros.FuncType <> f病案首页 Then
                    '按诊疗项目输入
                    strSql = "Select 0 as 末级,ID,上级ID,编码,名称,NULL as 规模" & _
                        " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                        " Union ALL " & _
                        " Select 1 as 末级,ID,分类ID as 上级ID,编码,名称,操作类型 as 规模" & _
                        " From 诊疗项目目录" & _
                        " Where 类别='F' And 服务对象 IN(2,3) And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
                        IIf(int性别 <> 0, " And Nvl(适用性别,0) IN(0,[2])", "") & _
                        " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)"
                Else
                    '按ICD9-CM3输入
                    strSql = _
                        " Select 0 as 末级,ID,上级ID," & _
                        " 类别||LPAD(序号,3,'0') as 编码," & _
                        " NULL as 附码,名称,简码,NULL as 说明" & _
                        " From 疾病编码分类 Where 类别='S'" & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                        " Union ALL " & _
                        " Select 1 as 末级,ID,分类ID as 上级ID,编码,附码,名称,简码,说明" & _
                        " From 疾病编码目录 Where 类别='S'" & _
                        IIf(int性别 <> 0, " And (性别限制=[1] Or 性别限制 is NULL)", "") & _
                        " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
                End If
                strInfoName = IIf(int手术输入方式 = 0 And gclsPros.FuncType <> f病案首页, "手术项目", "手术编码")
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, strInfoName, False, "", "", False, True, False, _
                            0, 0, 0, blnCancel, False, False, decode(int性别, 1, "男", 2, "女", ""), int性别)
            Case PI_麻醉方式 '该列为隐藏列
                If gclsPros.FuncType <> f病案首页 Then
                    strSql = "Select 0 as 末级,ID,上级ID,编码,名称,NULL as 麻醉类型" & _
                    " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union ALL " & _
                    " Select 1 as 末级,ID,分类ID as 上级ID,编码,名称,操作类型 as 麻醉类型" & _
                    " From 诊疗项目目录 Where 类别='G'" & _
                    " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                    " And (站点='" & gstrNodeNo & "' Or 站点 is Null)"
                    strInfoName = "麻醉项目"
                    Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, 2, strInfoName, , , , , True, , , , , blnCancel)
                End If
            Case PI_主刀医师, PI_助手1, PI_助手2, PI_麻醉医师, PI_助产护士
                strInfoName = IIf(LngCol = PI_助产护士, "护士", "医生")
                Set rsTmp = GetManData(strInfoName)
                Set textTmp = GetReplaceObject(vsOPS)
                blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "缺省,医生,护士,病案编码员,门诊,住院,管理,技术", rsTmp)
            Case PI_切口部位
                strSql = "Select Rownum As ID, A.编码, A.名称, A.简码 From 切口部位 A"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "首页读取切口部位")
                Set textTmp = GetReplaceObject(vsOPS)
                blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "切口部位", rsTmp)
        End Select
        '项目输入控制
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有" & strInfoName & "可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp)
            Call EnterNextCellOPS(vsOPS)
        End If
    End With
End Sub

Public Sub OPSComboDropDown(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_ComboDropDown事件
    Dim i As Long

    With vsOPS
        If LngCol = PI_切口愈合 Or LngCol = PI_麻醉类型 Or LngCol = PI_手术情况 Then
            For i = 0 To .ComboCount - 1
                If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Public Sub OPSDblClick(ByRef vsOPS As VSFlexGrid)
'vsOPS_DblClick事件
    Call OPSKeyPress(vsOPS, vbKeySpace)
End Sub

Public Sub OPSKeyDown(ByRef vsOPS As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsOPS_KeyDown事件
    Dim i As Long

    If vsOPS.Editable = flexEDNone Then Exit Sub
    With vsOPS
        If intKeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, PI_手术名称) <> "" Then
                If MsgBox("确实要删除该行手术吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call ChangeVSFHeight(vsOPS, True, 600, 2)
                    '设置诊断符合情况
                    Call SetDiagMatchInfo(BCC_术前与术后)
                End If
            End If
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call OPSKeyPress(vsOPS, intKeyCode)
        End If
    End With
End Sub

Public Sub OPSKeyPress(ByRef vsOPS As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsOPS_KeyPress事件
    If vsOPS.Editable = flexEDNone Then Exit Sub
    With vsOPS
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellOPS(vsOPS)
        Else
            If .ComboList = "..." Then
                If intKeyAscii = Asc("*") Then
                    intKeyAscii = 0
                    Call OPSCellButtonClick(vsOPS, .Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Public Sub OPSKeyPressEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsOPS_KeyPressEdit事件
    Dim strInput As String

    With vsOPS
        If intKeyAscii = vbKeyReturn Then
            gclsPros.IsReturn = True
            Select Case LngCol
                Case PI_切口愈合, PI_麻醉类型, PI_手术情况
                    intKeyAscii = 0
                    If .ComboIndex <> -1 Then
                        .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.ComboItem(.ComboIndex))
                        Call EnterNextCellOPS(vsOPS)
                    End If
            End Select
        Else
            gclsPros.IsReturn = False
            If LngCol = PI_手术日期 Or LngCol = PI_结束日期 Or LngCol = PI_麻醉开始时间 Or LngCol = PI_抗菌用药时间 Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            ElseIf LngCol = PI_准备天数 Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            ElseIf LngCol = PI_抗菌药天数 Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Public Sub OPSSetupEditWindow(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsOPS_SetupEditWindow事件
    With vsOPS
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub OPSStartEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_StartEdit事件
    If Not OPSCellEditable(LngRow, LngCol) Then
        blnCancel = True
    ElseIf LngCol = PI_切口部位 Or LngCol = PI_重返手术室目的 Then
        vsOPS.EditMaxLength = 100
    End If
End Sub

Public Sub OPSValidateEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim str性别 As String, int性别 As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim textTmp As TextBox
    Dim strTmp As String

    On Error GoTo errH
    With vsOPS
        Select Case LngCol
            Case PI_手术编码, PI_手术名称
                If gclsPros.FuncType = f病案首页 Then
                    If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        If LngCol = PI_手术名称 Or (LngCol = PI_手术编码 And .TextMatrix(LngRow, PI_手术名称) = "") Then
                            .EditText = ""
                        Else
                            .EditText = .Cell(flexcpData, LngRow, LngCol)
                        End If
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        '编码名称独立时，若编辑手术名称时，若编码不为空，则当作自由录入，其他当作匹配
                    ElseIf Not (LngCol = PI_手术名称 And .TextMatrix(LngRow, PI_手术编码) <> "" And gclsPros.CNIndent) Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str性别)
                        If str性别 = "男" Then
                            int性别 = 1
                        ElseIf str性别 = "女" Then
                            int性别 = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "手术项目", "手术编码"), False, "", "", False, True, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str性别, int性别)
                        If rsTmp Is Nothing Then
                            If Not blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                                MsgBox "没有找到您查找的手术项目。", vbInformation, gstrSysName
                                blnCancel = True
                            Else
                                blnCancel = True
                            End If
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        End If
                    End If
                Else
                    If .EditText = "" Then
                        .EditText = .Cell(flexcpData, LngRow, LngCol)
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    ElseIf LngCol = PI_手术名称 And .TextMatrix(LngRow, PI_手术编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, LngRow, LngCol) & "*" Then
                        '判断加了前缀后的名称是否存在其他的诊断编码
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str性别)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", str性别, int性别)
                        If rsTmp.RecordCount <> 1 Then
                            '允许在标准的名称前后输入附加信息
                            .TextMatrix(LngRow, PI_手术名称) = .EditText
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp)
                            .EditText = .Text '不处理.Cell(flexcpData, lngRow, lngCol)，以便修改内容时再次使用like判断
                        End If
                    ElseIf LngCol = PI_手术名称 And .TextMatrix(LngRow, PI_手术编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str性别)
                        If str性别 = "男" Then
                            int性别 = 1
                        ElseIf str性别 = "女" Then
                            int性别 = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "手术项目", "手术编码"), False, "", "", False, True, True, _
                                    vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str性别, int性别)
                        If blnInputCancel Then
                            blnCancel = True
                        Else
                            If rsTmp Is Nothing Then
                                .TextMatrix(LngRow, PI_手术名称) = .EditText
                            Else
                                Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            End If
                        End If
                    Else
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str性别)
                        If str性别 = "男" Then
                            int性别 = 1
                        ElseIf str性别 = "女" Then
                            int性别 = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "手术项目", "手术编码"), False, "", "", False, True, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str性别, int性别)
                        If rsTmp Is Nothing Then
                            If Not blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                                If Not gclsPros.OPSFree Or LngCol = PI_手术编码 Then
                                    MsgBox "没有找到您查找的手术项目。", vbInformation, gstrSysName
                                    blnCancel = True
                                Else
                                    .TextMatrix(LngRow, PI_手术编码) = ""
                                    .TextMatrix(LngRow, PI_诊疗项目ID) = ""
                                    .Cell(flexcpData, LngRow, PI_手术编码) = ""
                                    .TextMatrix(LngRow, PI_手术操作ID) = ""
                                    '输入后始终保持一新行
                                    If LngRow = .Rows - 1 Then .AddItem "": Call ChangeVSFHeight(vsOPS, True)
                                End If
                            Else
                                blnCancel = True
                            End If
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_手术日期, PI_结束日期, PI_麻醉开始时间, PI_抗菌用药时间
                If LngCol <> PI_抗菌用药时间 Then
                    strInput = Format(zlStr.FullDate(.EditText, , gclsPros.InTime, gclsPros.OutTime), "yyyy-mm-dd hh:mm")
                Else
                    strInput = Format(zlStr.FullDate(.EditText), "yyyy-mm-dd hh:mm")
                End If
                If IsDate(strInput) Then
                    '抗菌用药可能是在院外使用，因此不做检查
                    If Not CheckDateRange(strInput) And LngCol <> PI_抗菌用药时间 Then
                        MsgBox "您输入的时间必须在病人的住院期间。", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                        Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                If strInput = "" And (LngCol = PI_麻醉开始时间 Or LngCol = PI_抗菌用药时间) Then
                    .TextMatrix(LngRow, LngCol) = strInput
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    Call EnterNextCellOPS(vsOPS)
                End If
            Case PI_麻醉方式
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = _
                        " Select A.ID,A.编码,A.名称,A.操作类型 as 麻醉类型" & _
                        " From 诊疗项目目录 A,诊疗项目别名 B" & _
                        " Where A.类别='G' And A.ID=B.诊疗项目ID" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.编码 Like [1] Or A.名称 Like [2] Or B.简码 Like [2] Or B.名称 Like [2])" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"

                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "麻醉项目", False, "", "", False, True, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%")
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "没有找到匹配的麻醉项目！", vbInformation, gstrSysName
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_主刀医师, PI_助手1, PI_助手2, PI_麻醉医师
                If (LngCol = PI_助手1 Or LngCol = PI_助手2) And .EditText = "" Then
                    .TextMatrix(LngRow, LngCol) = "": .Cell(flexcpData, LngRow, LngCol) = ""
                    If LngCol = PI_助手1 Then
                        .TextMatrix(LngRow, PI_助手2) = "": .Cell(flexcpData, LngRow, PI_助手2) = ""
                    End If
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = "编码 Like '" & strInput & "*' OR 名称 Like '*" & strInput & "*' OR 简码 Like '*" & strInput & "*' OR 五笔简码 Like '*" & strInput & "*'"
                    Set rsTmp = Rec.FilterNew(GetManData("医生"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "缺省,医生,护士,病案编码员,门诊,住院,管理,技术", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            If (LngCol = PI_主刀医师 Or LngCol = PI_助手1 Or LngCol = PI_助手2 Or LngCol = PI_麻醉医师) And zlCommFun.IsCharChinese(.EditText) And Not gclsPros.IsOutDocCtrl Then
                                If MsgBox("没有找到匹配的本院医生，是否录入未在本院建档的医生？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                                    Exit Sub
                                End If
                            Else
                                MsgBox "没有找到匹配的医生！", vbInformation, gstrSysName
                            End If
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_助产护士
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = "编码 Like '" & strInput & "*' OR 名称 Like '*" & strInput & "*' OR 简码 Like '*" & strInput & "*' OR 五笔简码 Like '*" & strInput & "*'"
                    Set rsTmp = Rec.FilterNew(GetManData("护士"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "缺省,医生,护士,病案编码员,门诊,住院,管理,技术", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "没有找到匹配的护士！", vbInformation, gstrSysName
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_切口部位
                If .EditText <> "" And .EditText <> .Cell(flexcpData, LngRow, LngCol) Then
                    strInput = UCase(.EditText)
                    strSql = "编码 Like '" & strInput & "*' OR 名称 Like '*" & strInput & "*' OR 简码 Like '*" & strInput & "*'"
                    strTmp = "Select Rownum As ID, A.编码, A.名称, A.简码 From 切口部位 A"
                    Set rsTmp = Rec.FilterNew(zlDatabase.OpenSQLRecord(strTmp, "首页读取切口部位"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "切口部位", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If Not rsTmp Is Nothing Then
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                    End If
                End If
                gclsPros.IsReturn = False
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub picCopyClick()
'picCopy_Click 事件
    Dim vsOPS As VSFlexGrid
    Dim i As Long, LngRow As Long
    Set vsOPS = gclsPros.CurrentForm.vsOPS

    With vsOPS
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, PI_手术名称) = "" Then
                LngRow = i: Exit For
            End If
        Next
        If LngRow = 0 Then
            .Rows = .Rows + 1
            LngRow = .Rows - 1
            Call ChangeVSFHeight(vsOPS, True)
        End If
        For i = .FixedCols To .Cols - 1
            If i <> PI_手术名称 And i <> PI_手术编码 Then
                .TextMatrix(LngRow, i) = .TextMatrix(.Row, i)
            End If
        Next
    End With
End Sub

Private Sub SetCopyImage(ByRef vsOPS As VSFlexGrid)
    Dim blnShow As Boolean
    Dim lngRowHeight As Long, i As Long, lngHeight As Long
    With vsOPS
        blnShow = .TextMatrix(.Row, PI_手术名称) <> "" And .Row >= .FixedRows And .ColIsVisible(PI_Copy)
        If blnShow Then
            For i = 0 To .Row - 1
                lngRowHeight = .RowHeight(i)
                If .RowHeightMin <> 0 Then
                    If lngRowHeight < .RowHeightMin Then
                        lngRowHeight = .RowHeightMin
                    End If
                End If
                If .RowHeightMax <> 0 Then
                    If lngRowHeight > .RowHeightMax Then
                        lngRowHeight = .RowHeightMax
                    End If
                End If
                lngHeight = lngHeight + lngRowHeight
            Next
            gclsPros.CurrentForm.picCopy.Left = 0
            gclsPros.CurrentForm.picCopy.Top = lngHeight
        End If
        gclsPros.CurrentForm.picCopy.Visible = blnShow
        gclsPros.CurrentForm.picCopy.Enabled = blnShow
        gclsPros.CurrentForm.picCopy.ZOrder
    End With
End Sub

'vsKSS事件
Public Sub KSSAfterEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsKSS_AfterEdit事件
    With vsKSS
        Call .Select(.Row, .Col)
    End With
End Sub

Public Sub KSSAfterRowColChange(ByRef vsKSS As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsKSS_AfterRowColChange事件
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If lngNewCol = KI_抗菌药物名 Then
        vsKSS.ColComboList(KI_抗菌药物名) = "..."
    End If
End Sub

Public Sub KSSCellButtonClick(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsKSS_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String
    Dim bln共享 As Boolean
    Dim vPoint As POINTAPI

    With vsKSS
        If LngCol = KI_抗菌药物名 Then
'            If gclsPros.ReadPages Then
            If gclsPros.ShareMedRec Or gclsPros.FuncType = f医生首页 Then
                bln共享 = True
                strSQLItem = _
                    " From 诊疗项目目录 A,药品特性 B" & _
                    " Where A.ID=B.药名ID And A.类别='5' And A.服务对象 IN(2,3) And Nvl(b.抗生素, 0) <> 0" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) "
                strSql = "Select 0 as 末级,Max(Level) as 级ID,ID,上级ID,编码,名称,NULL as 单位" & _
                    " From 诊疗分类目录 Where 类型=1 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With ID In (Select A.分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID" & _
                    " Group by ID,上级ID,编码,名称"
                strSql = strSql & " Union ALL" & _
                    " Select 1 as 末级,1 as 级ID,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位" & _
                    strSQLItem & " Order By 末级,级ID Desc,编码"
            Else
                strSql = "Select Rownum As ID, A.编码, A.名称, A.简码" & vbNewLine & _
                    "From 抗生素药 A"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, IIf(bln共享, 2, 0), "抗菌药物", False, "", "", False, True, False, vPoint.X, vPoint.Y, IIf(bln共享, 0, .CellHeight), blnCancel, False, Not bln共享)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有抗菌药物数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                Call KSSEnterNextCell(vsKSS)
            End If
        End If
    End With
End Sub

Public Sub KSSKeyDown(ByRef vsKSS As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsKSS_KeyDown事件
    If vsKSS.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf intKeyCode = vbKeyDelete Then
        If MsgBox("确实要删除该行内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsKSS
                .RemoveItem .Row
                If .Rows < 4 Then .Rows = 4: Call ChangeVSFHeight(vsKSS, True)
                Call SetKSSSerial
            End With
        End If
    ElseIf intKeyCode > 127 Then
        '解决直接输入汉字的问题
        Call KSSKeyPress(vsKSS, intKeyCode)
    End If
End Sub

Public Sub KSSKeyPress(ByRef vsKSS As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsKSS_KeyPress事件
    With vsKSS
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call KSSEnterNextCell(vsKSS)
        ElseIf .Editable <> flexEDNone Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call KSSCellButtonClick(vsKSS, .Row, .Col)
            Else
                .ColComboList(KI_抗菌药物名) = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Public Sub KSSKeyPressEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsKSS_KeyPressEdit事件
    gclsPros.IsReturn = intKeyAscii = vbKeyReturn
    With vsKSS
        If LngCol = KI_使用天数 Then
            If .EditSelLength <> 0 Then Exit Sub
            If intKeyAscii = vbKeyBack Then Exit Sub
            If Len(.EditText) > 18 Then intKeyAscii = 0
        ElseIf LngCol = KI_用药目的 Then
            If .EditSelLength <> 0 Then Exit Sub
            If intKeyAscii = vbKeyBack Then Exit Sub
            If LenB(StrConv(.EditText, vbFromUnicode)) >= 200 Then intKeyAscii = 0
        End If
    End With
End Sub

Public Sub KSSSetupEditWindow(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsKSS_SetupEditWindow事件
    With vsKSS
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub KSSValidateEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsKSS_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsKSS
        If LngCol = KI_抗菌药物名 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, LngRow, LngCol)
                If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
            ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
            Else
                strInput = UCase(.EditText)
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If

                If gclsPros.ShareMedRec Or gclsPros.FuncType = f医生首页 Then
                    strSql = _
                        " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位" & _
                        " From 诊疗项目目录 A,诊疗项目别名 B,药品特性 C" & _
                        " Where A.ID=B.诊疗项目ID And A.ID=C.药名ID And Nvl(c.抗生素, 0) <> 0" & _
                        " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And A.类别='5' And A.服务对象 IN(2,3) And B.码类=[3] And (" & strSql & ")" & _
                        " Order by A.编码"
                Else
                    strSql = "Select Rownum As ID, A.编码, A.名称, A.简码" & vbNewLine & _
                        "From 抗生素药 A" & vbNewLine & _
                        "Where " & strSql & vbNewLine & _
                        "Order By 编码"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                    '判断是否有数据
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "没有找到指定的抗菌药物。", vbInformation, gstrSysName
                        blnCancel = True: .EditText = "": Exit Sub
                    End If
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                    End If
                    Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                    .EditText = .Text
                    If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "抗菌药物", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        '判断是否有数据
                        If rsTmp Is Nothing Then
                            MsgBox "没有找到指定的抗菌药物。", vbInformation, gstrSysName
                            blnCancel = True: .EditText = "": Exit Sub
                        End If
                        Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                        .EditText = .Text
                        If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
                    End If
                End If
            End If
            gclsPros.IsReturn = False
        ElseIf LngCol = KI_使用天数 Or LngCol = KI_DDD数 Then
            If (Not IsNumeric(.EditText) Or InStr(.EditText, "-") > 0 Or InStr(.EditText, "+") > 0) And .EditText <> "" Then
                MsgBox "请输入有效的数字。", vbInformation, gstrSysName
                blnCancel = True
            Else
                If Len(.EditText) > 12 Then
                    MsgBox "请输入12位以下的数字。", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            End If
        ElseIf LngCol = KI_使用阶段 Then
            '如果用户修改了，则提取的时候不影响这一项
            If .Cell(flexcpData, LngRow, LngCol) = "新增" Then .Cell(flexcpData, LngRow, LngCol) = ""
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsSpirit事件
Public Sub SpiritAfterRowColChange(ByRef vsSpirit As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsSpirit_AfterRowColChange事件
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsSpirit.FocusRect = flexFocusSolid
End Sub

Public Sub SpiritKeyDown(ByRef vsSpirit As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsSpirit_KeyDown事件
    Dim LngCol As Long
    If vsSpirit.Editable = flexEDNone Then Exit Sub
    With vsSpirit
        If intKeyCode = vbKeyDelete Then
            If MsgBox("你是否真的要删除该行的精神药品信息吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsSpirit, True)
            End If
            zlControl.ControlSetFocus vsSpirit, True
        ElseIf intKeyCode <> vbKeyReturn Then
            Exit Sub
        Else
            If .Row = .Rows - 1 Then
                If Trim(.TextMatrix(.Row, SI_药物名称)) = "" Then Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
        End If
    End With
End Sub

Public Sub SpiritKeyDownEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsSpirit_KeyDownEdit事件
    Dim strInput As String, blnCancel As Boolean
    Dim rsTemp As Recordset, strSql As String
    Dim vPoint  As POINTAPI

    If intKeyCode <> vbKeyReturn Then Exit Sub
    intKeyCode = 0
    With vsSpirit
        If LngCol = SI_药物名称 Then
            '多选
            .Cell(flexcpData, LngRow, LngCol) = ""
            strInput = UCase(Trim(.EditText))
            If strInput = "" Then Exit Sub
            strInput = gclsPros.LikeString & strInput & "%"
            strSql = "" & _
                "   SELECT S.药品id as ID, I.编码, I.名称, I.规格, I.产地, I.计算单位 AS 售价单位,S.批准文号, S.标识码,S.GMP认证,I.建档时间, I.撤档时间 " & _
                "   FROM 收费项目目录 I, 药品规格  S,药品特性 J  " & _
                "   WHERE I.ID=S.药品id   and I.类别 In ('5','6','7') And s.药名id=J.药名ID  And J.毒理分类 In ('精神I类','精神II类') " & _
                "           And (i.撤档时间 Is Null Or to_char(i.撤档时间,'yyyy-mm-dd')='3000-01-01') " & _
                "           And (i.编码 like [1]  Or i.名称 Like [1] Or   Exists(Select 1 From 收费项目别名 Where i.Id=收费细目id And 码类 = 3))"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTemp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "药品选择器", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, False, strInput)
            If blnCancel = True Then
                Exit Sub
            End If
            If rsTemp Is Nothing Then
                MsgBox "没有找到您查找的抗精神病药。", vbInformation, gstrSysName
                Exit Sub
            Else
                .EditText = NVL(rsTemp!名称)
                .TextMatrix(.Row, SI_药物名称) = NVL(rsTemp!名称)
                .Cell(flexcpData, .Row, SI_药物名称) = NVL(rsTemp!ID)
            End If
            If .TextMatrix(.Rows - 1, SI_药物名称) <> "" Then
                .Rows = .Rows + 1
                Call ChangeVSFHeight(vsSpirit, True)
            End If
        End If
        Call EnterNextCellSpirit
    End With
End Sub

Public Sub SpiritKeyPress(ByRef vsSpirit As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsSpirit_KeyPress事件
    If vsSpirit.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call EnterNextCellSpirit
    End If
End Sub

Public Sub SpiritStartEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSpirit_StartEdit事件
    If Trim(vsSpirit.TextMatrix(LngRow, SI_药物名称)) = "" And LngCol <> SI_药物名称 Then
        blnCancel = True
    End If
End Sub

Public Sub SpiritValidateEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSpirit_ValidateEdit事件
    Dim intMax As Integer
    If LngCol >= SI_药物名称 And LngCol <= SI_疗效 Then
        intMax = 50
        If LngCol = SI_药物名称 Then
            intMax = 200
        ElseIf LngCol = SI_特殊反应 Then
            intMax = 100
        End If
        If LenB(StrConv(vsSpirit.EditText, vbFromUnicode)) > intMax Then
            MsgBox "您输入的内容不能超过" & intMax \ 2 & "个汉字。", vbInformation, gstrSysName
            blnCancel = True
            Exit Sub
        End If
    End If
End Sub

Public Sub TxtGotFocus(ByRef objTextBox As Object, Optional ByVal blnChineseIn As Boolean, Optional ByVal blnSetInfect As Boolean)
'功能：实现控件获得焦点全选控件内容的功能
'参数：objTextBox=具有Text属性，且具有Sel方法，类似TextBox,ComboBox等控件

    '人员以及限定输入项目不用输汉字
    zlCommFun.OpenIme blnChineseIn
    If gclsPros.PatiType = PF_住院 And blnSetInfect Then
        '医院感染相关控件可见性以及位置
        If gclsPros.CurrentForm.picInfectInfo.Visible Then Call ShowInfectInfo(False)
    End If
    '控件内容选择
    Call zlControl.TxtSelAll(objTextBox)
End Sub

Public Sub ShowInfectInfo(Optional ByVal blnShow As Boolean = True, Optional ByRef objCtrl As Object, _
                            Optional ByVal lngLeft As Long, Optional ByVal lngTop As Long)
'功能：对感染信息显示并设置位置或隐藏
'参数：
'      blnShow=true,显示感染信息；false,隐藏感染信息
'      objCtrl=控件对象
'      lngLeft,lngTop=感染信息的位置，相对于首页信

    Dim blnExit As Boolean

    blnExit = True
    If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 And gclsPros.PatiType <> PF_门诊 Then
        With gclsPros.CurrentForm
            '相关控件状态设置
            If .picInfectInfo.Visible = blnShow Then Exit Sub
            .picInfectInfo.Visible = blnShow
            .picInfectInfo.Top = lngTop - frmMain.Top - frmMain.PicForm.Top - .picMain.Top - .Top + 150
            .picInfectInfo.Left = lngLeft - frmMain.Left - .picMain.Left - frmMain.PicDirectory.Width - 200
            .picInfectInfo.ZOrder
        End With
    End If
End Sub

Private Sub TxtMouseDown(ByRef objText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：TextBox的默认右键消息提示的修改
    If intButton = 2 And objText.Locked Then
        gclsPros.TXTProc = GetWindowLong(objText.hwnd, GWL_WNDPROC)
        Call SetWindowLong(objText.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TxtMouseUp(ByRef objText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：TextBox的默认右键消息提示的修改
    If intButton = 2 And objText.Locked Then
        Call SetWindowLong(objText.hwnd, GWL_WNDPROC, gclsPros.TXTProc)
    End If
End Sub

Private Sub GetDiagTypeScope(ByRef vsDiag As VSFlexGrid, ByVal lngType As Long, ByRef lngBgn As Long, ByRef lngEnd As Long)
'功能：获取当前诊断类型的范围
'参数：lngType=当前类型
'返回：lngBgn=当前类型的起始行
'      lngEnd=当前类型结束行
    Dim i As Long
    Dim LngCol As Long

    lngBgn = 0: lngEnd = 0
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, DI_诊断分类)) = lngType Then
                If lngBgn < .FixedRows Then lngBgn = i
                lngEnd = i
            End If
        Next
    End With
End Sub

Private Function DiagRowCanMove(ByVal intStep As Integer, ByVal lngType As Long, ByVal LngRow As Long) As Boolean
'功能：设置诊断移动控件状态
'参数：intStep=移动位置，1-向下移动，-1向上移动
'   lngType=当前诊断类型
'   lngRow=需判定的行，一般为当前行
    Dim lngBgn As Long, lngEnd As Long
    '根据当前行的位置设置移动诊断控件的可用性
    Call GetDiagTypeScope(IIf(lngType <= 10, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY), lngType, lngBgn, lngEnd)
    If lngBgn = lngEnd Then '只有一行诊断，则不可移动
        DiagRowCanMove = False
    ElseIf LngRow = lngBgn Then '当前行是本分类第一行，则只能下移
        DiagRowCanMove = intStep = 1 And gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    ElseIf LngRow = lngEnd Then '当前行是本分类最后一行，则只能上
        DiagRowCanMove = intStep = -1 And gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    Else  '当前行是本分类中间某一行，则可以上下移动
        DiagRowCanMove = gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    End If
End Function

Private Sub MoveDiagRows(ByRef vsDiag As VSFlexGrid, ByVal intStep As Integer)
'功能：移动诊断行
'参数：vsDiag=当前诊断表格
'      intStep=移动位置，1-向下移动，-1向上移动
    Dim strTmp As String
    Dim i As Long, LngRow As Long
    Dim bln西医 As Boolean, bln分化程度 As Boolean
    Dim blnJudge As Boolean

    bln西医 = vsDiag.Name = "vsDiagXY"

    With vsDiag
        If Not DiagRowCanMove(intStep, Val(.TextMatrix(.Row, DI_诊断分类)), .Row) Then Exit Sub
        If .Row < 0 Then
            Exit Sub
        ElseIf gclsPros.FuncType <> f病案首页 Then '不可编辑跟位置有关的诊断
            LngRow = IIf(intStep = 1, .Row, .Row + intStep)
            '正常完成的出院诊断不允许改
            If gclsPros.PathState = PS_正常结束 And gclsPros.PathOutTime Then
                If bln西医 Then
                    blnJudge = .TextMatrix(.Row, DI_诊断分类) = "出院诊断" And gclsPros.InPath <= DT_入院诊断XY
                Else
                    blnJudge = .TextMatrix(.Row, DI_诊断分类) = "出院诊断" And gclsPros.InPath >= DT_门诊诊断ZY
                End If
                If blnJudge Then Exit Sub
            End If
        End If
        For i = .FixedCols To .Cols - 1
            '交换界面数据
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '交换隐藏数据
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        '交换隐藏数据
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        Call SetDiagReletedInfo(vsDiag)
        .Row = .Row + intStep
    End With
End Sub

Private Function OPSRowCanMove(ByVal intStep As Integer, ByVal LngRow As Long) As Boolean
'功能：设置诊断移动控件状态
'参数：intStep=移动位置，1-向下移动，-1向上移动
'   lngType=当前诊断类型
'   lngRow=需判定的行，一般为当前行
    Dim lngBgn As Long, lngEnd As Long
    Dim vsOPS As VSFlexGrid

    '根据当前行的位置设置移动诊断控件的可用性
    Set vsOPS = gclsPros.CurrentForm.vsOPS
    lngBgn = vsOPS.FixedRows: lngEnd = vsOPS.Rows - 1
    If lngBgn = lngEnd Then '只有一行诊断，则不可移动
        OPSRowCanMove = False
    ElseIf LngRow = lngBgn Then '当前行是本分类第一行，则只能下移
        OPSRowCanMove = intStep = 1 And gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    ElseIf LngRow = lngEnd Then '当前行是本分类最后一行，则只能上
        OPSRowCanMove = intStep = -1 And gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    Else  '当前行是本分类中间某一行，则可以上下移动
        OPSRowCanMove = gclsPros.OpenMode <> EM_查阅 And gclsPros.Module <> p住院护士站
    End If
End Function

Private Sub MoveOPSRows(ByRef vsOPS As VSFlexGrid, ByVal intStep As Integer)
'功能：移动诊断行
'参数：vsDiag=当前诊断表格
'      intStep=移动位置，1-向下移动，-1向上移动
    Dim strTmp As String
    Dim i As Long, LngRow As Long
    Dim bln西医 As Boolean, bln分化程度 As Boolean
    Dim blnJudge As Boolean


    With vsOPS
        If Not OPSRowCanMove(intStep, .Row) Then Exit Sub
        If .Row < .FixedRows Then
            Exit Sub
        End If
        For i = .FixedCols To .Cols - 1
            '交换界面数据
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '交换隐藏数据
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        '交换隐藏数据
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        .Row = .Row + intStep
    End With
End Sub

Private Sub OPSSetInput(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, rsInput As ADODB.Recordset)
'功能：根据手术情况输入的情况，设置表格数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim int手术输入 As Integer
    Dim blnOPSLevel As Boolean

    With vsOPS
        Select Case LngCol
            Case PI_手术编码, PI_手术名称
                If Not rsInput Is Nothing Then
                    int手术输入 = Val(gclsPros.OPSInput)
                    If gclsPros.CNIndent And gclsPros.FuncType = f病案首页 Then
                        '功能:如果是相互独立的，如果存在手术名称，则不替换,一致的情况,就需同步更新
                        If .TextMatrix(LngRow, PI_手术名称) = .Cell(flexcpData, LngRow, PI_手术名称) Or Trim(.TextMatrix(LngRow, PI_手术名称)) = "" Then
                            .TextMatrix(LngRow, PI_手术名称) = rsInput!名称
                        End If
                    Else
                        .TextMatrix(LngRow, PI_手术名称) = rsInput!名称
                    End If
                     .Cell(flexcpData, LngRow, PI_手术名称) = .TextMatrix(LngRow, PI_手术名称)
                    .TextMatrix(LngRow, PI_手术编码) = rsInput!编码
                    .AutoSize PI_手术编码, PI_手术名称
                    If int手术输入 = 0 Then
                        .TextMatrix(LngRow, PI_诊疗项目ID) = rsInput!ID
                        .TextMatrix(LngRow, PI_手术操作ID) = ""
                        strSql = "Select A.疾病ID as ID, Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 手术级别 From 疾病诊断对照 A, 疾病编码目录 B Where A.疾病ID = B.ID and  A.手术ID=[1]"
                    Else
                        .TextMatrix(LngRow, PI_手术操作ID) = rsInput!ID
                        .TextMatrix(LngRow, PI_诊疗项目ID) = ""
                        strSql = "Select A.手术ID as ID, Decode(B.手术类型, '甲', '四级手术', '乙', '三级手术', '丙', '二级手术', '丁', '一级手术', '四级', '四级手术', '三级', '三级手术', '二级', '二级手术', '一级', '一级手术', Null) 手术级别 From 疾病诊断对照 A, 疾病编码目录 B Where A.疾病ID(+) = B.ID and  B.ID=[1]"
                    End If
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, Val(rsInput!ID))
                    If Not rsTmp.EOF Then
                        If int手术输入 = 0 Then
                            .TextMatrix(LngRow, PI_手术操作ID) = Val(rsTmp!ID)
                        Else
                            .TextMatrix(LngRow, PI_诊疗项目ID) = Val(rsTmp!ID & "")
                        End If
                         If NVL(rsTmp!手术级别) <> "" Then .TextMatrix(LngRow, PI_手术级别) = NVL(rsTmp!手术级别)
                         blnOPSLevel = NVL(rsTmp!手术级别) <> ""
                    End If
                Else
                    .TextMatrix(LngRow, LngCol) = .EditText
                    .TextMatrix(LngRow, PI_手术操作ID) = ""
                    .TextMatrix(LngRow, PI_诊疗项目ID) = ""
                End If
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)

                '手术日期相同时，其他输入内容默认与上一行相同
                If Not rsInput Is Nothing And LngRow > .FixedRows And LngRow = .Rows - 1 Then
                    If .TextMatrix(LngRow, PI_手术日期) = .TextMatrix(LngRow - 1, PI_手术日期) Then
                        .TextMatrix(LngRow, PI_主刀医师) = .TextMatrix(LngRow - 1, PI_主刀医师)
                        .TextMatrix(LngRow, PI_助产护士) = .TextMatrix(LngRow - 1, PI_助产护士)
                        .TextMatrix(LngRow, PI_助手1) = .TextMatrix(LngRow - 1, PI_助手1)
                        .TextMatrix(LngRow, PI_助手2) = .TextMatrix(LngRow - 1, PI_助手2)
                        .TextMatrix(LngRow, PI_麻醉方式) = .TextMatrix(LngRow - 1, PI_麻醉方式)
                        .TextMatrix(LngRow, PI_麻醉医师) = .TextMatrix(LngRow - 1, PI_麻醉医师)
                        .TextMatrix(LngRow, PI_切口愈合) = .TextMatrix(LngRow - 1, PI_切口愈合)
                        .TextMatrix(LngRow, PI_麻醉ID) = .TextMatrix(LngRow - 1, PI_麻醉ID)
                        .TextMatrix(LngRow, PI_麻醉类型) = .TextMatrix(LngRow - 1, PI_麻醉类型)

                        For i = PI_主刀医师 To .Cols - 1
                            .Cell(flexcpData, LngRow, i) = .TextMatrix(LngRow, i)
                        Next
                    End If
                End If
                .Cell(flexcpData, LngRow, PI_手术级别) = IIf(blnOPSLevel, 1, 0)
                '设置诊断符合情况
                Call SetDiagMatchInfo(BCC_术前与术后)
                '输入后始终保持一新行
                If LngRow = .Rows - 1 Then .AddItem "": Call ChangeVSFHeight(vsOPS, True)
            Case PI_麻醉方式 '该列为隐藏列
                .TextMatrix(LngRow, LngCol) = rsInput!名称
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                .TextMatrix(LngRow, PI_麻醉ID) = rsInput!ID
                .TextMatrix(LngRow, PI_麻醉类型) = NVL(rsInput!麻醉类型)
            Case PI_主刀医师, PI_助产护士, PI_助手1, PI_助手2, PI_麻醉医师
                .TextMatrix(LngRow, LngCol) = rsInput!名称
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            Case PI_切口部位
                .TextMatrix(LngRow, LngCol) = rsInput!名称
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetSign(ByVal intIndex As Integer, Optional ByVal blnUnSign As Boolean)
'功能：签名或取消签名
'参数：intIndex=签名按钮索引
'      blnUnSign=Fasle-签名，True-取消签名
    Dim strSql As String
    Dim rsTmp As Recordset, i As Long
    Dim bln手术 As Boolean    '是否填写了手术记录
    '说明：arrInfos，arrManIdxs，arrSgnIdxs三个数组的元素一一对应，人员级别从低到高
    Dim arrInfos() As Variant '各类签名的信息名
    Dim arrManIdxs() As Variant '签名人员下拉列表的Index
    Dim arrSgnIdxs() As Variant '签名按钮的Index
    Dim blnSign As Boolean '以前的签名状态
    Dim blnDiagnose As Boolean
    Dim strTmp As String, cboTmp As ComboBox

    With gclsPros.CurrentForm
        '判断是否启用数字签名
        If gintCA > 0 And CheckSign(1, 0, 0, gclsPros.出院科室ID, 2) Then
            If gobjESign Is Nothing Then
                On Error Resume Next
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Err.Clear: On Error GoTo 0
                If Not gobjESign Is Nothing Then
                    Call gobjESign.Initialize(gcnOracle, gclsPros.SysNo)
                End If
            End If
            If gobjESign Is Nothing Then
                MsgBox "电子签名部件未能正确安装，签名操作不能继续。", vbInformation, gstrSysName
                Exit Sub
            Else
                If Not gobjESign.CheckCertificate(UserInfo.DBUser) Then Exit Sub
            End If
        End If

        arrInfos = Array("住院医师", "主治医师", "主任医师", "科主任")
        arrManIdxs = Array(MC_住院医师, MC_主治医师, MC_主任或副主任, MC_科主任)
        arrSgnIdxs = Array(SL_住院医师, SL_主治医师, SL_主任医师, SL_科主任)
        If blnUnSign Then
            '并发检查病案是否编目或首页处于锁定状态
            If Not CheckMecRed(gclsPros.病人ID, gclsPros.主页ID, .Caption, "取消签名") Then Exit Sub
        Else
            '并发检查病案是否编目或首页处于锁定状态
'            If Not CheckMecRed(gclsPros.病人ID, gclsPros.主页ID, .Caption, "签名") Then Exit Sub
            '需要确定更高签名级别的人
            For i = UBound(arrSgnIdxs) To LBound(arrSgnIdxs) Step -1
                If i = LBound(arrSgnIdxs) Then Exit For '级别最低的住院医师
                If i <> UBound(arrSgnIdxs) Then
                    If .cboManInfo(arrManIdxs(i + 1)).Text = "" Then
                        If strTmp = "" Then
                             strTmp = "没有确定" & arrInfos(i + 1)
                             Set cboTmp = .cboManInfo(arrManIdxs(i + 1))
                        Else
                             strTmp = strTmp & "和" & arrInfos(i + 1)
                        End If
                    End If
                End If
            Next
            If strTmp <> "" Then
                Call ShowMessage(cboTmp, strTmp & "。")
                Exit Sub
            End If
            On Error GoTo errH
            '如果有手术记录，则提示是否继续
            bln手术 = False
            For i = 1 To .vsOPS.Rows - 1
                If Trim(.vsOPS.TextMatrix(i, PI_手术名称)) <> "" Then
                    bln手术 = True
                End If
            Next

            strSql = "Select Count(1) As 手术 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And 医嘱状态=8 And 诊疗类别='F'"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.病人ID, gclsPros.主页ID)
            If Val(rsTmp!手术 & "") > 0 And Not bln手术 Then
                .vsOPS.Row = .vsOPS.FixedRows: .vsOPS.Col = PI_手术编码
                If ShowMessage(.vsOPS, "该病人存在手术医嘱，但首页中没有添加手术记录，是否继续？", True) = vbNo Then Exit Sub
            End If

            '签名前自动保存
            If Not CheckMedPageData(blnDiagnose) Then
                gclsPros.IsCheckData = False
                Exit Sub
            End If

            If Not gclsMain.IsDiagInput And Not blnDiagnose And gclsPros.MustDiagType <> "" Then
                If MsgBox("要求的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            If Not SaveMedPageData() Then Exit Sub
        End If

        For i = LBound(arrSgnIdxs) To UBound(arrSgnIdxs)
            If arrSgnIdxs(i) = intIndex Then
                strSql = "Zl_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'" & arrInfos(i) & "签名'," & IIf(blnUnSign, "Null", "'" & UserInfo.姓名 & "'") & ")"
                Exit For
            End If
        Next
        Call zlDatabase.ExecuteProcedure(strSql, gclsPros.CurrentForm.Caption)
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.病人ID, gclsPros.主页ID)
        blnSign = gclsPros.IsSigned
        gclsPros.IsSigned = SetSignature
        If blnSign And Not gclsPros.IsSigned Then
            Call SetFaceInit(True)
        End If
        Call SetFaceEditable(gclsPros.IsSigned)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function PageOperate(ByVal mopType As MedRec_Operate, Optional ByVal intPage As Integer) As Boolean
'功能：打印或预览首页或点击确定保存首页
'参数：intType=2（打印），=1（预览）0=设置，3点击确定按钮保存首页
'      intPage=1-4打印的页数（格式）=5打印正面+附页1，=6打印反面+附页2
'      blnSavePage=True-点击确定按钮是否保存首页 ，False-打印或预览首页
'返回：是否成功
    Dim blnDiagnose As Boolean
    Dim blnPagePrint As Boolean, intPrint As Integer

    If mopType = MOP_确定 And (gclsPros.FuncType = f医生首页 And gclsPros.IsSigned Or gclsPros.OpenMode = EM_查阅) Then
        gclsPros.IsOK = True
        PageOperate = True
        gclsPros.IsDiagChange = Not blnDiagnose
        Exit Function
    End If

    '病案首页自然无签名锁定状态，住院首页若无签名锁定状态才保存
    If gclsPros.OpenMode <> EM_查阅 And Not gclsPros.IsSigned Then
        If gclsPros.FuncType = f病案首页 Then
            If Not ValidatePageNos(True) Then Exit Function
        End If
        If gclsPros.FuncType = f医生首页 And gclsPros.PatiInfo!病人性质 = 1 Then
            If Not Check留观 Then
                gclsPros.IsCheckData = False
                Exit Function
            End If
        Else
            If Not CheckMedPageData(blnDiagnose) Then
                gclsPros.IsCheckData = False
                Exit Function
            End If
        End If
        If Not gclsMain.IsDiagInput And Not blnDiagnose And gclsPros.MustDiagType <> "" Then
            If MsgBox("要求的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If

        If Not SaveMedPageData() Then Exit Function
        '门诊首页照片保存
        If gclsPros.FuncType = f医生首页 And gclsPros.PatiType = PF_门诊 Then
            Call SavePatPicture(gclsPros.病人ID)
        End If
    End If
    If gclsPros.FuncType = f病案首页 And mopType = MOP_确定 Then
        If gclsPros.OpenMode = EM_新增病案 Or gclsPros.OpenMode = EM_新增首页 Then
            If Val(zlDatabase.GetPara("录入病案存放位置", gclsPros.SysNo, gclsPros.Module)) = 1 Then
                Call gclsMain.MedRecSaveLocation(gclsPros.病人ID, gclsPros.主页ID)
            End If
        End If
    End If
    '修改首页或病案点击确定则退出
    If mopType = MOP_确定 And gclsPros.OpenMode = EM_编辑 Then
        If gclsPros.FuncType = f病案首页 Then Call gclsMain.SavePage(gclsPros.病人ID, gclsPros.主页ID)
        gclsPros.IsOK = True
        PageOperate = True
        gclsPros.IsDiagChange = Not blnDiagnose
    End If
    ' 住院首页打印首页，病案首页打印病案首页封面，病案首页只有确定一种状态
    If gclsPros.FuncType = f病案首页 Or mopType <> MOP_确定 And gclsPros.FuncType <> f病案首页 Then
        If gclsPros.FuncType <> f病案首页 Then
            blnPagePrint = True
        Else
            blnPagePrint = InStr(gclsPros.Privs, "档案袋打印") > 0
            If blnPagePrint Then
                intPrint = Val(zlDatabase.GetPara("病案档案袋封皮打印", gclsPros.SysNo, gclsPros.Module))
                blnPagePrint = intPrint <> 0
                If blnPagePrint And intPrint = 2 Then
                    blnPagePrint = MsgBox("是否打印病人病案档案袋封皮？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                End If
            End If
        End If
        If blnPagePrint Then
            Call PrintInMedRec(mopType, gclsPros.病人ID, gclsPros.主页ID, gclsPros.出院科室ID, intPage)
        End If
    End If
    '病案重新输入，需要清除界面数据，重新初始化,新增病案，新增首页属于病案特有
    If gclsPros.OpenMode = EM_新增病案 Or gclsPros.OpenMode = EM_新增首页 Then
        Call gclsMain.SavePage(gclsPros.病人ID, gclsPros.主页ID)
        '清除所输入内容,再增加一个用户
        gclsPros.InNo = ""
        Call ClearPageContent
        Call SetAllVSF(True)
        Call ChangePage(, 0)
    End If
    PageOperate = True
End Function

Private Sub KSSSetDiagInput(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, rsInput As ADODB.Recordset)
'功能：处理抗生素项目的输入
    With vsKSS
        If Not rsInput Is Nothing Then
            .TextMatrix(LngRow, KI_抗菌药物名) = NVL(rsInput!名称)
            .RowData(LngRow) = Val(rsInput!ID)
        Else
            .TextMatrix(LngRow, KI_抗菌药物名) = .EditText
        End If
        .Cell(flexcpData, LngRow, KI_抗菌药物名) = .TextMatrix(LngRow, KI_抗菌药物名)
    End With
End Sub

Private Sub KSSEnterNextCell(ByRef vsKSS As VSFlexGrid)
    With vsKSS
        If .Row = .Rows - 1 Then
            If .TextMatrix(.Row, KI_抗菌药物名) = "" Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True: Exit Sub
            ElseIf .Editable = flexEDNone Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True: Exit Sub
            Else
                .AddItem ""
                Call ChangeVSFHeight(vsKSS, True)
                Call SetKSSSerial
            End If
        End If

        If .Col = KI_联合用药 Then
            .Col = .FixedCols
            .Row = .Row + 1
        Else
            .Col = .Col + 1
        End If
        If .Row >= 0 And .Row < .Rows - 1 And .Col >= 0 And .Row < .Cols - 1 Then
            .ShowCell .Row, .Col
        End If
    End With
End Sub


Public Function AddOrDelFreeCols(ByRef vsFree As VSFlexGrid, ByVal strFreeNames As String, ByVal strFreeNum As String, ByVal blnAdd As Boolean) As Boolean
'功能：删除或者新增指定费用
'参数：vsFree=费用表格
'      strFreeNames=费用名
'      blnAdd=true-新增费用，False-删除费用
'      strFreeNum=费用数量
'返回：是否找到该费用
    Dim LngRow As Long, LngCol As Long, i As Long
    Dim blnFind As Boolean, j As Long
    Dim lngPreRow As Long, lngPreCol As Long

    With vsFree
        If blnAdd Then
            If .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col + 1)) = strFreeNames And .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col + 1, .Col)) <> "" Then
                If Not gclsPros.SameName Then
                    MsgBox "该费用已录入，请再选一种。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '检查最后一行中未录入的位置
            For i = (.Rows - 1) * 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) = "" Or .TextMatrix(LngRow, LngCol + 1) = "" Then
                    Exit For
                ElseIf LngCol = 4 Then '最后一栏已经填写,则增加一行
                    .Rows = .Rows + 1: LngRow = LngRow + 1: LngCol = 0: Call ChangeVSFHeight(vsFree, True): Exit For
                End If
            Next
            .TextMatrix(LngRow, LngCol) = strFreeNames
            .TextMatrix(LngRow, LngCol + 1) = strFreeNum
        Else
            '找到需要删除的位置
            For i = 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) = strFreeNames And .TextMatrix(LngRow, LngCol + 1) = strFreeNum Then
                    blnFind = True: Exit For
                End If
            Next
            '将此位置后的费用全部往前移动
            If blnFind Then
                For j = i + 1 To .Rows * 3 - 1
                    LngRow = j \ 3: LngCol = (j Mod 3) * 2
                    lngPreRow = (j - 1) \ 3: lngPreCol = ((j - 1) Mod 3) * 2
                    .TextMatrix(lngPreRow, lngPreCol) = .TextMatrix(LngRow, LngCol)
                    .TextMatrix(lngPreRow, lngPreCol + 1) = .TextMatrix(LngRow, LngCol + 1)
                Next
                '倒数第二行，最后一栏没有填写，则移除最后一行
                If .Rows > 2 Then
                    If .TextMatrix(.Rows - 2, 4) = "" Then
                        .Rows = .Rows - 1
                        Call ChangeVSFHeight(vsFree, True)
                    End If
                End If
            End If
        End If
        Call SumAndSetFrees
    End With

    AddOrDelFreeCols = True
End Function

Public Sub SetDiagInput(ByRef vsDiagTmp As VSFlexGrid, ByVal LngRow As Long, rsInput As ADODB.Recordset, Optional bln附码 As Boolean)
'功能：处理诊断项目的输入
'      bln附码=是否是附码输入
    Dim str性别 As String
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String, bln分化程度 As Boolean
    Dim bln西医 As Boolean, blnRCodeIn As Boolean
    Dim lngTmpRow As Long, lng出院Row As Long
    Dim lng原诊断ID As Long, int诊断次序 As Integer
    Dim blnSame病理诊断 As Boolean
    Dim rs病理 As New ADODB.Recordset
    Dim rsOutPut As ADODB.Recordset
    Dim blnGet附码 As Boolean

    blnGet附码 = gclsPros.GetExtraCode
    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If gclsPros.FuncType = f病案首页 Then
                    '病案此条逻辑不知什么原因，暂时保留,可能是重庆疾病编码的特殊需求
                    If rsInput!编码 Like "R*" Then
                        If blnRCodeIn Then
                            Exit For
                        Else
                            If MsgBox("你现在正使用R编码作为主要编码，是否输入？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                                blnRCodeIn = True: Exit For
                            End If
                        End If
                    End If
                End If
                '不是单独的附码输入
                If Not bln附码 Then
                    If i > 1 Then
                        '最后一种诊断类别（中医：其他诊断，西医：损伤中毒）选择多条时的处理
                        lng原诊断ID = 0
                        If LngRow = .Rows - 1 Then
                            .Rows = .Rows + 1
                            Call ChangeVSFHeight(vsDiagTmp, True)
                            .TextMatrix(.Rows - 1, DI_诊断分类) = .TextMatrix(LngRow, DI_诊断分类)
                            If gclsPros.PatiType = PF_门诊 Then .TextMatrix(.Rows - 1, DI_诊断类型) = .TextMatrix(LngRow, DI_诊断类型)
                        End If
                        '确定当前显示行
                        If Val(.TextMatrix(LngRow + 1, DI_诊断分类)) = Val(.TextMatrix(LngRow, DI_诊断分类)) Then
                            For j = LngRow + 1 To .Rows - 1
                                If Val(.TextMatrix(j, DI_诊断分类)) = Val(.TextMatrix(LngRow, DI_诊断分类)) Then
                                    LngRow = j
                                    If .TextMatrix(j, DI_诊断描述) = "" Then Exit For
                                Else
                                    Exit For
                                End If
                            Next
                            If .TextMatrix(LngRow, DI_诊断描述) <> "" Then
                                LngRow = LngRow + 1: .AddItem "", LngRow
                                Call ChangeVSFHeight(vsDiagTmp, True)
                                .TextMatrix(LngRow, DI_诊断分类) = .TextMatrix(LngRow - 1, DI_诊断分类)
                                If gclsPros.PatiType = PF_门诊 Then .TextMatrix(LngRow, DI_诊断类型) = .TextMatrix(LngRow - 1, DI_诊断类型)
                            End If
                        Else
                            LngRow = LngRow + 1: .AddItem "", LngRow
                            Call ChangeVSFHeight(vsDiagTmp, True)
                            .TextMatrix(LngRow, DI_诊断分类) = .TextMatrix(LngRow - 1, DI_诊断分类)
                            If gclsPros.PatiType = PF_门诊 Then .TextMatrix(LngRow, DI_诊断类型) = .TextMatrix(LngRow - 1, DI_诊断类型)
                        End If
                    Else
                        lng原诊断ID = Val(.TextMatrix(LngRow, DI_诊断ID))
                    End If

                    .TextMatrix(LngRow, DI_诊断编码) = rsInput!编码 & ""
                    If gclsPros.CNIndent And gclsPros.FuncType = f病案首页 Then
                        '功能:如果是相互独立的，如果存在诊断描述，则不替换
                        '一致的情况,就需同步更新疾病
                        If .TextMatrix(LngRow, DI_诊断描述) = .Cell(flexcpData, LngRow, DI_诊断描述) Or Trim(.TextMatrix(LngRow, DI_诊断描述)) = "" Then
                            .TextMatrix(LngRow, DI_诊断描述) = rsInput!名称
                        End If
                    Else
                        .TextMatrix(LngRow, DI_诊断描述) = rsInput!名称
                    End If
                    .Cell(flexcpData, LngRow, DI_诊断描述) = rsInput!名称 & ""  '保存原名
                    .Cell(flexcpData, LngRow, DI_诊断编码) = rsInput!编码 & ""
                    .AutoSize DI_诊断编码, DI_诊断描述
                    If .ColWidth(DI_诊断描述) < 3200 Then
                        .ColWidth(DI_诊断描述) = 3200
                    End If
                    If gclsPros.FuncType = f诊断选择 Then .TextMatrix(LngRow, DI_关联) = 1
                    .TextMatrix(LngRow, DI_诊断ID) = rsInput!诊断ID & ""
                    .TextMatrix(LngRow, DI_疾病ID) = rsInput!疾病id & ""
                    .TextMatrix(LngRow, DI_疗效限制) = rsInput!疗效限制 & ""
                    .TextMatrix(LngRow, DI_分娩信息) = IIf(Val(rsInput!分娩 & "") = 1, "1", "")
                    .TextMatrix(LngRow, DI_是否病人) = IIf(Val(rsInput!是否病人 & "") = 1, "1", "")
                    .TextMatrix(LngRow, DI_疾病编码) = rsInput!疾病编码 & ""
                    .TextMatrix(LngRow, DI_疾病类别) = rsInput!疾病类别 & ""
                    '并发症，院内感染出院情况默认为无
                    If Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_并发症 Or Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_院内感染 Then
                        .TextMatrix(LngRow, DI_入院病情) = "无"
                    End If
                    If blnGet附码 Then
                        If Not IsNull(rsInput!附码) Then
                            Set rsTmp = GetDiagExtraID(rsInput!附码 & "")
                            If rsTmp.RecordCount > 0 Then
                                .TextMatrix(LngRow, DI_附码ID) = rsTmp!ID & ""
                            Else
                                .TextMatrix(LngRow, DI_附码ID) = ""
                            End If
                        End If
                        .TextMatrix(LngRow, DI_ICD附码) = IIf(bln附码, rsInput!编码 & "", rsInput!附码 & "")
                        .Cell(flexcpData, LngRow, DI_ICD附码) = .TextMatrix(LngRow, DI_ICD附码)
                    End If
                    '如果启用了参数ICD附码必须填写时，录入的C00-D48则弹出要求录入肿瘤形态学编码；
                    If gclsPros.CheckICD附码 = 1 And (Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断XY) _
                        And (InStr("C", Left(.TextMatrix(LngRow, DI_诊断编码), 1)) > 0 Or (InStr("D", Left(.TextMatrix(LngRow, DI_诊断编码), 1)) > 0 And Val(Mid(.TextMatrix(LngRow, DI_诊断编码), 2, 2)) <= 48)) And Left(.TextMatrix(LngRow, DI_诊断编码), 1) <> "" Then
                        If frmZLInPut.ShowMe(gclsPros.CurrentForm, "   诊断[" & .Cell(flexcpData, LngRow, DI_诊断描述) & "]为肿瘤诊断，请输入肿瘤形态学编码！", rsOutPut) Then
                            Call SetDiagInput(vsDiagTmp, LngRow, rsOutPut, True)
                        End If
                    End If
                Else
                    .TextMatrix(LngRow, DI_附码ID) = rsInput!项目ID & ""
                    .TextMatrix(LngRow, DI_ICD附码) = rsInput!编码 & ""
                    .Cell(flexcpData, LngRow, DI_ICD附码) = .TextMatrix(LngRow, DI_ICD附码)
                End If
                
             
                
                '病案首页，出院主要诊断替换门诊主要诊断与入院主要诊断
                If gclsPros.FuncType = f病案首页 Then
                    '门诊、入院诊断不能设附码,中医不设置附码
                    If Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_门诊诊断XY And Val(.TextMatrix(LngRow, DI_诊断分类)) <> DT_入院诊断XY And bln西医 Then
                        '出院诊断附码是V,W,X,Y,则设置损伤中毒原因,这段代码好似特别针对重庆疾病编码
                        If Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_出院诊断XY And Val(.TextMatrix(LngRow, DI_附码ID)) <> 0 Then
                            If InStr("VWXY", Left(rsInput!编码 & "", 1)) > 0 Then
                                If gclsPros.Sex Like "*男*" Then
                                    str性别 = "男"
                                ElseIf gclsPros.Sex Like "*女*" Then
                                    str性别 = "女"
                                End If

                                strSql = "Select A.Id,A.Id As 项目id, A.编码, A.序号, A.附码, D.ID 附码ID, D.名称 附码名称, A.名称, A.说明, Null 编者, A.分类id, " & IIf(gclsPros.BriefCode = 0, "A.简码", "A.五笔码") & " as 简码, A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, Null 疾病id,A.类别 疾病类别, Null 诊断id" & vbNewLine & _
                                        "From 疾病编码目录 A, 疾病编码分类 C, 疾病编码目录 D " & vbNewLine & _
                                        "Where A.ID=[1] And A.附码=D.编码(+)  And A.分类id = C.Id(+)" & vbNewLine & _
                                        "  And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIf(str性别 <> "", " And (A.性别限制=[2] Or A.性别限制 is Null) ", " ")
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取出院诊断附码对应的损伤中毒码", Val(.TextMatrix(LngRow, DI_附码ID)), str性别)
                                '定位损伤中毒的空行，没有则新增一行
                                lngTmpRow = 0
                                For j = FindDiagRow(DT_损伤中毒码) To .Rows - 1
                                    If .TextMatrix(j, DI_诊断描述) = "" Then
                                        lngTmpRow = j
                                        Exit For
                                    End If
                                Next
                                If lngTmpRow = 0 Then .Rows = .Rows + 1: lngTmpRow = .Rows - 1: Call ChangeVSFHeight(vsDiagTmp, True)
                                '输入损伤中毒
                                Call SetDiagInput(vsDiagTmp, lngTmpRow, rsTmp)
                            End If
                        End If
                    End If

                    lng出院Row = FindDiagRow(IIf(bln西医, DT_出院诊断XY, DT_出院诊断ZY))
                    If LngRow = lng出院Row Then
                        '替换门诊诊断
                        lngTmpRow = FindDiagRow(IIf(bln西医, DT_门诊诊断XY, DT_门诊诊断ZY))
                        If .TextMatrix(lngTmpRow, DI_疾病ID) = "" Then
                             '功能:如果是相互独立的，如果存在诊断描述，则不替换
                            If gclsPros.CNIndent And Trim(.TextMatrix(lngTmpRow, DI_诊断描述)) = "" Or Not gclsPros.CNIndent Then
                                .TextMatrix(lngTmpRow, DI_疾病ID) = .TextMatrix(LngRow, DI_疾病ID)
                                .TextMatrix(lngTmpRow, DI_诊断编码) = .TextMatrix(LngRow, DI_诊断编码)
                                .TextMatrix(lngTmpRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                                .Cell(flexcpData, lngTmpRow, DI_诊断描述) = .Cell(flexcpData, LngRow, DI_诊断描述)
                                .Cell(flexcpData, lngTmpRow, DI_诊断编码) = .Cell(flexcpData, LngRow, DI_诊断编码)
                                '设置诊断相关信息
                                Call SetDiagReletedInfo(vsDiagTmp, lngTmpRow)
                            End If
                            .TextMatrix(lngTmpRow, DI_备注) = .TextMatrix(LngRow, DI_备注)
                        End If
                        '替换入院诊断
                        lngTmpRow = FindDiagRow(IIf(bln西医, DT_入院诊断XY, DT_入院诊断ZY))
                        If .TextMatrix(lngTmpRow, DI_疾病ID) = "" Then
                             '功能:如果是相互独立的，如果存在诊断描述，则不替换
                            If gclsPros.CNIndent And Trim(.TextMatrix(lngTmpRow, DI_诊断描述)) = "" Or Not gclsPros.CNIndent Then
                                .TextMatrix(lngTmpRow, DI_疾病ID) = .TextMatrix(LngRow, DI_疾病ID)
                                .TextMatrix(lngTmpRow, DI_诊断编码) = .TextMatrix(LngRow, DI_诊断编码)
                                .TextMatrix(lngTmpRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                                .Cell(flexcpData, lngTmpRow, DI_诊断描述) = .Cell(flexcpData, LngRow, DI_诊断描述)
                                .Cell(flexcpData, lngTmpRow, DI_诊断编码) = .Cell(flexcpData, LngRow, DI_诊断编码)
                                '设置诊断相关信息
                                Call SetDiagReletedInfo(vsDiagTmp, lngTmpRow)
                            End If
                            .TextMatrix(lngTmpRow, DI_备注) = .TextMatrix(LngRow, DI_备注)
                        End If

                        '删除损伤中毒原因
                        If bln西医 And Not .TextMatrix(LngRow, DI_诊断编码) Like "S*" And Not .TextMatrix(LngRow, DI_诊断编码) Like "T*" Then
                            lngTmpRow = FindDiagRow(DT_损伤中毒码)
                            If lngTmpRow < .Rows - 1 Then
                                .Rows = lngTmpRow + 1
                                Call ChangeVSFHeight(vsDiagTmp, True)
                            End If
                            .Cell(flexcpText, lngTmpRow, .FixedCols, lngTmpRow, .Cols - 1) = ""
                            .Cell(flexcpData, lngTmpRow, .FixedCols, lngTmpRow, .Cols - 1) = ""
                            .TextMatrix(lngTmpRow, DI_诊断分类) = DT_损伤中毒码
                            .RowData(lngTmpRow) = 0
                        End If
                    ElseIf Val(.TextMatrix(LngRow, DI_诊断分类)) = DT_损伤中毒码 Then
                        If .TextMatrix(LngRow, DI_出院情况) = "" Then
                            .TextMatrix(LngRow, DI_出院情况) = .TextMatrix(lng出院Row, DI_出院情况)
                        End If
                    End If
                End If

                If Not bln西医 Then
                    '中医根据疾病诊断参考取证候
                    Call Set中医证候(LngRow, Val(.TextMatrix(LngRow, DI_诊断ID)))
                End If
                If gclsPros.FuncType <> f病案首页 Then
                    If CreatePlugInOK(IIf(gclsPros.PatiType = PF_门诊, p门诊医生站, p住院医生站)) Then
                        int诊断次序 = 0
                        If gclsPros.PatiType = PF_住院 Then
                            For j = .FixedRows To LngRow
                                If .TextMatrix(j, DI_诊断分类) = .TextMatrix(LngRow, DI_诊断分类) Then
                                    int诊断次序 = int诊断次序 + 1
                                End If
                            Next
                        Else
                            int诊断次序 = IIf(LngRow = .FixedRows, -1, -2)
                        End If
                        On Error Resume Next
                        Select Case int诊断次序
                            Case -1
                                Call gobjPlugIn.DiagnosisEnter(gclsPros.SysNo, p门诊医生站, gclsPros.病人ID, gclsPros.主页ID, Val(rsInput!项目ID), .TextMatrix(LngRow, DI_诊断描述), lng原诊断ID)
                                Call zlPlugInErrH(Err, "DiagnosisEnter")
                            Case -2
                                Call gobjPlugIn.DiagnosisOtherEnter(gclsPros.SysNo, p门诊医生站, gclsPros.病人ID, gclsPros.主页ID, Val(rsInput!项目ID), .TextMatrix(LngRow, DI_诊断描述), lng原诊断ID)
                                Call zlPlugInErrH(Err, "DiagnosisOtherEnter")
                            Case Else
                                Call gobjPlugIn.DiagnosisEnterIn(gclsPros.SysNo, p住院医生站, gclsPros.病人ID, gclsPros.主页ID, Val(rsInput!项目ID), .TextMatrix(LngRow, DI_诊断描述), lng原诊断ID, _
                                    IIf(gclsPros.Is护士站, 1, 0), .TextMatrix(LngRow, DI_诊断分类), int诊断次序)
                                Call zlPlugInErrH(Err, "DiagnosisEnterIn")
                        End Select
                        Err.Clear: On Error GoTo errH
                    End If
                End If
                rsInput.MoveNext
            Next
        Else
            If Not bln附码 Then
                If gclsPros.CNIndent And gclsPros.FuncType = f病案首页 Or gclsPros.FuncType <> f病案首页 Then
                    .TextMatrix(LngRow, DI_诊断描述) = .EditText
                    If gclsPros.FuncType <> f病案首页 Then
                        .Cell(flexcpData, LngRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                        .TextMatrix(LngRow, DI_诊断编码) = ""
                         .Cell(flexcpData, LngRow, DI_诊断描述) = ""
                        .TextMatrix(LngRow, DI_诊断ID) = ""
                        .TextMatrix(LngRow, DI_疾病ID) = ""
                        .TextMatrix(LngRow, DI_证候ID) = ""
                    End If
                End If
            Else
                .TextMatrix(LngRow, DI_固定附码) = ""
                .TextMatrix(LngRow, DI_ICD附码) = ""
                .Cell(flexcpData, LngRow, DI_ICD附码) = ""
                .TextMatrix(LngRow, DI_附码ID) = ""
            End If
        End If
        .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
        '设置诊断符合情况
        If Not (gclsPros.PatiType = PF_门诊 And gclsPros.FuncType = f诊断选择) Then
            Call SetDiagReletedInfo(vsDiagTmp, LngRow)
        End If
        If gclsPros.FuncType <> f诊断选择 Then
            '设置诊断相关信息
            If gclsPros.Module = p门诊医生站 Then
                If gclsPros.CurrentForm.optState(OP_复诊).Value = False Then
                    If PatiReSeeDoctor Then
                        If MsgBox("病人就诊科室、医生、诊断与上次相同，要标记为复诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            gclsPros.CurrentForm.optState(OP_复诊).Value = True
                        End If
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Set中医证候(ByVal LngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'      blnFreeInput  true - 自由录入
'返回：是否有对应关系
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String

    On Error GoTo errH

    With gclsPros.CurrentForm.vsDiagZY
        If blnFreeInput Then
            .TextMatrix(LngRow, DI_证候ID) = ""
            .TextMatrix(LngRow, DI_证候编码) = ""
            .TextMatrix(LngRow, DI_中医证候) = .EditText
            .Cell(flexcpData, LngRow, DI_中医证候) = .TextMatrix(LngRow, DI_中医证候)
        Else
            If rsInput Is Nothing Then
                If lng诊断ID = 0 Then Exit Function
                strSql = "Select Distinct A.证候序号 As ID, A.证候id As 项目id, B.编码, B.附码, A.证候名称 名称," & IIf(gclsPros.BriefCode = 0, "B.简码", "B.五笔码 As 简码") & ", B.说明" & vbNewLine & _
                            "From 疾病诊断参考 A, 疾病编码目录 B" & vbNewLine & _
                            "Where A.证候id = B.Id(+) And A.诊断id = [1] And A.证候名称 Is Not Null" & vbNewLine & _
                            "Order By A.证候序号"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsInput = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                If rsInput Is Nothing Then
                    If Not blnCancel Then Exit Function
                    If .EditText <> "" Then .EditText = .Cell(flexcpData, LngRow, DI_中医证候)
                    Set中医证候 = True: Exit Function
                End If
            End If

            .TextMatrix(LngRow, DI_证候ID) = NVL(rsInput!项目ID)
            .TextMatrix(LngRow, DI_证候编码) = NVL(rsInput!编码)
            If Not IsNull(rsInput!名称) Then
                '去掉已有的证候
                If .TextMatrix(LngRow, DI_诊断描述) Like "?*(?*)" Then
                    strTmp = Mid(.TextMatrix(LngRow, DI_诊断描述), 1, InStrRev(.TextMatrix(LngRow, DI_诊断描述), "(") - 1)
                Else
                    strTmp = .TextMatrix(LngRow, DI_诊断描述)
                End If
                .TextMatrix(LngRow, DI_诊断描述) = strTmp
                .Cell(flexcpData, LngRow, DI_诊断描述) = .TextMatrix(LngRow, DI_诊断描述)
                .TextMatrix(LngRow, DI_中医证候) = NVL(rsInput!名称)
                .Cell(flexcpData, LngRow, DI_中医证候) = .TextMatrix(LngRow, DI_中医证候)
                If .EditText <> "" Then .EditText = .TextMatrix(LngRow, DI_中医证候)
            End If

            Set中医证候 = True
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetDiagReletedInfo(ByRef vsDiagTmp As VSFlexGrid, Optional ByVal LngRow As Long = -1)
'功能：设置诊断相关信息，可以根据某行诊断，设置诊断符合情况
    Dim bln西医 As Boolean
    Dim strDiagTypeName As String
    Dim strTmp As String, bln分化程度 As Boolean, blnOld分化程度 As Boolean
    Dim lngTmpRow As Long, i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnLockd As Boolean
    Dim n As Integer

    On Error GoTo errH
    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        If gclsPros.PatiType <> PF_门诊 Then
            If LngRow <> -1 Then
                lngBegin = LngRow: lngEnd = LngRow
            Else
                lngBegin = .FixedRows: lngEnd = .Rows - 1
            End If
            
            If gclsPros.FuncType = f病案首页 Then
                If bln西医 Then '分娩设置
                    lngTmpRow = FindDiagRow(DT_病理诊断): i = FindDiagRow(DT_出院诊断XY)
                    gclsPros.CurrentForm.cmdDeliceryInfo.Visible = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Tag = ""
                    For j = i To lngTmpRow - 1
                        If .TextMatrix(j, DI_分娩信息) = "1" Then
                            gclsPros.CurrentForm.cmdDeliceryInfo.Visible = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Next
                Else
                    lngTmpRow = .Rows: i = FindDiagRow(DT_出院诊断XY)
                End If
                '疾病编码设置的治疗效果只能是其他,死亡的出院情况不调整
                For j = i To lngTmpRow - 1
                    If .TextMatrix(j, DI_是否病人) <> "1" And Val(.TextMatrix(j, DI_疾病ID)) <> 0 Then
                        If zlStr.NeedName(.TextMatrix(j, DI_出院情况)) <> "死亡" Then .TextMatrix(j, DI_出院情况) = "其他"
                    End If
                Next
            End If

            If gclsPros.FuncType <> f诊断选择 Then
                For i = lngBegin To lngEnd
                    strDiagTypeName = .TextMatrix(i, DI_诊断类型)
                    If strDiagTypeName = "" And i >= 1 Then
                        For j = i To 1 Step -1
                            strDiagTypeName = .TextMatrix(j, DI_诊断类型)
                            If strDiagTypeName <> "" Then Exit For
                        Next
                    End If
                    Select Case strDiagTypeName
                        Case "门（急）诊诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_门诊与出院XY, BCC_门诊与出院ZY))
                            If bln西医 Then Call SetDiagMatchInfo(BCC_门诊与入院)
                        Case "入院诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_入院与出院XY, BCC_入院与出院ZY))
                            If bln西医 Then Call SetDiagMatchInfo(BCC_门诊与入院)
                        Case "其他诊断", "出院诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_门诊与出院XY, BCC_门诊与出院ZY))
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_入院与出院XY, BCC_入院与出院ZY))
                            If bln西医 And strDiagTypeName = "出院诊断" Then
                                '根据出院诊断相关信息设置相关控件属性
                                strTmp = UCase(Trim(.TextMatrix(i, DI_诊断编码)))
                                If .TextMatrix(i, DI_诊断类型) = "出院诊断" Then
                                    bln分化程度 = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
                                    blnOld分化程度 = gclsPros.CurrentForm.cboBaseInfo(BCC_分化程度).Locked
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_分化程度), Not bln分化程度, True)
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_最高诊断依据), Not bln分化程度, True)
                                    If gclsPros.CurrentForm.Visible And bln分化程度 And blnOld分化程度 Then
                                        Call SetCboDefaultValue(BCC_分化程度)
                                        Call SetCboDefaultValue(BCC_最高诊断依据)
                                    End If
                                End If
                            End If
                        Case "病理诊断" '西医诊断
                            Call SetDiagMatchInfo(BCC_放射与病理)
                            Call SetDiagMatchInfo(BCC_临床与病理)
                            'Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_病理号), .TextMatrix(i, DI_诊断描述) = "", True)
                        Case "院内感染" '西医诊断
                            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_病原学检查), .TextMatrix(i, DI_诊断描述) = "", True)
                            Call chkInfoClick(CHK_病原学检查)
                    End Select
                Next
            ElseIf gclsPros.FuncType = f诊断选择 And gclsPros.PatiType = PF_住院 Then
                For i = lngBegin To lngEnd
                    strDiagTypeName = .TextMatrix(i, DI_诊断类型)
                    If strDiagTypeName = "" And i >= 1 Then
                        For j = i To 1 Step -1
                            strDiagTypeName = .TextMatrix(j, DI_诊断类型)
                            If strDiagTypeName <> "" Then Exit For
                        Next
                    End If

                    Select Case strDiagTypeName
                        Case "门（急）诊诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_门诊与出院XY, BCC_门诊与出院ZY))
                            If bln西医 Then Call SetDiagMatchInfo(BCC_门诊与入院)
                        Case "入院诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_入院与出院XY, BCC_入院与出院ZY))
                            If bln西医 Then Call SetDiagMatchInfo(BCC_门诊与入院)
                        Case "其他诊断", "出院诊断"
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_门诊与出院XY, BCC_门诊与出院ZY))
                            Call SetDiagMatchInfo(IIf(bln西医, BCC_入院与出院XY, BCC_入院与出院ZY))
                            If bln西医 And strDiagTypeName = "出院诊断" Then
                                '根据出院诊断相关信息设置相关控件属性
                                strTmp = UCase(Trim(.TextMatrix(i, DI_诊断编码)))
                                If .TextMatrix(i, DI_诊断类型) = "出院诊断" Then
                                    bln分化程度 = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
                                    blnOld分化程度 = gclsPros.CurrentForm.cboBaseInfo(BCC_分化程度).Locked
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_分化程度), Not bln分化程度, True)
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_最高诊断依据), Not bln分化程度, True)
                                    If gclsPros.CurrentForm.Visible And bln分化程度 And blnOld分化程度 Then
                                        Call SetCboDefaultValue(BCC_分化程度)
                                        Call SetCboDefaultValue(BCC_最高诊断依据)
                                    End If
                                End If
                            End If
                        Case "病理诊断" '西医诊断
                            Call SetDiagMatchInfo(BCC_放射与病理)
                            Call SetDiagMatchInfo(BCC_临床与病理)
                    End Select
                Next
            End If
        Else
            '如果填写了发病时间，则下面的发病时间则不允许填写了
            blnLockd = IsDate(.TextMatrix(.FixedRows, DI_发病时间))
            If Not blnLockd Then
                If Not bln西医 Then
                    blnLockd = IsDate(gclsPros.CurrentForm.vsDiagXY.TextMatrix(gclsPros.CurrentForm.vsDiagXY.FixedRows, DI_发病时间))
                Else
                    blnLockd = IsDate(gclsPros.CurrentForm.vsDiagZY.TextMatrix(gclsPros.CurrentForm.vsDiagZY.FixedRows, DI_发病时间))
                End If
            End If
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病日期), blnLockd, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病时间), blnLockd, True)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ChangeOutInfo(Optional ByVal str出院情况 As String, Optional ByVal blnSet出院方式 As Boolean, Optional ByVal blnCheckAll As Boolean)
'功能：调整出院情况后同步调整其他诊断的出院情况
'参数：str出院情况=出院情况，当出院情况为空且是诊断编辑触发出院情况变动时，自动检查所有诊断的出院情况
'                      规则：1、存在死亡诊断，则不用处理
'                                2、不存在死亡诊断则自动解除出院情况的锁定状态
'          blnSet出院方式=是否在设置出院方式
    Dim vsDiagTmp As VSFlexGrid
    Dim i As Long, lngStart As Long, lngEnd As Long
    Dim blnHave死亡 As Boolean
    Dim intIndex As Integer, blnLocked As Boolean, strTmp As String

     '出院情况为死亡，则医院感染，并发症，其他诊断的出院情况为死亡
    Set vsDiagTmp = gclsPros.CurrentForm.vsDiagXY
    With vsDiagTmp
        lngEnd = FindDiagRow(DT_病理诊断): lngStart = FindDiagRow(DT_出院诊断XY)
        If str出院情况 = "死亡" Then
            For i = lngStart To lngEnd - 1
                If .TextMatrix(i, DI_诊断描述) <> "" Then
                    If InStr(gclsPros.CurrentForm.txtInfo(GC_出院科室).Text, "产科") > 0 Then
                        If .TextMatrix(i, DI_诊断类型) <> "其他诊断" And .TextMatrix(i, DI_诊断类型) <> "出院诊断" Then
                            If .TextMatrix(i, DI_诊断类型) = "" And .Cell(flexcpData, i, DI_诊断类型) <> "其他诊断" Then
                                .TextMatrix(i, DI_出院情况) = "死亡"
                            End If
                        End If
                    Else
                        .TextMatrix(i, DI_出院情况) = "死亡"
                    End If
                End If
            Next
        ElseIf str出院情况 <> "" And Not blnSet出院方式 Then '不是出院方式带来的出院情况改变
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then
                    If InStr(gclsPros.CurrentForm.txtInfo(GC_出院科室).Text, "产科") > 0 Then
                        If .TextMatrix(i, DI_诊断类型) <> "其他诊断" And .TextMatrix(i, DI_诊断类型) <> "出院诊断" Then
                            If .TextMatrix(i, DI_诊断类型) = "" And .Cell(flexcpData, i, DI_诊断类型) <> "其他诊断" Then
                                .TextMatrix(i, DI_出院情况) = str出院情况
                            End If
                        End If
                    Else
                        .TextMatrix(i, DI_出院情况) = str出院情况
                    End If
                End If
            Next
        ElseIf blnSet出院方式 Then '出院方式带来的出院情况改变,将死亡的出院情况清空
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then .TextMatrix(i, DI_出院情况) = ""
            Next
        Else '出院情况为空，则自动检查是否存在死亡出院情况
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then blnHave死亡 = True: Exit For
            Next
        End If
        '疾病编码设置的治疗效果只能是其他,死亡的出院情况不调整
        If gclsPros.FuncType = f病案首页 And str出院情况 <> "死亡" And str出院情况 <> "" Then
            For i = lngStart To lngEnd - 1
                If .TextMatrix(i, DI_是否病人) <> "1" And Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                    .TextMatrix(i, DI_出院情况) = "其他"
                End If
            Next
        End If
    End With
    '处理中医科时，中医诊断，西医诊断只有两者之一时死亡相关的数据可能会被清空
    If gclsPros.IsTCM Then
        Set vsDiagTmp = gclsPros.CurrentForm.vsDiagZY
        With vsDiagTmp
            lngEnd = .Rows: lngStart = FindDiagRow(DT_出院诊断ZY)
            If str出院情况 = "死亡" Then
                For i = lngStart To lngEnd - 1
                    If .TextMatrix(i, DI_诊断描述) <> "" Then .TextMatrix(i, DI_出院情况) = "死亡"
                Next
            ElseIf str出院情况 <> "" And Not blnSet出院方式 Then '不是出院方式带来的出院情况改变
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then .TextMatrix(i, DI_出院情况) = str出院情况
                Next
            ElseIf blnSet出院方式 Then '出院方式带来的出院情况改变,将死亡的出院情况清空
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then .TextMatrix(i, DI_出院情况) = ""
                Next
            ElseIf Not blnHave死亡 Then
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_出院情况)) = "死亡" Then blnHave死亡 = True: Exit For
                Next
            End If
            '疾病编码设置的治疗效果只能是其他,死亡的出院情况不调整
            If gclsPros.FuncType = f病案首页 And str出院情况 <> "死亡" And str出院情况 <> "" Then
                For i = lngStart To lngEnd - 1
                    If .TextMatrix(i, DI_是否病人) <> "1" And Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                        .TextMatrix(i, DI_出院情况) = "其他"
                    End If
                Next
            End If
        End With
    End If
    '诊断选择
    If gclsPros.FuncType = f诊断选择 And gclsPros.PatiType = PF_住院 Then
        If blnHave死亡 Then str出院情况 = "死亡"
        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_死亡时间), str出院情况 <> "死亡", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_死亡原因), str出院情况 <> "死亡", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_死亡原因), str出院情况 <> "死亡", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检), str出院情况 <> "死亡", True)
        If gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).ListIndex = -1 Then
            gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).ListIndex = 0
        End If
        Exit Sub
    End If
    '设置出院方式时为死亡不锁定出院方式，设置诊断时为死亡则锁定
    If Not blnSet出院方式 Then
        If str出院情况 <> "" Then blnHave死亡 = str出院情况 = "死亡"
        gblnSet = True
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式), blnHave死亡)
        If blnHave死亡 Then
            '如果是死亡，则出院情况必须为死亡
            intIndex = Cbo.FindIndex(gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式), "死亡")
            If intIndex = -1 Then
                gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).AddItem "死亡"
                intIndex = gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).NewIndex
            End If
            gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).ListIndex = intIndex
            blnLocked = True
            str出院情况 = "死亡"
        ElseIf str出院情况 <> "" And zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).Text) = "死亡" Then
            gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).ListIndex = -1
            blnLocked = True
        Else
            str出院情况 = zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_出院方式).Text)
            blnLocked = Not (str出院情况 Like "*转院*" Or str出院情况 Like "*转社区*")
        End If
        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_转入医疗机构), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_转入医疗机构), blnLocked, True)
        gblnSet = False
    End If
    '出院情况变动导致其他的控件状态改变
    Call ChangeOutInfoSub(str出院情况)
End Sub

Public Sub ChangeOutInfoSub(Optional ByVal str出院情况 As String)
    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_死亡时间), str出院情况 <> "死亡", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_死亡原因), str出院情况 <> "死亡", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_死亡原因), str出院情况 <> "死亡", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_死亡期间), str出院情况 <> "死亡", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检), str出院情况 <> "死亡", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_随诊), str出院情况 = "死亡", True)
    If str出院情况 = "死亡" Then
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).Clear
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "无"
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "有"
    Else
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).Clear
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).AddItem "-"
    End If
    
    If gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).ListIndex = -1 Then
        gclsPros.CurrentForm.cboBaseInfo(BCC_死亡患者尸检).ListIndex = 0
    End If
    Call chkInfoClick(CHK_随诊)
End Sub

Public Sub EnterNextCellDiag(ByRef vsDiagTmp As VSFlexGrid)
    Dim i As Long, j As Long

    With vsDiagTmp
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, DI_诊断编码) To DI_Del
                If Not .ColHidden(j) Then
                    If DiagCellEditable(vsDiagTmp, i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= DI_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > DI_Del And .TextMatrix(.Rows - 1, DI_诊断描述) <> "" Then
            .Rows = .Rows + 1
            Call ChangeVSFHeight(vsDiagTmp, True)
            .TextMatrix(.Rows - 1, DI_诊断分类) = .TextMatrix(.Rows - 2, DI_诊断分类)
            If gclsPros.PatiType = PF_门诊 Then .TextMatrix(.Rows - 1, DI_诊断类型) = .TextMatrix(.Rows - 2, DI_诊断类型)
            .ShowCell i, IIf(gclsPros.FuncType = f病案首页, DI_诊断编码, DI_诊断描述)
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Private Sub EnterNextCellSpirit()
    '------------------------------------------------------------------------------------------------------
    '功能:移动列
    '入参:
    '出参:
    '返回:
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim vsSpiritTmp As VSFlexGrid

    Set vsSpiritTmp = gclsPros.CurrentForm.vsSpirit
    With vsSpiritTmp
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then .Row = .FixedRows
         If .Col = .Cols - 1 Then
            .Col = SI_药物名称
            If .Row = .Rows - 1 Then
                If Trim(.TextMatrix(.Row, SI_药物名称)) <> "" Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    Call ChangeVSFHeight(vsSpiritTmp, True)
                End If
            Else
                .Row = .Row + 1
            End If
         Else
            .Col = .Col + 1
         End If
         If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
         End If
         If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
         End If
    End With
End Sub

Private Sub EnterNextCellOPS(ByRef vsOPS As VSFlexGrid)
    Dim i As Long, j As Long

    With vsOPS
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, PI_手术日期) To PI_切口愈合
                If OPSCellEditable(i, j) Then Exit For
            Next
            If j <= PI_切口愈合 Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Private Sub EnterNextCellFees(ByRef vsFree As VSFlexGrid)
    Dim i As Long, j As Long

    With vsFree
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = 0 To 5 Step 2
                If Not (j <= .Col \ 2 * 2 And i = .Row) Then
                    If Not FreeHaveLowLevel(i, j) Then Exit For
                    If .TextMatrix(i, j) <> "" Then Exit For
                End If
            Next
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Public Sub FormResize()
    On Error Resume Next
    With gclsPros.CurrentForm
        If .ScaleWidth < .picMain.Width Then
            .hsbMain.Visible = True
            .picMain.Left = .ScaleLeft + ((.ScaleWidth - .picMain.Width) * ((.hsbMain.Value) / 100))
        Else
            .hsbMain.Visible = False
            .picMain.Left = .ScaleLeft + (.ScaleWidth - .picMain.Width) / 2
        End If

        .vsbMain.Move .ScaleWidth - .vsbMain.Width, .ScaleTop, .vsbMain.Width, .ScaleHeight
        .vsbMain.LargeChange = 100
        .vsbMain.SmallChange = .vsbMain.LargeChange / 2
        
        .hsbMain.Top = .vsbMain.Top + .vsbMain.Height - 255
        .hsbMain.Left = .ScaleLeft
        .hsbMain.Width = .ScaleLeft + .ScaleWidth - 255
        .hsbMain.LargeChange = 100 / ((.picMain.Width) / .ScaleWidth)
        .hsbMain.SmallChange = 10

        .cmdTop.Move .ScaleWidth - .cmdTop.Width - .vsbMain.Width, .ScaleHeight - .cmdTop.Height - 400
        Call vsbMainChange
    End With
End Sub

Public Sub vsbMainChange()
    With gclsPros.CurrentForm
        .picMain.Top = 500 - ((.picMain.Height + 1100 - .ScaleHeight) * (.vsbMain.Value / 1000))
        If .vsbMain.Value > 0 Then
            .cmdTop.Visible = True
        Else
            .cmdTop.Visible = False
        End If
    End With
End Sub

Public Sub hsbMainChange()
    With gclsPros.CurrentForm
        .picMain.Left = .ScaleLeft + ((.ScaleWidth - .picMain.Width) * ((.hsbMain.Value) / 100))
    End With
End Sub

Public Sub cmdTopClick()
    gclsPros.CurrentForm.vsbMain.Value = 0
End Sub

Public Sub cmdTopGotFocus()
    Call ShowInfectInfo(False)
End Sub

Public Sub txtAdressInfoChange(ByRef Index As Integer)
    Call CheckValueChange(gclsPros.CurrentForm.txtAdressInfo(Index))
End Sub
Public Sub PicPageResize(ByVal Index As Integer)
    With gclsPros.CurrentForm
        If gclsPros.FuncType = f病案首页 Then
            If Index = PIC_基本信息 Then
                .PicOut.Move 0, .vsTransfer.Top + .vsTransfer.Height + 200
            ElseIf Index = PIC_手术记录 Then
                .PicOPS.Move 0, .vsOPS.Top + .vsOPS.Height + 200
            End If
        ElseIf gclsPros.FuncType = f医生首页 Then
            If gclsPros.MedPageSandard = ST_云南省标准 Then
                If Index = PIC_手术记录 Then
                    .PicOPS.Move 0, .vsOPS.Top + .vsOPS.Height + 200
                End If
            End If
        End If
        
        If gclsPros.MedPageSandard = ST_四川省标准 Then
            If Index = PIC_重症监护 Then
                .lblICUInstruments.Top = IIf(gclsPros.FuncType = f医生首页, 150, 250) + .vsFlxAddICU.Top + .vsFlxAddICU.Height
                .vsICUInstruments.Top = IIf(gclsPros.FuncType = f医生首页, 350, 500) + .vsFlxAddICU.Top + .vsFlxAddICU.Height
            End If
        End If
    End With
End Sub

Public Sub SubCMainWndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    '自定义的消息处理函数
    Dim wzDelta As Integer
    Select Case msg
        Case WM_MOUSEWHEEL   '滚动
            wzDelta = HIWORD(wParam)
            With gclsPros.CurrentForm
                If wzDelta > 0 Then        '向上滚动
                    Call ChangePage(False, , , False)
                Else                        '向下滚动
                    Call ChangePage(True, , , False)
                End If
            End With
    End Select
End Sub

Public Function SetErrObjectColor(ByVal strErrID As String, Optional ByVal blnOld As Boolean, Optional ByVal colorBack As Long) As Object
'功能：定位到控件所在的页面，设置控件的颜色
    Dim clsErrTmp As clsErrInfo
    Dim i As Long
    Dim objErrArr() As ERROBJ
    Dim objTmp As Object
    Dim vsfTmp As VSFlexGrid
    
    If InStr(strErrID, "Error-") > 0 Then
        Set clsErrTmp = gColErr.Item(strErrID)
    ElseIf InStr(strErrID, "Warn-") > 0 Then
        Set clsErrTmp = gColWarn.Item(strErrID)
    Else
        Exit Function
    End If
    
    If clsErrTmp Is Nothing Then
        Exit Function
    End If
    ReDim objErrArr(UBound(clsErrTmp.GetObjErr()) - LBound(clsErrTmp.GetObjErr()) + 1)
    For i = LBound(clsErrTmp.GetObjErr()) To UBound(clsErrTmp.GetObjErr())
        objErrArr(i) = clsErrTmp.GetObjErr(i)
        
        Select Case objErrArr(i).StrObjName
            Case "txtInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtSpecificInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtSpecificInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtDateInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtDateInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtAdressInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtAdressInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboBaseInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboBaseInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboSpecificInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboSpecificInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboManInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboManInfo(objErrArr(i).LngObjIndex)
                End If
            Case "chkInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.chkInfo(objErrArr(i).LngObjIndex)
                End If
            Case "mskDateInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.mskDateInfo(objErrArr(i).LngObjIndex)
                End If
            Case "padrInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.padrInfo(objErrArr(i).LngObjIndex)
                End If
            Case "lstInfection"
                Set objTmp = gclsPros.CurrentForm.lstInfection
            Case "lstAdvEvent"
                Set objTmp = gclsPros.CurrentForm.lstAdvEvent
            Case "vsDiagXY"
                Set objTmp = gclsPros.CurrentForm.vsDiagXY
            Case "vsDiagZY"
                Set objTmp = gclsPros.CurrentForm.vsDiagZY
            Case "vsAller"
                Set objTmp = gclsPros.CurrentForm.vsAller
            Case "vsOPS"
                Set objTmp = gclsPros.CurrentForm.vsOPS
            Case "vsChemoth"
                Set objTmp = gclsPros.CurrentForm.vsChemoth
            Case "vsRadioth"
                Set objTmp = gclsPros.CurrentForm.vsRadioth
            Case "vsSpirit"
                Set objTmp = gclsPros.CurrentForm.vsSpirit
            Case "vsKSS"
                Set objTmp = gclsPros.CurrentForm.vsKSS
            Case "vsFlxAddICU"
                Set objTmp = gclsPros.CurrentForm.vsFlxAddICU
            Case "vsICUInstruments"
                Set objTmp = gclsPros.CurrentForm.vsICUInstruments
            Case "vsfMain"
                Set objTmp = gclsPros.CurrentForm.vsfMain
            Case "vsTSJC"
                Set objTmp = gclsPros.CurrentForm.vsTSJC
            Case "vsTransfer"
                Set objTmp = gclsPros.CurrentForm.vsTransfer
            Case "vsFees"
                Set objTmp = gclsPros.CurrentForm.vsFees
            Case "vsInfect"
                Set objTmp = gclsPros.CurrentForm.vsInfect
            Case "vsSample"
                Set objTmp = gclsPros.CurrentForm.vsSample
            Case Else
                '加载外挂部件对象
                If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                    Err.Clear: On Error Resume Next
                    If objErrArr(i).LngObjIndex = -1 Then
                        Set objTmp = colErrTmp(objErrArr(i).StrObjName & objErrArr(i).PicIndex)
                    Else
                        Set objTmp = colErrTmp(objErrArr(i).StrObjName & objErrArr(i).PicIndex & objErrArr(i).LngObjIndex)
                    End If
                    Err.Clear: On Error GoTo 0
                End If
        End Select
        If Not objTmp Is Nothing Then
            If blnOld Then
                If TypeName(objTmp) = "VSFlexGrid" Then
                    Set vsfTmp = objTmp
                    If vsfTmp.Rows > objErrArr(i).LngRow And vsfTmp.Cols > objErrArr(i).LngCol Then
                        vsfTmp.Cell(flexcpBackColor, objErrArr(i).LngRow, objErrArr(i).LngCol, objErrArr(i).LngRow, objErrArr(i).LngCol) = objErrArr(i).OldColor
                    End If
                Else
                    objTmp.BackColor = objErrArr(i).OldColor
                End If
            Else
                If TypeName(objTmp) = "VSFlexGrid" Then
                    Set vsfTmp = objTmp
                    vsfTmp.Cell(flexcpBackColor, objErrArr(i).LngRow, objErrArr(i).LngCol, objErrArr(i).LngRow, objErrArr(i).LngCol) = colorBack
                    vsfTmp.Row = objErrArr(i).LngRow
                    vsfTmp.Col = objErrArr(i).LngCol
                    gclsPros.CurrentForm.picMain.SetFocus
                    Call LocateObjectPage(vsfTmp)
                Else
                    objTmp.BackColor = colorBack
                    gclsPros.CurrentForm.picMain.SetFocus
                    Call LocateObjectPage(objTmp)
                End If
            End If
        End If
    Next

End Function

Public Sub LocateObjectPage(ByRef objTmp As Object)
'功能：根据控件定位到该控件所在的那一页
'参数: objTmp - 要定位到所在页的控件
    Dim intIndex As Integer
    Dim picTmp As PictureBox
    Dim lngObjTop As Long
    Dim strName As String
    Dim i As Integer
    
On Error GoTo errH
    If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
        lngObjTop = gclsPros.CurrentForm.picMain.Top
        If objTmp.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Top
        ElseIf objTmp.Container.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Top
        ElseIf objTmp.Container.Container.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container.Container.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Container.Container.Top + objTmp.Top
        End If
        
        If Not picTmp Is Nothing Then
    '        在界面上看得见控件的话就不翻页
            If Not (lngObjTop > 0 And lngObjTop + objTmp.Height < gclsPros.CurrentForm.Height) Then
                intIndex = picTmp.Index
                Call ChangePage(, intIndex, objTmp)
            Else
                objTmp.SetFocus
            End If
        End If
    End If
    Exit Sub
errH:
    Err.Clear
    '定位外挂附页
    If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
            strName = ""
            For i = 1 To gfrmMecCol.Count
                strName = strName & "," & gfrmMecCol(i).Name
            Next
            If InStr(strName, objTmp.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Top
                
            ElseIf InStr(strName, objTmp.Container.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Top
                
            ElseIf InStr(strName, objTmp.Container.Container.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Container.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Container.Container.Top + objTmp.Top
            End If
            
        End If
        
        If Not picTmp Is Nothing Then
    '        在界面上看得见控件的话就不翻页
            If Not (lngObjTop > 0 And lngObjTop + objTmp.Height < gclsPros.CurrentForm.Height) Then
                intIndex = picTmp.Index
                Call ChangePage(, intIndex, objTmp)
            Else
                objTmp.SetFocus
            End If
        End If
    End If
End Sub

Public Sub VsErrClick(ByVal strErrID As String)
'功能：根据所点击的错误或者警告信息定位到具体控件，并设置控件的额颜色
    Dim clsErrTmp As clsErrInfo
    Dim i As Long
    Dim objErrArr() As ERROBJ
    Dim objTmp As Object
    Dim picTmp As PictureBox
    Dim intIndex As Integer
    Static strOldIndex  As String
    
    If strOldIndex <> "" And strOldIndex <> strErrID Then
        If InStr(strOldIndex, "Error-") > 0 Then
            Set clsErrTmp = gColErr.Item(strOldIndex)
        ElseIf InStr(strOldIndex, "Warn-") > 0 Then
            Set clsErrTmp = gColWarn.Item(strOldIndex)
        End If
        If Not clsErrTmp Is Nothing Then
            Call SetErrObjectColor(strOldIndex, True)
            Set clsErrTmp = Nothing
        End If
    End If
    
    
    If InStr(strErrID, "Error-") > 0 Then
        Set clsErrTmp = gColErr.Item(strErrID)
        strOldIndex = strErrID
    ElseIf InStr(strErrID, "Warn-") > 0 Then
        Set clsErrTmp = gColWarn.Item(strErrID)
        strOldIndex = strErrID
    Else
        strOldIndex = ""
        Exit Sub
    End If
    
    If Not clsErrTmp Is Nothing Then
        Call SetErrObjectColor(strErrID, False, vbRed)
    End If

End Sub

Public Function SetAllObject() As Boolean
'功能：设置一些控件的状态属性
    Dim objTmp As Object
    Dim strName As String
    
    If gclsPros.FuncType <> f医生首页 And gclsPros.FuncType <> f病案首页 Then
        Exit Function
    End If
    gclsPros.CurrentForm.picMain.Top = gclsPros.CurrentForm.ScaleTop + 500
    For Each objTmp In gclsPros.CurrentForm.Controls
        If InStr(",VScrollBar,HScrollBar,Subclass,Line,Image,PatiAddress,", "," & TypeName(objTmp) & ",") < 1 Then
            objTmp.Appearance = 0
        End If
        If InStr(",Frame,PictureBox,CheckBox,TextBox,OptionButton,", "," & TypeName(objTmp) & ",") > 0 Then
            If TypeName(objTmp) <> "TextBox" Then
                objTmp.BackColor = GPAGECOLOR
            Else
                If Not objTmp.Locked Then objTmp.BackColor = GPAGECOLOR
            End If
            If TypeName(objTmp) = "PictureBox" Then
                objTmp.AutoRedraw = True
            End If
        End If
        If TypeName(objTmp) = "Label" Then
            objTmp.BorderStyle = 0
            objTmp.BackStyle = 0
            objTmp.Appearance = 0
        ElseIf TypeName(objTmp) = "TextBox" Then
            objTmp.BorderStyle = 0
        ElseIf TypeName(objTmp) = "PictureBox" Then
            objTmp.TabStop = False
        End If
        If objTmp.Name = "mskDateInfo" Then
            If objTmp.Index = DC_确诊日期 Then
                strName = zlRegInfo("单位名称")
                If InStr(strName, "平果") > 0 Then
                    objTmp.Mask = "####-##-##"
                    objTmp.Tag = "####-##-##"
                End If
            End If
            If objTmp.Index = DC_死亡时间 Then
                objTmp.Mask = "####-##-## ##:##"
                objTmp.Tag = "####-##-## ##:##"
            End If
        End If
        If objTmp.Name = "lblTitle" Then
            objTmp.Move 0, 10
        ElseIf objTmp.Name = "lineH" Then
            objTmp.BorderStyle = 3
        End If
    Next
    With gclsPros.CurrentForm
        If gclsPros.FuncType = f病案首页 Then
            .lblSpecificInfo(SLC_住院号).ForeColor = vbBlue
        ElseIf gclsPros.FuncType = f医生首页 Then
            .lblAutoInfo.ForeColor = vbBlue
        End If
        .lblEdit(0).ForeColor = vbBlue
        .lblEdit(1).ForeColor = vbBlue
    End With
End Function

Public Function SetAllVSF(Optional blnPic As Boolean) As Boolean
'功能：在PictureBox 上面调整VSF控件的大小和位置
    Dim objTmp As Object
    Dim vsfTmp As VSFlexGrid
    Dim strVSFName As String
    
    If gclsPros.FuncType <> f医生首页 And gclsPros.FuncType <> f病案首页 Then
        Exit Function
    End If
    
    For Each objTmp In gclsPros.CurrentForm.Controls
        If TypeName(objTmp) = "VSFlexGrid" Then
            Set vsfTmp = objTmp
            strVSFName = vsfTmp.Name
            
            vsfTmp.SelectionMode = flexSelectionFree
            vsfTmp.FocusRect = flexFocusSolid
            vsfTmp.HighLight = flexHighlightWithFocus
            vsfTmp.BackColorSel = &H404040

            vsfTmp.Left = -10
            If strVSFName = "vsOPS" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 1000, 2)
                vsfTmp.Width = vsfTmp.Container.Width - 400
            ElseIf strVSFName = "vsfMain" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 30)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsDiagXY" Or strVSFName = "vsDiagZY" Then
                Call ChangeVSFHeight(vsfTmp, blnPic)
                vsfTmp.Width = vsfTmp.Container.Width - 400
            ElseIf strVSFName = "vsTSJC" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 0)
            ElseIf strVSFName = "vsInfect" Then
                 Call ChangeVSFHeight(vsfTmp, blnPic)
                 vsfTmp.Width = vsfTmp.Container.Width / 2 - 100
            ElseIf strVSFName = "vsSample" Then
                 Call ChangeVSFHeight(vsfTmp, blnPic)
                 vsfTmp.Left = vsfTmp.Container.Width / 2 + 50
                 vsfTmp.Width = vsfTmp.Container.Width / 2 - 40
            ElseIf strVSFName = "vsTransfer" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 20, 2)
                vsfTmp.Left = 1250
            ElseIf strVSFName = "vsAller" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 300, 3)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsChemoth" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, , 2)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsRadioth" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, , 2)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            Else
                Call ChangeVSFHeight(vsfTmp, blnPic)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            End If
        End If
    Next
End Function

Public Function ChangeVSFHeight(ByRef vsfTmp As VSFlexGrid, Optional ByVal blnPic As Boolean, Optional ByVal lngAddheight As Long = -1, Optional ByVal lngMinRows As Long = 3) As Boolean
'功能：在PictureBox 上面调整VSF的大小
    Dim i As Long
    Dim lngOldVSFHeight As Long
    Dim picContainer As PictureBox
    Dim lngRows As Long
    Dim lngVSFHeight As Long
    Dim lngRowHeight As Long
    Dim lngMaxHeight As Long
    
    If gclsPros.FuncType <> f医生首页 And gclsPros.FuncType <> f病案首页 Then
        Exit Function
    End If
    Call CheckValueChange(vsfTmp)
    lngOldVSFHeight = vsfTmp.Height
    Set picContainer = vsfTmp.Container
    lngRowHeight = IIf(vsfTmp.RowHeightMax < vsfTmp.RowHeightMin, vsfTmp.RowHeightMin, vsfTmp.RowHeightMax)

    lngRows = vsfTmp.Rows
    If lngRows < lngMinRows Then lngRows = lngMinRows: vsfTmp.Rows = lngMinRows
    For i = 0 To vsfTmp.Rows - 1
        lngVSFHeight = lngVSFHeight + vsfTmp.RowHeight(i)
    Next
    lngVSFHeight = IIf(lngVSFHeight < lngRows * lngRowHeight, lngRows * lngRowHeight, lngVSFHeight)
    If lngAddheight = -1 Then lngAddheight = lngRowHeight * 1.5
    vsfTmp.Height = lngVSFHeight + lngAddheight
    
    If vsfTmp.Name = "vsInfect" Or vsfTmp.Name = "vsSample" Then
        lngMaxHeight = IIf(gclsPros.CurrentForm.vsInfect.Height > gclsPros.CurrentForm.vsSample.Height, gclsPros.CurrentForm.vsInfect.Height, gclsPros.CurrentForm.vsSample.Height)
        If vsfTmp.Height - lngOldVSFHeight <> 0 Then
            picContainer.Height = vsfTmp.Top + lngMaxHeight + 300
            If blnPic Then
                Call SetPicPosition(True)
            End If
        End If
    ElseIf vsfTmp.Height - lngOldVSFHeight <> 0 Then
        picContainer.Height = picContainer.Height + (vsfTmp.Height - lngOldVSFHeight)
        If blnPic Then
            Call SetPicPosition(True)
        End If
    End If
End Function


Public Function SetPicPosition(Optional ByVal blnV As Boolean, Optional ByVal blnH As Boolean) As Boolean
'功能：在PictureBox 上面调整每个PicPage控件的位置
    Dim lngLeft As Long
    Dim i As Long, j As Long, lngHeight As Long
    
    With gclsPros.CurrentForm
        lngLeft = .picMain.ScaleLeft + ((.picMain.ScaleWidth - .PicPage(0).Width) / 2)

        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                .PicPage(i).Visible = True
                If i = .PicPage.LBound Then
                    .PicPage(i).Move lngLeft, .picMain.ScaleTop
                Else
                    .PicPage(i).Move lngLeft, .PicPage(j).Top + .PicPage(j).Height
                End If
                j = i
                lngHeight = lngHeight + .PicPage(i).Height
            Else
                .PicPage(i).Visible = False
            End If
        Next
        .picMain.Height = .PicPage(0).ScaleTop + lngHeight + 500
    
        Call DrawLine(blnV, blnH)
    End With
End Function

Private Sub DrawLine(Optional ByVal blnV As Boolean, Optional ByVal blnH As Boolean)
'功能：在整个页面画上边框，在每一个PicPage最上面画上分割线，TextBox控件下面面画直线,在ComboBox外面加套Frame
'参数: blnV -画竖着的边框线, blnH - 每一个PicPage最上面画上分割线，TextBox控件下面面画直线,在ComboBox外面加套Frame
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, j As Long
    Dim objText As Object
    Dim objPic As PictureBox
    Dim lngMin As Long, lngMax As Long
    Dim blnFra As Boolean
    Dim cboTmp As ComboBox
    On Error Resume Next
    
    With gclsPros.CurrentForm
         For i = .PicPage.LBound To .PicPage.UBound
            If blnH Then
                .PicPage(i).Cls
            End If
            If .PicPage(i).Tag = "true" Then
                lngMax = i
            End If
        Next
        lngMin = .PicPage.LBound + 1
        '整个页面的边框线
        If blnV Then
            .picMain.Cls
            '整个页面的左边竖线
            .picMain.DrawWidth = 1
            x1 = .PicPage(lngMin).Left - 15
            y1 = .PicPage(lngMin).Top
            x2 = x1
            y2 = .PicPage(lngMax).Top + .PicPage(lngMax).Height
            .picMain.Line (x1, y1)-(x2, y2)
            '整个页面的右边竖线
            x1 = .PicPage(lngMin).Left + .PicPage(lngMin).Width + 5
            y1 = .PicPage(lngMin).Top
            x2 = x1
            y2 = .PicPage(lngMax).Top + .PicPage(lngMax).Height
            .picMain.Line (x1, y1)-(x2, y2)
        End If
        
        '每个PictureBox的上边画一条横线
        If blnH Then
            For i = lngMin To lngMax
                If .PicPage(i).Tag = "true" Then
                    x1 = .PicPage(i).ScaleLeft
                    y1 = .PicPage(i).ScaleTop
                    x2 = .PicPage(i).ScaleLeft + .PicPage(i).ScaleWidth
                    y2 = y1
                    .PicPage(i).DrawWidth = 1
                    .PicPage(i).Line (x1, y1)-(x2, y2)
                    If i = lngMax Then
                        .PicPage(i).Line (x1, y1 + .PicPage(i).ScaleHeight - 10)-(x2, y2 + .PicPage(i).ScaleHeight - 10)
                    End If
                End If
            Next
'            下面这句代码设置在ComboBox外面加套Frame 只执行一次
            blnFra = (.fraCbo.UBound = 0)
            
            For Each objText In .Controls
                '在每个TextBox 下面画一条线
                If TypeName(objText) = "TextBox" Then
                    If objText.Name <> "txtAdressInfo" Then
                        DrawLineCTL objText
                    ElseIf objText.Name = "txtAdressInfo" Then
                        If gclsPros.IsStructAdress Then
                            If objText.Index = ADRC_单位地址 Then
                                If gclsPros.MedPageSandard <> ST_四川省标准 Then
                                    DrawLineCTL objText
                                End If
                            ElseIf objText.Index = ADRC_病人区域 Then
                                DrawLineCTL objText
                            End If
                        Else
                            DrawLineCTL objText
                        End If
                    End If
                ElseIf TypeName(objText) = "ComboBox" Then  '在每一个ComboBox的外面套一个Frame，使之看起来像平面的
                    If blnFra And TypeName(objText.Container) = "PictureBox" Then
                        Set cboTmp = objText
                        j = j + 1
                        Load .fraCbo(j)
                        Set .fraCbo(j).Container = cboTmp.Container
                        .fraCbo(j).Left = cboTmp.Left
                        .fraCbo(j).Top = cboTmp.Top + 25
                        .fraCbo(j).Width = cboTmp.Width
                        .fraCbo(j).Height = IIf(gclsPros.FuncType = f病案首页, 250, 225)
                        .fraCbo(j).BackColor = GPAGECOLOR
                        If cboTmp.Tag = "年龄" Then
                            .fraCbo(j).Visible = False
                        Else
                            .fraCbo(j).Visible = True
                        End If
                        Set cboTmp.Container = .fraCbo(j)
                        cboTmp.Width = cboTmp.Width + 50
                        cboTmp.Left = -25
                        cboTmp.Top = -25
                    End If
                End If
            Next
        
            For j = .fraCbo.LBound + 1 To .fraCbo.UBound
                DrawLineCTL .fraCbo(j)
            Next
        End If
    End With
End Sub

Private Sub DrawLineCTL(ByRef objCtl As Object, Optional ByVal bytModel As Byte = 0)
'功能:给指定对象画一条线或清除此原有线条
'objCtl-传入控件对象，根据该控件对象获取对应坐标值
'bytModel=0-画线;1-清除线
    Dim objPic As Object  '容器
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    Select Case TypeName(objCtl)
    Case "Frame"
        'FraCbo下面画一条线
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width - 20
        y2 = y1
    Case "TextBox"
        '在每个TextBox 下面画一条线
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width
        y2 = y1
    End Select
    Set objPic = objCtl.Container
    objPic.DrawWidth = 1
    If bytModel = 0 Then
        objPic.Line (x1, y1)-(x2, y2)
    Else
        objPic.Line (x1, y1)-(x2, y2), objPic.BackColor '清除线条
    End If
End Sub

Public Sub LoadVsErrData()
'功能：将错误信息和警告信息加载到界面上
    Dim clsErr As clsErrInfo
    If gColErr.Count <= 0 And gColWarn.Count <= 0 Then Exit Sub
    frmMain.dkpMain.FindPane(Pane_检查).Closed = False
    With frmMain.vsErr
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = 0
        .Rows = .FixedRows
        If gColErr.Count > 0 Then
            .Rows = .Rows + 1
            .MergeCells = flexMergeFree
            .TextMatrix(.Rows - 1, ERR_ID) = "错误（" & CStr(gColErr.Count) & "个）"
            .TextMatrix(.Rows - 1, ERR_类型) = "错误（" & CStr(gColErr.Count) & "个）"
            .TextMatrix(.Rows - 1, ERR_信息) = "错误（" & CStr(gColErr.Count) & "个）"
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpAlignment, .Rows - 1, ERR_ID, .Rows - 1, ERR_信息) = flexAlignLeftCenter
            .Cell(flexcpFontBold, .Rows - 1, ERR_ID, .Rows - 1, ERR_信息) = True
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 0
            For Each clsErr In gColErr
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, ERR_ID) = "  " & Mid(clsErr.strErrID, InStr(clsErr.strErrID, "-") + 1)
                .Cell(flexcpData, .Rows - 1, ERR_ID) = clsErr.strErrID
                .TextMatrix(.Rows - 1, ERR_类型) = "错误"
                .TextMatrix(.Rows - 1, ERR_信息) = clsErr.StrErrInfo
                Set .Cell(flexcpPicture, .Rows - 1, 0) = frmMain.imgError.Picture
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 0) = flexAlignLeftCenter
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
            Next
        End If
        
        If gColWarn.Count > 0 Then
            .Rows = .Rows + 1
            .MergeCells = flexMergeFree
            .TextMatrix(.Rows - 1, ERR_ID) = "警告（" & CStr(gColWarn.Count) & "个）"
            .TextMatrix(.Rows - 1, ERR_类型) = "警告（" & CStr(gColWarn.Count) & "个）"
            .TextMatrix(.Rows - 1, ERR_信息) = "警告（" & CStr(gColWarn.Count) & "个）"
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpAlignment, .Rows - 1, ERR_ID, .Rows - 1, ERR_信息) = flexAlignLeftCenter
            .Cell(flexcpFontBold, .Rows - 1, ERR_ID, .Rows - 1, ERR_信息) = True
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 0
            For Each clsErr In gColWarn
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, ERR_ID) = "  " & Mid(clsErr.strErrID, InStr(clsErr.strErrID, "-") + 1)
                .Cell(flexcpData, .Rows - 1, ERR_ID) = clsErr.strErrID
                .TextMatrix(.Rows - 1, ERR_类型) = "警告"
                .TextMatrix(.Rows - 1, ERR_信息) = clsErr.StrErrInfo
                Set .Cell(flexcpPicture, .Rows - 1, 0) = frmMain.imgWarn.Picture
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 0) = flexAlignLeftCenter
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
            Next
        End If
        
        .Cell(flexcpForeColor, .FixedRows, ERR_ID, .Rows - 1, ERR_类型) = vbRed
        .Cell(flexcpForeColor, .FixedRows, ERR_信息, .Rows - 1, ERR_信息) = vbBlue
    End With
End Sub

Public Sub ClearErrCol()
'功能：清除掉错误信息和警告信息
    Dim i As Long
    
    Call VsErrClick("")
    
    If gColErr.Count > 0 Then
        For i = 1 To gColErr.Count
            gColErr.Remove 1
        Next
    End If
    If gColWarn.Count > 0 Then
        For i = 1 To gColWarn.Count
            gColWarn.Remove 1
        Next
    End If
    
    frmMain.vsErr.Rows = frmMain.vsErr.FixedRows
    frmMain.dkpMain.FindPane(Pane_检查).Closed = True
End Sub

Public Sub menuPageOperate(ByVal mopType As MedRec_Operate, Optional ByVal intPage As Integer)
    Dim strMsg As String
    Select Case mopType
        Case MOP_预览
            strMsg = "预览"
        Case MOP_打印
            strMsg = "打印"
        Case MOP_确定
            strMsg = "保存"
    End Select
    If PageOperate(mopType, intPage) Then
        strMsg = strMsg & "成功！"
    Else
        If gColErr.Count > 0 Then
            strMsg = strMsg & "失败，发现" & CStr(gColErr.Count) & "个错误，" & CStr(gColWarn.Count) & "个警告！"
        Else
            strMsg = strMsg & " 失败！"
        End If
    End If
    frmMain.stbThis.Panels(2).Text = strMsg
End Sub

Private Sub SetComboBoxProperty(ByVal blnMask As Boolean)
'功能：界面加载完毕之后发现有Combobox控件的内容处于选中状态，取消这种选中状态
'参数：blnMask - true 屏蔽掉ComboBox控件的鼠标滚轮事件，false 取消对ComboBox控件的鼠标滚轮事件的屏蔽
    Dim cboTmp As ComboBox
    Dim objTmp As Object
    
    If gclsPros.CboMask = blnMask Then
        Exit Sub
    End If
    gclsPros.CboMask = blnMask
    For Each objTmp In gclsPros.CurrentForm.Controls
        If TypeName(objTmp) = "ComboBox" Then
            Set cboTmp = objTmp
            If blnMask Then
                Call CallHook(cboTmp.hwnd)
                If cboTmp.Style = 0 Then cboTmp.SelLength = 0
            Else
                Call CallUnhook(cboTmp.hwnd)
            End If
        End If
    Next
End Sub



Public Sub SetYoubian(Index As Integer, intLevel As Integer, rsReturn As ADODB.Recordset)
'功能：在输入病人结构化地址的时候,加载邮编
    If (Not rsReturn Is Nothing) And intLevel = 2 Then
        If Index = ADRC_现住址 Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_家庭邮编).Text = rsReturn!邮编 & ""
        ElseIf Index = ADRC_户口地址 Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_户口邮编).Text = rsReturn!邮编 & ""
        ElseIf Index = ADRC_单位地址 Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_单位邮编).Text = rsReturn!邮编 & ""
        End If
    End If
End Sub

Public Sub DiagMouseDown(ByRef vsDiag As VSFlexGrid, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
    Dim LngRow As Long
    If intButton = 2 Then
        vsDiag.SetFocus
        LngRow = vsDiag.MouseRow
        If LngRow >= vsDiag.FixedRows And LngRow <= vsDiag.Rows - 1 Then
            If Not vsDiag.RowHidden(LngRow) Then vsDiag.Row = LngRow
        End If
    End If
End Sub

Public Sub DiagMouseUp(ByRef vsDiag As VSFlexGrid, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean

    If intButton = 2 Then
        Set mobjDiag = Nothing
        If frmMain.cbsMain Is Nothing Then Exit Sub
        Set objPopup = frmMain.cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If gobjPlugIn Is Nothing And blnDo Then Exit Sub '当弹出没有菜单项目时会显示一个空白小方块
        If Not objPopup Is Nothing Then
            Set mobjDiag = vsDiag
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub ExeDiagPlugIn(ByVal strName As String)
'功能：执行诊断外挂功能
    Dim lngID As String
    Dim strXML As String
    If CreatePlugInOK(gclsPros.Module) And (Not mobjDiag Is Nothing) Then
        With mobjDiag
            lngID = Val(.RowData(.Row))
            strXML = "<ROOT><诊断ID>" & .TextMatrix(.Row, DI_诊断ID) & "</诊断ID><疾病ID>" & .TextMatrix(.Row, DI_疾病ID) & "</疾病ID></ROOT>"
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(gclsPros.SysNo, gclsPros.Module, strName, gclsPros.病人ID, gclsPros.主页ID, lngID, .TextMatrix(.Row, DI_诊断描述), 6, strXML)
            Call zlPlugInErrH(Err, "ExecuteFunc")
            Err.Clear: On Error GoTo 0
        End With
    End If
End Sub

Public Sub ChangeCtl()
    '将获得焦点的控件 置于屏幕显示位置
On Error GoTo errH
    If mblnReturn = True Then
        If Not gclsPros.CurrentForm.ActiveControl Is Nothing Then
            If Not gclsPros.CurrentForm.ActiveControl.Container Is Nothing Then
                If Not gclsPros.CurrentForm.ActiveControl.Container Is Nothing Then
                    If gclsPros.CurrentForm.ActiveControl.Container.Name = "PicPage" Or gclsPros.CurrentForm.ActiveControl.Container.Name = "fraCbo" Then
                        Call LocateObjectPage(gclsPros.CurrentForm.ActiveControl)
                        mblnReturn = False
                    End If
                End If
            End If
        End If
    End If
errH:
    Err.Clear
End Sub
