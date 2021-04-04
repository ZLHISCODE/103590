VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'诊断类型枚举
Public Enum DiagType
    DT_门诊诊断XY = 1
    DT_入院诊断XY = 2
    DT_出院诊断XY = 3 '出院诊断(其他诊断)
    DT_院内感染 = 5
    DT_病理诊断 = 6
    DT_损伤中毒码 = 7
    DT_术前诊断 = 8
    DT_术后诊断 = 9
    DT_并发症 = 10
    DT_门诊诊断ZY = 11
    DT_入院诊断ZY = 12
    DT_出院诊断ZY = 13
    DT_病原学诊断 = 21
    DT_影像学诊断 = 22
End Enum
'功能枚举
Public Enum FuncEnum
    f医生首页 = 1
    f病案首页 = 2
    f电子病案 = 3
    f诊断选择 = 4
End Enum
'编辑模式
Public Enum Edit_Mode
    EM_编辑 = 0
    EM_查阅 = 1
    EM_新增病案 = 2
    EM_新增首页 = 3
End Enum
'住院号类型
Public Enum InNo_Type
    IT_NotVili = 0 '未经过验证
    IT_New = 1 '新的住院号
    IT_Old = 2 '以前的住院号
    IT_NewMed = 3 '新增首页的住院号
End Enum
'首页标准
Public Enum MedPage_Standard
    ST_卫生部标准 = 1
    ST_四川省标准 = 2
    ST_云南省标准 = 3
    ST_湖南省标准 = 4
    ST_门诊首页 = 5
End Enum
'临床路径状态
Public Enum Path_State
    PS_未导入 = -1
    PS_不符合导入 = 0
    PS_执行中 = 1
    PS_正常结束 = 2
    PS_变异结束 = 3
End Enum
'病人来源
Public Enum PatiFrom
    PF_门诊 = 1
    PF_住院 = 2
End Enum
'模块枚举
Public Enum Enum_Inside_Program
    p病案管理 = 200
    p病案评分 = 301
    p电子病案查阅 = 1259
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    P新版护士站 = 1265
    p电子病案审查 = 1560
    p手麻接口 = 2425
    p病人信息公共部件 = 9003
End Enum
'签名级别控件
Public Enum SignLevel
    SL_科主任 = 0
    SL_主任医师 = 1
    SL_主治医师 = 2
    SL_住院医师 = 3
End Enum
'信息的编辑方式
Public Enum InfoSave
    IS_全保存 = 0 '不论控件是否可用已经可见均搜集信息进行保存
    IS_只展示 = 1 '控件只进行展示信息作用
    IS_半保存 = 2 '只有可用可见时才搜集信息
End Enum
'隔离各个类的数据环境
Private mfeFunc As FuncEnum                            'FuncType:功能类型
Private mlngSys As Long                                       'SysNo:系统号
Private mlngModule As Long                               'Module：调用的模块号,
Private mfrmCurrent As Object                            'CurrentForm:当前窗体对象
Private mfrmMain As Object                                'MainForm:当前窗体对象的父窗体
Private mstrPrivs As String                                    'Privs: 功能权限
Private mrsMainInfo As ADODB.Recordset          'MainInfoRec:保存界面控件信息注册的主记录，通过该记录可以访问界面控件组或控件
Private mrsSecdInfo  As ADODB.Recordset          'SecdInfoRec:保存界面控件信息注册的次级记录，通过该记录可以访问界面控件组中具体控件或控件本身的内部元素或控件本身
Private mrsBaseCode As ADODB.Recordset          'BaseCode:基础字典表记录集
Private mrsManInfo As ADODB.Recordset            'ManInfo:医护等人员信息
Private mstrLoadMans As String                            'LoadMans:已经缓存的人员类别
Private mrsDiagConn As ADODB.Recordset         ' DiagConn:诊断与申请单关联关系
Private mblnComfirmInfec As Boolean                 'IsComfirmInfect:是否确认传染病，确认传染病功能打开的诊断选择器
Private mstMedPage As MedPage_Standard        'MedPageSandard:首页标准
Private mblnShareMedRec As Boolean                 'ShareMedRec:病案共享
Private mblnReadPages As Boolean                      'ReadPages:可以读取病案项目
Private memOpenMode As Edit_Mode                 'OpenMode:打开首页的模式
Private mlngFontSize As Long                               'FontSize：界面字体大小（一般用于表格）
Private mlng病人ID As Long                                  '病人ID:当前病人ID
Private mlng主页ID As Long                                  '主页ID:当前病人主页ID,或挂号单ID
Private mstrRegistNo As String                            'RegistNo:当前病人挂号单NO
Private mpfPatiType As PatiFrom                         'PatiType :当前显示的该病人的住院信息还是门诊信息
Private mbln护士站 As Boolean                             'Is护士站 :当前是护士站操作
Private mlng入院科室ID As Long                           '门诊:入院科室ID为病人科室ID
Private mlng出院科室ID As Long                           '门诊:入院科室ID为病人科室ID,住院：出院科室ID为病人当前科室ID或出院科室ID
Private mblnTCM As Boolean                                'IsTCM：病人所属科室是否具有中医科性质
Private mbln产科 As Boolean                                 'Is产科：科室是否具有产科性质
Private mpsCur As Path_State                                'PathState：病人该次住院的路径状态
Private mstrMustDiagType As String                      'MustDiagType:要求必须输入的诊断
Private mintInsureType As Integer                          'InsureType:该病人本次住院使用的保险类别
Private mlngValueLen As Long                                'ValueLen:记录病案主页从表“信息值”字段长度
Private mblnOK As Boolean                                    'IsOK:是否点击确定
Private mblnHave中医 As Boolean                         'Have中医:该用户具有中医编辑权限，可以编辑中医部分
Private mblnMoved As Boolean                             'Moved:是否从历史表中读取数据
Private mblnInfosChange As Boolean                    'InfosChange：界面信息是否改变
Private mlngTXTProc As Long                                 'TXTProc :保存默认的消息函数的地址
Private mlngDateIndex As Long                              'DateIndex:输入时间的控件的Index
Private mblnReturn As Boolean                              'IsReturn:是否回车
Private mbln手术 As Boolean                                  'Have手术:病人是否进行过手术

Private mstrPatiName As String                              'PatiName:当前病人姓名
Private mstrSex As String                                        'Sex:当前病人性别
Private mstr入院时间 As String                                'InTime:当前病人的入院时间
Private mstr出院时间 As String                                'OutTime:当前病人的出院时间
Private mdtInPath As DiagType                               'InPath:导入路径时的诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
Private mstrPathDiag As String                               'PathDiag:病人一次住院第二条路径以后导入诊断列表：诊断类型1|疾病ID1|诊断ID1,诊断类型2|疾病ID2|诊断ID2```
Private mblnIsPathOutTime As Boolean                 'PathOutTime:完成路径的时间是否比出院诊断记录时间大
Private mstr疾病IDs As String                                  'DiseaseIDs:用于保存诊断表格中存在的疾病ID,多个ID以逗号分割,在Closed事件中传递给父窗体
Private mstr诊断IDs As String                                  'DiagIDs:用于保存诊断表格中存在的诊断ID,多个ID以逗号分割,在Closed事件中传递给父窗体
Private mrsAuxiInfo  As ADODB.Recordset             'AuxiInfo:病人信息从表或病案主页从表
Private mrsPatiInfo  As ADODB.Recordset              'PatiInfo:病案主页信息以及病人信息
Private mrsPatiOut   As New ADODB.Recordset             'PatiOut:病人信息从表或病案主页从表
Private mrsFeesOut  As New ADODB.Recordset              'FeesOut:病案主页信息以及病人信息
Private mrsAdressInfo As ADODB.Recordset          'AdressInfo:病人地址信息

Private mblnUseTYT As Boolean                             'UseTYT:是否启用太元通接口

Private mint社区 As Integer                                     'CommunityID:病人的社区,门诊首页
Private mstr社区号 As String                                    'CommunityNO:病人的社区号,门诊首页
Private mstrPictureFile As String                             'PictureFile:病人照片,门诊首页

Private mstr住院号 As String                                   'InNo:当前病人的住院号
Private mbln编目 As Boolean                                  'Is编目:是否对已有主页编目
Private mblnChange编码员 As Boolean                  'Change编码员:是否可以修改病案编码员

Private mitInNoType As InNo_Type                        'NoType：0-当前病人的住院号是没经过验证的；1-住院号是新的；2-住院号是以前的:3-住院号是新增首页的
Private mblnSelPati As Boolean                              'IsSelPati：'保存是否从列表中选取得来
Private mblnExistPati As Boolean                            'IsExistPati：'保存是否从列表中选取得来
Private mblnNewPageNo As Boolean                     'NewPageNo：'病案号是否是新获取的
Private mblnOnlyPatiInfo As Boolean                     'OnlyPatiInfo:只有病人信息
Private mblnCheckData As Boolean                        'IsCheckData:是否在数据检查中

Private mblnSigned As Boolean                             'IsSigned:是否签名
Private mbln复诊 As Boolean                                  'Is复诊：是否复诊
Private mblnLastDiag As Boolean                           'IsLastDiag:读取的是病人上次挂号的诊断
Private mblnLoad As Boolean                                 'IsLoad :数据是否已经加载
'诊断选择器
Private mlngAplyMark As Long                               'AplyMark:当前申请单的标识
Private mlngAdviceID As Long                                'AdviceID:当前申请单对应的医嘱ID
Private mstrDiagRowIDs As String                          'DiagRowIDs:已经选择关联的诊断ID
Private mstrDiagNames As String                           'DiagNames:诊断描述串
Private mblnDiagSel As Boolean                              'DiagSel:有诊断勾选
'公共对象：
Private mclsMipModule As zl9ComLib.clsMipModule '消息对象
'外部类的接口属性
Private mblnOpen As Boolean                             'IsOpen:首页是否打开
Private mblnMerged As Boolean                         'IsMerged:病人首页是否经过合并
Private mblnDiagChange As Boolean                  'IsDiagChange:出院首要诊断是否发生改变
Private mblnDiagInput As Boolean                      'IsDiagInput:指定诊断的首页诊断是否输入
Private mstrPatiIDs As String                                '病案首页在本次打开关闭过程中修改的病人的ID集合

'参数属性（系统参数、模块参数）
'1、全局参数
Private mblnDo As Boolean                                    'IsDo:是否使用个性化设置
Private mstrLike As String                                       'LikeString:如果是双向匹配，则为%
Private mint简码 As Integer                                     'BriefCode:简码匹配方式：0-拼音,1-五笔
'2、系统参数
Private mbytPass As Byte                                         'PassType:合理用药接口类型,0-未使用,1-美康,2-大通,3-太元通
Private mintAllerSource As Integer                         'AllerSource:太元通过敏来源控制：0-医生选择，1-按药品目录输入，2-按过敏源输入
Private mintDiagSourceMZ As Integer                   'DiagSource MZ:表示门诊相关诊断:1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Private mintDiagSourceZY As Integer                    'DiagSource ZY:表示住院相关诊断:1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Private mintDiagInput As Integer                           'DiagInput:1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Private mblnMaskID As Boolean                            'IsMaskID:是否身份证掩码现实
Private mblnOutDocCtrl As Boolean                       'IsOutDocCtrl :外院医生必须建档
Private mblnFreeInput As Boolean                        'FreeInput : 首页编辑诊断描述和手术名称自由调整
'病案首页相关参数
Private mblnNewInNo As Boolean                          'NewInNo:病人每次住院使用一个新产生的住院号
Private mintFreeDeci As Integer                             'FreeDeci:表示录入病人费用和报盘时的小数位
Private mblnEditPageNo As Boolean                      'EditPageNo:允许用户修改病案号
Private mblnUseFileRules As Boolean                     'UseFileRules:档案号会根据设置的规则自动生成
Private mblnSinPageNo As Boolean                        'SinPageNo:每一份病案主页都应该有个病案号
Private mstrOutFile As String                                   'OutFile:通过该文件从其它系统得到病人的基本信息与费用情况
Private mblnOnLine As Boolean                               'OnLine:是否允许从ZLHIS处提数据
Private mblnOnLineNew As Boolean                        'OnLineNew:病案与HIS联机使用时仍允许新增病案，与参数32(gclsPros.OnLine)配合使用
Private mblnInputOutList As Boolean                       'InputOutList:要求在住院科室日报中输入出院病人清单
Private mstrFreeFormat As String                              'FreeFormat:费用格式字符串
'3、模块参数
Private mintCheck损伤中毒 As Integer                   'Check损伤中毒:主要出院诊断编码为S、T类时损伤中毒诊断：1-必须填写（不是ST开头的则禁止填写），2-提示是否填写，0-不检查
Private mintCheck病理诊断 As Integer                   'Check病理诊断:主要出院诊断编码为C、D类时病理诊断：1-必须填写（不是CD开头的则禁止填写），2-提示是否填写，0-不检查
Private mintCheckICD附码 As Integer                    'mintCheckICD附码:主要出院诊断编码为C、D类时ICD附码：1-必须填写，2-提示是否填写，0-不检查
Private mblnStructAdress   As Boolean                  'IsStructAdress:设置住院首页中病人地址是否启用结构化显示和存储
Private mblnShowTown   As Boolean                     'IsShowTown:设置住院首页中病人地址中乡镇地址是否结构化输入
Private mblnNotUseXYItems As Boolean               'NotUseXYItems中医首页里不使用的项目：病例分型，输液反应，输血反应，输红细胞，输血小板，输血浆，输全血，自体回收，输其他，输血前的9项检查，HBsAg，HCV-Ab，HIV-Ab，示教病案，科研病案，随诊，随诊期限，呼吸机使用，研究生医师
Private mintCheck区域     As Integer                       'Check区域 :首页中保存时区域：1-必须填写，2-提示是否填写，0-不检查
Private mblnUseOPSEndTime As Boolean              'UseOPSEndTime:是否使用手术结束时间
Private mblnSeparateEdit     As Boolean                 'SeparateEdit:不良事件的项目、输液反应、引发药物、临床表现、住院期间身体约束、离院时透析(血透、腹透)尿素氮值等信息在启用该参数时只能由护士填写
Private mintDiagInputZY As Integer                        'DiagInputZY:系统允许选择诊断输入方式（DiagInput=1）时住院中医诊断：0-根据诊断标准输入,1-根据疾病编码输入
Private mintDiagInputXY As Integer                        'DiagInputXY：系统允许选择诊断输入方式（DiagInput=1）时住院西医诊断：0-根据诊断标准输入,1-根据疾病编码输入
Private mintDiagInputMZ As Integer                       'DiagInputMZ：系统允许选择诊断输入方式（DiagInput=1）时门诊诊断：0-根据诊断标准输入,1-根据疾病编码输入
Private mintAllerInput As Integer                            'AllerInput:过敏输入来源：0-按药品目录输入，1-按过敏源输入
Private mintOPSInput As Integer                             'OPSInput:选择手术输入方式时：0-根据诊疗项目输入,1-根据ICD9-CM3输入
Private mblnOPSFree As Boolean                            'OPSFree:未找到时手术名称是否允许自由录入
Private mblnPathVCauses As Boolean                        'PathVCauses:路径变异原因从字典表中选取；不允许自由录入；
Private mblnM病理 As Boolean                             'mblnM病理:病理诊断只允许录入M打头的肿瘤形态学编码
'病案首页相关参数
Private mblnDefautADD As Boolean                       'DefautADD:录入出生地后,户口地址及联系人地址默认为与出生地相同
Private mblnCNIndent As Boolean                          'CNIndent:此参数是控制是否允许编辑诊断描述和ICD编码之间的输入关系:如果为true,则输入ICD编码并且存诊断描述时，则不替换诊断描述,如诊断描述为空,则替换诊断描述，同时可以修改诊断描述
Private mblnAuto病理 As Boolean                          'Auto病理:在病案首页编辑界面中控制手术病人自动填充出院诊断为病理诊断
Private mblnSameName As Boolean                      'SameName:在病案首页编辑界面中控制在输入病人费用的费用名时是否允许输入重复的费用名
Private mblnAutoDaig As Boolean                          'AutoDaig:设置病案首页在录入时,是否自动将西医诊断和中医诊断的诊断描述和编码提取出来作为默认值
Private mblnDaigFree As Boolean                           'DaigFree:设置病案首页在录入中医诊断时，允许疾病编码值为空
Private mblnNull出院情况 As Boolean                     'Null出院情况：在病案首页编辑界面中控制输入诊断但没有输入出院情况是否允许保存
Private mblnEditUnrecive As Boolean                     'EditUnrecive:在病案首页编辑界面中控制是否允许对未接收的病人病案进行编目操作
Private mblnEditPayType As Boolean                        'EditPayType:在病案首页编辑界面中控制是否将光标移动到医院付款方式上去
Private mblnTabPayType As Boolean                        'TabPayType:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到付款方式上面去输入付款方式信息
Private mblnTabAgeUnit As Boolean                        'TabAgeUnit:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到年龄单位上面去
Private mblnTabNation As Boolean                           'TabNation:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到国籍上面去输入国籍信息
Private mblnTabFileNo As Boolean                           'TabFileNo:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到档案号上面去输入档案号信息        '
Private mblnTabReadm As Boolean                           'TabReadm:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到再入院上面去输入再入院信息
Private mblnTabXRaysNo As Boolean                        'TabXRaysNo:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到X线号上面去输入X线号信息
Private mblnLoadFinish As Boolean

'与VSF表格增加行有关的
Private mblnAddAller As Boolean
Private mblnAddDiag As Boolean

Private mblnCboMask As Boolean
Private mblnGet附码 As Boolean                         '诊断自动提取附码
Private mblnGet医疗机构 As Boolean                     '医疗机构自有录入
Private mbln病理号 As Boolean                         '病理诊断填写后可以不录入病理号

'==============================================================
'缓存属性，缓存变量，用来隔离各个类的数据环境
'==============================================================
'是否正确的加载完成
Public Property Get LoadFinish() As Boolean
    LoadFinish = mblnLoadFinish
End Property
Public Property Let LoadFinish(ByVal blnLoadFinish As Boolean)
    mblnLoadFinish = blnLoadFinish
End Property

'FuncType:功能类别
Public Property Get FuncType() As FuncEnum
    FuncType = mfeFunc
End Property
'系统号
Public Property Get SysNo() As Long
    SysNo = mlngSys
End Property
Public Property Let SysNo(ByVal lngSys As Long)
    mlngSys = lngSys
End Property
'调用模块号
Public Property Get Module() As Long
    Module = mlngModule
End Property
Public Property Let Module(ByVal lngModule As Long)
    mlngModule = lngModule
End Property
'当前展示窗体或活动窗体
Public Property Get CurrentForm() As Object
    Set CurrentForm = mfrmCurrent
End Property
Public Property Set CurrentForm(frmCurrent As Object)
    Set mfrmCurrent = frmCurrent
End Property
'当前展示窗体或活动窗体的父窗体
Public Property Get MainForm() As Object
    Set MainForm = mfrmMain
End Property
Public Property Set MainForm(frmMain As Object)
    Set mfrmMain = frmMain
End Property
'用户模块的功能权限：Privs
Public Property Get Privs() As String
    Privs = mstrPrivs
End Property
Public Property Let Privs(strPrivs As String)
    If strPrivs = "" Then
        mstrPrivs = GetPrivFunc(Me.SysNo, Me.Module)
    Else
        mstrPrivs = strPrivs
    End If
End Property
'主信息记录
Public Property Get MainInfoRec() As ADODB.Recordset
    Set MainInfoRec = mrsMainInfo
End Property
Public Property Set MainInfoRec(rsMainInfo As ADODB.Recordset)
    Set mrsMainInfo = rsMainInfo
End Property
'次级信息记录
Public Property Get SecdInfoRec() As ADODB.Recordset
    Set SecdInfoRec = mrsSecdInfo
End Property
Public Property Set SecdInfoRec(rsSecdInfo As ADODB.Recordset)
    Set mrsSecdInfo = rsSecdInfo
End Property
'BaseCode:基础字典表记录集
Public Property Get BaseCode() As ADODB.Recordset
    Set BaseCode = mrsBaseCode
End Property
Public Property Set BaseCode(rsBaseCode As ADODB.Recordset)
    Set mrsBaseCode = rsBaseCode
End Property
'ManInfo:医护等人员信息
Public Property Get ManInfo() As ADODB.Recordset
    Set ManInfo = mrsManInfo
End Property
Public Property Set ManInfo(rsManInfo As ADODB.Recordset)
    Set mrsManInfo = rsManInfo
    If rsManInfo Is Nothing Then
        mstrLoadMans = ""
    End If
End Property
'LoadMans:已经缓存的人员类别
Public Property Get LoadMans() As String
    LoadMans = mstrLoadMans
End Property
Public Property Let LoadMans(strLoadMans As String)
    mstrLoadMans = strLoadMans
End Property

'DiagConn:诊断与申请单关联关系
Public Property Get DiagConn() As ADODB.Recordset
    Set DiagConn = mrsDiagConn
End Property
Public Property Set DiagConn(rsDiagConn As ADODB.Recordset)
    Set mrsDiagConn = rsDiagConn
End Property
'IsComfirmInfect:是否确认传染病，确认传染病功能打开的诊断选择器
Public Property Get IsComfirmInfect() As Boolean
    IsComfirmInfect = mblnComfirmInfec
End Property

Public Property Let IsComfirmInfect(ByVal blnComfirmInfect As Boolean)
    mblnComfirmInfec = blnComfirmInfect
End Property

'首页标准
Public Property Get MedPageSandard() As MedPage_Standard
    MedPageSandard = mstMedPage
End Property

Public Property Let MedPageSandard(ByVal stMedPage As MedPage_Standard)
    If stMedPage <> mstMedPage Then
        Set mrsMainInfo = Nothing
        Set mrsSecdInfo = Nothing
        mstMedPage = stMedPage
    End If
    Call InitCacheRecInfo
End Property
'ShareMedRec:病案共享
Public Property Get ShareMedRec() As Boolean
    '病案首页可以直接读取病案项目
    ShareMedRec = mblnShareMedRec
End Property
'ReadPages:可以读取病案项目
Public Property Get ReadPages() As Boolean
    '病案首页可以直接读取病案项目
    ReadPages = mfeFunc = f病案首页 Or mfeFunc <> f诊断选择 And mblnShareMedRec
End Property
'OpenMode ：打开方式
Public Property Get OpenMode() As Edit_Mode
    OpenMode = Decode(mfeFunc, f电子病案, EM_查阅, f诊断选择, EM_编辑, memOpenMode)
End Property

Public Property Let OpenMode(ByVal emOpenMode As Edit_Mode)
    memOpenMode = Decode(mfeFunc, f电子病案, EM_查阅, f诊断选择, EM_编辑, emOpenMode)
End Property
'FontSize：界面字体大小
Public Property Get FontSize() As Long
'病案首页五号字体为10.5，但是公共方法为长整形字体，因此当作11号字体
    FontSize = Decode(mfeFunc, f病案首页, 11, f诊断选择, mlngFontSize, 9)
End Property

Public Property Let FontSize(ByVal lngFontSize As Long)
    mlngFontSize = lngFontSize
End Property
'病人ID
Public Property Get 病人ID() As Long
    病人ID = mlng病人ID
End Property

Public Property Let 病人ID(ByVal lng病人ID As Long)
    mlng病人ID = lng病人ID
End Property
'主页ID
Public Property Get 主页ID() As Long
    主页ID = mlng主页ID
End Property

Public Property Let 主页ID(ByVal lng主页ID As Long)
    mlng主页ID = lng主页ID
End Property
'RegistNo:当前病人挂号单NO
Public Property Get RegistNo() As String
    RegistNo = mstrRegistNo
End Property

Public Property Let RegistNo(ByVal strRegistNo As String)
    mstrRegistNo = strRegistNo
End Property
'PatiType :当前显示的该病人的住院信息还是门诊信息
Public Property Get PatiType() As PatiFrom
    If Me.FuncType = f病案首页 Then mpfPatiType = PF_住院
    PatiType = mpfPatiType
End Property

Public Property Let PatiType(ByVal pfPatiType As PatiFrom)
    If Me.FuncType = f病案首页 Then pfPatiType = PF_住院
    If pfPatiType <> mpfPatiType Then
        mpfPatiType = pfPatiType
        Call Me.RefreshPara
    Else
        mpfPatiType = pfPatiType
    End If
End Property
'Is护士站 :当前是护士站操作
Public Property Get Is护士站() As Boolean
    mbln护士站 = mlngModule = p住院护士站 Or mlngModule = P新版护士站
    Is护士站 = mbln护士站
End Property
'入院科室ID
Public Property Get 入院科室ID() As Long
    入院科室ID = mlng入院科室ID
End Property

Public Property Let 入院科室ID(ByVal lng入院科室ID As Long)
    Dim blnTCM As Boolean, bln产科 As Boolean

    If mlng入院科室ID <> lng入院科室ID Then
        '只有病案需要判断入院科室
        If mfeFunc = f病案首页 Then
            If lng入院科室ID <> 0 Then
                blnTCM = sys.DeptHaveProperty(lng入院科室ID, "中医科")
                bln产科 = sys.DeptHaveProperty(lng入院科室ID, "产科")
            End If
        End If
        If Me.出院科室ID <> 0 Then
            If Not blnTCM Then blnTCM = sys.DeptHaveProperty(Me.出院科室ID, "中医科")
            '诊断选择器只需判断中医科
            If Not bln产科 And gclsPros.FuncType <> f诊断选择 Then bln产科 = sys.DeptHaveProperty(Me.出院科室ID, "产科")
        End If
        If blnTCM <> Me.IsTCM Or bln产科 <> Me.Is产科 Then
            If bln产科 <> Me.Is产科 Then
                mbln产科 = bln产科
                If mfeFunc = f病案首页 Then SetLstBoxFromRec ("不良事件")
            Else
                mbln产科 = bln产科
            End If
            mblnTCM = blnTCM
             If mfeFunc = f病案首页 Or mfeFunc = f医生首页 Then Call SetPageVisible
        End If
        mlng入院科室ID = lng入院科室ID
    End If
End Property
'出院科室ID
Public Property Get 出院科室ID() As Long
    出院科室ID = mlng出院科室ID
End Property

Public Property Let 出院科室ID(ByVal lng出院科室ID As Long)
    Dim blnTCM As Boolean, bln产科 As Boolean

    If mlng出院科室ID <> lng出院科室ID Then
        If lng出院科室ID <> 0 Then
            blnTCM = sys.DeptHaveProperty(lng出院科室ID, "中医科")
'            诊断选择器只需判断中医科
            If gclsPros.FuncType <> f诊断选择 Then bln产科 = sys.DeptHaveProperty(lng出院科室ID, "产科")
        End If
'        病案同时需要判断入院科室
        If Me.入院科室ID <> 0 And mfeFunc = f病案首页 Then
            If Not blnTCM Then blnTCM = sys.DeptHaveProperty(Me.入院科室ID, "中医科")
            If Not bln产科 Then bln产科 = sys.DeptHaveProperty(Me.入院科室ID, "产科")
        End If
        '中医科属参数判断
        If (mfeFunc = f电子病案 Or mfeFunc = f诊断选择) And mpfPatiType = PF_门诊 Then
            blnTCM = blnTCM Or Val(zlDatabase.GetPara("门诊西医科允许录入中医诊断", Me.SysNo, 1252, "0")) = 1
        End If
        If blnTCM <> Me.IsTCM Or bln产科 <> Me.Is产科 Then
            If bln产科 <> Me.Is产科 Then
                mbln产科 = bln产科
                If mfeFunc = f病案首页 Then SetLstBoxFromRec ("不良事件")
            Else
                mbln产科 = bln产科
            End If
            mblnTCM = blnTCM
            If (mfeFunc = f病案首页 Or mfeFunc = f医生首页) And Me.IsOpen Then Call SetPageVisible
        End If
        mlng出院科室ID = lng出院科室ID
    End If
End Property
'IsTCM：病人所属科室是否具有中医科性质
Public Property Get IsTCM() As Boolean
    Call TestCode
    IsTCM = mblnTCM
End Property
'Is产科：病人所属科室是否具有产科性质
Public Property Get Is产科() As Boolean
    Is产科 = mbln产科
End Property
'PathState：病人该次住院的路径状态
Public Property Get PathState() As Path_State
    PathState = mpsCur
End Property

Public Property Let PathState(ByVal psInput As Path_State)
    mpsCur = psInput
End Property
'MustDiagType:要求必须输入的诊断
Public Property Get MustDiagType() As String
    MustDiagType = mstrMustDiagType
End Property

Public Property Let MustDiagType(ByVal strMustDiagType As String)
    mstrMustDiagType = strMustDiagType
End Property
'InsureType:该病人本次住院使用的保险类别
Public Property Get InsureType() As Integer
    InsureType = mintInsureType
End Property

Public Property Let InsureType(ByVal intInsureType As Integer)
    mintInsureType = intInsureType
End Property
'ValueLen:记录病案主页从表“信息值”字段长度
Public Property Get ValueLen() As Long
    ValueLen = mlngValueLen
End Property

Public Property Let ValueLen(ByVal lngValueLen As Long)
    mlngValueLen = lngValueLen
End Property
'IsOK:是否点击确定
Public Property Get IsOK() As Boolean
    IsOK = mblnOK
End Property

Public Property Let IsOK(ByVal blnOK As Boolean)
    mblnOK = blnOK
End Property
'Have中医:该用户具有中医编辑权限，可以编辑中医部分
Public Property Get Have中医() As Boolean
    If Me.FuncType = f病案首页 Then
        '新增病案或者新增首页时，如果没有输入入院科室与出院科室，则默认显示中医诊断页
        If (Me.OpenMode = EM_新增病案 Or Me.OpenMode = EM_新增首页) And Me.入院科室ID = 0 And Me.出院科室ID = 0 Then
            mblnHave中医 = InStr(";" & Me.Privs & ";", ";中医;") > 0
        Else
            mblnHave中医 = InStr(";" & Me.Privs & ";", ";中医;") > 0 And Me.IsTCM
        End If
    Else
        mblnHave中医 = Me.IsTCM
    End If
    Have中医 = mblnHave中医
End Property

' Moved:是否从历史表中读取数据
Public Property Get Moved() As Boolean
    Moved = mblnMoved
End Property
Public Property Let Moved(ByVal blnMoved As Boolean)
    mblnMoved = blnMoved
End Property
'InfosChange界面信息是否改变
Public Property Get InfosChange() As Boolean
    InfosChange = mblnInfosChange
End Property
Public Property Let InfosChange(ByVal blnInfosChange As Boolean)
    mblnInfosChange = blnInfosChange
End Property
'TXTProc :保存默认的消息函数的地址
Public Property Get TXTProc() As Long
    TXTProc = mlngTXTProc
End Property
Public Property Let TXTProc(ByVal lngTXTProc As Long)
    mlngTXTProc = lngTXTProc
End Property

'DateIndex:输入时间的控件的Index
Public Property Get DateIndex() As Long
    DateIndex = mlngDateIndex
End Property
Public Property Let DateIndex(ByVal lngDateIndex As Long)
    mlngDateIndex = lngDateIndex
End Property
'IsReturn:是否回车
Public Property Get IsReturn() As Boolean
    IsReturn = mblnReturn
End Property
Public Property Let IsReturn(ByVal blnReturn As Boolean)
    mblnReturn = blnReturn
End Property
'Have手术:病人是否进行过手术
Public Property Get Have手术() As Boolean
    Have手术 = mbln手术
End Property
Public Property Let Have手术(ByVal bln手术 As Boolean)
    mbln手术 = bln手术
End Property
'Sex:当前病人性别
Public Property Get Sex() As String
    Sex = mstrSex
End Property

Public Property Let Sex(ByVal strSex As String)
    If strSex Like "*男*" Then
        mstrSex = "男"
    ElseIf strSex Like "*女*" Then
        mstrSex = "女"
    End If
End Property
'PatiName:当前病人姓名
Public Property Get PatiName() As String
    PatiName = mstrPatiName
End Property

Public Property Let PatiName(ByVal strPatiName As String)
    mstrPatiName = strPatiName
End Property
'InTime:当前病人的入院时间
Public Property Get InTime() As String
    InTime = mstr入院时间
End Property

Public Property Let InTime(ByVal str入院时间 As String)
    Dim objTxt As TextBox
    If mstr入院时间 <> str入院时间 And gclsPros.FuncType <> f诊断选择 And gclsPros.PatiType = PF_住院 Then
        If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
            Set objTxt = Me.CurrentForm.txtSpecificInfo(SLC_住院天数)
        Else
            Set objTxt = Me.CurrentForm.txtInfo(GCA_住院天数)
        End If
        If str入院时间 <> "" And Me.OutTime <> "" Then
            objTxt.Text = DateDiff("d", CDate(str入院时间), CDate(Me.OutTime))
        ElseIf str入院时间 <> "" Then
            objTxt.Text = DateDiff("d", CDate(str入院时间), zlDatabase.Currentdate)
        End If
        If Val(objTxt.Text) = 0 Then objTxt.Text = "1"
    End If
    mstr入院时间 = str入院时间
End Property
'OutTime:当前病人的出院时间
Public Property Get OutTime() As String
    OutTime = mstr出院时间
End Property

Public Property Let OutTime(ByVal str出院时间 As String)
    Dim objTxt As TextBox
    If mstr出院时间 <> str出院时间 And gclsPros.FuncType <> f诊断选择 And gclsPros.PatiType = PF_住院 Then
        If gclsPros.FuncType = f病案首页 Or gclsPros.FuncType = f医生首页 Then
            Set objTxt = Me.CurrentForm.txtSpecificInfo(SLC_住院天数)
        Else
            Set objTxt = Me.CurrentForm.txtInfo(GCA_住院天数)
        End If
        If str出院时间 <> "" And Me.InTime <> "" Then
            objTxt.Text = DateDiff("d", CDate(Me.InTime), CDate(str出院时间))
        ElseIf Me.InTime <> "" Then
            objTxt.Text = DateDiff("d", CDate(Me.InTime), zlDatabase.Currentdate)
        End If
        If Val(objTxt.Text) = 0 Then objTxt.Text = "1"
    End If
    mstr出院时间 = str出院时间
End Property
'InPath:导入路径时的诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
Public Property Get InPath() As DiagType
    InPath = mdtInPath
End Property

Public Property Let InPath(ByVal dtInPath As DiagType)
    mdtInPath = dtInPath
End Property
'PathDiag:病人一次住院第二条路径以后导入诊断列表：诊断类型1|疾病ID1|诊断ID1,诊断类型2|疾病ID2|诊断ID2```
Public Property Get PathDiag() As String
    PathDiag = mstrPathDiag
End Property

Public Property Let PathDiag(ByVal strPathDiag As String)
    mstrPathDiag = strPathDiag
End Property
'PathOutTime:完成路径的时间是否比出院诊断记录时间大
Public Property Get PathOutTime() As Boolean
    PathOutTime = mblnIsPathOutTime
End Property

Public Property Let PathOutTime(ByVal blnIsPathOutTime As Boolean)
    mblnIsPathOutTime = blnIsPathOutTime
End Property
'DiseaseIDs:用于保存诊断表格中存在的疾病ID,多个ID以逗号分割,在Closed事件中传递给父窗体
Public Property Get DiseaseIDs() As String
    DiseaseIDs = mstr疾病IDs
End Property

Public Property Let DiseaseIDs(ByVal str疾病IDs As String)
    mstr疾病IDs = str疾病IDs
End Property
'DiagIDs:用于保存诊断表格中存在的诊断ID,多个ID以逗号分割,在Closed事件中传递给父窗体
Public Property Get DiagIDs() As String
    DiagIDs = mstr诊断IDs
End Property

Public Property Let DiagIDs(ByVal str诊断IDs As String)
    mstr诊断IDs = str诊断IDs
End Property
'AuxiInfo:病人信息从表或病案主页从表
Public Property Get AuxiInfo() As ADODB.Recordset
    Set AuxiInfo = mrsAuxiInfo
End Property

Public Property Set AuxiInfo(ByVal rsAuxiInfo As ADODB.Recordset)
    Set mrsAuxiInfo = rsAuxiInfo
    Me.ValueLen = rsAuxiInfo.Fields("信息值").DefinedSize
End Property
'PatiInfo:病案主页信息以及病人信息
Public Property Get PatiInfo() As ADODB.Recordset
    Set PatiInfo = mrsPatiInfo
End Property

Public Property Set PatiInfo(ByVal rsPatiInfo As ADODB.Recordset)
    Set mrsPatiInfo = rsPatiInfo
    If Not rsPatiInfo.EOF Then
        Me.Sex = rsPatiInfo!性别 & ""
        Me.InsureType = Val(rsPatiInfo!险类 & "")
        If Me.PatiType = PF_门诊 Then
            Me.主页ID = rsPatiInfo!挂号ID
            Me.病人ID = rsPatiInfo!病人ID
            Me.出院科室ID = Val(rsPatiInfo!科室id & "")
            Me.CommunityID = Val(rsPatiInfo!社区 & "")
            Me.CommunityNO = rsPatiInfo!社区号 & ""
        Else
            Me.OutTime = Format(rsPatiInfo!出院日期 & "", "yyyy-MM-dd hh:mm:ss")
            Me.InTime = Format(rsPatiInfo!入院日期 & "", "yyyy-MM-dd hh:mm:ss")
            Me.入院科室ID = Val(rsPatiInfo!入院科室ID & "")
            Me.出院科室ID = Val(rsPatiInfo!出院科室ID & "")
        End If
     End If
End Property
'AuxiInfo:病人信息从表或病案主页从表
Public Property Get PatiOut() As ADODB.Recordset
    Set PatiOut = mrsPatiOut
End Property

Public Property Set PatiOut(ByVal rsPatiOut As ADODB.Recordset)
    Set mrsPatiOut = rsPatiOut
End Property
'FreeOut:病人信息从表或病案主页从表
Public Property Get FeesOut() As ADODB.Recordset
    Set FeesOut = mrsFeesOut
End Property

Public Property Set FeesOut(ByVal rsFeesOut As ADODB.Recordset)
    Set mrsFeesOut = rsFeesOut
End Property
'AdressInfo:病人地址信息
Public Property Get AdressInfo() As ADODB.Recordset
    Set AdressInfo = mrsAdressInfo
End Property

Public Property Set AdressInfo(ByVal rsAdressInfo As ADODB.Recordset)
    Set mrsAdressInfo = rsAdressInfo
End Property
'UseTYT:是否启用太元通接口
Public Property Get UseTYT() As Boolean
    UseTYT = mblnUseTYT
End Property

Public Property Let UseTYT(ByVal blnUseTYT As Boolean)
    mblnUseTYT = blnUseTYT
End Property
'CommunityID:病人的社区,门诊首页
Public Property Get CommunityID() As Integer
    CommunityID = mint社区
End Property

Public Property Let CommunityID(ByVal int社区 As Integer)
    mint社区 = int社区
End Property
'CommunityNO:病人的社区号,门诊首页
Public Property Get CommunityNO() As String
    CommunityNO = mstr社区号
End Property

Public Property Let CommunityNO(ByVal str社区号 As String)
    mstr社区号 = str社区号
End Property
'PictureFile:病人照片,门诊首页
Public Property Get PictureFile() As String
    PictureFile = mstrPictureFile
End Property

Public Property Let PictureFile(ByVal strPictureFile As String)
    mstrPictureFile = strPictureFile
End Property
'InNo:当前病人的住院号
Public Property Get InNo() As String
    InNo = mstr住院号
End Property

Public Property Let InNo(ByVal str住院号 As String)
    mstr住院号 = str住院号
End Property
'Is编目:是否对已有主页编目
Public Property Get Is编目() As Boolean
    Is编目 = mbln编目
End Property

Public Property Let Is编目(ByVal bln编目 As Boolean)
    mbln编目 = bln编目
End Property
 'NoType：0-当前病人的住院号是没经过验证的；1-住院号是新的；2-住院号是以前的:3-住院号是新增首页的
Public Property Get NoType() As InNo_Type
    NoType = mitInNoType
End Property

Public Property Let NoType(ByVal itInNoType As InNo_Type)
    mitInNoType = itInNoType
End Property
'IsSelPati：'保存是否从列表中选取得来
Public Property Get IsSelPati() As Boolean
    IsSelPati = mblnSelPati
End Property

Public Property Let IsSelPati(ByVal blnSelPati As Boolean)
    mblnSelPati = blnSelPati
End Property
'IsExistPati：'保存是否从列表中选取得来
Public Property Get IsExistPati() As Boolean
    IsExistPati = mblnExistPati
End Property

Public Property Let IsExistPati(ByVal blnExistPati As Boolean)
    mblnExistPati = blnExistPati
End Property
'NewPageNo：'病案号是否是新获取的
Public Property Get NewPageNo() As Boolean
    NewPageNo = mblnNewPageNo
End Property

Public Property Let NewPageNo(ByVal blnNewPageNo As Boolean)
    mblnNewPageNo = blnNewPageNo
End Property
'OnlyPatiInfo:只有病人信息
Public Property Get OnlyPatiInfo() As Boolean
    OnlyPatiInfo = mblnOnlyPatiInfo
End Property

Public Property Let OnlyPatiInfo(ByVal blnOnlyPatiInfo As Boolean)
    mblnOnlyPatiInfo = blnOnlyPatiInfo
End Property

'IsCheckData:是否在数据检查中
Public Property Get IsCheckData() As Boolean
    IsCheckData = mblnCheckData
End Property

Public Property Let IsCheckData(ByVal blnCheckData As Boolean)
    mblnCheckData = blnCheckData
End Property

'Change编码员:是否可以修改病案编码员
Public Property Get Change编码员() As Boolean
    Change编码员 = mblnChange编码员
End Property

Public Property Let Change编码员(ByVal blnChange编码员 As Boolean)
    mblnChange编码员 = blnChange编码员
End Property
'IsSigned:是否签名
Public Property Get IsSigned() As Boolean
    IsSigned = mblnSigned
End Property

Public Property Let IsSigned(ByVal blnSigned As Boolean)
    mblnSigned = blnSigned
End Property
'Is复诊：是否复诊
Public Property Get Is复诊() As Boolean
    Is复诊 = mbln复诊
End Property

Public Property Let Is复诊(ByVal bln复诊 As Boolean)
    mbln复诊 = bln复诊
End Property
'IsLastDiag:读取的是病人上次挂号的诊断
Public Property Get IsLastDiag() As Boolean
    IsLastDiag = mblnLastDiag
End Property

Public Property Let IsLastDiag(ByVal blnLastDiag As Boolean)
    mblnLastDiag = blnLastDiag
End Property
'IsLoad :数据是否已经加载
Public Property Get IsLoad() As Boolean
    IsLoad = mblnLoad
End Property

Public Property Let IsLoad(ByVal blnLoad As Boolean)
    mblnLoad = blnLoad
End Property
'AplyMark:当前申请单的标识
Public Property Get AplyMark() As Long
    AplyMark = mlngAplyMark
End Property
Public Property Let AplyMark(ByVal lngAplyMark As Long)
    mlngAplyMark = lngAplyMark
End Property
'AdviceID:当前申请单对应的医嘱ID
Public Property Get AdviceID() As Long
    AdviceID = mlngAdviceID
End Property
Public Property Let AdviceID(ByVal lngAdviceID As Long)
    mlngAdviceID = lngAdviceID
End Property
'DiagRowIDs:已经选择关联的诊断ID
Public Property Get DiagRowIDs() As String
    DiagRowIDs = mstrDiagRowIDs
End Property
Public Property Let DiagRowIDs(ByVal strDiagRowIDs As String)
    mstrDiagRowIDs = strDiagRowIDs
End Property
'DiagNames:诊断描述串
Public Property Get DiagNames() As String
    DiagNames = mstrDiagNames
End Property
Public Property Let DiagNames(ByVal strDiagNames As String)
    mstrDiagNames = strDiagNames
End Property
'DiagSel:有诊断勾选
Public Property Get DiagSel() As Boolean
    DiagSel = mblnDiagSel
End Property

Public Property Let DiagSel(ByVal blnDiagSel As Boolean)
    mblnDiagSel = blnDiagSel
End Property

'病案参数
'DefautADD:录入出生地后,户口地址及联系人地址默认为与出生地相同
Public Property Get DefautADD() As Boolean
    DefautADD = mblnDefautADD
End Property

Public Property Let DefautADD(ByVal blnDefautADD As Boolean)
    mblnDefautADD = blnDefautADD
End Property
 'CNIndent:此参数是控制是否允许编辑诊断描述和ICD编码之间的输入关系:如果为true,则输入ICD编码并且存诊断描述时，则不替换诊断描述,如诊断描述为空,则替换诊断描述，同时可以修改诊断描述
Public Property Get CNIndent() As Boolean
    CNIndent = mblnCNIndent
End Property

Public Property Let CNIndent(ByVal blnCNIndent As Boolean)
    mblnCNIndent = blnCNIndent
End Property

'Auto病理:在病案首页编辑界面中控制手术病人自动填充出院诊断为病理诊断
Public Property Get Auto病理() As Boolean
    Auto病理 = mblnAuto病理
End Property

Public Property Let Auto病理(ByVal blnAuto病理 As Boolean)
    mblnAuto病理 = blnAuto病理
End Property
'SameName:在病案首页编辑界面中控制在输入病人费用的费用名时是否允许输入重复的费用名
Public Property Get SameName() As Boolean
    SameName = mblnSameName
End Property

Public Property Let SameName(ByVal blnSameName As Boolean)
    mblnSameName = blnSameName
End Property
'AutoDaig:设置病案首页在录入时,是否自动将西医诊断和中医诊断的诊断描述和编码提取出来作为默认值
Public Property Get AutoDaig() As Boolean
    AutoDaig = mblnAutoDaig
End Property

Public Property Let AutoDaig(ByVal blnAutoDaig As Boolean)
    mblnAutoDaig = blnAutoDaig
End Property
'AutoDaig:设置病案首页在录入时,是否自动将西医诊断和中医诊断的诊断描述和编码提取出来作为默认值
Public Property Get DaigFree() As Boolean
    DaigFree = mblnDaigFree
End Property

Public Property Let DaigFree(ByVal blnDaigFree As Boolean)
    mblnDaigFree = blnDaigFree
End Property
'Null出院情况：在病案首页编辑界面中控制输入诊断但没有输入出院情况是否允许保存
Public Property Get Null出院情况() As Boolean
    Null出院情况 = mblnNull出院情况
End Property

Public Property Let Null出院情况(ByVal blnNull出院情况 As Boolean)
    mblnNull出院情况 = blnNull出院情况
End Property
'EditUnrecive:在病案首页编辑界面中控制是否允许对未接收的病人病案进行编目操作
Public Property Get EditUnrecive() As Boolean
    EditUnrecive = mblnEditUnrecive
End Property

Public Property Let EditUnrecive(ByVal blnEditUnrecive As Boolean)
    mblnEditUnrecive = blnEditUnrecive
End Property
'EditPayType:在病案首页编辑界面中控制是否将光标移动到医院付款方式上去
Public Property Get EditPayType() As Boolean
    EditPayType = mblnEditPayType
End Property

Public Property Let EditPayType(ByVal blnEditPayType As Boolean)
    mblnEditPayType = blnEditPayType
End Property
'TabPayType:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到付款方式上面去输入付款方式信息
Public Property Get TabPayType() As Boolean
    TabPayType = mblnTabPayType
End Property

Public Property Let TabPayType(ByVal blnTabPayType As Boolean)
    mblnTabPayType = blnTabPayType
End Property
'TabAgeUnit:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到年龄单位上面去
Public Property Get TabAgeUnit() As Boolean
    TabAgeUnit = mblnTabAgeUnit
End Property

Public Property Let TabAgeUnit(ByVal blnTabAgeUnit As Boolean)
    mblnTabAgeUnit = blnTabAgeUnit
End Property
'TabNation:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到国籍上面去输入国籍信息
Public Property Get TabNation() As Boolean
    TabNation = mblnTabNation
End Property

Public Property Let TabNation(ByVal blnTabNation As Boolean)
    mblnTabNation = blnTabNation
End Property
'TabFileNo:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到档案号上面去输入档案号信息
Public Property Get TabFileNo() As Boolean
    TabFileNo = mblnTabFileNo
End Property

Public Property Let TabFileNo(ByVal blnTabFileNo As Boolean)
    mblnTabFileNo = blnTabFileNo
End Property
'TabReadm:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到再入院上面去输入再入院信息
Public Property Get TabReadm() As Boolean
    TabReadm = mblnTabReadm
End Property

Public Property Let TabReadm(ByVal blnTabReadm As Boolean)
    mblnTabReadm = blnTabReadm
End Property
'TabXRaysNo:在病案首页编辑界面中控制在输入回车键时是否允许将光标跳到X线号上面去输入X线号信息
Public Property Get TabXRaysNo() As Boolean
    TabXRaysNo = mblnTabXRaysNo
End Property

Public Property Let TabXRaysNo(ByVal blnTabXRaysNo As Boolean)
    mblnTabXRaysNo = blnTabXRaysNo
End Property

'==============================================================
'外部接口属性
'==============================================================
'首页是否是打开状态，只读属性
Public Property Get IsOpen() As Boolean
    IsOpen = mblnOpen
End Property

Public Property Let IsOpen(ByVal blnOpen As Boolean)
    mblnOpen = blnOpen
End Property
'首页是否合并过,只读属性
Public Property Get IsMerged() As Boolean
    IsMerged = mblnMerged
End Property

Public Property Let IsMerged(ByVal blnMerged As Boolean)
    mblnMerged = blnMerged
End Property
'出院首要诊断是否发生改变,只读属性
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mblnDiagChange
End Property

Public Property Let IsDiagChange(ByVal blnDiagChange As Boolean)
    mblnDiagChange = blnDiagChange
End Property
'IsDiagInput:指定诊断的首页诊断是否输入
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mblnDiagInput
End Property

Public Property Let IsDiagInput(ByVal blnDiagInput As Boolean)
    mblnDiagInput = blnDiagInput
End Property

'==============================================================
'参数属性
'==============================================================
'1、全局参数：
'IsDo:是否使用个性化设置
Public Property Get IsDo() As Boolean
    IsDo = mblnDo
End Property
'LikeString:如果是双向匹配，则为%
Public Property Get LikeString() As String
    LikeString = mstrLike
End Property
'BriefCode:简码匹配方式：0-拼音,1-五笔
Public Property Get BriefCode() As String
    BriefCode = mint简码
End Property
'2、系统参数
'PassType:合理用药接口类型,0-未使用,1-美康,2-大通,3-太元通
Public Property Get PassType() As Byte
    PassType = mbytPass
End Property
Friend Property Let PassType(ByVal bytPass As Byte)
    mbytPass = bytPass
End Property
'AllerSource:太元通过敏来源控制：0-医生选择，1-按药品目录输入，2-按过敏源输入
Public Property Get AllerSource() As Integer
    AllerSource = mintAllerSource
End Property
'DiagInputMZ:表示门诊相关诊断:1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public Property Get DiagSourceMZ() As Integer
    DiagSourceMZ = mintDiagSourceMZ
End Property
'DiagInputZY:表示住院相关诊断:1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public Property Get DiagSourceZY() As Integer
    DiagSourceZY = mintDiagSourceZY
End Property
'DiagInput:1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Public Property Get DiagInput() As Integer
    DiagInput = mintDiagInput
End Property
'IsMaskID:是否身份证掩码现实
Public Property Get IsMaskID() As Boolean
    IsMaskID = mblnMaskID
End Property
'IsOutDocCtrl:外院医生必须建档
Public Property Get IsOutDocCtrl() As Boolean
    IsOutDocCtrl = mblnOutDocCtrl
End Property
'FreeInput:诊断描述和手术名称自由调整
Public Property Get FreeInput() As Boolean
    FreeInput = mblnFreeInput
End Property
'NewInNo:病人每次住院使用一个新产生的住院号
Public Property Get NewInNo() As Boolean
    NewInNo = mblnNewInNo
End Property
Friend Property Let NewInNo(ByVal blnNewInNo As Boolean)
    mblnNewInNo = blnNewInNo
    If blnNewInNo Then
        mblnSinPageNo = blnNewInNo
    Else
        mblnSinPageNo = Val(zlDatabase.GetPara(39, Me.SysNo, 0)) = 1
    End If
End Property
'FreeDeci:表示录入病人费用和报盘时的小数位
Public Property Get FreeDeci() As Integer
    FreeDeci = mintFreeDeci
End Property
Friend Property Let FreeDeci(ByVal intFreeDeci As Integer)
    mintFreeDeci = intFreeDeci
    mstrFreeFormat = "###0." & String(mintFreeDeci, "0") & ";-" & "###0." & String(mintFreeDeci, "0") & ";0;0"
End Property
'EditPageNo:允许用户修改病案号
Public Property Get EditPageNo() As Boolean
    EditPageNo = mblnEditPageNo
End Property
'UseFileRules:档案号会根据设置的规则自动生成
Public Property Get UseFileRules() As Boolean
    UseFileRules = mblnUseFileRules
End Property
'SinPageNo:每一份病案主页都应该有个病案号
Public Property Get SinPageNo() As Boolean
    SinPageNo = mblnSinPageNo
End Property
'OutFile:通过该文件从其它系统得到病人的基本信息与费用情况
Public Property Get OutFile() As String
    OutFile = mstrOutFile
End Property
'OnLine:是否允许从ZLHIS处提数据
Public Property Get OnLine() As Boolean
    OnLine = mblnOnLine
End Property
Friend Property Let OnLine(ByVal blnOnLine As Boolean)
    mblnOnLine = blnOnLine
    If blnOnLine Then
        Me.OnLineNew = Val(zlDatabase.GetPara(38, Me.SysNo, "0")) = 1
    End If
End Property
'OnLineNew:病案与HIS联机使用时仍允许新增病案，与参数32(OnLine)配合使用
Public Property Get OnLineNew() As Boolean
    OnLineNew = mblnOnLineNew
End Property
Friend Property Let OnLineNew(ByVal blnOnLineNew As Boolean)
    mblnOnLineNew = blnOnLineNew
End Property
'InputOutList:要求在住院科室日报中输入出院病人清单
Public Property Get InputOutList() As Boolean
    InputOutList = mblnInputOutList
End Property
'FreeFormat:费用格式字符串
Public Property Get FreeFormat() As String
    FreeFormat = mstrFreeFormat
End Property

'3、模块参数：
 'Check损伤中毒:主要出院诊断编码为S、T类时损伤中毒诊断：1-必须填写（不是ST开头的则禁止填写），2-提示是否填写，0-不检查
Public Property Get Check损伤中毒() As Integer
    Check损伤中毒 = mintCheck损伤中毒
End Property
'Check病理诊断:主要出院诊断编码为C、D类时病理诊断：1-必须填写（不是CD开头的则禁止填写），2-提示是否填写，0-不检查
Public Property Get Check病理诊断() As Integer
    Check病理诊断 = mintCheck病理诊断
End Property

Public Property Get CheckICD附码() As Integer
    CheckICD附码 = mintCheckICD附码
End Property

 'IsStructAdress:设置住院首页中病人地址是否启用结构化显示和存储
Public Property Get IsStructAdress() As Boolean
    IsStructAdress = mblnStructAdress
End Property
Public Property Get M病理() As Boolean
    M病理 = mblnM病理
End Property
'IsShowTown:设置住院首页中病人地址中乡镇地址是否结构化输入
Public Property Get IsShowTown() As Boolean
    IsShowTown = mblnShowTown
End Property
'NotUseXYItems:中医首页里不使用的项目：病例分型，输液反应，输血反应，输红细胞，输血小板，输血浆，输全血，自体回收，输其他，输血前的9项检查，HBsAg，HCV-Ab，HIV-Ab，示教病案，科研病案，随诊，随诊期限，呼吸机使用，研究生医师
Public Property Get NotUseXYItems() As Boolean
    NotUseXYItems = mblnNotUseXYItems
End Property
'Check区域:首页中保存时区域：1-必须填写，2-提示是否填写，0-不检查
Public Property Get Check区域() As Integer
    Check区域 = mintCheck区域
End Property
'UseOPSEndTime:是否使用手术结束时间
Public Property Get UseOPSEndTime() As Boolean
    UseOPSEndTime = mblnUseOPSEndTime
End Property
'SeparateEdit:不良事件的项目、输液反应、引发药物、临床表现、住院期间身体约束、离院时透析(血透、腹透)尿素氮值等信息在启用该参数时只能由护士填写
Public Property Get SeparateEdit() As Boolean
    SeparateEdit = mblnSeparateEdit
End Property
'DiagInputZY:系统允许选择诊断输入方式（DiagInput=1）时住院中医诊断输入：0-根据诊断标准输入,1-根据疾病编码输入
Public Property Get DiagInputZY() As Integer
    '病案首页只读取疾病编码
    DiagInputZY = IIf(Me.FuncType = f病案首页, 1, mintDiagInputZY)
End Property
Public Property Let DiagInputZY(ByVal intDiagInputZY As Integer)
    mintDiagInputZY = intDiagInputZY
    If Me.PatiType = PF_门诊 Then
        mintDiagInputXY = mintDiagInputZY
    End If
End Property
'DiagInputXY:系统允许选择诊断输入方式（DiagInput=1）时住院西医诊断输入：0-根据诊断标准输入,1-根据疾病编码输入
Public Property Get DiagInputXY() As Integer
    '病案首页只读取疾病编码
    DiagInputXY = IIf(Me.FuncType = f病案首页, 1, mintDiagInputXY)
End Property
Public Property Let DiagInputXY(ByVal intDiagInputXY As Integer)
    mintDiagInputXY = intDiagInputXY
    If Me.PatiType = PF_门诊 Then
        mintDiagInputZY = mintDiagInputXY
    End If
End Property
'DiagInputMZ:系统允许选择诊断输入方式（DiagInput=1）时门诊诊断输入：0-根据诊断标准输入,1-根据疾病编码输入
Public Property Get DiagInputMZ() As Integer
    DiagInputMZ = mintDiagInputMZ
End Property
Public Property Let DiagInputMZ(ByVal intDiagInputMZ As Integer)
    mintDiagInputMZ = intDiagInputMZ
End Property
'AllerInput:模块参数，过敏输入来源：0-按药品目录输入，1-按过敏源输入
Public Property Get AllerInput() As Integer
    AllerInput = mintAllerInput
End Property
Public Property Let AllerInput(ByVal intAllerInput As Integer)
    mintAllerInput = intAllerInput
End Property
'OPSInput:选择手术输入方式时：0-根据诊疗项目输入,1-根据ICD9-CM3输入
Public Property Get OPSInput() As Integer
    OPSInput = IIf(mfeFunc = f病案首页, 1, mintOPSInput)
End Property
Public Property Let OPSInput(ByVal intOPSInput As Integer)
    mintOPSInput = intOPSInput
End Property
'OPSFree:未找到时手术名称是否允许自由录入
Public Property Get OPSFree() As Boolean
    OPSFree = mblnOPSFree
End Property
Public Property Let OPSFree(ByVal blnOPSFree As Boolean)
    mblnOPSFree = blnOPSFree
End Property

'PathVCauses:路径变异原因从字典表中选取；不允许自由录入；
Public Property Get PathVCauses() As Boolean
    PathVCauses = mblnPathVCauses
End Property
Public Property Let PathVCauses(ByVal blnPathVCauses As Boolean)
    mblnPathVCauses = blnPathVCauses
End Property

Public Property Get GetExtraCode() As Boolean
    GetExtraCode = mblnGet附码
End Property
Friend Property Let GetExtraCode(ByVal blnGet附码 As Boolean)
    mblnGet附码 = blnGet附码
End Property

Public Property Get GetMedical() As Boolean
    GetMedical = mblnGet医疗机构
End Property
Friend Property Let GetMedical(ByVal blnGet医疗机构 As Boolean)
    mblnGet医疗机构 = blnGet医疗机构
End Property

Public Property Get AddPathologic() As Boolean
    AddPathologic = mbln病理号
End Property
Friend Property Let AddPathologic(ByVal bln病理号 As Boolean)
    mbln病理号 = bln病理号
End Property

Private Sub TestCode()
'    mblnTCM = True
'    mstrPrivs = ""
'    memOpenMode = EM_查阅
End Sub

Public Function InitProperty(ByVal feCurrent As FuncEnum, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule, Optional objCommunity As Object, Optional objInsure As clsInsure) As Boolean
'功能：初始化属性类
    mstMedPage = -1
    mfeFunc = feCurrent
    mlngSys = lngSys
    mlngModule = lngModule
    Call GetUserInfo
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    mblnShareMedRec = CheckShare(300, 100)
    If mfeFunc = f诊断选择 Then
        Set mrsDiagConn = New ADODB.Recordset
        '记录申请单关联诊断
        With mrsDiagConn
            .Fields.Append "诊断ID", adInteger '病人诊断记录ID
            .Fields.Append "标识ID", adInteger
            .Fields.Append "疾病目录ID", adInteger '疾病ID
            .Fields.Append "诊断目录ID", adInteger '诊断ID
            .Fields.Append "类型", adInteger '0,Null,诊断选择数据；1-传染病数据
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenStatic
            .Open
        End With
        Call InitCacheRecInfo
    End If
    InitProperty = True
End Function

Public Function RefreshPara() As Boolean
'功能：刷新参数
'说明：Me.属性名=Value,这种属性赋值主要用于参数赋值时有其他特殊处理
    Dim lngModule As Long
    Dim strTmp As String

    Call ClearVarValues
    '全局参数：
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    mstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    mblnDo = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    '系统参数
    If Me.FuncType <> f病案首页 And Me.FuncType <> f诊断选择 Then
        mblnMaskID = Val(zlDatabase.GetPara(247, Me.SysNo, , 0)) = 1
    End If
    If Me.FuncType <> f电子病案 Then
        If Me.FuncType = f医生首页 Then
            '合理用药接口类型，0-未启用，1-四川美通，2-大通，3-太元通
            mbytPass = Val(zlDatabase.GetPara(30, Me.SysNo))
            '过敏输入来源控制
            mintAllerSource = Val(zlDatabase.GetPara(224, Me.SysNo))
            '诊断输入来源
            mintDiagInput = Val(zlDatabase.GetPara(55, Me.SysNo, , 1))
            '电子签名认证中心
            gintCA = Val(zlDatabase.GetPara(25, gclsPros.SysNo))
            '电子签名控制场合
            gstrESign = zlDatabase.GetPara(26, gclsPros.SysNo)
            '读取部门启用数据
            Set grsSign = New ADODB.Recordset
            grsSign.Fields.Append "部门ID", adBigInt
            grsSign.Fields.Append "场合", adBigInt
            grsSign.Fields.Append "是否启用", adBigInt
            grsSign.CursorLocation = adUseClient
            grsSign.LockType = adLockOptimistic
            grsSign.CursorType = adOpenStatic
            grsSign.Open
        ElseIf Me.FuncType = f诊断选择 Then
            '诊断输入来源
            mintDiagInput = Val(zlDatabase.GetPara(55, Me.SysNo, , 1))
        End If
        If Me.FuncType <> f病案首页 Then
            '诊断输入方式
            strTmp = Trim(zlDatabase.GetPara(65, gclsPros.SysNo) & "") & "11"
            mintDiagSourceMZ = Val(Mid(strTmp, 1, 1)): mintDiagSourceMZ = IIf(mintDiagSourceMZ = 0, 1, mintDiagSourceMZ)
            mintDiagSourceZY = Val(Mid(strTmp, 2, 1)): mintDiagSourceZY = IIf(mintDiagSourceZY = 0, 1, mintDiagSourceZY)
            mblnFreeInput = Val(zlDatabase.GetPara("诊断手术名称自由调整", gclsPros.SysNo, 0)) = 1
        Else
            Me.NewInNo = Val(zlDatabase.GetPara(145, IIf(Me.ShareMedRec, 100, Me.SysNo), , 0)) = 1
            Me.FreeDeci = Val(zlDatabase.GetPara(86, Me.SysNo, , 2))
            mblnEditPageNo = Val(zlDatabase.GetPara(40, Me.SysNo, , 0)) = 1
            mblnUseFileRules = (Val(zlDatabase.GetPara(41, Me.SysNo, , 0)) = 1)
            mblnInputOutList = Val(zlDatabase.GetPara(34, Me.SysNo, , "1")) = 1
            mstrOutFile = zlDatabase.GetPara(37, Me.SysNo)
            Me.OnLine = Val(zlDatabase.GetPara(32, Me.SysNo, , "1")) = 1
        End If
    End If
    '模块参数
    If Me.FuncType = f病案首页 Then
        '设置本源模块值
        lngModule = p病案管理
        mintCheck损伤中毒 = Val(zlDatabase.GetPara("损伤中毒检查", Me.SysNo, lngModule, 0))
        mintCheck病理诊断 = Val(zlDatabase.GetPara("病理诊断检查", Me.SysNo, lngModule, 0))
        mintCheckICD附码 = Val(zlDatabase.GetPara("ICD附码检查", Me.SysNo, lngModule, 0))
        mblnM病理 = (Val(zlDatabase.GetPara("病理诊断只允许录入肿瘤形态学编码", Me.SysNo, lngModule, 0)) = 1)
        If Not Me.ShareMedRec Then
            mblnStructAdress = Val(zlDatabase.GetPara(251, Me.SysNo)) <> 0 '病人地址结构化录入
            mblnShowTown = Val(zlDatabase.GetPara(252, Me.SysNo)) <> 0 '乡镇地址结构化录入
            mblnOutDocCtrl = Val(zlDatabase.GetPara(253, Me.SysNo)) <> 0 '外院医生必须建档
            mblnNotUseXYItems = Val(zlDatabase.GetPara("中医科室不使用西医病案首页项目", Me.SysNo, lngModule, 0)) <> 0
            mblnPathVCauses = Val(zlDatabase.GetPara("路径变异原因从字典表中选取", Me.SysNo)) <> 0
            mblnGet附码 = Val(zlDatabase.GetPara("诊断自动提取附码", Me.SysNo)) = 1
            mblnGet医疗机构 = Val(zlDatabase.GetPara("医疗机构不允许自由录入", Me.SysNo)) = 1
            mbln病理号 = Val(zlDatabase.GetPara("不填写病理号", Me.SysNo)) = 1
        Else
            mblnStructAdress = Val(zlDatabase.GetPara(251, 100)) <> 0 '病人地址结构化录入
            mblnShowTown = Val(zlDatabase.GetPara(252, 100)) <> 0 '乡镇地址结构化录入
            mblnOutDocCtrl = Val(zlDatabase.GetPara(253, 100)) <> 0 '外院医生必须建档
            mblnNotUseXYItems = Val(zlDatabase.GetPara("中医科室不使用西医病案首页项目", 100, p住院医生站, 0)) <> 0
            mblnPathVCauses = Val(zlDatabase.GetPara("路径变异原因从字典表中选取", 100)) <> 0
            mblnGet附码 = Val(zlDatabase.GetPara("诊断自动提取附码", 100)) = 1
            mblnGet医疗机构 = Val(zlDatabase.GetPara("医疗机构不允许自由录入", 100)) = 1
            mbln病理号 = Val(zlDatabase.GetPara("不填写病理号", 100)) = 1
        End If
        mblnDefautADD = Val(zlDatabase.GetPara("户口及联系人地址默认与出生地相同", Me.SysNo, lngModule)) = 1
        mblnCNIndent = Val(zlDatabase.GetPara("诊断编码和描述互相独立录入", Me.SysNo, lngModule)) = 1
        mblnAuto病理 = Val(zlDatabase.GetPara("手术病人自动填写病理诊断", Me.SysNo, lngModule)) = 1
        mblnSameName = Val(zlDatabase.GetPara("允许费用名重复", Me.SysNo, lngModule)) = 1
        mblnAutoDaig = Val(zlDatabase.GetPara("自动提取西医或中医诊断", Me.SysNo, lngModule)) = 1
        mblnDaigFree = Val(zlDatabase.GetPara("允许中医诊断疾病编码为空", Me.SysNo, lngModule)) = 1
        mblnNull出院情况 = Val(zlDatabase.GetPara("允许不填写诊断出院情况", Me.SysNo, lngModule)) = 1
        mblnEditUnrecive = Val(zlDatabase.GetPara("只允许编目已接收的病案", Me.SysNo, lngModule)) <> 1
        mblnEditPayType = Val(zlDatabase.GetPara("编辑医疗付款方式", Me.SysNo, lngModule)) = 1
        mblnTabPayType = Val(zlDatabase.GetPara("录入付款方式", Me.SysNo, lngModule)) = 1
        mblnTabAgeUnit = Val(zlDatabase.GetPara("编辑年龄单位", Me.SysNo, lngModule)) = 1
        mblnTabNation = Val(zlDatabase.GetPara("编辑国籍", Me.SysNo, lngModule)) = 1
        mblnTabFileNo = Val(zlDatabase.GetPara("录入档案号", Me.SysNo, lngModule)) = 1
        mblnTabReadm = Val(zlDatabase.GetPara("录入再入院", Me.SysNo, lngModule)) = 1
        mblnTabXRaysNo = Val(zlDatabase.GetPara("录入X线号", Me.SysNo, lngModule)) = 1
    ElseIf Me.FuncType = f医生首页 Or Me.FuncType = f诊断选择 Then
        '设置本源模块值
        lngModule = IIf(Me.PatiType = PF_住院, p住院医生站, p门诊医生站)
        If Me.PatiType = PF_住院 Then
            mintCheck损伤中毒 = Val(zlDatabase.GetPara("损伤中毒检查", Me.SysNo, lngModule, 0))
            mintCheck病理诊断 = Val(zlDatabase.GetPara("病理诊断检查", Me.SysNo, lngModule, 0))
            mintCheckICD附码 = Val(zlDatabase.GetPara("ICD附码检查", Me.SysNo, lngModule, 0))
            mblnM病理 = (Val(zlDatabase.GetPara("病理诊断只允许录入肿瘤形态学编码", Me.SysNo, lngModule, 0)) = 1)
            mblnGet附码 = Val(zlDatabase.GetPara("诊断自动提取附码", Me.SysNo)) = 1
            mblnGet医疗机构 = Val(zlDatabase.GetPara("医疗机构不允许自由录入", Me.SysNo)) = 1
            mbln病理号 = Val(zlDatabase.GetPara("不填写病理号", Me.SysNo)) = 1
            If Me.FuncType <> f诊断选择 Then
                mblnStructAdress = Val(zlDatabase.GetPara(251, Me.SysNo)) <> 0 '病人地址结构化录入
                mblnShowTown = Val(zlDatabase.GetPara(252, Me.SysNo)) <> 0 '乡镇地址结构化录入
                mblnOutDocCtrl = Val(zlDatabase.GetPara(253, Me.SysNo)) <> 0 '外院医生必须建档
                mblnNotUseXYItems = Val(zlDatabase.GetPara("中医科室不使用西医病案首页项目", Me.SysNo, lngModule, 0)) <> 0
                mintCheck区域 = Val(zlDatabase.GetPara("区域检查", Me.SysNo, lngModule, 1))
                mblnPathVCauses = Val(zlDatabase.GetPara("路径变异原因从字典表中选取", Me.SysNo)) <> 0
            End If
        End If
    ElseIf Me.FuncType = f电子病案 Then
        '设置本源模块值
        lngModule = p电子病案查阅
    End If

    If Me.FuncType <> f诊断选择 And Me.PatiType = PF_住院 Then
        mblnUseOPSEndTime = Val(zlDatabase.GetPara("使用手术结束时间", IIf(Me.FuncType = f病案首页, 100, Me.SysNo), p住院医生站, 0)) <> 0
        If Me.FuncType = f医生首页 Then
            mblnSeparateEdit = Val(zlDatabase.GetPara("医生和护士分别填写病案首页", Me.SysNo, p住院医生站, 0)) = 1
        End If
    End If
End Function

Private Sub Class_Terminate()
    Call ClearVarValues(True)
End Sub

Private Sub ClearVarValues(Optional ByVal blnClassTerminate As Boolean)
'清空变量
    If blnClassTerminate Then
        mfeFunc = 0
        mlngSys = 0
        mlngModule = 0
        mstrPrivs = ""
        If Not mfrmCurrent Is Nothing Then
            Unload mfrmCurrent
            Set mfrmCurrent = Nothing
        End If
        Set mfrmMain = Nothing
        Set mrsMainInfo = Nothing
        Set mrsSecdInfo = Nothing
        Set mrsBaseCode = Nothing
        Set mrsManInfo = Nothing
        Set mrsDiagConn = Nothing
        Set mrsPatiOut = Nothing
        Set mrsFeesOut = Nothing
        mstrLoadMans = ""
        mstMedPage = 0
        mblnShareMedRec = False
        mblnReadPages = False
        mlng病人ID = 0
        mlng主页ID = 0
        mstrRegistNo = ""
        mpfPatiType = 0
        mlngValueLen = 0
        mblnOK = False
        Set mclsMipModule = Nothing
        mblnOpen = False
    End If
    mblnLastDiag = False
    mblnExistPati = False
    mblnOnlyPatiInfo = False
    mitInNoType = IT_NotVili
    If mfeFunc <> f病案首页 Then memOpenMode = EM_编辑
    mlngFontSize = 9
    mbln护士站 = False
    mlng入院科室ID = 0: mlng出院科室ID = 0
    mblnTCM = False: mbln产科 = False
    mpsCur = PS_未导入
    mstrMustDiagType = ""
    mintInsureType = 0
    mblnHave中医 = False
    mblnMoved = False
    mblnInfosChange = False
    mlngTXTProc = 0
    mlngDateIndex = 0
    mblnReturn = False
    mbln手术 = False
    mstrPatiName = "": mstrSex = ""
    mstr入院时间 = "": mstr出院时间 = ""
    mdtInPath = 0
    mstrPathDiag = ""
    mblnIsPathOutTime = False
    mstr疾病IDs = "": mstr诊断IDs = ""
    Set mrsAuxiInfo = Nothing: Set mrsPatiInfo = Nothing
    Set mrsAdressInfo = Nothing
    mblnUseTYT = False
    mint社区 = 0: mstr社区号 = ""
    mstr住院号 = ""
    mbln编目 = False
    mblnChange编码员 = False
    mblnSigned = False
    mbln复诊 = False
    mlngAplyMark = 0: mlngAdviceID = 0
    mstrDiagRowIDs = "": mstrDiagNames = ""
    mblnMerged = False
    mblnDiagChange = False
    mblnDiagInput = False
    mblnMaskID = False
    mblnLoadFinish = False
End Sub

Public Sub InitFacePara()
'功能：初始化首页界面以及诊断选择界面的参数
    Dim lngModule As Long, strTmp As String
    Dim bln参数设置 As Boolean
    Dim lngTmp As Integer

    bln参数设置 = InStr(Me.Privs, "参数设置") > 0
    If Me.FuncType = f医生首页 Or Me.FuncType = f诊断选择 Then
        '设置本源模块值
        With Me.CurrentForm
            lngModule = IIf(Me.PatiType = PF_住院, p住院医生站, p门诊医生站)
            If Me.PatiType = PF_住院 Then
                If Me.FuncType <> f诊断选择 Then
                    strTmp = zlDatabase.GetPara("手术情况输入", Me.SysNo, lngModule, "00", Array(.OptParaOPSInfo(PC_按诊疗项目输入), .OptParaOPSInfo(PC_按ICDCM9编码输入), .chkParaOPSInfo(PC_未找到时自由录入)), bln参数设置) & "00"
                    mintOPSInput = Val(Mid(strTmp, 1, 1)): mblnOPSFree = Val(Mid(strTmp, 2, 1)) <> 0
                    .OptParaOPSInfo(IIf(mintOPSInput <> 0, PC_按ICDCM9编码输入, PC_按诊疗项目输入)).Value = True: .chkParaOPSInfo(PC_未找到时自由录入).Value = IIf(mblnOPSFree, 1, 0)
                End If
            End If
            '设置控件颜色状态
            mintDiagInputXY = Val(zlDatabase.GetPara(IIf(Me.PatiType = PF_住院, "西医诊断输入", "门诊诊断输入"), Me.SysNo, lngModule, 0, Array(.optDiag(PC_XY按诊断输入), .optDiag(PC_XY按疾病编码输入)), bln参数设置))
            lngTmp = Me.DiagInputXY
            Call SetInputRoot(0, Me.DiagInput, lngTmp, .optDiag(PC_XY按诊断输入), .optDiag(PC_XY按疾病编码输入))
            Me.DiagInputXY = lngTmp
            If Me.FuncType = f诊断选择 Or Me.PatiType = PF_住院 Then
                mintDiagInputZY = Val(zlDatabase.GetPara(IIf(Me.PatiType = PF_住院, "西医诊断输入", "门诊诊断输入"), Me.SysNo, lngModule, 0, Array(.optDiag(PC_ZY按诊断输入), .optDiag(PC_ZY按疾病编码输入)), bln参数设置))
                lngTmp = Me.DiagInputXY
                Call SetInputRoot(1, Me.DiagInput, lngTmp, .optDiag(PC_ZY按诊断输入), .optDiag(PC_ZY按疾病编码输入))
                Me.DiagInputZY = lngTmp
            End If
            If Me.FuncType <> f诊断选择 Then
                mintAllerInput = zlDatabase.GetPara("过敏输入来源", Me.SysNo, lngModule, 0, Array(.optAller(PC_按药品目录输入), .optAller(PC_按过敏源输入)))
                If Me.PassType = 3 And (Me.AllerSource = 2 Or Me.AllerSource = 0) Then
                    If gobjPass Is Nothing Then
                        On Error Resume Next
                        Set gobjPass = CreateObject("zlPassInterface.clsPass")
                        If Err.Number <> 0 Then Err.Clear: Me.PassType = 0
                        On Error GoTo 0
                        If gobjPass Is Nothing Then
                            Me.PassType = 0
                        Else
                            If gobjPass.zlPassInit(gcnOracle, Me.SysNo, 5) = False Then
                                Me.PassType = 0
                                Set gobjPass = Nothing
                            End If
                        End If
                    End If
                End If
                If Not gobjPass Is Nothing Then
                    Call SetInputRoot(2, Me.AllerSource, mintAllerInput, .optAller(PC_按药品目录输入), .optAller(PC_按过敏源输入))
                Else
                    Call SetInputRoot(mintAllerInput, 1, mintAllerInput, .optAller(PC_按药品目录输入), .optAller(PC_按过敏源输入))
                End If
            End If
        End With
    End If
End Sub

Public Function InitCacheRecInfo(Optional ByVal blnMedSaved As Boolean) As Boolean
'功能：初始化首页信息记录集
'参数：blnMedSaved=是否是首页保存后初始化缓存记录集，False=初始加载数据初始化，True=保存首页后初始化
'返回：是否初始化成功
'说明：保存首页后初始化:主要是清空删除的行记录信息，并将信息现值赋值给信息原值，初始化改变状态
'         初始加载数据初始化:主要是清空动态加载型控件的行记录，并清空原值与现值，初始化改变状态
    Dim strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long

    On Error GoTo errH
    '首页标准未改变，不用初始化记录集
    If Not Me.MainInfoRec Is Nothing Then
        '病案首页可以在首页界面新增病案与首页，点击确定需要将缓存记录集回复初始装
        '这几个信息Tag用来辅助判断是否保存，因此需要这样更新
        strTmp = "西医诊断;中医诊断;过敏药物;手麻情况"
        arrTmp = Split(strTmp, ";")
        For i = LBound(arrTmp) To UBound(arrTmp)
            gclsPros.MainInfoRec.Filter = "信息名='" & arrTmp(i) & "'"
            If Not gclsPros.MainInfoRec.EOF Then
                Call Rec.Update(gclsPros.SecdInfoRec, "序号=" & gclsPros.MainInfoRec!序号, "Tag", IIf(gclsPros.FuncType = f病案首页, 4, 3))
            End If
        Next
        '住院首页在预览打印时需要保存，此时原有信息不能被清空
        If blnMedSaved And gclsPros.OpenMode <> EM_新增病案 And gclsPros.OpenMode <> EM_新增首页 Then
            Call Rec.Delete(gclsPros.SecdInfoRec, "改变状态=" & CS_删除行)

            gclsPros.SecdInfoRec.Filter = "改变状态 = 3"
            If gclsPros.SecdInfoRec.RecordCount > 0 Then
                gclsPros.SecdInfoRec.MoveFirst
                For j = 1 To gclsPros.SecdInfoRec.RecordCount
                    gclsPros.SecdInfoRec.Update "改变状态", 0
                    gclsPros.SecdInfoRec.MoveNext
                Next
            End If
            
            gclsPros.SecdInfoRec.Filter = ""
            gclsPros.SecdInfoRec.Sort = "Sort"
            If gclsPros.SecdInfoRec.RecordCount > 0 Then
                gclsPros.SecdInfoRec.MoveFirst
                For j = 1 To gclsPros.SecdInfoRec.RecordCount
                    With gclsPros.SecdInfoRec
                        Call gclsPros.SecdInfoRec.Update(Array("信息原值", "主信息原值", "信息现值", "主信息现值", "改变状态"), Array(!信息现值, !主信息现值, Null, Null, CS_未改变))
                        .MoveNext
                    End With
                Next
            End If
            '之所以用上面这两段代码不用下面这一句，是因为用下面这一句在抗菌药物记录会出问题
'            Call rec.Update(gclsPros.SecdInfoRec, "LJH", "信息原值", "!信息现值", "主信息原值", "!主信息现值", "信息现值", Null, "主信息现值", Null, "改变状态", CS_未改变)
            Call Rec.Update(gclsPros.MainInfoRec, "", "信息原值", "!信息现值", "是否改变", 0)
        Else
            '恢复初始状态
            Me.MainInfoRec.Filter = "": Me.MainInfoRec.Sort = "序号"
            Do While Not Me.MainInfoRec.EOF
                If Me.MainInfoRec!ExpState = 2 Then
                    Call Rec.Delete(Me.SecdInfoRec, "序号=" & Me.MainInfoRec!序号)
                End If
                Me.MainInfoRec.Update Array("信息原值", "信息现值", "是否改变"), Array(Null, Null, 0)
                Me.MainInfoRec.MoveNext
            Loop
            Me.SecdInfoRec.Filter = "": Me.SecdInfoRec.Sort = "Sort"
            'Tag值不用处理因为只有加载扩展的项目才使用Tag，而此类项目已经删除
            Do While Not Me.SecdInfoRec.EOF
                Me.SecdInfoRec.Update Array("信息原值", "信息现值", "主信息原值", "主信息现值", "改变状态"), Array(Null, Null, Null, Null, CS_未改变)
                Me.SecdInfoRec.MoveNext
            Loop
        End If
        InitCacheRecInfo = True
        Exit Function
    End If

    '首页标准改变，初始化记录集
    '1、主记录结构定义
    Set Me.MainInfoRec = New ADODB.Recordset
    With Me.MainInfoRec
        .Fields.Append "序号", adInteger, , adFldKeyColumn              '主键，标识信息
        .Fields.Append "信息名", adVarChar, 100, adFldKeyColumn   '信息名称
        '该记录集仅记录一个信息对应一个控件的情况或多个信息对应一个控件，其他情况不填写
        .Fields.Append "控件名", adVarChar, 100, adFldIsNullable      '展示信息的控件名称
        .Fields.Append "Index", adInteger, , adFldIsNullable                '为空时表示不是控件数组
        .Fields.Append "ExpState", adInteger                                        '信息扩展状态，0-不扩展，1-初始扩展，2-加载扩展
        .Fields.Append "页码", adInteger                                                '信息所在的页码
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable  '信息在首页加载时的值
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable  '信息在首页检查时的值
        .Fields.Append "Edit", adInteger                                                 '0-可编辑,1-不可编辑，只用于展示,2-不可编辑不保存
        .Fields.Append "是否改变", adInteger                                          '信息是否有改变0-未改变，1-改变了
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    '2、次级信息记录集结构定义
    Set Me.SecdInfoRec = New ADODB.Recordset
    With Me.SecdInfoRec
        .Fields.Append "Sort", adInteger                                  '本记录集的主键
        .Fields.Append "序号", adInteger                                  '标识信息，引用主记录集
        .Fields.Append "控件名", adVarChar, 100                           '展示信息的控件名称
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable            '行号或控件数组Index
        .Fields.Append "页码", adInteger                                  '信息所在的页码
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable      '信息在首页加载时的值
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable      '信息在首页检查时的值
        .Fields.Append "主信息原值", adVarChar, 2000, adFldIsNullable    '信息的主要部分，标识一个信息是否被彻底改变，信息在首页加载时的值
        .Fields.Append "主信息现值", adVarChar, 2000, adFldIsNullable    '信息在首页检查时的值
        .Fields.Append "Edit", adInteger                                                      '0-可编辑,1-不可编辑，只用于展示,2-不可编辑不保存
        .Fields.Append "改变状态", adInteger                                              '信息改变程度0-未改变，1-次级信息改变，2-主信息改变,3-新增,-1，删除
        .Fields.Append "ID", adBigInt, , adFldIsNullable                             '信息行在数据库中的ID,一般表格类控件使用
        .Fields.Append "Tag", adVarChar, 2000                                           '存储额外数据
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    '3、定义信息记录集内容
    If Me.FuncType = f医生首页 Then
        If Me.PatiType = PF_门诊 Then
            If Not GetOutMedRecCache Then Exit Function
        Else
            If Not GetInMedRecCache Then Exit Function
        End If
    ElseIf Me.FuncType = f电子病案 Then
        If Me.PatiType = PF_门诊 Then
            If Not GetOutArchiveRecCache Then Exit Function
        Else
            If Not GetInArchiveRecCache Then Exit Function
        End If
    ElseIf Me.FuncType = f病案首页 Then
        If Not GetPageMedRecCache Then Exit Function
    ElseIf Me.FuncType = f诊断选择 Then
        If Not GetDiagSelRecCache Then Exit Function
    End If

    InitCacheRecInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function RollBackCacheRecInfo() As Boolean
'功能：回滚初首页信息记录集
'参数：回滚信息保存

    On Error GoTo errH
    '恢复初始状态
    Me.MainInfoRec.Filter = "": Me.MainInfoRec.Sort = "序号"
    Do While Not Me.MainInfoRec.EOF
        If Me.MainInfoRec!ExpState = 2 Then
             '将信息原值不为空的恢复为未改变状态，这种中可能有新增行
            Call Rec.Update(Me.SecdInfoRec, "序号=" & Me.MainInfoRec!序号 & " And  信息原值<>Null", "改变状态", CS_未改变)
            '将新增行删除
            Call Rec.Delete(Me.SecdInfoRec, "序号=" & Me.MainInfoRec!序号 & " And 改变状态=" & CS_新增行)
        End If
        Me.MainInfoRec.Update Array("信息现值", "是否改变"), Array(Null, 0)
        Me.MainInfoRec.MoveNext
    Loop
    '清空保存值
    Me.SecdInfoRec.Filter = "": Me.SecdInfoRec.Sort = "Sort"
    'Tag值不用处理因为只有加载扩展的项目才使用Tag，而此类项目已经删除
    Do While Not Me.SecdInfoRec.EOF
        Me.SecdInfoRec.Update Array("信息现值", "主信息现值", "改变状态"), Array(Null, Null, CS_未改变)
        Me.SecdInfoRec.MoveNext
    Loop
    RollBackCacheRecInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function GetInMedRecCache() As Boolean
'获取住院首页缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long

    On Error GoTo errH
    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        '基本信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "住院号", "txtSpecificInfo", SLC_住院号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "主页ID", "txtSpecificInfo", SLC_入院次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "再入院", "chkInfo", CHK_再入院)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "医疗付款方式", "cboBaseInfo", BCC_付款方式)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "姓名", "txtInfo", GC_姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "性别", "cboBaseInfo", BCC_性别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生日期", "mskDateInfo", DC_出生日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "年龄", "txtSpecificInfo", SLC_年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "身高", "txtSpecificInfo", SLC_身高)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "体重", "txtSpecificInfo", SLC_体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "婚姻状况", "cboBaseInfo", BCC_婚姻)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "不足周岁年龄", "txtSpecificInfo", SLC_婴幼儿年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "新生儿出生体重", "txtSpecificInfo", SLC_新生儿出生体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "新生儿入院体重", "txtSpecificInfo", SLC_新生儿入院体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "国籍", "cboBaseInfo", BCC_国籍)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "民族", "cboBaseInfo", BCC_民族)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "职业", "cboBaseInfo", BCC_职业)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出生地点", "txtAdressInfo", ADRC_出生地点)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "籍贯", "txtAdressInfo", ADRC_籍贯)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号状态", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "外籍身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "其他证件", "txtInfo", GC_其他证件)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址", "txtAdressInfo", ADRC_现住址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭电话", "txtSpecificInfo", SLC_家庭电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址邮编", "txtSpecificInfo", SLC_家庭邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址", "txtAdressInfo", ADRC_户口地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址邮编", "txtSpecificInfo", SLC_户口邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位地址", "txtAdressInfo", ADRC_单位地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位电话", "txtSpecificInfo", SLC_单位电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位邮编", "txtSpecificInfo", SLC_单位邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人姓名", "txtInfo", GC_联系人姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人关系", "cboBaseInfo", BCC_关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人附加信息", "txtInfo", GC_其他关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人电话", "txtSpecificInfo", SLC_联系人电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人地址", "txtAdressInfo", ADRC_联系人地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "区域", "txtAdressInfo", ADRC_病人区域)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "健康号", "txtInfo", BCC_健康卡号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "监护人身份证号", "txtInfo", GC_监护人身份证号)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "Email", "txtInfo", GC_Email)
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "Qq", "txtSpecificInfo", SLC_QQ)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院日期", "mskDateInfo", DC_入院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院科室", "txtInfo", GC_入院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院病室", "txtInfo", GC_入院病房)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院方式", "cboBaseInfo", BCC_入院途径)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院转入", "txtInfo", GC_入院转入)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院前经外院治疗", "chkInfo", CHK_入院前外院治疗)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_初始扩展, IS_全保存, "转科记录", "txtInfo", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "txtInfo", GC_转科1)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "txtInfo", GC_转科2)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "txtInfo", GC_转科3)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出院日期", "mskDateInfo", DC_出院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出院科室", "txtInfo", GC_出院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出院病室", "txtInfo", GC_出院病房)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "住院天数", "txtSpecificInfo", SLC_住院天数)
        '西医诊断（包含中医诊断诊断符合情况）
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "感染与死亡关系", "cboBaseInfo", BCC_感染与死亡关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "感染部位", "lstInfectParts", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "入院病况", "cboBaseInfo", BCC_入院情况)
        If Me.MedPageSandard = ST_四川省标准 Then
             .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "告病重病危", "chkInfo", CHK_住院出现危重)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "是否确诊", "chkInfo", CHK_是否确诊)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "确诊日期", "mskDateInfo", DC_确诊日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "病理号", "txtInfo", GC_病理号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "分化程度", "cboBaseInfo", BCC_分化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "最高诊断依据", "cboBaseInfo", BCC_最高诊断依据)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_全保存, "诊断符合情况", "cboBaseInfo", Null)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_门诊与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_入院与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_门诊与入院)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_术前与术后)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_放射与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_临床与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_临床与尸检)
             '中医诊断
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_门诊与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_入院与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_辩证)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_治法)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_方药)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "死亡时间", "mskDateInfo", DC_死亡时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "死亡根本原因", "txtInfo", GC_死亡原因)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "住院死亡期间", "cboBaseInfo", BCC_死亡期间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "病原学检查", "chkInfo", CHK_病原学检查)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "尸检标志", "cboBaseInfo", BCC_死亡患者尸检)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "新发肿瘤", "chkInfo", CHK_新发肿瘤)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "病原学诊断", "txtInfo", GC_病原学诊断)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "抢救次数", "txtSpecificInfo", SLC_抢救次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "成功次数", "txtSpecificInfo", SLC_成功次数)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "疑难病历", "chkInfo", CHK_疑难病例)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "抢救病因", "txtInfo", GC_抢救病因)
        '中医诊断
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_加载扩展, IS_全保存, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医危重", "chkInfo", CHK_危重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医急症", "chkInfo", CHK_急症)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医疑难", "chkInfo", CHK_疑难)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医治疗类别", "cboBaseInfo", BCC_治疗类别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医抢救方法", "cboBaseInfo", BCC_抢救方法)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "自制中药制剂", "cboBaseInfo", BCC_自制中药制剂)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医设备", "cboBaseInfo", BCC_中医诊疗设备)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医技术", "cboBaseInfo", BCC_中医诊疗技术)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "辨证施护", "cboBaseInfo", BCC_辨证施护)
        '过敏与手术
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "无过敏记录", "chkInfo", CHK_无过敏记录)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_全保存, "过敏药物", "vsAller", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_全保存, "手麻情况", "vsOPS", Null)
        If Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "围术期死亡", "chkInfo", CHK_围术期死亡)
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "术后猝死", "chkInfo", CHK_术后猝死)
        End If
        '住院情况
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "病例分型", "cboBaseInfo", BCC_病例分型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "血型", "cboBaseInfo", BCC_血型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "RH", "cboBaseInfo", BCC_RH)
        If Me.MedPageSandard <> ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输液反应", "cboBaseInfo", BCC_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血反应", "cboBaseInfo", BCC_输血反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输其他", "txtInfo", GC_输其他)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "医学警示", "txtInfo", GC_医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输红细胞", "txtSpecificInfo", SLC_输红细胞)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血小板", "txtSpecificInfo", SLC_输血小板)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血浆", "txtSpecificInfo", SLC_输血浆)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输全血", "txtSpecificInfo", SLC_输全血)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "自体回收", "txtSpecificInfo", SLC_自体回收)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血检查", "cboBaseInfo", BCC_输血前9项检查)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "其他医学警示", "txtInfo", GC_其他医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HBSAG", "cboBaseInfo", BCC_HBsAg)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HCV-AB", "cboBaseInfo", BCC_HCVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HIV-AB", "cboBaseInfo", BCC_HIVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "示教病案", "chkInfo", CHK_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "科研病案", "chkInfo", CHK_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "疑难病历", "chkInfo", CHK_疑难病例)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "发病时间", "mskDateInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病日期)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "生育状况", "cboBaseInfo", BCC_生育状况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院方式", "cboBaseInfo", BCC_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院转入", "txtInfo", GC_转入医疗机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "再入院计划天数", "cboBaseInfo", BCC_再入院计划天数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "31天内再住院", Null, Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院无)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院有)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_半保存, Null, 4, "txtInfo", GC_31天内再住院)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "昏迷时间", "txtSpecificInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_分钟)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_分钟)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "呼吸机使用时间", "txtSpecificInfo", SLC_呼吸机使用)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊标志", "chkInfo", CHK_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊期限", "txtSpecificInfo", SLC_随诊期限)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血反应", "cboBaseInfo", BCC_输血反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血检查", "cboBaseInfo", BCC_输血前9项检查)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "医学警示", "txtInfo", GC_医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "其他医学警示", "txtInfo", GC_其他医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输红细胞", "txtSpecificInfo", SLC_输红细胞)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血小板", "txtSpecificInfo", SLC_输血小板)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血浆", "txtSpecificInfo", SLC_输血浆)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输全血", "txtSpecificInfo", SLC_输全血)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "自体回收", "txtSpecificInfo", SLC_自体回收)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输白蛋白", "txtSpecificInfo", SLC_输白蛋白)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输其他", "txtInfo", GC_输其他)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "临床路径", "chkInfo", CHK_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "退出原因", "txtInfo", GC_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "变异原因", "txtInfo", GC_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 0)  'CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 1)  'PETCT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 2)  '双源CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 3)  'X片
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 4)  'B超
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 5)  '超声心动图
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 6)  'MRI
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 7)  '同位素检查
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "示教病案", "chkInfo", CHK_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "科研病案", "chkInfo", CHK_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院方式", "cboBaseInfo", BCC_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院转入", "txtInfo", GC_转入医疗机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊标志", "chkInfo", CHK_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊期限", "txtSpecificInfo", SLC_随诊期限)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "院内会诊", "txtSpecificInfo", SLC_院内会诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "外院会诊", "txtSpecificInfo", SLC_外院会诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "会诊情况", "txtInfo", GC_其他会诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "再入院计划天数", "cboBaseInfo", BCC_再入院计划天数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "31天内再住院", Null, Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院无)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院有)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_半保存, Null, 4, "txtInfo", GC_31天内再住院)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "昏迷时间", "txtSpecificInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_分钟)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_分钟)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "呼吸机使用时间", "txtSpecificInfo", SLC_呼吸机使用)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "发病时间", "mskDateInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病日期)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "生育状况", "cboBaseInfo", BCC_生育状况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤分期", "cboBaseInfo", BCC_肿瘤分期)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "门诊医师", "cboManInfo", MC_门诊医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "科主任", "cboManInfo", MC_科主任)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "主任医师", "cboManInfo", MC_主任或副主任)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "进修医师", "cboManInfo", MC_进修医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "主治医师", "cboManInfo", MC_主治医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "住院医师", "cboManInfo", MC_住院医师)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "主诊医师", "cboManInfo", MC_主诊医师)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "研究生实习医师", "cboManInfo", MC_研究生医师)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "实习医师", "cboManInfo", MC_实习医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "质控医师", "cboManInfo", MC_质控医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "质控护士", "cboManInfo", MC_质控护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "责任护士", "cboManInfo", MC_责任护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "病案质量", "cboBaseInfo", BCC_病案质量)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "主页质量日期", "mskDateInfo", DC_质控日期)
        '放疗与化疗
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_全保存, "病案化疗记录", "vsChemoth", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_全保存, "病案放疗记录", "vsRadioth", Null)
        '特定药品=====附页1
        .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人抗生素记录", "vsKSS", Null)

        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病案精神治疗", "vsSpirit", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "临床路径", "chkInfo", CHK_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "退出原因", "txtInfo", GC_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "变异原因", "txtInfo", GC_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_半保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_半保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_半保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_半保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "告病重病危", "chkInfo", CHK_住院出现危重)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病案重症监护情况", "vsFlxAddICU", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "器械导管使用情况", "vsICUInstruments", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人感染记录", "vsInfect", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人病原学检查", "vsSample", Null)
        End If
        '其他====附页2
        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案重症监护情况", "vsFlxAddICU", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "监护室名称", "txtInfo", GC_重症监护室名称)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "人工气道脱出", "chkInfo", CHK_人工气道脱出)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "重返重症医学科", "chkInfo", CHK_重返重症医学科)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "重返间隔时间", "cboBaseInfo", BCC_重返间隔时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "CT", "chkInfo", CHK_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "MRI", "chkInfo", CHK_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "彩色多普勒", "chkInfo", CHK_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 0)  '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 1)  '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 2)  '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "身体约束", "chkInfo", CHK_住院物理约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束总时间", "txtSpecificInfo", SLC_约束总时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束方式", "cboBaseInfo", BCC_约束方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束工具", "cboBaseInfo", BCC_约束工具)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束原因", "cboBaseInfo", BCC_约束原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "新生儿离院方式", "cboBaseInfo", BCC_新生儿离院方式)
        ElseIf Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "重症监护天数", "txtSpecificInfo", SLC_重症监护天)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "重症监护小时", "txtSpecificInfo", SLC_重症监护小时)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "单病种", "chkInfo", CHK_单病种管理)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "临床路径", "cboBaseInfo", BCC_临床路径管理) '是否修改为临床路径管理
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤分期", "cboBaseInfo", BCC_肿瘤分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤T", "txtInfo", GC_肿瘤T)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤M", "txtInfo", GC_肿瘤M)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤N", "txtInfo", GC_肿瘤N)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "传染病", "cboBaseInfo", BCC_法定传染病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "APGAR", "txtSpecificInfo", SLC_Apgar)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "标本送检", "chkInfo", CHK_细菌标本送检)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "DRGS", "cboBaseInfo", BCC_实施DGRS管理)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液反应", "cboBaseInfo", BCC_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液药物", "txtInfo", GC_引发药物)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液表现", "txtInfo", GC_临床表现)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "身体约束", "chkInfo", CHK_住院物理约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "透析尿素氮值", "txtInfo", GC_透析尿素氮值)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "距上一次住本院时间", "txtSpecificInfo", SLC_距上次住院时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "是否因同一疾病", "chkInfo", CHK_是否同一疾病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
        End If

        If Me.MedPageSandard = ST_卫生部标准 Or Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "CT", "chkInfo", CHK_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "MRI", "chkInfo", CHK_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "彩色多普勒", "chkInfo", CHK_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 0)  '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 1)  '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 2)  '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
        End If
    End With

    GetInMedRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDiagSelRecCache() As Boolean
'获取诊断选择缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long
    On Error GoTo errH

    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_加载扩展, IS_全保存, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_半保存, "分化程度", "cboBaseInfo", BCC_分化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_半保存, "最高诊断依据", "cboBaseInfo", BCC_最高诊断依据)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_初始扩展, IS_全保存, "诊断符合情况", "cboBaseInfo", Null)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_门诊与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_入院与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_门诊与入院)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_放射与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_临床与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "cboBaseInfo", BCC_临床与尸检)
             '中医诊断
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_门诊与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_入院与出院ZY)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_半保存, "死亡时间", "mskDateInfo", DC_死亡时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_半保存, "死亡根本原因", "txtInfo", GC_死亡原因)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_半保存, "尸检标志", "cboBaseInfo", BCC_死亡患者尸检)
    End With
    GetDiagSelRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOutMedRecCache() As Boolean
'获取门诊首页缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long

    On Error GoTo errH

    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        '基本信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "姓名", "txtInfo", GC_姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "性别", "cboBaseInfo", BCC_性别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "门诊号", "txtInfo", GC_门诊号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生日期", "mskDateInfo", DC_出生日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "年龄", "txtSpecificInfo", SLC_年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "医疗付款方式", "cboBaseInfo", BCC_付款方式)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "国籍", "cboBaseInfo", BCC_国籍)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "民族", "cboBaseInfo", BCC_民族)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "婚姻状况", "cboBaseInfo", BCC_婚姻)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "职业", "cboBaseInfo", BCC_职业)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号状态", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "外籍身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "区域", "txtAdressInfo", ADRC_病人区域)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "其他证件", "txtInfo", GC_其他证件)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "籍贯", "txtAdressInfo", ADRC_籍贯)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出生地点", "txtAdressInfo", ADRC_出生地点)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位地址", "txtAdressInfo", ADRC_单位地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位电话", "txtSpecificInfo", SLC_单位电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位邮编", "txtSpecificInfo", SLC_单位邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址", "txtAdressInfo", ADRC_现住址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭电话", "txtSpecificInfo", SLC_家庭电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址邮编", "txtSpecificInfo", SLC_家庭邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址", "txtAdressInfo", ADRC_户口地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址邮编", "txtSpecificInfo", SLC_户口邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "监护人", "txtInfo", GC_监护人)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "文化程度", "cboBaseInfo", BCC_文化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "生育状况", "cboBaseInfo", BCC_生育状况)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "血型", "cboBaseInfo", BCC_血型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "RH", "cboBaseInfo", BCC_RH)
        '就诊信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "无过敏记录", "chkInfo", CHK_无过敏记录)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "过敏药物", "vsAller", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "摘要", "txtInfo", GC_摘要)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_全保存, "复诊", Null, Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "optState", OP_初诊)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "optState", OP_复诊)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "传染病上传", "chkInfo", CHK_传染病上传)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "去向", "cboBaseInfo", BCC_去向)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "发病地址", "txtInfo", GC_发病地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_半保存, "发病时间", "mskDateInfo", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_半保存, Null, 1, "mskDateInfo", DC_发病日期)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_半保存, Null, 1, "mskDateInfo", DC_发病时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "医学警示", "txtInfo", GC_医学警示)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "其他医学警示", "txtInfo", GC_其他医学警示)
    End With

    GetOutMedRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPageMedRecCache() As Boolean
'获取病案首页缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long

    On Error GoTo errH

    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        '基本信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "病案号", "txtInfo", GC_病案号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "档案号", "txtInfo", GC_档案号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "医疗付款方式", "cboBaseInfo", BCC_付款方式)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "住院次数", "txtSpecificInfo", SLC_入院次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "再入院", "chkInfo", CHK_再入院)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "主页X线号", "txtInfo", GC_X线号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "住院号", "txtSpecificInfo", SLC_住院号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "姓名", "txtInfo", GC_姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "年龄", "txtSpecificInfo", SLC_年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "性别", "cboBaseInfo", BCC_性别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出生日期", "mskDateInfo", DC_出生日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "身高", "txtSpecificInfo", SLC_身高)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "体重", "txtSpecificInfo", SLC_体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "婚姻状况", "cboBaseInfo", BCC_婚姻)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "职业", "cboBaseInfo", BCC_职业)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "国籍", "cboBaseInfo", BCC_国籍)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "民族", "cboBaseInfo", BCC_民族)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭电话", "txtSpecificInfo", SLC_家庭电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出生地点", "txtAdressInfo", ADRC_出生地点)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "籍贯", "txtAdressInfo", ADRC_籍贯)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号状态", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "外籍身份证号", "cboBaseInfo", BCC_身份证)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "其他证件", "txtInfo", GC_其他证件)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "不足周岁年龄", "txtSpecificInfo", SLC_婴幼儿年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "新生儿出生体重", "txtSpecificInfo", SLC_新生儿出生体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "新生儿入院体重", "txtSpecificInfo", SLC_新生儿入院体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址", "txtAdressInfo", ADRC_现住址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "家庭地址邮编", "txtSpecificInfo", SLC_家庭邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位地址", "txtAdressInfo", ADRC_单位地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位电话", "txtSpecificInfo", SLC_单位电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "单位邮编", "txtSpecificInfo", SLC_单位邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址", "txtAdressInfo", ADRC_户口地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "户口地址邮编", "txtSpecificInfo", SLC_户口邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人姓名", "txtInfo", GC_联系人姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人关系", "cboBaseInfo", BCC_关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人附加信息", "txtInfo", GC_其他关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人电话", "txtSpecificInfo", SLC_联系人电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "联系人地址", "txtAdressInfo", ADRC_联系人地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "区域", "txtAdressInfo", ADRC_病人区域)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "健康号", "txtInfo", BCC_健康卡号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "监护人身份证号", "txtInfo", GC_监护人身份证号)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "Email", "txtInfo", GC_Email)
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "Qq", "txtSpecificInfo", SLC_QQ)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院日期", "mskDateInfo", DC_入院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院科室", "txtInfo", GC_入院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院病室", "txtInfo", GC_入院病房)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院方式", "cboBaseInfo", BCC_入院途径)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院转入", "txtInfo", GC_入院转入)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "入院前经外院治疗", "chkInfo", CHK_入院前外院治疗)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_初始扩展, IS_全保存, "转科记录", "vsTransfer", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 1)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 2)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 3)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 4)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 5)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 6)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_初始扩展, 0, "转科时间", "vsTransfer", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 1)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 2)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 3)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 4)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 5)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 0, "vsTransfer", 6)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出院日期", "mskDateInfo", DC_出院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出院科室", "txtInfo", GC_出院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "出院病室", "txtInfo", GC_出院病房)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "住院天数", "txtSpecificInfo", SLC_住院天数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_全保存, "医保号", "txtInfo", GC_医保号)
        '西医诊断（包含中医诊断诊断符合情况）
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_全保存, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "感染与死亡关系", "cboBaseInfo", BCC_感染与死亡关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "感染部位", "lstInfectParts", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "入院病况", "cboBaseInfo", BCC_入院情况)
        If Me.MedPageSandard = ST_四川省标准 Then
             .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "告病重病危", "chkInfo", CHK_住院出现危重)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "是否确诊", "chkInfo", CHK_是否确诊)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "确诊日期", "mskDateInfo", DC_确诊日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "病理号", "txtInfo", GC_病理号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "分化程度", "cboBaseInfo", BCC_分化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "最高诊断依据", "cboBaseInfo", BCC_最高诊断依据)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_全保存, "诊断符合情况", "cboBaseInfo", Null)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_门诊与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_入院与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_门诊与入院)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_术前与术后)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_放射与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_临床与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 1, "cboBaseInfo", BCC_临床与尸检)
             '中医诊断
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_门诊与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_入院与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_辩证)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_治法)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 2, "cboBaseInfo", BCC_方药)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "死亡根本原因", "txtInfo", GC_死亡原因)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "死亡时间", "mskDateInfo", DC_死亡时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "住院死亡期间", "cboBaseInfo", BCC_死亡期间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "病原学检查", "chkInfo", CHK_病原学检查)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "病原学诊断", "txtInfo", GC_病原学诊断)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "尸检标志", "cboBaseInfo", BCC_死亡患者尸检)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "新发肿瘤", "chkInfo", CHK_新发肿瘤)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "抢救次数", "txtSpecificInfo", SLC_抢救次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "成功次数", "txtSpecificInfo", SLC_成功次数)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_全保存, "疑难病历", "chkInfo", CHK_疑难病例)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "抢救病因", "txtInfo", GC_抢救病因)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_半保存, "分娩信息", "cmdDeliceryInfo", Null)
        '中医诊断
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_加载扩展, IS_全保存, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医危重", "chkInfo", CHK_危重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医急症", "chkInfo", CHK_急症)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医疑难", "chkInfo", CHK_疑难)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医治疗类别", "cboBaseInfo", BCC_治疗类别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医抢救方法", "cboBaseInfo", BCC_抢救方法)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "自制中药制剂", "cboBaseInfo", BCC_自制中药制剂)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医设备", "cboBaseInfo", BCC_中医诊疗设备)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "中医技术", "cboBaseInfo", BCC_中医诊疗技术)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_全保存, "辨证施护", "cboBaseInfo", BCC_辨证施护)
        '过敏与手术
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "无过敏记录", "chkInfo", CHK_无过敏记录)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_半保存, "过敏药物", "vsAller", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "门诊医师", "cboManInfo", MC_门诊医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "科主任", "cboManInfo", MC_科主任)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "主任医师", "cboManInfo", MC_主任或副主任)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "进修医师", "cboManInfo", MC_进修医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "主治医师", "cboManInfo", MC_主治医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "住院医师", "cboManInfo", MC_住院医师)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "主诊医师", "cboManInfo", MC_主诊医师)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "研究生实习医师", "cboManInfo", MC_研究生医师)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "实习医师", "cboManInfo", MC_实习医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "质控医师", "cboManInfo", MC_质控医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "质控护士", "cboManInfo", MC_质控护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "责任护士", "cboManInfo", MC_责任护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "病案质量", "cboBaseInfo", BCC_病案质量)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "主页质量日期", "mskDateInfo", DC_质控日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_全保存, "手麻情况", "vsOPS", Null)
        If Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "围术期死亡", "chkInfo", CHK_围术期死亡)
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "术后猝死", "chkInfo", CHK_术后猝死)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "特级护理天数", "txtSpecificInfo", SLC_特护)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "一级护理天数", "txtSpecificInfo", SLC_一级护理)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "二级护理天数", "txtSpecificInfo", SLC_二级护理)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "三级护理天数", "txtSpecificInfo", SLC_三级护理)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "ICU天数", "txtSpecificInfo", SLC_ICU)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_全保存, "CCU天数", "txtSpecificInfo", SLC_CCU)
        '住院情况
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "费用和", "txtSpecificInfo", SLC_费用和)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 2, ES_加载扩展, IS_半保存, "病人费用", "vsFees", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "病例分型", "cboBaseInfo", BCC_病例分型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "血型", "cboBaseInfo", BCC_血型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "RH", "cboBaseInfo", BCC_RH)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血反应", "cboBaseInfo", BCC_输血反应)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输其他", "txtInfo", GC_输其他)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输红细胞", "txtSpecificInfo", SLC_输红细胞)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血小板", "txtSpecificInfo", SLC_输血小板)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血浆", "txtSpecificInfo", SLC_输血浆)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输全血", "txtSpecificInfo", SLC_输全血)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "自体回收", "txtSpecificInfo", SLC_自体回收)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输血检查", "cboBaseInfo", BCC_输血前9项检查)
        
        If Me.MedPageSandard <> ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输液反应", "cboBaseInfo", BCC_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HBSAG", "cboBaseInfo", BCC_HBsAg)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HCV-AB", "cboBaseInfo", BCC_HCVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "HIV-AB", "cboBaseInfo", BCC_HIVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "示教病案", "chkInfo", CHK_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "科研病案", "chkInfo", CHK_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "疑难病历", "chkInfo", CHK_疑难病例)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "疑难病历", "chkInfo", CHK_疑难病例)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "发病时间", "mskDateInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病日期)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "生育状况", "cboBaseInfo", BCC_生育状况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院方式", "cboBaseInfo", BCC_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_半保存, "出院转入", "txtInfo", GC_转入医疗机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊标志", "chkInfo", CHK_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_半保存, "随诊期限", "txtSpecificInfo", SLC_随诊期限)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "发病时间", "mskDateInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病日期)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "mskDateInfo", DC_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "输白蛋白", "txtSpecificInfo", SLC_输白蛋白)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤分期", "cboBaseInfo", BCC_肿瘤分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "生育状况", "cboBaseInfo", BCC_生育状况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "临床路径", "chkInfo", CHK_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "退出原因", "txtInfo", GC_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "变异原因", "txtInfo", GC_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 0) 'CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 1) 'PETCT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 2) '双源CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 3) 'X片
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 4) 'B超
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 5) '超声心动图
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 6) 'MRI
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "vsTSJC", 7) '同位素检查
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "示教病案", "chkInfo", CHK_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "科研病案", "chkInfo", CHK_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院方式", "cboBaseInfo", BCC_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "出院转入", "txtInfo", GC_转入医疗机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "随诊标志", "chkInfo", CHK_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_半保存, "随诊期限", "txtSpecificInfo", SLC_随诊期限)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "院内会诊", "txtSpecificInfo", SLC_院内会诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "外院会诊", "txtSpecificInfo", SLC_外院会诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "会诊情况", "txtInfo", GC_其他会诊)
        End If
        
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "再入院计划天数", "cboBaseInfo", BCC_再入院计划天数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_半保存, "31天内再住院", Null, Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院无)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "optInput", OP_再住院有)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_半保存, Null, 4, "txtInfo", GC_31天内再住院)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "呼吸机使用时间", "txtSpecificInfo", SLC_呼吸机使用)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_全保存, "昏迷时间", "txtSpecificInfo", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_天)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_小时)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院前_分钟)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_天)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_小时)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 4, "txtSpecificInfo", SLC_昏迷时间入院后_分钟)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "编目员姓名", "cboManInfo", MC_编目员)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "编目日期", "mskDateInfo", DC_编目日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_全保存, "收回日期", "mskDateInfo", DC_收回日期)
        '放疗与化疗
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_全保存, "病案化疗记录", "vsChemoth", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_全保存, "病案放疗记录", "vsRadioth", Null)
        '特定药品
        .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人抗生素记录", "vsKSS", Null)
        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病案精神治疗", "vsSpirit", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "临床路径", "chkInfo", CHK_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "退出原因", "txtInfo", GC_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "变异原因", "txtInfo", GC_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "告病重病危", "chkInfo", CHK_住院出现危重)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "距上一次住本院时间", "cboBaseInfo", BCC_距上次住院时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_全保存, "是否因同一疾病", "chkInfo", CHK_是否同一疾病)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病案重症监护情况", "vsFlxAddICU", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "器械导管使用情况", "vsICUInstruments", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人感染记录", "vsInfect", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_全保存, "病人病原学检查", "vsSample", Null)
        End If
       '其他====附页2
        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案重症监护情况", "vsFlxAddICU", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "监护室名称", "txtInfo", GC_重症监护室名称)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "人工气道脱出", "chkInfo", CHK_人工气道脱出)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "重返重症医学科", "chkInfo", CHK_重返重症医学科)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "重返间隔时间", "cboBaseInfo", BCC_重返间隔时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "CT", "chkInfo", CHK_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "MRI", "chkInfo", CHK_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "彩色多普勒", "chkInfo", CHK_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 0) '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 1) '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 2) '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "身体约束", "chkInfo", CHK_住院物理约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束总时间", "txtSpecificInfo", SLC_约束总时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束方式", "cboBaseInfo", BCC_约束方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束工具", "cboBaseInfo", BCC_约束工具)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "约束原因", "cboBaseInfo", BCC_约束原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "新生儿离院方式", "cboBaseInfo", BCC_新生儿离院方式)
        ElseIf Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "重症监护天数", "txtSpecificInfo", SLC_重症监护天)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "重症监护小时", "txtSpecificInfo", SLC_重症监护小时)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "单病种", "chkInfo", CHK_单病种管理)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "临床路径", "cboBaseInfo", BCC_临床路径管理) '是否修改为临床路径管理
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤分期", "cboBaseInfo", BCC_肿瘤分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤T", "txtInfo", GC_肿瘤T)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤M", "txtInfo", GC_肿瘤M)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "肿瘤N", "txtInfo", GC_肿瘤N)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "传染病", "cboBaseInfo", BCC_法定传染病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "APGAR", "txtSpecificInfo", SLC_Apgar)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "标本送检", "chkInfo", CHK_细菌标本送检)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "DRGS", "cboBaseInfo", BCC_实施DGRS管理)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液反应", "cboBaseInfo", BCC_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液药物", "txtInfo", GC_引发药物)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "输液表现", "txtInfo", GC_临床表现)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "身体约束", "chkInfo", CHK_住院物理约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "透析尿素氮值", "txtInfo", GC_透析尿素氮值)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "距上一次住本院时间", "txtSpecificInfo", SLC_距上次住院时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "是否因同一疾病", "chkInfo", CHK_是否同一疾病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
        End If

        If Me.MedPageSandard = ST_卫生部标准 Or Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_全保存, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "CT", "chkInfo", CHK_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "MRI", "chkInfo", CHK_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "彩色多普勒", "chkInfo", CHK_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_全保存, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 0) '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 1) '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_全保存, Null, 7, "vsTSJC", 2) '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_全保存, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮发生期间", "cboBaseInfo", BCC_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "压疮分期", "cboBaseInfo", BCC_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床伤害", "cboBaseInfo", BCC_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_半保存, "跌倒或坠床原因", "cboBaseInfo", BCC_跌倒或坠床原因)
        End If
    End With

    GetPageMedRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetInArchiveRecCache() As Boolean
'获取住院首页缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long

    On Error GoTo errH

    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        '基本信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "医疗付款方式", "txtInfo", GCA_付款方式)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "健康号", "txtInfo", GCA_健康卡号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "主页ID", "txtInfo", GCA_住院次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "再入院", "chkInfo", CHKA_再入院)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "病案号", "txtInfo", GCA_病案号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "姓名", "txtInfo", GCA_姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "性别", "txtInfo", GCA_性别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生日期", "txtInfo", GCA_出生日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "年龄", "txtInfo", GCA_年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "住院号", "txtInfo", GCA_住院号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身高", "txtInfo", GCA_身高)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "体重", "txtInfo", GCA_体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "职业", "txtInfo", GCA_职业)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "婚姻状况", "txtInfo", GCA_婚姻)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "民族", "txtInfo", GCA_民族)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "不足周岁年龄", "txtInfo", GCA_不足周岁年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "新生儿出生体重", "txtInfo", GCA_新生儿体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "新生儿入院体重", "txtInfo", GCA_新生儿入院体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生地点", "txtInfo", GCA_出生地点)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "籍贯", "txtInfo", GCA_籍贯)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "国籍", "txtInfo", GCA_国籍)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号状态", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "外籍身份证号", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "其他证件", "txtInfo", GCA_其他证件)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭地址", "txtInfo", GCA_家庭地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭电话", "txtInfo", GCA_家庭电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭地址邮编", "txtInfo", GCA_家庭邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "户口地址", "txtInfo", GCA_户口地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "户口地址邮编", "txtInfo", GCA_户口邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位地址", "txtInfo", GCA_单位地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位电话", "txtInfo", GCA_单位电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位邮编", "txtInfo", GCA_单位邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "联系人姓名", "txtInfo", GCA_联系人姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "联系人关系", "txtInfo", GCA_联系人关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "联系人附加信息", "txtInfo", GCA_其他关系)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "联系人电话", "txtInfo", GCA_联系人电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "联系人地址", "txtInfo", GCA_联系人地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "区域", "txtInfo", GCA_区域)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "监护人身份证号", "txtInfo", GCA_监护人身份证号)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "Email", "txtInfo", GCA_Email)
            .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "Qq", "txtInfo", GCA_QQ)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院方式", "txtInfo", GCA_入院途径)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院转入", "txtInfo", GCA_入院转入)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院日期", "txtInfo", GCA_入院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院科室", "txtInfo", GCA_入院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院病室", "txtInfo", GCA_入院病室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "入院前经外院治疗", "chkInfo", CHKA_入院前经外院治疗)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_初始扩展, IS_只展示, "转科记录", "txtInfo", Null)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 0, "txtInfo", GCA_转科1)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 0, "txtInfo", GCA_转科2)
            Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 0, "txtInfo", GCA_转科3)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出院日期", "txtInfo", GCA_出院时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出院科室", "txtInfo", GCA_出院科室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出院病室", "txtInfo", GCA_出院病室)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "住院天数", "txtInfo", GCA_住院天数)
        '西医诊断（包含中医诊断诊断符合情况）
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_只展示, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "入院病况", "txtInfo", GCA_入院情况)
        If Me.MedPageSandard = ST_四川省标准 Then
             .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "告病重病危", "chkInfo", CHKA_住院期间告病重或病危)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "是否确诊", "chkInfo", CHKA_是否确诊)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "确诊日期", "txtInfo", GCA_确诊日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "病理号", "txtInfo", GCA_病理号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "分化程度", "txtInfo", GCA_分化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "最高诊断依据", "txtInfo", GCA_最高诊断依据)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_只展示, "诊断符合情况", "txtInfo", Null)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_门诊与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_入院与出院XY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_门诊与入院)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_术前与术后)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_放射与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_临床与病理)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 1, "txtInfo", GCA_临床与尸检)
             '中医诊断
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 2, "txtInfo", GCA_门诊与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 2, "txtInfo", GCA_入院与出院ZY)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 2, "txtInfo", GCA_辨证)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 2, "txtInfo", GCA_治法)
             Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 2, "txtInfo", GCA_方药)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "死亡时间", "txtInfo", GCA_死亡时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "死亡根本原因", "txtInfo", GCA_死亡原因)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "病原学检查", "chkInfo", CHKA_医院感染作病原学检查)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "尸检标志", "txtInfo", GCA_死亡患者尸检)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "新发肿瘤", "chkInfo", CHKA_新发肿瘤)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "病原学诊断", "txtInfo", GCA_医院感染病原学诊断)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "抢救次数", "txtInfo", GCA_抢救次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "成功次数", "txtInfo", GCA_成功次数)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "感染与死亡关系", "txtInfo", GCA_感染与死亡)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "感染部位", "txtInfo", GCA_感染部位)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "疑难病历", "chkInfo", CHKA_疑难病例)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "抢救病因", "txtInfo", GCA_抢救原因)
        '中医诊断
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_加载扩展, IS_只展示, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医危重", "chkInfo", CHKA_危重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医急症", "chkInfo", CHKA_急症)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医疑难", "chkInfo", CHKA_疑难)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医治疗类别", "txtInfo", GCA_治疗类别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医抢救方法", "txtInfo", GCA_抢救方法)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "自制中药制剂", "txtInfo", GCA_自制中药)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医设备", "txtInfo", GCA_中医设备)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "中医技术", "txtInfo", GCA_中医技术)
        .AddNew arrMainFileds, Array(Identity(lngCount), 2, 0, ES_不用扩展, IS_只展示, "辨证施护", "txtInfo", GCA_辨证施护)
        '过敏与手术
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_只展示, "过敏药物", "vsAller", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_加载扩展, IS_只展示, "手麻情况", "vsOPS", Null)
        If Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_只展示, "围术期死亡", "chkInfo", CHKA_围术期死亡)
            .AddNew arrMainFileds, Array(Identity(lngCount), 3, 0, ES_不用扩展, IS_只展示, "术后猝死", "chkInfo", CHKA_术后猝死)
        End If
        '住院情况
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "病例分型", "txtInfo", GCA_病例分型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "血型", "txtInfo", GCA_血型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "RH", "txtInfo", GCA_RH)
        If Me.MedPageSandard <> ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输液反应", "txtInfo", GCA_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血反应", "txtInfo", GCA_输血反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输其他", "txtInfo", GCA_输其他)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "医学警示", "txtInfo", GCA_医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输红细胞", "txtInfo", GCA_输红细胞)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血小板", "txtInfo", GCA_输血小板)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血浆", "txtInfo", GCA_输血浆)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输全血", "txtInfo", GCA_输全血)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "自体回收", "txtInfo", GCA_自体回收)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血检查", "txtInfo", GCA_输血前9项检查)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "其他医学警示", "txtInfo", GCA_其他医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "HBSAG", "txtInfo", GCA_HBsAg)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "HCV-AB", "txtInfo", GCA_HCVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "HIV-AB", "txtInfo", GCA_HIVAb)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "示教病案", "chkInfo", CHKA_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "科研病案", "chkInfo", CHKA_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "疑难病历", "chkInfo", CHKA_疑难病例)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "发病时间", "txtInfo", GCA_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "生育状况", "txtInfo", GCA_生育状况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "出院方式", "txtInfo", GCA_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "出院转入", "txtInfo", GCA_转入机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "再入院计划天数", "txtInfo", GCA_再入院天数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "31天内再住院", "txtInfo", GCA_31天目的)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_只展示, "昏迷时间", "txtInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前分钟)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后分钟)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "呼吸机使用时间", "txtInfo", GCA_呼吸机小时)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "随诊标志", "chkInfo", CHKA_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "随诊期限", "txtInfo", GCA_随诊期限)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血反应", "txtInfo", GCA_输血反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血检查", "txtInfo", GCA_输血前9项检查)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "医学警示", "txtInfo", GCA_医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "其他医学警示", "txtInfo", GCA_其他医学警示)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输红细胞", "txtInfo", GCA_输红细胞)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血小板", "txtInfo", GCA_输血小板)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输血浆", "txtInfo", GCA_输血浆)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输全血", "txtInfo", GCA_输全血)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "自体回收", "txtInfo", GCA_自体回收)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输白蛋白", "txtInfo", GCA_输白蛋白)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "输其他", "txtInfo", GCA_输其他)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "临床路径", "chkInfo", CHKA_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "退出原因", "txtInfo", GCA_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "变异原因", "txtInfo", GCA_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_只展示, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 0)   'CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 1)   'PETCT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 2)   '双源CT
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 3)   'X片
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 4)   'B超
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 5)   '超声心动图
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 6)   'MRI
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "vsTSJC", 7)   '同位素检查
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "示教病案", "chkInfo", CHKA_示教病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "科研病案", "chkInfo", CHKA_科研病案)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "出院方式", "txtInfo", GCA_出院方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "出院转入", "txtInfo", GCA_转入机构)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "肿瘤分期", "txtInfo", GCA_肿瘤分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "随诊标志", "chkInfo", CHKA_随诊)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "随诊期限", "txtInfo", GCA_随诊期限)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "院内会诊", "txtInfo", GCA_院内会诊次数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "外院会诊", "txtInfo", GCA_外院会诊次数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "会诊情况", "txtInfo", GCA_其他会诊情况)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "再入院计划天数", "txtInfo", GCA_再入院天数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "31天内再住院", "txtInfo", GCA_31天目的)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_初始扩展, IS_只展示, "昏迷时间", "txtInfo", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院前分钟)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后天)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后小时)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 4, "txtInfo", GCA_入院后分钟)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "呼吸机使用时间", "txtInfo", GCA_呼吸机小时)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "发病时间", "txtInfo", GCA_发病时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "生育状况", "txtInfo", GCA_生育状况)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "门诊医师", "txtInfo", GCA_门诊医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "科主任", "txtInfo", GCA_科主任)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "主任医师", "txtInfo", GCA_主任医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "进修医师", "txtInfo", GCA_进修医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "主治医师", "txtInfo", GCA_主治医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "住院医师", "txtInfo", GCA_住院医师)
        If Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "主诊医师", "txtInfo", GCA_主诊医师)
        Else
            .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "研究生实习医师", "txtInfo", GCA_研究生医师)
        End If
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "实习医师", "txtInfo", GCA_实习医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "质控医师", "txtInfo", GCA_质控医师)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "质控护士", "txtInfo", GCA_质控护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "责任护士", "txtInfo", GCA_责任护士)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "病案质量", "txtInfo", GCA_病案质量)
        .AddNew arrMainFileds, Array(Identity(lngCount), 4, 0, ES_不用扩展, IS_只展示, "主页质量日期", "txtInfo", GCA_质控日期)
        '放疗与化疗
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_只展示, "病案化疗记录", "vsChemoth", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 5, 0, ES_加载扩展, IS_只展示, "病案放疗记录", "vsRadioth", Null)
        '特定药品=====附页1
        .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "病人抗生素记录", "vsKSS", Null)

        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "病案精神治疗", "vsSpirit", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "临床路径", "chkInfo", CHKA_进入路径)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "退出原因", "txtInfo", GCA_退出原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "变异原因", "txtInfo", GCA_变异原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "压疮发生期间", "txtInfo", GCA_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "压疮分期", "txtInfo", GCA_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "跌倒或坠床伤害", "txtInfo", GCA_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "跌倒或坠床原因", "txtInfo", GCA_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_不用扩展, IS_只展示, "告病重病危", "chkInfo", CHKA_住院期间告病重或病危)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "病案重症监护情况", "vsFlxAddICU", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "器械导管使用情况", "vsICUInstruments", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "病人感染记录", "vsInfect", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 6, 0, ES_加载扩展, IS_只展示, "病人病原学检查", "vsSample", Null)
        End If
        '其他====附页2
        If Me.MedPageSandard = ST_卫生部标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_只展示, "病案重症监护情况", "vsFlxAddICU", Null)
        ElseIf Me.MedPageSandard = ST_云南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "监护室名称", "txtInfo", GCA_重症监护室)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "人工气道脱出", "chkInfo", CHKA_人工气道脱出)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "重返重症医学科", "chkInfo", CHKA_重返重症医学科)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "重返间隔时间", "txtInfo", GCA_重返间隔时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_只展示, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "CT", "chkInfo", CHKA_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "MRI", "chkInfo", CHKA_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "彩色多普勒", "chkInfo", CHKA_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_只展示, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 0)   '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 1)   '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 2)   '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "身体约束", "chkInfo", CHKA_住院期间身体约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "约束总时间", "txtInfo", GCA_约束总时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "约束方式", "txtInfo", GCA_约束方式)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "约束工具", "txtInfo", GCA_约束工具)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "约束原因", "txtInfo", GCA_约束原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "新生儿离院方式", "txtInfo", GCA_新生儿离院方式)
        ElseIf Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "重症监护天数", "txtInfo", GCA_重症监护天数)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "重症监护小时", "txtInfo", GCA_重症监护小时)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "单病种", "chkInfo", CHKA_单病种)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "临床路径", "txtInfo", GCA_临床路径管理)   '是否修改为临床路径管理
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "肿瘤分期", "txtInfo", GCA_肿瘤分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "肿瘤T", "txtInfo", GCA_肿瘤分期T)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "肿瘤M", "txtInfo", GCA_肿瘤分期M)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "肿瘤N", "txtInfo", GCA_肿瘤分期N)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "传染病", "txtInfo", GCA_传染病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "APGAR", "txtInfo", GCA_Apgar)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "标本送检", "chkInfo", CHKA_细菌培养标本送检)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "DRGS", "txtInfo", GCA_DrGs管理)
        ElseIf Me.MedPageSandard = ST_四川省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "压疮发生期间", "txtInfo", GCA_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "压疮分期", "txtInfo", GCA_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "跌倒或坠床伤害", "txtInfo", GCA_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "跌倒或坠床原因", "txtInfo", GCA_跌倒或坠床原因)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "输液反应", "txtInfo", GCA_输液反应)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "输液药物", "txtInfo", GCA_引发药物)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "输液表现", "txtInfo", GCA_临床表现)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "身体约束", "chkInfo", CHKA_住院期间身体约束)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "透析尿素氮值", "txtInfo", GCA_离院透析尿素氮值)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "距上一次住本院时间", "txtInfo", GCA_距上一次住本院时间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "是否因同一疾病", "chkInfo", CHKA_是否因同一疾病)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_只展示, "病案项目", "vsfMain", Null)
        End If

        If Me.MedPageSandard = ST_卫生部标准 Or Me.MedPageSandard = ST_湖南省标准 Then
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_加载扩展, IS_只展示, "病案项目", "vsfMain", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "CT", "chkInfo", CHKA_CT)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "MRI", "chkInfo", CHKA_MRI)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "彩色多普勒", "chkInfo", CHKA_彩色多普勒)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_初始扩展, IS_只展示, "特殊检查", "vsTSJC", Null)
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 0)   '特殊检查4
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 1)   '特殊检查5
                Me.SecdInfoRec.AddNew arrSecdFileds, Array(Identity(lngSort), lngCount, CS_未改变, IS_只展示, Null, 7, "vsTSJC", 2)   '特殊检查6
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "感染因素", "lstInfection", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "不良事件", "lstAdvEvent", Null)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "压疮发生期间", "txtInfo", GCA_压疮发生期间)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "压疮分期", "txtInfo", GCA_压疮分期)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "跌倒或坠床伤害", "txtInfo", GCA_跌倒或坠床伤害)
            .AddNew arrMainFileds, Array(Identity(lngCount), 7, 0, ES_不用扩展, IS_只展示, "跌倒或坠床原因", "txtInfo", GCA_跌倒或坠床原因)
        End If
    End With

    GetInArchiveRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOutArchiveRecCache() As Boolean
'获取门诊首页缓存记录集
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant
    Dim lngCount As Long, lngSort As Long

    On Error GoTo errH

    With Me.MainInfoRec
        arrMainFileds = Array("序号", "页码", "是否改变", "ExpState", "Edit", "信息名", "控件名", "Index")
        arrSecdFileds = Array("Sort", "序号", "改变状态", "Edit", "ID", "页码", "控件名", "IndexEx")
        lngCount = 0: lngSort = 0
        '基本信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "姓名", "txtInfo", GCA_姓名)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "性别", "txtInfo", GCA_性别)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "门诊号", "txtInfo", GCA_门诊号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生日期", "txtInfo", GCA_出生日期)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "年龄", "txtInfo", GCA_年龄)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "医疗付款方式", "txtInfo", GCA_付款方式)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "国籍", "txtInfo", GCA_国籍)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "民族", "txtInfo", GCA_民族)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "婚姻状况", "txtInfo", GCA_婚姻)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "职业", "txtInfo", GCA_职业)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "身份证号状态", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "外籍身份证号", "txtInfo", GCA_身份证号)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "区域", "txtInfo", GCA_区域)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "其他证件", "txtInfo", GCA_其他证件)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "籍贯", "txtInfo", GCA_籍贯)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "出生地点", "txtInfo", GCA_出生地点)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位地址", "txtInfo", GCA_单位地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位电话", "txtInfo", GCA_单位电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "单位邮编", "txtInfo", GCA_单位邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭地址", "txtInfo", GCA_家庭地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭电话", "txtInfo", GCA_家庭电话)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "家庭地址邮编", "txtInfo", GCA_家庭邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "户口地址", "txtInfo", GCA_户口地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "户口地址邮编", "txtInfo", GCA_户口邮编)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "监护人", "txtInfo", GCA_监护人)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "文化程度", "txtInfo", GCA_文化程度)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "生育状况", "txtInfo", GCA_生育状况)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "血型", "txtInfo", GCA_血型)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "RH", "txtInfo", GCA_RH)
        .AddNew arrMainFileds, Array(Identity(lngCount), 0, 0, ES_不用扩展, IS_只展示, "监护人身份证号", "txtInfo", GCA_监护人身份证号)
        '就诊信息页
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "无过敏记录", "chkInfo", CHKA_无过敏记录)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_只展示, "过敏药物", "vsAller", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "摘要", "txtInfo", GCA_就诊摘要)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_只展示, "西医诊断", "vsDiagXY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_加载扩展, IS_只展示, "中医诊断", "vsDiagZY", Null)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "复诊", "chkInfo", CHKA_传染病上传)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "传染病上传", "chkInfo", CHKA_传染病上传)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "去向", "txtInfo", GCA_去向)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "发病地址", "txtInfo", GCA_发病地址)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_只展示, "发病时间", "txtInfo", GCA_发病时间)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "身高", "txtInfo", GCA_身高)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_只展示, "体重", "txtInfo", GCA_体重)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "体温", "txtInfo", GCA_体温)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "脉搏", "txtInfo", GCA_脉搏)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_初始扩展, IS_只展示, "呼吸", "txtInfo", GCA_呼吸)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "血压", "txtInfo", GCA_血压)  '收缩压/舒张压 血压单位
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "医学警示", "txtInfo", GCA_医学警示)
        .AddNew arrMainFileds, Array(Identity(lngCount), 1, 0, ES_不用扩展, IS_只展示, "其他医学警示", "txtInfo", GCA_其他医学警示)
    End With
    GetOutArchiveRecCache = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Property Get AddAller() As Boolean
    AddAller = mblnAddAller
End Property
Public Property Let AddAller(ByVal blnAddAller As Boolean)
    mblnAddAller = blnAddAller
End Property

Public Property Get AddDiag() As Boolean
    AddDiag = mblnAddDiag
End Property
Public Property Let AddDiag(ByVal blnAddDiag As Boolean)
    mblnAddDiag = blnAddDiag
End Property

Public Property Get CboMask() As Boolean
    CboMask = mblnCboMask
End Property

Public Property Let CboMask(ByVal blnCboMask As Boolean)
    mblnCboMask = blnCboMask
End Property
