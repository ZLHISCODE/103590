Attribute VB_Name = "mdlCISKernel"
Option Explicit
Public gfrmMain As Object                   '导航台窗体
Public gclsInsure As New clsInsure          '医保变量
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gobjCISBase As Object                '诊疗基础部件
Public gobjInExse As Object                 '住院费用部件
Public gobjPath As Object                   '临床路径部件(通过医生站或护士站传入)
Public gobjPathOut As Object                '门诊临床路径部件(通过医生站或护士站传入)
Public gobjLIS As Object                    'LIS申请部件
Public gobjExchange As Object               'HL7数据交换部件
Public gobjSquareCard As Object             '一卡通交易部件(由于需调初始化接口，由门诊医生站或医技站传入，并且由他们销毁)
Public gobjEmrInterface As Object           '新版病历申请附项读取部件
Public gobjRecipeAudit As Object            '处方审查系统对象
Public gobjPublicExpense As Object          '费用公共部件
Public gobjPublicDrug As Object             '药品公共部件
Public gobjPublicBlood As Object            '血库公共部件
Public gobjPublicPatient As Object          '病人信息公共部件

Public gstrTsPrivsMZ As String              '门诊医生特殊医嘱权限字符串
Public gstrTsPrivsZY As String              '住院医生特殊医嘱权限字符串

Public gcolPrivs As Collection              '记录内部模块的权限
Public gMainPrivs As String                 '调用主界面所具有的权限,注意非内部模块权限
Public gstrSysName As String                '系统名称
Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String            'OEM产品名称
Public glngSys As Long

Public gstrSQL As String
Public gbln加班加价 As Boolean              '状态判断临时变量
Public grsDuty As ADODB.Recordset           '存放医生职务
Public gstr动态费别 As String               '存放门诊当前科室可用动态费别,在公共函数中使用,使用时才赋值:CalcDrugPrice,CalcPrice
Public gblnKSSStrict As Boolean             '是否启用抗菌药物严格控制
Public gbln抗菌药物使用自备药 As Boolean

Public grsSkinTest As ADODB.Recordset       '存放皮试项目的阴阳性定义
Public grsTube As ADODB.Recordset           '存放试管相关信息
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码输入方式
Public grs医疗付款方式 As ADODB.Recordset
Public gbyt病人审核方式 As Byte '49501:病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
Public gbln未入科禁止记账  As Boolean

Public gbln手术分级管理 As Boolean  '是否启用手术分级管理
Public gbln手术授权管理 As Boolean  '是否启用手术按医师授权管理
Public gbln手术等级管理 As Boolean  '是否启用参数：主刀医师达到手术等级无需审核
Public gbln输血分级管理 As Boolean  '是否启用输血分级管理
Public gbln输血申请中级以上 As Boolean  '输血申请只能由中级及以上医师提出
Public gbln输血申请三级审核 As Boolean  '输血申请三级审核
Public gint医嘱执行有效天数 As Integer '允许修改n天内登记的医嘱执行记录
Public gbyt转科时未审核销帐单据检查 As Byte  '转科时是否检查病人存在未审核的销帐单据:0-不检查,1-检查并提示,2-检查并禁止
Public gbln血库系统 As Boolean  '是否安装血库系统
Public gbln显示血液库存 As Boolean '输血申请是否显示血液库存
Public gbln下达用血申请确定血液信息 As Boolean '申请单的方式下达用血申请，是否要勾选具体的血液才能申请
Public gbyt超量原因 As Byte '医嘱超量时必须输入原因 0-不用填写原因，1－必须填写
Public gstr不录超量科室 As String
Public gbln特殊药品分开发送 As Boolean '特殊药品分开发送
Public gbln医嘱终止原因 As Boolean '系统参数：停嘱时录入原因
Public gstr可不填停嘱原因科室 As String '系统参数：可不填停嘱原因科室
Public gobjDrugExplain As Object '药品说明书部件
'--------------------
'申请单启用环节 参数，启用申请单后必须使用申请单下达医嘱 参数
Public gstrInUseApp As String '住院
Public gstrOutUseApp As String '门诊
Public gblnIn必用 As Boolean '住院
Public gblnOut必用 As Boolean '门诊
'-------------------
Public gbln反算天数 As Boolean '药品录入时根据总量与单量反算用药天数
Public gbln启用影像信息系统接口 As Boolean
Public gbln启用影像信息系统预约 As Boolean
Public gbln科室药房对照按本机参数设置 As Boolean
Public gbln会诊科室下达医嘱处理 As Boolean '系统参数：会诊科室下达医嘱由会诊申请科室处理
Public gbln审方系统 As Boolean '中联合理用药审方功能启用开关，通过读取服务配置来判断  三方服务配置目录

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p住院记帐操作 = 1150
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p临床路径应用 = 1256
    P门诊路径应用 = 1248
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    p输血审核管理 = 1268
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    pXWPACS观片 = 1288
    p观片工具管理 = 1289
    p输液配置中心 = 1345
    p新版门诊病历 = 2251
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    用药级别 As Long
    手术等级 As String
End Type
Public UserInfo As TYPE_USER_INFO

'电子签名
Public gintCA As Integer '电子签名认证中心
Public gstrESign As String '电子签名控制场合
Public gobjESign As Object '电子签名接口部件
Public grsSign As Recordset  '电子签名启用部门（缓存）

'外挂功能
Public gobjPlugIn As Object

'RIS接口部件
Public gobjRis As Object

'CIS系统参数
Public gbln药品按规格下医嘱 As Boolean
Public gint过敏登记有效天数 As Integer
Public gbln长期医嘱次日生效 As Boolean
Public gstr住院发送划价单 As String
Public gstr门诊发送划价单 As String
Public gstr输液配置中心 As String
Public glng补录时限 As Long
Public gbln只允许补录临嘱 As Boolean
Public gblnShowOrigin As Boolean '是否显示产地

Public gbln执行前先结算 As Boolean '一卡通执行前先收费或记帐审核

Public gintRXCount As Integer
Public gint补录间隔 As Integer '自动识别为补录医嘱的时间间隔(分钟)
Public gbln发送生成条形码 As Boolean '是否在检验医嘱发送时生成条形码
Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
                                                      '为负数(-N)时表示,N元内免密支付,表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码
Public gbln指定医嘱在其他科室执行 As Boolean '指定医嘱在其他科室执行
Public gstr医嘱核对 As String    '输血皮试医嘱需要核对 按位存取11，第一位为 输血医嘱，第二位为 皮试医嘱
Public gint门诊新开医嘱间隔  As Integer '门诊新开医嘱间隔
Public gbln开单后立即收费或记帐审核 As Boolean '项目开单后立即收费或记帐审核

'HIS系统参数
Public gbytMediOutMode As Byte '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按金额小数位数计算的格式化串,如"0.0000"
Public gbytDecPrice As Byte '费用单价的小数点位数
Public gstrDecPrice As String '价格按小数位数计算的格式化串,如"0.0000"

Public gint普通挂号天数 As Integer '普通挂号单有效天数
Public gint急诊挂号天数 As Integer '急诊挂号单有效天数
Public gint诊疗编码 As Integer '0-顺序编号,1-种类+分类号+顺序编号
Public gbyt药品名称显示 As Byte '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
Public gbyt输入药品显示 As Byte '0-按输入匹配显示，1-固定显示通用名和商品名
Public gbln简码匹配方式切换 As Boolean '允许在窗口界面的工具栏切换简码匹配方式切换

Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gint诊断来源 As Integer '1-由医生选择输入来源,2-按照诊断标准输入,3-按照疾病编码输入
Public gstr诊断输入 As String '1门诊/2住院：1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Public gstrMatchMode As String '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
Public gbyt出院检查未执行 As Byte '出院时是否检查有未执行项目:0-不检查,1-检查并提示,2-检查并禁止
Public gbyt转科检查未执行 As Byte '转科时是否检查有未执行项目:0-不检查,1-检查并提示,2-检查并禁止
Public gbyt出院检查未发药 As Byte '出院时是否检查有未发药品:0-不检查,1-检查并提示,2-检查并禁止
Public gbyt转科检查未发药 As Byte '转科时是否检查有未发药品:0-不检查,1-检查并提示,2-检查并禁止

Public gint医保对码 As Integer '是否对住院医保病人的项目对码情况进行检查:0-不检查,1-检查并提醒,2-检查并禁止
Public gcurMaxMoney As Currency '单笔费用最大提醒金额

'医技工作站系统费用参数
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gbln收费类别 As Boolean '是否首先输入类别
Public gblnFeeKindCode As Boolean '不输类别时,首位当作收费类别简码

Public gbyt住院自动发料 As Byte  '住院记帐完成后是否自动发料 0-不自动发料，1-自动发料，2-本科室开单时自动发料
Public gbln门诊自动发料 As Boolean '门诊记帐完成后是否自动发料
Public gbln收费后自动发药 As Boolean '

Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用
Public gstr医保费用类型 As String '医保病人允许的费用类型
Public gstr公费费用类型 As String '公费病人允许的费用类型
Private mlng部门编码平均长度 As Long


Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤消出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29 '出院接口中是否要与接口商进行交易
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    
    support医生确定处方类型 = 48
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support实时监控 = 60
    
    support上传门诊档案 = 70                    '在门诊医嘱发送时，是否调用TranElecDossier函数完成门诊病人电子卷宗/电子档案的上传
    support上传住院档案 = 70                    '在住院医嘱发送时，是否调用TranElecDossier 特殊说明：门诊住院使用同一个业务序号
End Enum

'Pass
Public gobjPass As Object  'PASS 接口

'接口类型枚举
Public Enum G_PASS_TYPE
    UNPASS = 0          '未启用
    MK = 1              '美康
    DT = 2              '大通
    TYT = 3             '太元通
    YWS = 4             '保进药卫士
    HZYY = 5            '杭州逸曜
End Enum
'调用模块编号
Public Enum PASS_MODEL
    PM_门诊编辑 = 0
    PM_住院编辑 = 1
    PM_住院医嘱清单 = 2
    PM_护士校对 = 3
    PM_门诊医嘱清单 = 4
End Enum


Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = NVL(rsTmp!专业技术职务)
            UserInfo.手术等级 = NVL(rsTmp!手术等级)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.用户名
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    If str姓名 <> "" Then
        strSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str姓名)
    Else
        strSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditName(ByVal strName As String) As String
'功能：从"审核医生/实习医生"中取审核医生名
    GetAuditName = Mid(strName, 1, IIF(InStr(strName, "/") > 0, InStr(strName, "/") - 1, Len(strName)))
End Function

Public Function HaveAuditPriv(Optional ByVal str姓名 As String) As Boolean
'功能：判断当前人员是否具有"执业医师"的资格
'参数：str姓名=判断指定人员，不传入时判断当前人员
    Dim rsTmp As New ADODB.Recordset
    Dim strHave As String, strNone As String
    Dim strSQL As String
    
    If str姓名 = "" Then str姓名 = UserInfo.姓名
    
    If InStr(strHave & "|", "|" & str姓名 & "|") > 0 Then
        HaveAuditPriv = True: Exit Function
    ElseIf InStr(strNone & "|", "|" & str姓名 & "|") > 0 Then
        HaveAuditPriv = False: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select B.分类 From 人员表 A,执业类别 B Where A.姓名=[1] And A.执业类别=B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "HaveAuditPriv", str姓名)
    If Not rsTmp.EOF Then
        If rsTmp!分类 = "执业医师" Or rsTmp!分类 = "执业助理医师" Then HaveAuditPriv = True
    End If
    If HaveAuditPriv Then
        strHave = strHave & "|" & str姓名
    Else
        strNone = strNone & "|" & str姓名
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get部门性质(ByVal lng部门ID As Long) As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 工作性质 From 部门性质说明 Where 部门ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get部门性质", lng部门ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!工作性质
        rsTmp.MoveNext
    Loop
    Get部门性质 = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean, Optional ByVal lngSys As Long) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    If lngSys = 0 Then lngSys = glngSys
    On Error Resume Next
    strPrivs = gcolPrivs(lngSys & "_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove lngSys & "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, lngSys & "_" & lngProg
    End If
    GetInsidePrivs = IIF(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function GetPatiUnitID(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
'功能：根据病人获取对应的病区ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 当前病区ID as 病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    GetPatiUnitID = NVL(rsTmp!病区ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiRsByUnit(ByVal lngUnitID As Long, ByVal lng病人ID As Long, _
    ByVal bln适用病人 As Boolean, ByVal bln剩余款 As Boolean, ByVal bln包含出院病人 As Boolean, _
    Optional ByVal blnIsPreOut As Boolean, Optional ByVal lng医嘱处理范围 As Long) As ADODB.Recordset
'功能：获取当前病区的在院病人列表，以及当前病人自动加入列表（转出，预出院，出院）
'参数：bln包含出院病人=护士站超期收回调用时，显示最近出院的病人
'     lng医嘱处理范围=-1所有病人包括婴儿，0病人，1婴儿
    Dim strSQL As String, intBedLen As Integer, strIF As String
    Dim curDate As Date, intDay As Integer, dtOutEnd As Date, dtOutBegin As Date
    Dim strPreOut As String
    Dim str剩余款 As String
    Dim str适用病人 As String
    Dim str范围所有 As String '-1
    Dim str范围婴儿 As String '1
    Dim strOther As String '1 or -1
    
    On Error GoTo errH
    intBedLen = GetMaxBedLen(lngUnitID, False)
    strPreOut = " And Nvl(b.状态,0)<>3 "
    'Union是为了利用索引(自动去掉重复的)，不用In。用0+[2]，是为了避免错误:表达式必须具有与对应表达式相同的数据类型
    strIF = "Select a.病人ID" & vbNewLine & _
            " From 病人信息 A, 病案主页 B,在院病人 R" & vbNewLine & _
            " Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID And (R.病区id = [1] or B.婴儿病区ID = [1]) " & IIF(blnIsPreOut, "", strPreOut) & vbNewLine & _
            " Union" & vbNewLine & _
            " Select 0+[2] as 病人ID From Dual"

    
    If bln包含出院病人 Then
         '出院病人时间范围
        curDate = zlDatabase.Currentdate
        intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, p住院护士站, 0))
        dtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, p住院护士站, 1))
        dtOutBegin = Format(curDate - intDay, "yyyy-MM-dd 00:00:00")
    
        strIF = strIF & " Union" & vbNewLine & _
                " Select a.病人id" & vbNewLine & _
                " From 病人信息 A, 病案主页 B" & vbNewLine & _
                " Where a.病人id = b.病人id And a.主页id = b.主页id And (b.当前病区id + 0 = [1] Or b.婴儿病区id + 0 = [1]) And B.出院日期 Between [3] And [4]"
    End If
    
    If bln剩余款 Then
        str剩余款 = "Nvl(E.预交余额,0)-Nvl(E.费用余额,0)+Decode(B.险类,Null,0,(Select Nvl(Sum(金额),0) From 保险模拟结算 F" & _
            " Where B.病人ID=F.病人ID And B.主页ID=F.主页ID))"
    Else
        str剩余款 = "Null"
    End If
    
    str适用病人 = IIF(bln适用病人, "zl_PatiWarnScheme(A.病人ID,B.主页ID)", "Null")
    
    If lng医嘱处理范围 = 1 Then str范围婴儿 = " And exists(select 1 from 病人新生儿记录 Z Where z.病人id=b.病人ID And z.主页id=b.主页ID)"
    
    If lng医嘱处理范围 = -1 Then
        str范围所有 = " Select 病人id, 主页id, 姓名, 住院号, 床号, 担保额, 剩余款,  适用病人, 险类, 住院医师, 费别, 护理等级, 科室,科室id," & _
            " 入院日期, 出院日期, 病人类型, 性别, 审核标志,婴儿科室ID,婴儿病区ID,Null as 婴儿姓名,Null as 婴儿序号,病人状态,留观号 From Pati Union All "
    End If
    
    If lng医嘱处理范围 = 1 Or lng医嘱处理范围 = -1 Then
        strOther = "select a.病人id,a.主页id,a.姓名,a.住院号,a.床号,a.担保额,a.剩余款, a.适用病人,a.险类,a.住院医师,a.费别,a.护理等级,a.科室,a.科室id," & _
            " a.入院日期, a.出院日期, a.病人类型, a.性别, a.审核标志,a.婴儿科室ID,a.婴儿病区ID,b.婴儿姓名,B.序号 AS 婴儿序号,病人状态,a.留观号 from Pati A,病人新生儿记录 B" & _
            " where A.病人id=b.病人id and A.主页ID=b.主页id "
    Else
        strOther = "Select 病人id, 主页id, 姓名, 住院号, 床号, 担保额, 剩余款,  适用病人, 险类, 住院医师, 费别, 护理等级, 科室, 科室id," & _
            " 入院日期, 出院日期, 病人类型, 性别, 审核标志,婴儿科室ID,婴儿病区ID,Null as 婴儿姓名,Null as 婴儿序号,病人状态,留观号 From Pati "
    End If
    
    strSQL = "Select * from (" & _
        " With Pati as (Select A.病人ID,B.主页ID,A.姓名,B.住院号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,A.担保额," & _
        str剩余款 & " as 剩余款," & str适用病人 & " as 适用病人,B.险类,B.住院医师,B.费别,D.名称 as 护理等级,C.名称 as 科室," & _
        " c.id as 科室id,B.入院日期,B.出院日期,B.病人类型,A.性别,b.审核标志,B.婴儿科室ID,B.婴儿病区ID,Decode(B.状态,3,1,Decode(B.出院日期,Null,0,2)) as 病人状态,B.留观号" & _
        " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 D,病人余额 E,(" & strIF & ") F" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.护理等级ID=D.ID(+)" & _
        " And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
        " And A.病人ID=E.病人ID(+) And E.性质(+)=1 And E.类型(+) = 2 And A.病人ID=F.病人ID" & _
        str范围婴儿 & " Order by 床号) " & _
        str范围所有 & _
        strOther & _
        ") order by 床号,NVL(婴儿序号,0)"
        
    Set GetPatiRsByUnit = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngUnitID, lng病人ID, dtOutBegin, dtOutEnd)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiTurnLimit(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal lng科室id As Long, _
    ByRef datTurn As Date, ByVal intPState As TYPE_PATI_State) As Boolean
'功能：检查出院、预出院、转科、转病区的病人的补录时限，如果允许，则返回该操作的时间
    Dim strSQL As String, strMsg As String
    Dim rsTmp As ADODB.Recordset
    
    If intPState = ps出院 Then
        strSQL = "Select 出院日期 as 终止时间 From 病案主页 Where 病人id = [1] and 主页ID=[2]"
        strMsg = "出院"
    Else
        If intPState = ps最近转出 Then
            strSQL = " And (终止原因 =3 and 科室ID=[4] or 终止原因 =15 and 病区ID=[3])"
            strMsg = "转科或转病区"
        ElseIf intPState = ps预出 Then
            strSQL = " And 终止原因 = 10"
            strMsg = "办理预出院"
        End If
        strSQL = "Select Max(终止时间) as 终止时间 From 病人变动记录" & _
                " Where 病人id = [1] and 主页ID=[2]   And 终止时间 is Not Null" & strSQL
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "补录的时限检查", lng病人ID, lng主页ID, lng病区ID, lng科室id)
    If IsNull(rsTmp!终止时间) Then
        MsgBox "因数据异常，没有找到该病人的变动信息，无法进行医嘱补录操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    datTurn = CDate(Format(rsTmp!终止时间, "yyyy-mm-dd HH:MM:SS"))
    If glng补录时限 = 0 Then
        ShowMsgBox "注意:" & vbCrLf & "    该病人已" & strMsg & ",系统设置为不允许进行医嘱补录操作。"
        Exit Function
    Else
        If datTurn + 1 / 24 * glng补录时限 < zlDatabase.Currentdate Then
            ShowMsgBox "注意:" & vbCrLf & "    该病人" & strMsg & "已经超过了" & glng补录时限 & "小时,不允许进行医嘱补录操作。"
            Exit Function
        Else
            '检查科室是否是操作员的可用科室
            If Get开嘱科室ID(UserInfo.ID, 0, lng科室id) <> lng科室id Then
                ShowMsgBox "注意:您不是当前科室的医生,不允许进行医嘱补录操作。"
                Exit Function
            End If
        End If
    End If
        
    CheckPatiTurnLimit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'功能：初始化系统参数
'返回：真-处理成功
    Dim strTmp As String
    gstrLike = IIF(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gbyt药品名称显示 = zlDatabase.GetPara("药品名称显示", , , 2)
    gbyt输入药品显示 = zlDatabase.GetPara("输入药品显示", , , 0)
    gbln简码匹配方式切换 = zlDatabase.GetPara("简码匹配方式切换", , , 1)

    '补录医嘱识别时间间隔(分钟)
    gint补录间隔 = Val(zlDatabase.GetPara(5, glngSys, , 30))
    
    '门诊新开医嘱间隔
    gint门诊新开医嘱间隔 = Val(zlDatabase.GetPara(223, glngSys, , 1))
    
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    gbytDecPrice = Val(zlDatabase.GetPara(157, glngSys, , 5))
    gstrDecPrice = "0." & String(gbytDecPrice, "0")
        
    '指定药房时限制库存
    gblnStock = Val(zlDatabase.GetPara(18, glngSys)) <> 0
        
    '挂号有效天数
    strTmp = zlDatabase.GetPara(21, glngSys)
    If Len(strTmp) = 1 Then strTmp = strTmp & strTmp
    gint普通挂号天数 = Val(Mid(strTmp, 1, 1))
    gint急诊挂号天数 = Val(Mid(strTmp, 2, 1))
    
    '检查未执行项目
    gbyt出院检查未执行 = Val(zlDatabase.GetPara(22, glngSys))
    gbyt转科检查未执行 = Val(zlDatabase.GetPara(32, glngSys))
    
    gbyt出院检查未发药 = Val(zlDatabase.GetPara(154, glngSys))
    gbyt转科检查未发药 = Val(zlDatabase.GetPara(155, glngSys))
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    '电子签名认证中心
    gintCA = Val(zlDatabase.GetPara(25, glngSys))
    
    '电子签名控制场合
    gstrESign = zlDatabase.GetPara(26, glngSys)
    
    '读取部门启用数据
    Set grsSign = New ADODB.Recordset
    grsSign.Fields.Append "部门ID", adBigInt
    grsSign.Fields.Append "场合", adBigInt
    grsSign.Fields.Append "是否启用", adBigInt
    grsSign.CursorLocation = adUseClient
    grsSign.LockType = adLockOptimistic
    grsSign.CursorType = adOpenStatic
    grsSign.Open
    
    '一卡通消费验证
    strTmp = zlDatabase.GetPara(28, glngSys) & "|"
    gdbl预存款消费验卡 = Val(Split(strTmp, "|")(0))
   
    '指定医嘱在其他科室执行
    gbln指定医嘱在其他科室执行 = Val(zlDatabase.GetPara(34, glngSys)) <> 0
    
    '医保费用类型
    gstr医保费用类型 = "'" & Replace(zlDatabase.GetPara(41, glngSys), "|", "','") & "'"

    '公费费用类型
    gstr公费费用类型 = "'" & Replace(zlDatabase.GetPara(42, glngSys), "|", "','") & "'"
    
    '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    '诊断输入来源
    gint诊断来源 = Val(zlDatabase.GetPara(55, glngSys, , 1))
    
    '门诊处方条数限制
    gintRXCount = Val(zlDatabase.GetPara(56, glngSys))
    
    '医保对码检查
    gint医保对码 = Val(zlDatabase.GetPara(59, glngSys, , 1))
    
    '单笔费用最大提醒金额
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    
    '诊疗编码递增模式
    gint诊疗编码 = Val(zlDatabase.GetPara(61, glngSys))
    
    '住院自动发料
    gbyt住院自动发料 = Val(zlDatabase.GetPara(63, glngSys))
    
    '诊断输入方式
    gstr诊断输入 = zlDatabase.GetPara(65, glngSys, , "11")
    
    '药品按规格下医嘱
    gbln药品按规格下医嘱 = Val(zlDatabase.GetPara(69, glngSys)) = 1
    
    '皮试结果有效时间
    gint过敏登记有效天数 = Val(zlDatabase.GetPara(70, glngSys))
    
    '长期医嘱次日生效
    gbln长期医嘱次日生效 = Val(zlDatabase.GetPara(71, glngSys)) = 1
    
    '是否要求首先输入类别
    gbln收费类别 = Val(zlDatabase.GetPara(72, glngSys, , 1)) <> 0
    
    
    '医嘱发送生成划价单的类别
    gstr住院发送划价单 = zlDatabase.GetPara(80, glngSys)
    gstr门诊发送划价单 = zlDatabase.GetPara(86, glngSys)

    '门诊自动发料
    gbln门诊自动发料 = Val(zlDatabase.GetPara(92, glngSys)) <> 0
    
    '自动发药退药
    gbln收费后自动发药 = zlDatabase.GetPara(45, glngSys) = "1"
    
    '从属项目汇总计算折扣
    gbln从项汇总折扣 = Val(zlDatabase.GetPara(93, glngSys)) <> 0
    
    '记帐报警包含划价费用
    gbln报警包含划价费用 = Val(zlDatabase.GetPara(98, glngSys)) <> 0
    
    '检验医嘱发送时生成条形码
    gbln发送生成条形码 = Val(zlDatabase.GetPara(143, glngSys)) <> 0
    
    '当不输类别时,输入费用项目时,首位当作类别简码
    gblnFeeKindCode = Val(zlDatabase.GetPara(144, glngSys)) <> 0 And Not gbln收费类别
    
    '分批药品出库方式
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    
    '输液配置中心(性质为“配制中心”的药房)
    gstr输液配置中心 = Get输液配置中心
    gbyt病人审核方式 = Val(zlDatabase.GetPara(185, glngSys))    '49501
    gbln未入科禁止记账 = Val(zlDatabase.GetPara(215, glngSys)) = 1 '51612

    '医嘱补录时是否只允许补录临嘱
    gbln只允许补录临嘱 = Val(zlDatabase.GetPara(191, glngSys)) <> 0
    '转科病人的补录时限
    glng补录时限 = Val(zlDatabase.GetPara(158, glngSys, , "24"))
    
    '下达医嘱时显示产地
    gblnShowOrigin = Val(zlDatabase.GetPara(162, glngSys, , "1")) <> 0
    
    '门诊一卡通,项目执行前必须先收费或先记帐审核
    gbln执行前先结算 = Val(zlDatabase.GetPara(163, glngSys)) <> 0
    
    
    '输血和皮试医嘱执行后需要核对
    gstr医嘱核对 = zlDatabase.GetPara(186, glngSys)
        
    '抗菌药物分级管理
    gblnKSSStrict = Val(zlDatabase.GetPara(187, glngSys)) <> 0
    gbln抗菌药物使用自备药 = Val(zlDatabase.GetPara(188, glngSys)) <> 0
    
    '是否启用手术分级管理
    gbln手术分级管理 = Val(zlDatabase.GetPara(209, glngSys)) <> 0
    
    '是否启用手术按医师授权管理
    gbln手术授权管理 = Val(zlDatabase.GetPara(217, glngSys)) <> 0
    
    '是否启用参数：主刀医师达到手术等级无需审核
    gbln手术等级管理 = Val(zlDatabase.GetPara(254, glngSys)) <> 0
    
    '是否启用输血分级管理
    gbln输血分级管理 = Val(zlDatabase.GetPara(216, glngSys)) <> 0
    '输血申请三级审核
    gbln输血申请三级审核 = Val(zlDatabase.GetPara(218, glngSys)) <> 0
    '输血申请只能由中级及以上医师提出
    gbln输血申请中级以上 = Val(zlDatabase.GetPara(219, glngSys)) <> 0
    
    '是否安装血库系统
    gbln血库系统 = (IsSysSetUp(2200) And Val(zlDatabase.GetPara(236, glngSys)) <> 0)
    gbln显示血液库存 = Val(zlDatabase.GetPara(286, glngSys)) = 0 And gbln血库系统 = True
    gbln下达用血申请确定血液信息 = Val(zlDatabase.GetPara(293, glngSys)) <> 0 And gbln血库系统 = True
    '医嘱超量时必须输入原因
    gbyt超量原因 = Val(zlDatabase.GetPara(230, glngSys, , 1))
    
    '当病人当前科室或者挂号科室在这之内的下医嘱时可以不录入超量说明，格式为科室id串竖线分割
    gstr不录超量科室 = zlDatabase.GetPara(233, glngSys, , "")
    gstr不录超量科室 = "," & Replace(gstr不录超量科室, "|", ",") & ","
    
    '当病人当前科室在这之内的停嘱时可以不录入停嘱说明，格式为科室id串竖线分割
    gstr可不填停嘱原因科室 = zlDatabase.GetPara(285, glngSys, , "")
    gstr可不填停嘱原因科室 = "," & Replace(gstr可不填停嘱原因科室, "|", ",") & ","
    
    '允许修改n天内登记的医嘱执行记录
    gint医嘱执行有效天数 = Val(zlDatabase.GetPara(220, glngSys))
    
    '转科时未审核销帐单据检查
    gbyt转科时未审核销帐单据检查 = Val(zlDatabase.GetPara(227, glngSys))
    
    '项目开单后立即收费或记帐审核
    gbln开单后立即收费或记帐审核 = Val(zlDatabase.GetPara(232, glngSys)) = 1
    
    '申请单启用环节，启用申请单后必须使用申请单下达医嘱
    Call Get申请单相关参数
    
    gbln反算天数 = Val(zlDatabase.GetPara(240, glngSys)) = 1
    gbln启用影像信息系统接口 = Val(zlDatabase.GetPara(255, glngSys)) = 1
    gbln特殊药品分开发送 = Val(zlDatabase.GetPara(262, glngSys)) <> 0
    gbln医嘱终止原因 = Val(zlDatabase.GetPara(271, glngSys)) = 1
    gbln科室药房对照按本机参数设置 = Val(zlDatabase.GetPara(274, glngSys)) = 1
    gbln会诊科室下达医嘱处理 = Val(zlDatabase.GetPara(302, glngSys)) = 1
    gbln审方系统 = "" <> GetParaURL("药师处方审查", "审查结果查询")
    
    On Error Resume Next
    If gobjDrugExplain Is Nothing Then Set gobjDrugExplain = CreateObject("zlKnowledgeConvert.CallView")
    err.Clear: On Error GoTo errH
    
    InitSysPar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiYear(lng病人ID As Long) As Integer
'功能：获取病人的准确年龄
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intYear As Integer
    
    On Error GoTo errH
    
    strSQL = "Select Sysdate as 当前,出生日期,年龄 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!出生日期) Then
            intYear = Year(rsTmp!当前) - Year(rsTmp!出生日期)
            If Format(rsTmp!当前, "MMdd") < Format(rsTmp!出生日期, "MMdd") Then
                intYear = intYear - 1
            End If
            If intYear < 0 Then intYear = 0
        Else
            intYear = Val(NVL(rsTmp!年龄))
        End If
    End If
    GetPatiYear = intYear
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省用法ID(ByVal int类型 As Integer, ByVal int来源 As Integer, Optional ByVal strFilter As String = "", Optional ByVal lng病人性质 As Long) As Long
'功能：返回缺省的给药途径或中药煎法
'参数：int类型=2-给药途径,3-中药煎法,4-中药用法,6-采集方法(检验)
'      int来源=1-门诊,2-住院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str服务对象 As String
    
    If lng病人性质 = 1 Then
        str服务对象 = ",1,2,3,"
    Else
        str服务对象 = "," & int来源 & ",3,"
    End If
    
    strSQL = "Select ID From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型=[1] " & strFilter & _
        " And (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
        " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
        " And Instr([2],','||服务对象||',')>0 And Rownum<100" & _
        " Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", CStr(int类型), str服务对象)
    If Not rsTmp.EOF Then
        Get缺省用法ID = rsTmp!ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Check适用用法(ByVal lng用法ID As Long, ByVal lng项目id As Long, ByVal int来源 As Integer, Optional ByVal lng收费细目ID As Long, Optional ByVal lng病人性质 As Long) As Boolean
'功能：检查指定的用法是否适用于指定的项目
'参数：int来源=1-门诊,2-住院
'      lng病人性质   0-普通住院病人,1-门诊留观病人,2-住院留观病人
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnCheckUseM As Boolean
    Dim bln总数 As Boolean
    Dim str服务对象 As String
    
    On Error GoTo errH
    
    If lng病人性质 = 1 Then
        str服务对象 = ",1,2,3,"
    Else
        str服务对象 = "," & int来源 & ",3,"
    End If
    
     '中西成药获取用法是否严格控制
    blnCheckUseM = CheckDrugUseM(lng收费细目ID, lng项目id)

    '检查项目用法用量
    strSQL = "Select Count(A.用法ID) as 总数,Max(Decode(A.用法ID,[2],1,0)) as 指定" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And Instr([3],','||B.服务对象||',')>0 And A.项目ID=[1] And A.性质>0" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, lng用法ID, str服务对象)
    
    If Not blnCheckUseM Then
        If NVL(rsTmp!总数, 0) <= 1 Then
            Check适用用法 = True
    
        ElseIf NVL(rsTmp!指定, 0) = 1 Then
            Check适用用法 = True
        End If
    Else
        If NVL(rsTmp!总数, 0) = 1 And (Not blnCheckUseM) Then
            Check适用用法 = True
            Exit Function
        ElseIf NVL(rsTmp!总数, 0) = 0 Then
            If (Not blnCheckUseM) And lng收费细目ID = 0 Then
                Check适用用法 = True
                Exit Function
            Else
                bln总数 = True
            End If
        ElseIf NVL(rsTmp!指定, 0) = 1 Then
            Check适用用法 = True
            Exit Function
        End If
    
        '检查药品用法用量
        If blnCheckUseM And lng收费细目ID <> 0 Then
            strSQL = "Select Count(A.用法ID) as 总数,Max(Decode(A.用法ID,[2],1,0)) as 指定" & _
                " From 药品用法用量 A,诊疗项目目录 B" & _
                " Where A.用法ID=B.ID And Instr([3],','||B.服务对象||',')>0 And A.药品ID=[1] And A.性质>0" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng收费细目ID, lng用法ID, str服务对象)
            If NVL(rsTmp!总数, 0) = 0 And bln总数 Then
                Check适用用法 = True
                Exit Function
            ElseIf NVL(rsTmp!指定, 0) = 1 Then
                Check适用用法 = True
                Exit Function
            End If
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDrugUseM(ByVal lng药品ID As Long, ByVal lng药名ID As Long) As Boolean
    '判断药品是否启用严格控制用法用量
    On Error GoTo errH
    
    '使用药品ID查询
    If lng药品ID <> 0 Then
        If Val(Sys.RowValue("药品规格", lng药品ID, "严格控制用法用量", "药品ID") & "") = 1 Then
            CheckDrugUseM = True
            Exit Function
        Else
            '读取药名ID
            If lng药名ID = 0 Then lng药名ID = Val(Sys.RowValue("药品规格", lng药品ID, "药名ID", "药品ID") & "")
        End If
    End If
    
    '使用药名ID查询
    If lng药名ID <> 0 Then
        If Val(Sys.RowValue("药品特性", lng药名ID, "严格控制用法用量", "药名ID") & "") = 1 Then
            CheckDrugUseM = True
            Exit Function
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Check上班安排(ByVal bln药房 As Boolean) As Boolean
'功能：检查医院的科室是否使用了上班安排
'参数：bln药房=是检查药房上班还是其它科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Static bln药房Load As Boolean
    Static bln药房Last As Boolean
    Static bln非药Load As Boolean
    Static bln非药Last As Boolean
    
    If bln药房 Then '是否有安排只需读取一次
        If bln药房Load Then Check上班安排 = bln药房Last: Exit Function
    Else
        If bln非药Load Then Check上班安排 = bln非药Last: Exit Function
    End If
    
    On Error GoTo errH
    
    If bln药房 Then
        strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 IN('西药房','成药房','中药房') And Rownum<2"
    Else
        strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 Not IN('西药房','成药房','中药房') And Rownum<2"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "Check上班安排")
    Check上班安排 = rsTmp.RecordCount > 0
    
    If bln药房 Then
        bln药房Load = True: bln药房Last = Check上班安排
    Else
        bln非药Load = True: bln非药Last = Check上班安排
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get操作员部门ID(ByVal int服务对象 As Integer, Optional ByVal lng默认部门 As Long) As Long
'功能：取操作员所属服务对指定对象的部门，缺省部门优先
    Static rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    If rsTmp Is Nothing Then
        blnNew = True
    Else
        blnNew = (rsTmp.State = adStateClosed)
    End If
    
    If blnNew Then
        strSQL = "Select Distinct B.部门ID,Nvl(B.缺省,0) as 缺省,C.服务对象 From 部门人员 B,部门性质说明 C" & _
            " Where B.人员ID = [1] And B.部门ID=C.部门ID" & _
            " Order by 缺省 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID)
    End If
    
    '74794,冉俊明,2014-7-18,护士在记账时候调用成套方案时未使用成套方案内的执行科室
    If lng默认部门 <> 0 Then
        rsTmp.Filter = "(服务对象 = 3 and 部门ID = " & lng默认部门 & ") " & _
                    "or (服务对象 = " & int服务对象 & " and 部门ID = " & lng默认部门 & ")"
        If Not rsTmp.EOF Then Get操作员部门ID = rsTmp!部门ID: Exit Function
    End If
    
    rsTmp.Filter = "服务对象 = 3 or 服务对象 = " & int服务对象
    
    If Not rsTmp.EOF Then
        Get操作员部门ID = rsTmp!部门ID
    Else
        Get操作员部门ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal lng病人ID As Long, lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目id As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal lng执行科室ID As Long, _
    Optional ByVal bytMode As Byte, Optional ByVal bytCallBy As Byte, _
    Optional ByVal int调用场合 As Integer = 1, _
    Optional lng成套缺省执行科室 As Long = 0) As Long
'功能：获取非药收费项目的执行科室
'参数：int范围=1.门诊,2-住院
'      lng执行科室ID=指定的缺省执行科室ID(用于药品和卫材)
'      bytMode=1-要返回缺省值,0-其它
'      bytCallBy=0-医嘱程序调用,1-附费程序调用
'      int调用场合=1-门诊,2-住院
'      lng成套缺省执行科室-缺省执行科室ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim lng病人病区ID As Long, bytDay As Byte
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        lng药房 = Val(zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省发料部门", glngSys, _
            IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        
        If lng主页ID > 0 Then
            lng病人病区ID = GetPatiUnitID(lng病人ID, lng主页ID)
        End If
        '有执行科室设置时
        strSQL = _
            " Select Distinct" & _
            "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            " And (A.开单科室ID is NULL Or instr([2],','||a.开单科室id||',')>0)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int范围, "," & lng病人科室ID & "," & lng病人病区ID & ",", lng项目id)
        If Not rsTmp.EOF Then
            If bytMode = 1 Then Get收费执行科室ID = rsTmp!执行科室ID  '如果都没有，则返回第一个可用的执行科室
            
            '1:缺省为指定的(医嘱的)执行科室,不管是否服务于病人科室
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            
            '2.缺省为参数指定的缺省科室
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            
            '3:其它可服务于病人科室的执行科室
            If rsTmp.EOF Then
                '2.0 如果成套中存在缺省的执行科室,则缺省为成套指定的缺省科室
                If lng成套缺省执行科室 <> 0 Then
                    rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                    If Not rsTmp.EOF Then
                            Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                    End If
                End If
                '2.1:尝试缺省为病人科室
                If lng执行科室ID <> lng病人科室ID And lng药房 <> lng病人科室ID Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人科室ID
                End If
                '3.2:尝试缺省为病人病区
                If rsTmp.EOF And lng主页ID <> 0 Then
                    If lng病人病区ID <> 0 And lng病人病区ID <> lng病人科室ID And lng病人病区ID <> lng执行科室ID And lng病人病区ID <> lng药房 Then
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人病区ID
                    End If
                End If
            End If
            '3.3:可服务于病人科室的一个执行科室
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            
            '3.4可服务于所有科室的当前病人科室执行
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0 And 执行科室ID=" & lng病人科室ID
            
            '4:如果都没有，则返回0用于检查
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省西药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省成药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int范围 = 2 Or int调用场合 = 2, "住院", "门诊") & "缺省中药房", glngSys, _
                IIF(bytCallBy = 1, p医嘱附费管理, IIF(int范围 = 2 Or int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达)), , , , , lng病人科室ID))
        End If
        
        '药品从系统指定的储备药房中找
        If Not Check上班安排(True) Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[5]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                " And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str药房, int范围, lng病人科室ID, lng项目id, bytDay)
        If Not rsTmp.EOF Then
            If lng成套缺省执行科室 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                If Not rsTmp.EOF Then
                        Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                End If
            End If
            Get收费执行科室ID = rsTmp!执行科室ID
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无明确科室
                If lng成套缺省执行科室 <> 0 Then
                    Get收费执行科室ID = lng成套缺省执行科室: Exit Function
                End If
                Get收费执行科室ID = Get操作员部门ID(int范围)
            Case 1 '1-病人所在科室
                Get收费执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng病人科室ID
                Else
                    Get收费执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
                Get收费执行科室ID = Get操作员部门ID(int范围, lng成套缺省执行科室)
            Case 4 '4-指定科室
                strSQL = "Select Distinct Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.收费细目ID=[1] And A.执行科室ID=B.部门ID" & _
                    " And B.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                    " And A.执行科室ID=C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " Order by 排序" '默认科室优先
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                    If lng成套缺省执行科室 <> 0 Then
                         rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                         If Not rsTmp.EOF Then
                                 Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                         End If
                     End If
                    Get收费执行科室ID = rsTmp!执行科室ID
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 6 '6-开单人所在科室
                Get收费执行科室ID = lng开单科室ID
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = Get操作员部门ID(int范围)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗执行科室ID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str类别 As String, ByVal lng项目id As Long, ByVal lng药品ID As Long, ByVal int执行科室 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, ByVal int期效 As Integer, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal blnBy缺省 As Boolean, Optional ByVal lngPreID As Long, Optional ByVal int调用场合 As Integer = 1) As Long
'功能：根据诊疗项目执行科室信息返回缺省的执行科室ID
'参数：lng药品ID=药品ID,确定到规格时要用
'      int执行科室=项目执行科室标志
'      lng病人科室ID=病人科室ID
'      lng西药房,lng成药房,lng中药房=药品缺省药房,药品类时需要
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
'      blnBy缺省=获取缺省药房时，如果本地有指定，是否按本地缺省指定的药房来，没有则不返回(如果仅有一个则仍返回)
'      lngPreID=前面医嘱的执行科室，如果传了该值，则带有优先缺省的性质；目前只用于药品
'      int调用场合=1-门诊,2-住院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str药房 As String, lng药房 As Long, str药房IDs As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean, strStock As String
    
    On Error GoTo errH
    
    '由用户特定情况直接返回
    strSQL = "Select zl_ClinicExeDept([1]," & IIF(lng主页ID = 0, "Null", "[2]") & ",[3],[4]," & IIF(lng药品ID = 0, "Null", "[5]") & ",[6]," & IIF(lng开嘱科室ID = 0, "Null", "[7]") & ") as 执行科室ID From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get诊疗执行科室ID", lng病人ID, lng主页ID, str类别, lng项目id, lng药品ID, lng病人科室ID, lng开嘱科室ID)
    If Not rsTmp.EOF Then
        If NVL(rsTmp!执行科室ID, 0) <> 0 Then
            Get诊疗执行科室ID = rsTmp!执行科室ID
            Exit Function
        End If
    End If
    
    '未定制返回的情况下根据程序规则处理
    If InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "缺省西药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID))
            str药房IDs = zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "可用西药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "缺省成药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID))
            str药房IDs = zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "可用成药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = Val(zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "缺省中药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID))
            str药房IDs = zlDatabase.GetPara(IIF(int调用场合 = 2, "住院", "门诊") & "可用中药房", glngSys, IIF(int调用场合 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        End If
        
        '药品库存限制
        If bln规格 And str药房IDs <> "" Then
            If gblnStock Then
                strStock = " And Exists(" & _
                    " Select 1 From 药品库存" & _
                    " Where (Nvl(批次,0)=0 Or 效期 Is Null Or 效期>Trunc(Sysdate))" & _
                    " And 性质=1 And 药品ID=[4] And 库房ID=A.执行科室ID" & _
                    " And 可用数量>0 And Instr('," & str药房IDs & ",',','||库房ID||',')>0)"
            Else
                strStock = " And Instr('," & str药房IDs & ",',','||A.执行科室ID||',')>0"
            End If
        End If
        
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                 IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[6]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str药房, int范围, lng病人科室ID, lng药品ID, lng项目id, bytDay)
        If Not rsTmp.EOF Then
            If blnBy缺省 And (lng药房 <> 0 Or lngPreID <> 0) Then
                If rsTmp.RecordCount > 1 Then
                    rsTmp.Filter = "执行科室ID=" & lngPreID
                    If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
                End If
            Else
                Get诊疗执行科室ID = rsTmp!执行科室ID
                rsTmp.Filter = "执行科室ID=" & lngPreID
                If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng药房
                If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            End If
            If Not rsTmp.EOF Then Get诊疗执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室
            Case 0 '0-无执行的叮嘱
                Exit Function
            Case 1 '1-病人所在科室
                Get诊疗执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get诊疗执行科室ID = lng病人科室ID
                Else
                    Get诊疗执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                End If
            Case 3 '3-操作员所在科室
                Get诊疗执行科室ID = Get操作员部门ID(int范围)
            Case 4 '4-指定科室
                If lng项目id = 0 Then
                    If int范围 = 1 Then
                        Get诊疗执行科室ID = lng病人科室ID
                    ElseIf int范围 = 2 Then
                        Get诊疗执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                    End If
                Else
                    If int期效 = 1 Then bln上班安排 = Check上班安排(False)
                    If Not bln上班安排 Then
                        strSQL = "Select Distinct Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                            " From 诊疗执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID=B.部门ID And A.诊疗项目ID=[1]" & _
                            " And B.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                            " And A.执行科室ID=C.ID And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " Order by 排序" '默认科室优先
                    Else
                        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                        strSQL = _
                            " Select Distinct Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID,Decode(A.病人来源,Null,2,1) as 排序" & _
                            " From 诊疗执行科室 A,部门安排 B,部门性质说明 C,部门表 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.星期=[4]" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                            " And A.执行科室ID=C.部门ID And C.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                            " And A.执行科室ID=D.ID And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & _
                            " And (D.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL)" & _
                            " And A.诊疗项目ID=[1]" & _
                            " Order by 排序"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, int范围, lng病人科室ID, bytDay)
                    If Not rsTmp.EOF Then
                        Get诊疗执行科室ID = rsTmp!执行科室ID
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                        If rsTmp.EOF Then rsTmp.Filter = "执行科室ID=" & lng病人科室ID
                        If rsTmp.EOF And int范围 = 2 Then rsTmp.Filter = "执行科室ID=" & GetPatiUnitID(lng病人ID, lng主页ID)
                        If Not rsTmp.EOF Then Get诊疗执行科室ID = rsTmp!执行科室ID
                    ElseIf gbln指定医嘱在其他科室执行 Then
                        If int范围 = 1 Then
                            Get诊疗执行科室ID = lng病人科室ID
                        ElseIf int范围 = 2 Then
                            Get诊疗执行科室ID = GetPatiUnitID(lng病人ID, lng主页ID)
                        End If
                    End If
                End If
            Case 5 '5-院外执行
                Exit Function
            Case 6 '6-开单人所在科室
                Get诊疗执行科室ID = lng开嘱科室ID
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExecDeptValidate(ByVal lng执行科室ID As Long, ByVal lng病人科室ID As Long, ByVal int范围 As Integer, ByVal lng诊疗项目ID As Long) As Boolean
'功能：检查指定的执行科室是否有效
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 1" & vbNewLine & _
            "From 部门表 A, 部门性质说明 C" & vbNewLine & _
            "Where a.Id = [1] And a.Id = c.部门id And c.服务对象 In ([3], 3) And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)" & vbNewLine & _
            "  And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (Exists" & vbNewLine & _
            " (Select 1 From 诊疗执行科室 B Where b.诊疗项目ID =[4] And b.执行科室ID = a.Id And (b.病人来源 Is Null Or b.病人来源 = [3]) And (b.开单科室ID Is Null Or b.开单科室ID = [2]))" & vbNewLine & _
            ") And Rownum<2"
            'Or这一句问题号：41496，为了成套中设置了执行科室，调用成套时的时候保证这个科室可用。
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckExecDeptValidate", lng执行科室ID, lng病人科室ID, int范围, lng诊疗项目ID)
    CheckExecDeptValidate = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get可用药房IDs(ByVal str类别 As String, ByVal lng项目id As Long, _
    ByVal lng药品ID As Long, ByVal lng科室id As Long, Optional ByVal int范围 As Integer = 2) As String
'功能：获取药品的有效诊疗执行科室ID串,用于判断缺省执行科室
'参数：lng科室ID=病人科室ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str药房 As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim str药房IDs As String, str可用药房 As String
    
    '系统可以指定药品执行科室,这里提取所有可选的供再选择
    If str类别 = "5" Then
        str药房 = "西药房"
        str可用药房 = zlDatabase.GetPara(Decode(int范围, 1, "门诊", 2, "住院", "") & "可用西药房", glngSys, Decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , lng科室id)
    ElseIf str类别 = "6" Then
        str药房 = "成药房"
        str可用药房 = zlDatabase.GetPara(Decode(int范围, 1, "门诊", 2, "住院", "") & "可用成药房", glngSys, Decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , lng科室id)
    ElseIf str类别 = "7" Then
        str药房 = "中药房"
        str可用药房 = zlDatabase.GetPara(Decode(int范围, 1, "门诊", 2, "住院", "") & "可用中药房", glngSys, Decode(int范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , lng科室id)
    End If
            
    '药品从系统指定的储备药房中找
    If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
    If Not bln上班安排 Then
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            IIF(int范围 <> 3, " And (A.病人来源 is NULL Or A.病人来源=[2])", "") & _
            IIF(lng科室id <> 0, " And (A.开单科室ID is NULL Or A.开单科室ID=[3])", "") & _
            IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    Else
        bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
        strSQL = _
            " Select Distinct C.ID" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And D.部门ID=C.ID And D.星期=[6]" & _
            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
            IIF(int范围 <> 3, " And (A.病人来源 is NULL Or A.病人来源=[2])", "") & _
            IIF(lng科室id <> 0, " And (A.开单科室ID is NULL Or A.开单科室ID=[3])", "") & _
            IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]")
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str药房, int范围, lng科室id, lng药品ID, lng项目id, bytDay)
    Do While Not rsTmp.EOF
        If str可用药房 = "" Then
            str药房IDs = str药房IDs & "," & rsTmp!ID
        ElseIf InStr("," & str可用药房 & ",", "," & rsTmp!ID & ",") > 0 Then
            str药房IDs = str药房IDs & "," & rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Get可用药房IDs = Mid(str药房IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get可用发料部门IDs(ByVal lng材料ID As Long, ByVal lng科室id As Long, Optional ByVal int范围 As Integer = 2, Optional ByVal lng诊疗项目ID As Long) As String
'功能：获取卫材的有效诊疗执行科室ID串,用于判断缺省执行科室
'参数：lng科室ID=病人科室ID
'      int范围=1-门诊,2-住院(缺省)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str发料部门IDs As String
    
    strSQL = _
        " Select Distinct C.ID" & _
        " From " & IIF(lng材料ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
        " And B.服务对象 IN([1],3) And B.部门ID=C.ID " & IIF(lng材料ID <> 0, " And A.收费细目ID=[3]", " And A.诊疗项目ID=[4]") & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
        IIF(int范围 <> 3, " And (A.病人来源 is NULL Or A.病人来源=[1])", "") & _
        IIF(lng科室id <> 0, " And (A.开单科室ID is NULL Or A.开单科室ID=[2])", "")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int范围, lng科室id, lng材料ID, lng诊疗项目ID)
    Do While Not rsTmp.EOF
        str发料部门IDs = str发料部门IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get可用发料部门IDs = Mid(str发料部门IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗执行科室(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    objCbo As Object, ByVal str类别 As String, ByVal lng项目id As Long, ByVal lng药品ID As Long, _
    ByVal int执行科室 As Integer, ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, _
    ByVal lng当前执行ID As Long, ByVal int期效 As Integer, Optional ByVal int范围 As Integer = 2, _
    Optional ByVal bln输液 As Boolean, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal bln屏蔽输液中心 As Boolean, Optional ByVal lng用法ID As Long, Optional ByVal bln静脉营养 As Boolean, Optional ByVal lng病区ID As Long, Optional ByRef str执行科室ids As String, Optional ByVal lng病人性质 As Long) As Boolean
'功能：根据诊疗项目执行科室信息返回可用的执行科室在指定下拉框中
'参数：int执行科室=项目执行科室标志
'      lng病人科室ID=病人科室ID
'      lng当前执行ID=医嘱当前的执行科室ID
'      int期效=0-长嘱,1-临嘱,临嘱可能要判断上班时间
'      int范围=1-门诊,2-住院(缺省)
'      bln输液=当前药品是否属于输液类的（给药途径为输液）
'      blnEditable=false表示当前医嘱不可编辑
'      bln屏蔽输液中心=如果一组输液其他医嘱的执行科室不是输液配置中心，则此条医嘱也不能设置为输液配置中心
'      lng用法ID=如果是药品，传入药品对应的给药途径ID,如果是输血检验，则传入采集和输血途径
'      bln静脉营养=静脉营养的药品，只能在配药中心配
'      str执行科室ids 出参返当前的所有可选科室IDs串，逗号分割,暂未处理
'      lng病人性质 0-普通住院病人,1-门诊留观病人,2-住院留观病人 (给药途径，中药煎法，中药服法，检验采集，输血途径，输血采集) 门诊留观病人不区分服务对象，开单科室
'说明：对非药医嘱,当前的执行科室可能是强行选择出来的,需要显示在选择框中;另选择框中增加一个其它供选择
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str药房 As String, str药房IDs As String
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean, i As Long
    Dim strStock As String, lng操作员科室ID As Long
    Dim str输液配置科室 As String
    Dim int输液配置医嘱期效 As Integer
    Dim bln不默认配置中心 As Boolean
    Dim bln指定输液配置中心 As Boolean
    Dim str配液给药途径 As String
    Dim strFilter As String
    Dim bln特殊输液药 As Boolean '自备、不取、离院带
    Dim strTmp As String
    Dim str默认药房 As String
    Dim strSQLOther As String
    
    If str类别 = "4" Then
        strSQL = _
            " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
            " From " & IIF(lng药品ID <> 0, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
            " And (A.开单科室ID is NULL Or A.开单科室ID=[3] or exists (Select 1 From 病区科室对应 x Where x.科室id=[3] and x.病区id=A.开单科室ID))" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            IIF(lng药品ID <> 0, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & _
            " Order by B.服务对象,C.编码"
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        bln规格 = ((int期效 = 1 Or gbln药品按规格下医嘱) And lng药品ID <> 0) Or lng药品ID <> 0
        
        '系统可以指定药品执行科室,这里提取所有可选的供再选择
        If str类别 = "5" Then
            str药房 = "西药房"
            str药房IDs = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用西药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
            str默认药房 = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "缺省西药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            str药房IDs = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用成药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
            str默认药房 = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "缺省成药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            str药房IDs = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "可用中药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
            str默认药房 = zlDatabase.GetPara(IIF(int范围 = 2, "住院", "门诊") & "缺省中药房", glngSys, IIF(int范围 = 2, p住院医嘱下达, p门诊医嘱下达), , , , , lng病人科室ID)
        End If
            
        '药品库存限制
        If bln规格 And str药房IDs <> "" And blnEditable Then
            If gblnStock Then
                strStock = " And Exists(" & _
                    " Select 1 From 药品库存" & _
                    " Where (Nvl(批次,0)=0 Or 效期 Is Null Or 效期>Trunc(Sysdate))" & _
                    " And 性质=1 And 药品ID=[4] And 库房ID=A.执行科室ID" & _
                    " And 可用数量>0 And Instr('," & str药房IDs & ",',','||库房ID||',')>0)"
            Else
                strStock = " And Instr('," & str药房IDs & ",',','||A.执行科室ID||',')>0"
            End If
        End If
        
        '是否启用输液配置中心
        'bln特殊输液药  自备、不取、离院带 当3个参数不启用时为不启用
        strTmp = ""
        bln特殊输液药 = Val(zlDatabase.GetPara("自备药允许发往静配中心", glngSys, p输液配置中心, "0")) = 1
        strTmp = strTmp & IIF(bln特殊输液药, "1", "0")
        bln特殊输液药 = Val(zlDatabase.GetPara("不取药允许发往静配中心", glngSys, p输液配置中心, "0")) = 1
        strTmp = strTmp & IIF(bln特殊输液药, "1", "0")
        bln特殊输液药 = Val(zlDatabase.GetPara("离院带药允许发往静配中心", glngSys, p输液配置中心, "0")) = 1
        strTmp = strTmp & IIF(bln特殊输液药, "1", "0")
         
        bln特殊输液药 = strTmp = "000"
       
        If bln屏蔽输液中心 And bln特殊输液药 Then
            '71101当选择自备药或其他执行性质时，也允许选择输液配置中心，因为有可能输液配置中心同时兼任普通药房的工作。
            'strStock = strStock & " And A.执行科室ID <> [12]  "
        Else
            If gstr输液配置中心 <> "" And blnEditable Then
                If bln输液 And (int范围 = 2 Or lng病人性质 = 1) Then
                    str输液配置科室 = zlDatabase.GetPara("来源病区", glngSys, p输液配置中心, "")
                    int输液配置医嘱期效 = Val(zlDatabase.GetPara("医嘱类型", glngSys, p输液配置中心, "1")) - 1
                    str配液给药途径 = zlDatabase.GetPara("输液给药途径", glngSys, p输液配置中心)
                    If (str输液配置科室 = "" Or InStr("," & str输液配置科室 & ",", "," & lng病区ID & ",") > 0) And (int期效 = int输液配置医嘱期效 Or int输液配置医嘱期效 = -1) And _
                            (InStr("," & str配液给药途径 & ",", "," & lng用法ID & ",") > 0 Or str配液给药途径 = "") Or bln静脉营养 Then
                        bln指定输液配置中心 = True
                    Else
                        'bln不默认配置中心=启用了输液配置中心，但是期效\科室\给药途径参数不匹配，把配置中心放最后。
                        bln不默认配置中心 = True
                    End If
                Else
                    '不是输液类的或范围不是住院的，把输液配置中心放最后
                    bln不默认配置中心 = True
                End If
            End If
        End If
        
        '药品从系统指定的储备药房中找
        If int范围 = 1 Then bln上班安排 = Check上班安排(True) '住院医嘱不管药房上班安排
        If Not bln上班安排 Then
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象,Decode(c.Id,[13],0,1) as 默认部门" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                IIF(bln不默认配置中心, " Order By 默认部门,Decode(instr('," & gstr输液配置中心 & ",',',' || C.ID || ','),0,B.服务对象,9),C.编码", _
                " Order by 默认部门,B.服务对象,C.编码")
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象,Decode(c.Id,[13],0,1) as 默认部门" & _
                " From " & IIF(bln规格, "收费执行科室", "诊疗执行科室") & " A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                " And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                " And D.部门ID=C.ID And D.星期=[8]" & _
                " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                " And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                " And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                IIF(bln规格, " And A.收费细目ID=[4]", " And A.诊疗项目ID=[5]") & strStock & _
                IIF(bln不默认配置中心, " Order By 默认部门,Decode(instr('," & gstr输液配置中心 & ",',',' || C.ID || ','),0,B.服务对象,9),C.编码", _
                " Order by 默认部门,B.服务对象,C.编码")
        End If
    Else
        Select Case int执行科室
            Case 0, 5 '0-无执行的叮嘱,5-院外执行
                Get诊疗执行科室 = True: Exit Function
            Case 1 '1-病人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([3],[6]) Order by 编码"
                Else
                    strSQL = _
                        " Select A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,病案主页 B" & _
                        " Where A.ID=B.当前病区ID And B.病人ID=[9] And B.主页ID=[10]" & _
                        " Union " & _
                        " Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
            Case 3 '3-操作员所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([7],[6]) Order by 编码"
                lng操作员科室ID = Get操作员部门ID(int范围)
            Case 4 '4-指定科室
                If int期效 = 1 Then bln上班安排 = Check上班安排(False)
                If Not bln上班安排 Then
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                        " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
                        " Where A.ID=B.执行科室ID And B.诊疗项目ID=[5] And A.ID=C.部门ID" & _
                        " And C.服务对象 IN([2],3) And (B.病人来源 is NULL Or B.病人来源=[2])" & _
                        " And (B.开单科室ID is NULL Or B.开单科室ID=[3])" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                Else
                    bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                        " From 诊疗执行科室 A,部门安排 B,部门表 C,部门性质说明 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.部门ID=C.ID And B.星期=[8]" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                        " And C.ID=D.部门ID And D.服务对象 IN([2],3) And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[3]) And A.诊疗项目ID=[5]" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                        " Order by 编码"
                End If
                '对于门诊留观病人不区分各种条件，上班安排，服务对象，开单科室，病人来源
                If lng病人性质 = 1 Then
                    Set rsTmp = Get诊疗项目记录(lng项目id)
                    If "E" = rsTmp!类别 & "" And InStr(",2,3,4,6,8,9,", "," & rsTmp!操作类型 & ",") > 0 Then
                        strSQL = " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                            " From 诊疗执行科室 A,部门表 C" & _
                            " Where A.执行科室ID+0=C.ID And A.诊疗项目ID=[5]" & _
                            " Union Select ID,编码,简码,名称 From 部门表 Where ID=[6]" & _
                            " Order by 编码"
                    End If
                End If
            Case 6 '6-开单人所在科室
                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([11],[6]) Order by 编码"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str药房, int范围, lng病人科室ID, lng药品ID, lng项目id, _
         lng当前执行ID, lng操作员科室ID, bytDay, lng病人ID, lng主页ID, lng开嘱科室ID, gstr输液配置中心, Val(str默认药房))
         
    '输液药品发送到输液配置中心
    If bln指定输液配置中心 Then
        For i = 0 To UBound(Split(gstr输液配置中心, ","))
            If str药房IDs = "" Or InStr("," & str药房IDs & ",", "," & Split(gstr输液配置中心, ",")(i) & ",") > 0 Then
                strFilter = strFilter & " Or ID=" & Split(gstr输液配置中心, ",")(i)
            End If
        Next
        rsTmp.Filter = Mid(strFilter, 5)
        If rsTmp.RecordCount = 0 Then
            '如果输液配置中心不是存储库房，则可以选择其他存储库房为发药药房(65111)
            rsTmp.Filter = 0
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
    End If
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        '使用API快速加入,不然可能有点慢
        AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, rsTmp!编码 & "-" & rsTmp!名称
        SetComboData objCbo.hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If InStr(",5,6,7,", str类别) > 0 And Val(str默认药房) <> 0 And Val(str默认药房) = Val(rsTmp!ID & "") And objCbo.Text = "" Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        If lng当前执行ID = rsTmp!ID Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '仅非药、非卫材医嘱可以选择
    If InStr(",4,5,6,7,", str类别) = 0 And gbln指定医嘱在其他科室执行 And IIF(gbln血库系统 = True And str类别 = "K", 0, 1) = 1 Then
        AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[其它...]"
        SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get频率信息_编码(ByVal str编码 As String, str频率 As String, _
    int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String) As Boolean
'功能：返回频率的相关信息
'参数：str编码=频率编码
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    str频率 = ""
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select 编码,名称,简码,英文名称,频率次数,频率间隔,间隔单位,适用范围 From 诊疗频率项目 Where 编码=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str编码)
    If Not rsTmp.EOF Then
        str频率 = NVL(rsTmp!名称)
        int频率次数 = NVL(rsTmp!频率次数, 0)
        int频率间隔 = NVL(rsTmp!频率间隔, 0)
        str间隔单位 = NVL(rsTmp!间隔单位)
    End If
    Get频率信息_编码 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get频率信息_名称(ByVal str频率 As String, int频率次数 As Integer, _
    int频率间隔 As Integer, str间隔单位 As String, str范围 As String, Optional str频率编码 As String) As Boolean
'功能：返回频率的相关信息
'参数：str频率=频率名称
'      str范围=1-西医,2-中医,-1-一次性,-2-持续性
'返回：当按名称取到时，返回True，否则返回False
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    strSQL = "Select 频率次数,频率间隔,间隔单位,编码 From 诊疗频率项目 Where 名称=[1] And Instr([2],','||适用范围||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str频率, "," & str范围 & ",")
    If Not rsTmp.EOF Then
        int频率次数 = NVL(rsTmp!频率次数, 0)
        int频率间隔 = NVL(rsTmp!频率间隔, 0)
        str间隔单位 = NVL(rsTmp!间隔单位)
        str频率编码 = "" & rsTmp!编码
        Get频率信息_名称 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省频率(ByVal lng项目id As Long, ByVal int范围 As Integer, str频率 As String, _
    int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String) As Boolean
'功能：从所有适用频率项目中取一个作为缺省频率
'参数：int范围=1-西医,2-中医,-1-一次性,-2-持续性
'返回：缺省频率信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnLoad As Boolean
    
    On Error GoTo errH
    
    str频率 = ""
    int频率次数 = 0
    int频率间隔 = 0
    str间隔单位 = ""
    
    '先取诊疗常用频率
    blnLoad = True
    If lng项目id <> 0 And int范围 = 1 Then
        strSQL = "Select B.名称,B.频率次数,B.频率间隔,B.间隔单位 From 诊疗用法用量 A,诊疗频率项目 B" & _
            " Where A.项目ID=[1] And A.用法ID Is Null And A.频次=B.编码 And B.适用范围=[2]" & _
            " Order by A.性质"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get缺省频率", lng项目id, int范围)
        If Not rsTmp.EOF Then blnLoad = False
    End If
    If blnLoad Then
        strSQL = "Select 名称,频率次数,频率间隔,间隔单位 From 诊疗频率项目 Where 适用范围=[1] Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get缺省频率", int范围)
    End If
    If Not rsTmp.EOF Then
        str频率 = NVL(rsTmp!名称)
        int频率次数 = NVL(rsTmp!频率次数, 0)
        int频率间隔 = NVL(rsTmp!频率间隔, 0)
        str间隔单位 = NVL(rsTmp!间隔单位)
    End If
    Get缺省频率 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check频率可用(ByVal lng项目id As Long, ByVal int范围 As Integer, ByVal str频率 As String) As Boolean
'功能：检查指定频率是否适用于项目的常用频率
'参数：int范围=1-西医,2-中医,-1-一次性,-2-持续性
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng项目id = 0 Then Check频率可用 = True: Exit Function
    
    strSQL = "Select B.名称 From 诊疗用法用量 A,诊疗频率项目 B" & _
        " Where A.项目ID=[1] And A.用法ID Is Null And A.频次=B.编码 And B.适用范围=[2]" & _
        " Order by A.性质"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Check频率可用", lng项目id, int范围)
    If rsTmp.EOF Then
        '没有设置常用频率，则没限制
        Check频率可用 = True
    ElseIf rsTmp.RecordCount = 1 Then
        '只设置了一个常用频率，只是缺省，也没限制
        Check频率可用 = True
    ElseIf str频率 <> "" Then
        rsTmp.Filter = "名称='" & str频率 & "'"
        If Not rsTmp.EOF Then Check频率可用 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get缺省时间(int范围 As Integer, str频率 As String, Optional lng给药途径ID As Long) As String
'功能：获取指定执行频率缺省的执行时间方案
'参数：int范围=1-西医;2-中医;-1-一次性;-2-持续性
'      lng给药途径ID=是否可以按指定给药途径优先取,否则取不确定给药途径的方案
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(A.给药途径ID,0) as 用法,A.时间方案" & _
        " From 诊疗频率时间 A,诊疗频率项目 B" & _
        " Where A.执行频率=B.编码 And B.适用范围=[1]" & _
        " And (A.给药途径ID is NULL Or A.给药途径ID=[2]) And B.名称=[3]" & _
        " Order by 用法 Desc,A.方案序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", int范围, lng给药途径ID, str频率)
    If Not rsTmp.EOF Then Get缺省时间 = NVL(rsTmp!时间方案)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get时间方案(objCbo As Object, int范围 As Integer, str频率 As String, Optional lng给药途径ID As Long) As Boolean
'功能：读取指定频率可用的诊疗频率时间方案在指定下拉框中,并设置缺省项(或保持原有值)
'参数：int范围=1-西医;2-中医;-1-一次性;-2-持续性;-3-必要时;-5-需要时
'      str频率=诊疗频率项目名称
'      lng给药途径ID=是否只读取指定给药途径的时间方案,否则为非药品的执行时间方案
'      str执行时间=缺省执行时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '按理不同方案(不管是否指定给药途径)的执行时间应该不相同,否则会重复出现
    strSQL = "Select A.方案序号,A.时间方案" & _
        " From 诊疗频率时间 A,诊疗频率项目 B" & _
        " Where A.执行频率=B.编码 And B.名称=[1]" & _
        " And (A.给药途径ID is NULL Or A.给药途径ID=[2])" & _
        " And B.适用范围=[3]" & _
        " Order by A.方案序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str频率, lng给药途径ID, int范围)
    strSQL = objCbo.Text: objCbo.Clear 'Clear会导致Text清空
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!时间方案
        rsTmp.MoveNext
    Next
    objCbo.Text = strSQL: objCbo.Tag = ""
    Get时间方案 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get开嘱医生(ByVal lng病人科室ID As Long, ByVal bln护士站 As Boolean, str缺省医生 As String, lng医生ID As Long, _
    Optional objCbo As Object, Optional ByVal int范围 As Integer = 2, Optional ByVal blnOnlyDefault As Boolean) As Boolean
'功能：获取可用的开嘱医生在指定的下拉框中
'参数：lng病人科室ID=病人所在科室ID
'      bln护士站=是否由护士代医生下医嘱
'      objCbo=要加入医生清单的下拉框
'      str缺省医生=缺省定位的医生,如果不传objCbo,则先优先定位,再返回缺省医生和医生ID
'      int范围=1-门诊,2-住院(缺省)
'      blnOnlyDefault=当指定了缺省医生时，是否只读取该医生的信息，此时应传入"str缺省医生"，"bln护士站=True"。
'                     如果同时传入了objCbo对象，则将当前缺省医生追加到该列表控件中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
        
    If bln护士站 Then
        If blnOnlyDefault And str缺省医生 <> "" Then
            strSQL = "Select ID,编号,姓名,简码 From 人员表 Where 姓名=[4]"
        Else
            '病人所在科室的医生
            strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIF(objCbo Is Nothing, ",B.部门ID", "") & _
                " From 人员表 A,部门人员 B,人员性质说明 C" & _
                " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
                " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And B.部门ID=[1]" & _
                " Order by A.简码"
            '全院住院科室的医生
            strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN([2],3)"
            strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & IIF(objCbo Is Nothing, ",B.部门ID", "") & _
                " From 人员表 A,部门人员 B,人员性质说明 C" & _
                " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
                " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And B.部门ID IN(" & strSQL & ")" & _
                " Order by A.简码"
        End If
    Else '医生下医嘱时,限制为只能为医生本人
        strSQL = "Select ID,编号,姓名,简码 From 人员表 Where ID=[3]"
        '医生下达医嘱时，在选择别人下达的医嘱进行修改时，开嘱医生应该加载，否则在编辑界面上选择其它医生下达的医嘱时，下方选项卡中的开叫医生为空。
        If str缺省医生 <> "" And str缺省医生 <> UserInfo.姓名 Then
            strSQL = strSQL & " union all Select ID,编号,姓名,简码 From 人员表 Where 姓名=[4]"
        End If
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get开嘱医生", lng病人科室ID, int范围, UserInfo.ID, str缺省医生)
    If objCbo Is Nothing Then
        If Not rsTmp.EOF Then
            If Not bln护士站 Then
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            ElseIf bln护士站 Then
                If str缺省医生 <> "" Then
                    '缺省医生(住院医师)优先
                    rsTmp.Filter = "姓名='" & str缺省医生 & "'"
                Else
                    '病人科室的医生优先
                    rsTmp.Filter = "部门ID=" & lng病人科室ID
                End If
                If rsTmp.EOF Then rsTmp.Filter = 0
                lng医生ID = rsTmp!ID
                str缺省医生 = rsTmp!姓名
            End If
        End If
    Else
        If blnOnlyDefault Then
            '先删除"其它"
            i = Cbo.FindIndex(objCbo, -1)
            If i <> -1 Then objCbo.RemoveItem objCbo.ListCount - 1
            
            '定位或加入其它选项
            If Not rsTmp.EOF Then
                i = Cbo.FindIndex(objCbo, rsTmp!ID)
                If i = -1 Then
                    objCbo.AddItem NVL(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                Else
                    Call Cbo.SetIndex(objCbo.hwnd, i)
                End If
            End If
            
            '加入"其它"供选择
            AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[其它...]"
            SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        Else
            '全部新加入
            objCbo.Clear
            For i = 1 To rsTmp.RecordCount
                AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, NVL(rsTmp!简码) & "-" & Chr(13) & rsTmp!姓名
                SetComboData objCbo.hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
                If rsTmp!姓名 = str缺省医生 Then
                    Call Cbo.SetIndex(objCbo.hwnd, i - 1)
                End If
                rsTmp.MoveNext
            Next
        End If
    End If
    Get开嘱医生 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check本科执行(ByVal lng执行科室ID As Long) As Boolean
'功能：确定指定的执行科室是否本科(医生科室)
'参数：lng执行科室ID=医嘱的执行科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim arr科室ID(1 To 4) As Long
    
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1] And 部门ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", UserInfo.ID, lng执行科室ID)
    Check本科执行 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lng药品ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal int范围 As Integer = 2, _
        Optional ByVal strDepartments As String, Optional ByVal lng总量 As Double) As Double
'功能：获取指定库房指定药品不分批库存(以门诊或住院单位)
'参数：int范围=1-门诊,2-住院(缺省),0-表示按售价
'      strDepartments可用执行科室字符串，用于批量查询库存
'      lng总量 如果lng总量不为空，则查询是否有库存大于这个总量
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    '获取药品库存(不分批或分批药品),药房不分批药品不管效期
    If int范围 = 0 Or int范围 = 3 Then
        strSQL = _
            " Select Nvl(Sum(A.可用数量),0) as 库存" & _
            " From 药品库存 A" & _
            " Where A.性质=1" & _
            " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0 Group By A.库房ID"
    Else
        strTmp = IIF(int范围 = 1, "门诊", "住院")
        strSQL = _
            " Select Nvl(Sum(A.可用数量),0)/Nvl(B." & strTmp & "包装,1) as 库存" & _
            " From 药品库存 A,药品规格 B" & _
            " Where A.药品ID=B.药品ID(+) And A.性质=1" & _
            " And (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.药品ID=[1] And Instr([2],',' || a.库房id || ',')>0" & _
            " Group by Nvl(B." & strTmp & "包装,1),A.库房ID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng药品ID, IIF(strDepartments = "", "," & lng库房ID & ",", "," & strDepartments & ","))
    
    Do While Not rsTmp.EOF
    
        If strDepartments = "" Then
            GetStock = Format(rsTmp!库存, "0.00000")
            Exit Function
        Else
            If Val(rsTmp!库存) & "" > lng总量 Then
                GetStock = Format(rsTmp!库存, "0.00000")
                Exit Function
            End If
        End If
        rsTmp.MoveNext
    
    Loop
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupCount(ByVal lng组合ID As Long, ByVal int来源 As Integer, Optional bln期效 As Boolean = True, Optional ByVal lng病人性质 As Long) As Long
'功能：获取组合项目中的项目数
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str服务对象 As String
    Dim strWhere As String
     
    On Error GoTo errH
    
    str服务对象 = "," & int来源 & ",3,"
    
    If lng病人性质 = 1 Then
        strWhere = "  And (Instr([2],','||B.服务对象||',')>0  or instr(',E2,E3,E4,E6,E8,E9,',','||b.类别||b.操作类型||',')>0)"
    Else
        strWhere = "  And Instr([2],','||B.服务对象||',')>0 "
    End If
    
    strSQL = "Select Count(*) as NUM" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,收费项目目录 C" & _
        " Where A.诊疗项目ID=B.ID(+) And A.收费细目ID=C.ID(+) And A.诊疗组合ID=[1]" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL) " & strWhere & _
        " And (A.收费细目ID is NULL Or (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL) And Instr([2],','||c.服务对象||',')>0 )" & _
        IIF(bln期效 And int来源 <> 2, " And A.期效=1", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng组合ID, str服务对象)
    If Not rsTmp.EOF Then GetGroupCount = NVL(rsTmp!Num, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetGroupNone(ByVal lng配方ID As Long, ByVal int来源 As Integer) As String
'功能：读取指定配方中无效的组成中药提示
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = "Select B.名称" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,药品规格 C,收费项目目录 D" & _
        " Where A.诊疗项目ID=B.ID And B.ID=C.药名ID And C.药品ID=D.ID And A.诊疗组合ID=[1]" & _
        " And (Not (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is NULL)" & _
        " Or Nvl(B.服务对象,0) Not IN([2],3))" & _
        " And (Not (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 is NULL)" & _
        " Or Nvl(D.服务对象,0) Not IN([2],3))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng配方ID, int来源)
    Do While Not rsTmp.EOF
        strMsg = strMsg & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    GetGroupNone = Mid(strMsg, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcDrugPrice(ByVal lng药品ID As Long, lng药房ID As Long, ByVal dbl数量 As Double, _
    Optional ByVal str费别 As String, Optional ByVal blnNone加班加价 As Boolean, Optional ByVal int场合 As Integer, Optional ByVal str药品价格等级 As String, Optional ByVal str卫材价格等级 As String, Optional ByVal str普通项目价格等级 As String) As Double
'功能：计算药品实价(即然要计算实价,药品则肯定为变价)，传入费别时，则计算实收金额
'参数：dbl数量=售价数量,按费别打折时计算的是实收金额
'      str费别=是否按费别计算打折的价格,主要在直接计算药品的金额而不显示单价时用
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
'      int场合  0－住院，1－门诊

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim dbl时价 As Double
    
    If dbl数量 = 0 Then Exit Function
    '不区分门诊住院，统一调用Zl_Fun_Getprice 公共接口
    On Error GoTo errH
    strSQL = "select Zl_Fun_Getprice([1],[2],[3],0,null) as 结果 from dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcDrugPrice", lng药品ID, lng药房ID, dbl数量)
    strSQL = rsTmp!结果 & ""
    If InStr(strSQL, "|") > 0 Then dbl时价 = Val(Split(strSQL, "|")(0))
    '当有费别参数时，是结合数量计算打折实收金额
    If str费别 <> "" And dbl时价 <> 0 Then
        dbl时价 = Format(dbl时价 * dbl数量, gstrDec)
        strSQL = _
            " Select A.屏蔽费别,B.收入项目ID" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcDrugPrice", lng药品ID)
        If rsTmp.EOF Then Exit Function
        
        '根据费别重新计算实收金额
        If Not (NVL(rsTmp!屏蔽费别, 0) = 1) Then
            dbl时价 = ActualMoney(str费别 & IIF(gstr动态费别 <> "", "," & gstr动态费别, ""), rsTmp!收入项目ID, dbl时价, lng药品ID, lng药房ID, dbl数量)
        End If
    End If
    CalcDrugPrice = dbl时价
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CalcPrice(ByVal lng项目id As Long, Optional ByVal str费别 As String, _
    Optional ByVal dbl数量 As Double, Optional ByVal blnNone加班加价 As Boolean, _
    Optional ByVal lng执行科室ID As Long, Optional ByVal lng卫材医嘱ID As Long, Optional ByVal str药品价格等级 As String, Optional ByVal str卫材价格等级 As String, Optional ByVal str普通项目价格等级 As String) As Double
'功能：获取收费细目的当前售价价格金额,变价返回缺省价格
'参数：str费别=是否按费别计算打折的实收金额
'      dbl数量=按费别计算时,必须要传入数量(按售价单位),这时计算的是实收金额
'      lng执行科室ID=当传入了费别时需要,可能按成本加打折计算
'      gbln加班加价=发送时计算才有用,其它地方都为False
'      blnNone加班加价=为真时,"gbln加班加价"无效
'      lng卫材医嘱ID=如果传入该参数，表示非跟踪在用的时价卫材医嘱的价格从医嘱计价中读取
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim dbl单价 As Double, dbl金额 As Double
    
    On Error GoTo errH
    
    If lng卫材医嘱ID <> 0 Then
        strSQL = "Select 单价 From 病人医嘱计价 Where 医嘱ID=[1] And 收费细目ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CalcPrice", lng卫材医嘱ID, lng项目id)
        If Not rsTmp.EOF Then dbl单价 = NVL(rsTmp!单价, 0)
    End If
    
    If str费别 = "" Then
        strSQL = _
            " Select Sum(Decode(Nvl(A.是否变价,0),1,Decode([2],0,B.缺省价格,[2]),B.现价)" & _
                IIF(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & ") as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            GetPriceGradeSQL(str药品价格等级, str卫材价格等级, str普通项目价格等级, "A", "B", "3", "4", "5") & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目id, dbl单价, str药品价格等级, str卫材价格等级, str普通项目价格等级)
        If Not rsTmp.EOF Then dbl金额 = NVL(rsTmp!金额, 0)
    Else
        '本来可以将ActualMoney函数的SQL一起写在这里，但费别可能被删除而求不出数据
        strSQL = _
            " Select A.屏蔽费别,A.加班加价,B.加班加价率,B.收入项目ID,Decode(Nvl(A.是否变价,0),1,Decode([2],0,B.缺省价格,[2]),B.现价)" & _
                IIF(gbln加班加价 And Not blnNone加班加价, "*Decode(A.加班加价,1,1+Nvl(B.加班加价率,0)/100,1)", "") & " as 金额" & _
            " From 收费项目目录 A,收费价目 B" & _
            " Where A.ID=B.收费细目ID And A.ID=[1]" & _
            GetPriceGradeSQL(str药品价格等级, str卫材价格等级, str普通项目价格等级, "A", "B", "3", "4", "5") & _
            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng项目id, dbl单价, str药品价格等级, str卫材价格等级, str普通项目价格等级)
        For i = 1 To rsTmp.RecordCount
            If NVL(rsTmp!屏蔽费别, 0) = 1 Then
                dbl金额 = dbl金额 + Format(dbl数量 * Format(NVL(rsTmp!金额, 0), "0.00000"), gstrDec)
            Else
                dbl金额 = dbl金额 + ActualMoney(str费别 & IIF(gstr动态费别 <> "", "," & gstr动态费别, ""), rsTmp!收入项目ID, Format(dbl数量 * Format(NVL(rsTmp!金额, 0), "0.00000"), gstrDec), _
                    lng项目id, lng执行科室ID, dbl数量, IIF(gbln加班加价 And Not blnNone加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0))
            End If
            rsTmp.MoveNext
        Next
    End If
    CalcPrice = dbl金额
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Check护理等级变动交叉(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strDate As String) As Boolean
'功能：检查病人如果当前无有效护理等级医嘱，且有已停止的护理等级医嘱，且停止时间在当前开始时间之后，则提示禁止。
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    Check护理等级变动交叉 = False
    If Not IsDate(strDate) Then Exit Function
    
    strSQL = "Select 1" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And a.诊疗类别 = 'H' And a.诊疗项目id = b.Id And b.操作类型 = '1' And a.医嘱状态 In (8, 9) And" & vbNewLine & _
            "      a.执行终止时间 > [3] And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID, CDate(strDate))
    If rsTmp.RecordCount > 0 Then
        '如果当前存在有效的护理等级，新医嘱发送后会自动停止旧的医嘱，变动记录的开始和结束时间就不会交叉
        strSQL = "Select 1" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And a.诊疗类别 = 'H' And a.诊疗项目id = b.Id And b.操作类型 = '1' And a.医嘱状态 In (3, 5, 7)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
        If rsTmp.RecordCount = 0 Then
            Check护理等级变动交叉 = True
            MsgBox "护理等级医嘱的开始时间不允许在已停止的护理等级医嘱之前。", vbInformation, gstrSysName
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ActualMoney(str费别 As String, ByVal lng收入项目ID As Long, ByVal cur应收金额 As Currency, _
    Optional ByVal lng收费细目ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal dbl数量 As Double, Optional ByVal dbl加班加价率 As Double) As Currency
'功能：根据收费细目ID或收入项目ID(前者优先),应收金额,按费别设置的分段比例打折规则计算实收金额；
'       或对药品按成本加收比例规则计算实收金额
'参数：str费别=病人费别；如果是按动态费别,传入格式为"病人费别,动态费别1,动态费别2,..."
'      lng库房ID,dbl数量,对药品类项目按成本价加收打折时才需要传入
'      dbl数量=包含付数在内的售价数量
'      dbl加班加价率=小数比率,传入的应收金额已按加班加价计算时需要，用于还原及重算
'返回：按打折规则和比例计算的实收金额,如果是动态费别,则"str费别"返回最优惠费别(注意如果未打折计算,可能原样返回,也可能返回第一个)
'说明：
'按成本价加收比例打折的两种计算方法(实际是一种)：
'1.打折金额 = 成本金额 * (1 + 加收比例)
'2.打折金额 = 成本价 * (1 + 加收比例) * 零售数量
'相关的计算公式：
'      成本价 = 药品售价 * (1 - 差价率)
'      成本金额 = 售价金额 * (1 - 差价率) = 成本价 * 零售数量
'      有库存金额时:差价率 = 库存差价 / 库存金额,否则:差价率 = 指导差价率
'      对于分批药品，应每个出库批次分别计算成本价和成本金额
'      对于时价分批，"药品售价=Nvl(零售价,实际金额/实际数量)"；分批或时价药品库存不足时，不予打折计算。
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str费别, lng收费细目ID, lng收入项目ID, cur应收金额 / (1 + dbl加班加价率), dbl数量, lng库房ID)
        
    str费别 = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl加班加价率), gstrDec)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Load动态费别(lng科室id As Long) As String
'功能：权限指定科室读取当前有效的动态费别(目前只用于门诊)
'返回：费别串="三八节,五一节"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select 编码,简码,名称 From 费别" & _
        " Where Nvl(属性,1)=2 And Nvl(适用科室,1)=1 And Nvl(服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Union ALL" & _
        " Select Distinct A.编码,A.简码,A.名称" & _
        " From 费别 A,费别适用科室 B" & _
        " Where A.名称=B.费别 And B.科室ID=[1]" & _
        " And Nvl(A.属性,1)=2 And Nvl(A.适用科室,1)=2 And Nvl(A.服务对象,3) IN(1,3)" & _
        " And Trunc(Sysdate) Between Nvl(A.有效开始,To_Date('1900-01-01','YYYY-MM-DD'))" & _
        " And Nvl(A.有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Load动态费别", lng科室id)
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    Load动态费别 = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckUserEmpower(ByVal lng诊疗项目ID As Long) As Boolean
'功能：检查操作员是否具有手术项目的开单权
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select Count(*) as 权限 From 人员手术权限 Where 人员id = [1] And 诊疗项目id = [2] And 记录性质 = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", UserInfo.ID, lng诊疗项目ID)
    CheckUserEmpower = rsTmp!权限 > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDocEmpower(ByVal lng诊疗项目ID As Long, ByVal strAppend As String) As Boolean
'功能：检查操作员是否具有手术项目的执行权
'参数：strAppend=当前申请附项的填写情况串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    strSQL = "select A.ID from 诊治所见项目 A,诊治所见分类 B where a.分类id=b.id and b.编码='06' and A.中文名='主刀医生'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpower")
    If rsTmp.RecordCount > 0 Then
        lngID = rsTmp!ID
        arrItem = Split(strAppend, "<Split1>")
        For i = 0 To UBound(arrItem)
            arrSub = Split(arrItem(i), "<Split2>")
            If Val(arrSub(2)) = lngID Then
                If Trim(arrSub(3)) <> "" Then
                    strDoc = Trim(arrSub(3))
                End If
                Exit For
            End If
        Next
    End If
    If strDoc = "" Then strDoc = UserInfo.姓名
    strSQL = "Select Count(*) as 权限 From 人员手术权限 A,人员表 B Where A.人员id = B.ID And B.姓名=[1] And A.诊疗项目id = [2] And A.记录性质 = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng诊疗项目ID)
    CheckDocEmpower = Val(rsTmp!权限 & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get临床科室(ByVal int范围 As Integer, Optional ByVal lng病人科室ID As Long, _
    Optional lng缺省科室ID As Long, Optional objCbo As Object, Optional ByVal blnBed As Boolean, Optional ByVal blnNode As Boolean = True, _
    Optional bln仅操作员科室 As Boolean, Optional ByVal int类别 As Integer) As Boolean
'功能：返回临床科室清单或缺省临床科室
'参数：int范围=1-门诊,2-住院,3-门诊或住院
'      lng病人科室ID=病人当前的科室,可能要排开该科室
'      objCbo=要加入科室清单的下拉框,不传时,返回缺省科室
'      lng缺省科室ID=有objCbo时,为缺省定位的科室；否则为要返回的缺省科室
'      blnBed=是否只取有床位的科室
'      blnNode=是否限制为当前站点的科室，转科医嘱调用时不限制
'      bln仅操作员科室=入院或留观医嘱，缺省科室为医生所属的住院科室
'      int类别 医嘱的类别，1－表示当前医嘱为门诊转住院医嘱Z2项目，非1为其它
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String, blnHave As Boolean
        
    On Error GoTo errH
    
    If int范围 = 1 Then
        strTmp = "1,3"
    ElseIf int范围 = 2 Then
        strTmp = "2,3"
    ElseIf int范围 = 3 Then
        strTmp = "1,2,3"
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,A.简码 " & _
        " From 部门表 A,部门性质说明 B " & IIF(bln仅操作员科室, ",部门人员 C ", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        IIF(blnNode, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
        " And A.ID=B.部门ID And Instr([1],','||B.服务对象||',')>0 And B.工作性质='临床'" & _
        IIF(lng病人科室ID <> 0, " And A.ID<>[2]", "") & _
        IIF(bln仅操作员科室, " And A.id = C.部门id And c.人员id =[3]", "") & _
        IIF(blnBed, " And (Exists(Select 科室ID From 床位状况记录 Where 科室ID=A.ID) Or Exists(Select 科室ID From 病区科室对应 Where 科室ID=A.ID))", "") & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", "," & strTmp & ",", lng病人科室ID, UserInfo.ID)
    
    If Not objCbo Is Nothing Then
        objCbo.Clear
        For i = 1 To rsTmp.RecordCount
            objCbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
            If rsTmp!ID = lng缺省科室ID Then
                Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                blnHave = True
            End If
            rsTmp.MoveNext
        Next
        
        If int类别 = 1 Then
            If lng缺省科室ID <> 0 And Not blnHave Then
                strSQL = "Select A.ID,A.编码,A.名称,A.简码" & _
                    " From 部门表 A,部门性质说明 B Where A.ID=B.部门ID And B.服务对象 IN(2,3) and a.id=[1]" & _
                    IIF(blnNode, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
                    " And B.工作性质='临床' And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng缺省科室ID)
                If Not rsTmp.EOF Then
                    objCbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
                    objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
                    Call Cbo.SetIndex(objCbo.hwnd, objCbo.NewIndex)
                End If
            End If
        End If
        
        If bln仅操作员科室 Then
            AddComboItem objCbo.hwnd, CB_ADDSTRING, 0, "[其它...]"
            SetComboData objCbo.hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
        End If
    ElseIf Not rsTmp.EOF Then
        lng缺省科室ID = rsTmp!ID
    End If
    Get临床科室 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗项目ID(ByVal lng医嘱ID As Long, bln中药检验 As Boolean) As Long
'功能：读取指定医嘱的诊疗项目ID
'参数：lng医嘱ID=相关ID为NULL的医嘱ID(成药,检查项目,主要手术,中药用法,及独立医嘱)
'      bln中药检验=医嘱ID是否为中药用法或采集方法的ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bln中药检验 Then
        strSQL = "Select 序号,诊疗项目ID From 病人医嘱记录 Where 诊疗类别 IN('7','C') And 相关ID=[1]"
    Else
        strSQL = "Select 序号,诊疗项目ID From 病人医嘱记录 Where ID=[1]"
    End If
    strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    strSQL = strSQL & " Order by 序号"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
    If Not rsTmp.EOF Then Get诊疗项目ID = NVL(rsTmp!诊疗项目ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get诊疗项目记录(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'功能：读取指定诊疗项目ID的记录
'参数：
    Dim strSQL As String
    
    strSQL = "Select /*+ rule*/ 计算规则,站点,类别,分类ID,ID,编码,名称,标本部位,计算单位,计算方式,执行频率,适用性别,单独应用,组合项目,操作类型,执行安排,执行科室,服务对象,计价性质,参考目录ID,人员ID,建档时间,撤档时间,录入限量,试管编码,执行分类,执行标记" & _
            " From 诊疗项目目录 Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN (Select Column_Value From Table(f_Num2list([1])))"
        Set Get诊疗项目记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get诊疗项目记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费项目记录(ByVal lngID As Long, Optional ByVal strIDs As String) As ADODB.Recordset
'功能：读取指定收费项目ID的记录
'参数：
    Dim strSQL As String
    
    strSQL = "Select /*+ rule*/ 类别,分类ID,ID,编码,名称,规格,产地,计算单位,说明,项目特性,费用类型,服务对象,屏蔽费别,是否变价,加班加价,补充摘要,费用确认,执行科室,标识主码,标识子码,备选码,最低限价,最低限价,建档时间,撤档时间,录入限量,计算方式,站点,启用原因,停用原因,病案费目" & _
            " From 收费项目目录 Where ID"
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = strSQL & " IN(Select Column_Value From Table(f_Num2list([1])))"
        Set Get收费项目记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs)
    Else
        strSQL = strSQL & " = [1]"
        Set Get收费项目记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDrugInfo(lng药名ID As Long, lng药品ID As Long, lng药房ID As Long, _
    Optional ByVal int范围 As Integer = 2, Optional ByVal bln停用 As Boolean = True) As ADODB.Recordset
'功能：获取指定药品相关信息
'参数：int范围=1-门诊,2-住院(缺省)
'      bln停用=是否排开已停用药品,用于长嘱药品发送处理
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    strTmp = IIF(int范围 = 1, "门诊", "住院")
    
    strSQL = _
        " Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
        " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
        " And 性质 = 1 And 库房ID=[1]" & IIF(lng药品ID <> 0, " And 药品ID=[2]", "") & _
        " Group by 药品ID Having Sum(Nvl(可用数量,0))<>0"
    strSQL = "Select A.药名ID,A.药品ID,A.剂量系数,A." & strTmp & "包装,A." & strTmp & "单位,A." & strTmp & "可否分零 As 可否分零,A.动态分零," & _
        " A.药房分批,B.是否变价,C.库存/A." & strTmp & "包装 as 库存,B.编码,Nvl(D.名称,B.名称) as 名称,B.规格,B.产地,B.撤档时间,B.服务对象,a.是否摆药" & _
        " From 药品规格 A,收费项目目录 B,(" & strSQL & ") C,收费项目别名 D" & _
        " Where A.药品ID=B.ID And A.药品ID=C.药品ID(+)" & _
        " And B.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[5]" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
        IIF(bln停用, " And B.服务对象 IN([3],3) And (B.撤档时间 is NULL Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))", "") & _
        " And A.药名ID=[4]" & IIF(lng药品ID <> 0, " And A.药品ID=[2]", "") & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng药房ID, lng药品ID, int范围, lng药名ID, IIF(gbyt药品名称显示 = 0, 1, 3))
    Set GetDrugInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = NVL(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetClinicBillID(ByVal lng项目id As Long, ByVal int场合 As Integer) As Long
'功能：获取诊疗项目对应的诊疗单据(不管附项,用于生成发送NO)
'参数：int场合=1-门诊,2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病历文件ID From 病历单据应用 Where 诊疗项目ID=[1] And 应用场合=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, int场合)
    If Not rsTmp.EOF Then GetClinicBillID = NVL(rsTmp!病历文件ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptIsWoman(ByVal lng科室id As Long, Optional ByVal str科室IDs As String) As Boolean
'功能：判断指定科室是否产科
'参数：str科室IDs-传入多个科室ID，判断是否其中有产科
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If str科室IDs = "" Then
        strSQL = "Select 工作性质,部门ID,服务对象 From 部门性质说明 Where 工作性质='产科' And 部门ID=[1]"
    Else
        strSQL = "Select /*+ Rule*/ 工作性质,部门ID,服务对象 From 部门性质说明 Where 工作性质='产科' And 部门ID In (Select Column_Value From Table(Cast(f_Str2List([2]) As zlTools.t_StrList)))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng科室id, str科室IDs)
    DeptIsWoman = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get科室IDs(Optional ByVal lng病区ID As Long) As String
'功能：根据病区ID获得病区对应的科室ID
'参数：是否取所属病区下的科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long

    strSQL = _
            " Select B.科室ID From 病区科室对应 B" & _
            " Where B.病区ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get科室IDs", lng病区ID)
    Get科室IDs = lng病区ID
    For i = 1 To rsTmp.RecordCount
        If InStr("," & Get科室IDs & ",", "," & rsTmp!科室ID & ",") = 0 Then
            Get科室IDs = Get科室IDs & "," & rsTmp!科室ID
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As String
'功能：检查病人在医技科室是否还有未执行完成(未执行或正在执行)的项目
'返回：医技科室名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(2,[1],[2],[3]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitExe", lng病人ID, lng主页ID, int婴儿)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = NVL(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As String
'功能：检查病人在药房是否还有未发药的药品或卫材
'返回：药房和发料部门名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],[3]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng病人ID, lng主页ID, int婴儿)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = NVL(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePause(ByVal lng医嘱ID As Long, Optional ByVal lng组ID As Long) As String
'功能：获取指定医嘱的暂停时间段记录
'返回："暂停时间,开始时间;...."
'注意：本方法利用了静态变进行缓存使用时注意先清一次缓存方式 Call GetAdvicePause(0)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    Static strLastPause As String
    Static lng相关ID As Long
    
    On Error GoTo errH
    
    If lng相关ID = lng组ID And lng组ID <> 0 Then GetAdvicePause = strLastPause: Exit Function
    If lng医嘱ID <> 0 Then
        strSQL = "Select 操作类型,操作时间 From 病人医嘱状态" & _
            " Where 操作类型 IN(6,7) And 医嘱ID=[1]" & _
            " Order by 操作时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
        For i = 1 To rsTmp.RecordCount
            If rsTmp!操作类型 = 6 Then
                strTmp = strTmp & ";" & Format(rsTmp!操作时间, "yyyy-MM-dd HH:mm:ss") & ","
            ElseIf rsTmp!操作类型 = 7 Then
                '启用的那一秒不在暂停的范围之内
                strTmp = strTmp & Format(DateAdd("s", -1, rsTmp!操作时间), "yyyy-MM-dd HH:mm:ss")
            End If
            rsTmp.MoveNext
        Next
    End If
    lng相关ID = lng组ID
    strLastPause = Mid(strTmp, 2)
    GetAdvicePause = Mid(strTmp, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlClinicCodeRepeat(str编码 As String, Optional lng项目id As Long) As Boolean
'功能：检查诊疗项目编码的是否与现有编码重复，重复则给出提示
'入参：str编码-输入的编码；lng项目ID-自己的ID号，当修改时，需要将自身除开才能判断
'出参：重复返回True；否则反馈Flase
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
        " From 诊疗项目目录 I,诊疗项目类别 K" & _
        " Where I.类别=K.编码 And I.编码=[1] And I.ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str编码, lng项目id)
    If Not rsTmp.EOF Then
        MsgBox "该项目编码与“" & rsTmp!名称 & "”的编码重复！", vbInformation, gstrSysName
        zlClinicCodeRepeat = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMax序号(ByVal strNO As String, ByVal int记录性质 As Integer, str登记时间 As String, int病人来源 As Integer) As Integer
'功能：获取指定单据当前的最大序号+1
'参数：str登记时间=组合医嘱只生成了部份主费用时，将要新生成的收费划价单(NO相同)的时间与已生成的一致。
'      int病人来源:1-门诊，2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    
    strTab = IIF(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")
    On Error GoTo errH
    str登记时间 = ""
    strSQL = "Select Max(序号) as 序号,Max(登记时间) as 时间 From " & strTab & " Where NO=[1] And 记录性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strNO, int记录性质)
    If Not rsTmp.EOF Then
        GetBillMax序号 = NVL(rsTmp!序号, 0) + 1
        If Not IsNull(rsTmp!时间) Then
            str登记时间 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax序号 = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DeptExist(ByVal str工作性质 As String, ByVal int服务对象 As Integer) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 部门ID From 部门性质说明 Where 工作性质=[1] And 服务对象 IN([2],3) And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str工作性质, int服务对象)
    DeptExist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedBySend(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, byt来源 As Byte) As Boolean
'功能：检查某次发送的医嘱中的费用是否已经执行了数据转出
'参数：lng发送号=因为门诊医嘱只有一次发送,可以不传入,byt来源:1-门诊，2-住院
'说明：1.在医嘱未转出的情况下，执行回退或作废操作时，如果包含已转出的费用，则禁止
'      2.对于住院长嘱有多次发送的情况，只判断当前要回退的这次医嘱发送费用
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")
    
    strSQL = "Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1]"
    strSQL = "Select B.NO From 病人医嘱发送 A,H" & strTab & " B" & _
        " Where A.记录性质=B.记录性质 And A.NO=B.NO" & _
        IIF(lng发送号 <> 0, " And A.发送号+0=[2]", "") & _
        " And A.医嘱ID IN(" & strSQL & ")" & _
        " Group by B.NO"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, lng发送号)
    If Not rsTmp.EOF Then MovedBySend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医疗付款码(ByVal str名称 As String) As String
'功能：根据医疗付款方式名称获取医疗付款性质  1-医保  2-公费
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If str名称 = "" Then Exit Function
    
    strSQL = "Select Decode(是否医保,1,1,Decode(是否公费,1,2,0)) 性质 From 医疗付款方式 Where 名称=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str名称)
    If Not rsTmp.EOF Then Get医疗付款码 = NVL(rsTmp!性质)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiDataMoved(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：判断指定病人的数据是否已转出
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 数据转出 From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查转出", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        CheckPatiDataMoved = Val("" & rsTmp!数据转出) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckOneDuty(ByVal str医嘱 As String, ByVal str职务 As String, ByVal str医生 As String, ByVal bln医保 As Boolean) As String
'功能：检查当前指定药品处方职务是否符合
'参数：str医嘱=药品医嘱提示内容
'      str职务=药品处方职务
'      str医生=开嘱医生
'      bln医保=是否公费或医保病人
'      grsDuty=记录医生职务缓存
'返回：职务不满足的提示信息，如果满足则返回空。
    Const STR_职务 = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim int职务A As Integer, int职务B As Integer
    
    If Len(str职务) <> 2 Or str医生 = "" Then Exit Function
    
    '取药品处方职务
    If bln医保 Then
        int职务B = Val(Right(str职务, 1))
    Else
        int职务B = Val(Left(str职务, 1))
    End If
    If int职务B = 0 Then Exit Function '不限制
    
    '取医生职务
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "医生", adVarChar, 50
        grsDuty.Fields.Append "职务", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.Filter = "医生='" & str医生 & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select 姓名,Nvl(聘任技术职务,0) as 职务 From 人员表 Where 姓名=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str医生)
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!医生 = rsTmp!姓名
            grsDuty!职务 = rsTmp!职务
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        int职务A = grsDuty!职务
    End If
        
    '检查职务要求
    If int职务A = 0 Then
        '医生未设置职务的情况
        strMsg = """" & str医嘱 & """要求的" & IIF(bln医保, "医保", "处方") & "职务不满足：" & vbCrLf & vbCrLf & IIF(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """未设置职务。"
    ElseIf int职务B < int职务A Then
        '数值越小职务越高
        strMsg = """" & str医嘱 & """要求的" & IIF(bln医保, "医保", "处方") & "职务不满足：" & vbCrLf & vbCrLf & IIF(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """的职务为""" & Split(STR_职务, ",")(int职务A - 1) & """。"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsDiagNoses(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str类型 As String) As Boolean
'功能：检查病人指定的诊断是否存在
'参数：lng就诊ID=门诊病人为挂号ID,住院病人为主页ID
'      str类型=诊断类型,如"1,11"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 记录来源,疾病ID,诊断ID,诊断描述,是否疑诊 From 病人诊断记录" & _
        " Where 病人ID=[1] And Nvl(主页ID,0)=[2] And Instr([3],','||诊断类型||',')>0 And 取消时间 Is Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng就诊ID, "," & str类型 & ",")
    ExistsDiagNoses = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人诊断记录(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str类型 As String) As ADODB.Recordset
'功能：获取病人诊断记录
'参数：lng就诊ID：门诊病人传挂号ID，住院病人传主页ID
'       诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'       记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select a.疾病id, a.诊断id, a.诊断描述, a.诊断次序, Nvl(b.编码, c.编码) As 编码, Nvl(b.名称, c.名称) 名称" & vbNewLine & _
             "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & vbNewLine & _
             "Where a.病人id = [1] And a.主页id = [2] And 取消时间 Is Null And 记录来源 IN (1, 3)  And NVL(A.编码序号,1) = 1 And Instr(',' ||[3]|| ',', ',' || 诊断类型 || ',') > 0 And a.疾病id = b.Id(+) And" & vbNewLine & _
             "      a.诊断id = c.Id(+)" & vbNewLine & _
             "Order By 记录来源, 诊断类型, 诊断次序"
    Set Get病人诊断记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng就诊ID, str类型)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人过敏记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人过敏记录
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng主页ID = 0 Then
        strSQL = "Select Distinct 药物ID,药物名,过敏源编码 From 病人过敏记录 Where 病人ID=[1] And 结果=1 And Nvl(过敏时间,记录时间)>Trunc(Sysdate-[3])"
    Else
        strSQL = "Select Distinct 药物ID,药物名,过敏源编码 From 病人过敏记录 Where 病人ID=[1] And 主页ID=[2] And 结果=1"
    End If
    Set Get病人过敏记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng主页ID, gint过敏登记有效天数)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadAdviceSignSource(ByVal int操作类型 As Integer, _
    ByVal lng病人ID As Long, ByVal varTime As Variant, strIDs As String, _
    ByVal lng签名id As Long, ByVal blnMoved As Boolean, strSource As String, _
    Optional ByVal str前提IDs As String, Optional ByVal colSomeTime As Collection, _
    Optional ByRef ColIDs As Collection, Optional ByRef ColSource As Collection) As Integer
'功能：获取病人用于电子签名/验证的医嘱源文内容
'参数：
'  int操作类型=要签名/验证签名的医嘱状态
'  签名时传入：
'    lng病人ID
'    varTime=病人挂号单号或主页ID
'    strIDs=指定要签名的医嘱ID序列(组ID)
'    str前提IDs=新开医嘱要签名的医嘱来源(是否医技)
'    colSomeTime=某医嘱的时间数据，如停止医嘱签名时，传入包含医嘱执行终止时间的数据，校对时传入校对时间数据
'  验证签名时：
'    lng签名ID=签名记录的ID
'    blnMoved=是否医嘱数据已转出
'返回：签名/验证签名的源文生成规则
'      strIDs=签名/验证签名的医嘱ID序列(每个明细ID)
'      strSource=签名/验证签名的医嘱源文
'      ColIDs=如果每条医嘱签名一次，则返回按每条医嘱返回医嘱ID集合
'      ColstrSource=如果每条医嘱签名一次，则返回按每条医嘱返回医嘱源文集合
    Dim rsTmp As New ADODB.Recordset
    Dim str组IDs As String, strSQL As String, i As Long
    Dim arrField As Variant, strField As String
    Dim strLine As String, intRule As Integer
    Dim bln每组医嘱单独签名 As Boolean
    Dim strID As String, str相关ID As String
    Dim strSourceTmp As String, strIDsTmp As String
    Dim strWhere As String
    
    On Error GoTo errH
    
    str组IDs = strIDs
    strSource = "": strIDs = ""
    intRule = 1 '这是最新的医嘱签名源文生成规则编号
    Set ColIDs = New Collection
    Set ColSource = New Collection
    bln每组医嘱单独签名 = Val(zlDatabase.GetPara(239, glngSys) & "") <> 0
    
    If lng签名id = 0 Then
        '签名时
        If int操作类型 = 1 Then
            If gbln血库系统 Then
                strWhere = " And a.诊疗项目id = c.Id(+) And" & vbNewLine & _
                    "      (a.诊疗项目id Is Null Or" & vbNewLine & _
                    "      Not (Nvl(a.审核状态, 0) <> 2 And (a.诊疗类别 = 'K' And Exists (Select 1" & vbNewLine & _
                    "                                                              From 病人医嘱记录 X, 诊疗项目目录 Y" & vbNewLine & _
                    "                                                              Where x.相关id = a.Id And x.诊疗项目id = y.Id And x.诊疗类别 = 'E' And" & vbNewLine & _
                    "                                                                    y.操作类型 = '8' And Nvl(y.执行分类, 0) = 0) Or" & vbNewLine & _
                    "        a.诊疗类别 = 'E' And c.操作类型 = '8' And Nvl(c.执行分类, 0) = 0)))"
            End If
            '对新开的医嘱进行签名：本次就诊/住院当前医生新下达的未签名医嘱；针对输血医嘱新开签名时，如果启用血库系统，只能对审核通的新开输血医嘱签名（审核状态＝2）。
            strSQL = _
                " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & IIF(gbln血库系统, ",诊疗项目目录 C", "") & " Where A.ID=B.医嘱ID And B.签名ID is Null And B.操作类型=1" & _
                " And A.医嘱状态=1 And Nvl(A.前提ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X)" & _
                " And Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))=[3]" & _
                " And Exists(Select M.姓名 From 人员表 M,执业类别 N" & _
                " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
                " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))" & strWhere & _
                IIF(TypeName(varTime) = "String", " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
                IIF(str组IDs <> "", " And Nvl(A.相关ID,A.ID) IN (Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([4])) X)", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, varTime, UserInfo.姓名, str组IDs, IIF("" = str前提IDs, "0", str前提IDs))
        Else
            '对要作废、停止、校对的医嘱进行签名：新开时签了名的指定医嘱，不一定是当前医生下达
            strSQL = _
                " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID is Not Null And B.操作类型=1" & _
                IIF(TypeName(varTime) = "String", " And A.病人ID+0=[1] And A.挂号单=[2]", " And A.病人ID=[1] And A.主页ID=[2]") & _
                IIF(str组IDs <> "", " And Nvl(A.相关ID,A.ID) IN(Select /*+cardinality(x,10)*/ x.Column_Value From Table(f_Num2list([3])) X)", "") & _
                " Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, varTime, str组IDs)
        End If
    Else
        '验证签名时:先读取签名时的源文生成规则
        strSQL = "Select 签名规则 From 医嘱签名记录 Where ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng签名id)
        If Not rsTmp.EOF Then intRule = NVL(rsTmp!签名规则, 1)
        '--
        strSQL = _
            " Select A.* From 病人医嘱记录 A,病人医嘱状态 B" & _
                " Where A.ID=B.医嘱ID And B.签名ID=[1] Order by A.婴儿,Nvl(A.相关ID,A.ID),A.序号"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng签名id)
    End If
    
    '医嘱源文的不同生成规则
    If intRule = 1 Then
        If int操作类型 = 3 Then
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,校对时间,执行性质,紧急标志,开嘱医生,开嘱时间"
        ElseIf int操作类型 = 8 Then
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行终止时间,执行性质,紧急标志,开嘱医生,开嘱时间"
        Else
            strField = "ID,相关ID,姓名,性别,年龄,婴儿,医嘱期效,开始执行时间,医嘱内容,标本部位,单次用量,总给予量," & _
                "医生嘱托,执行频次,频率次数,频率间隔,间隔单位,执行时间方案,执行性质,紧急标志,开嘱医生,开嘱时间"
        End If
    End If
    arrField = Split(strField, ",")
        
    '生成医嘱签名源文
    Do While Not rsTmp.EOF
        strLine = ""
        For i = 0 To UBound(arrField)
            If lng签名id = 0 And int操作类型 = 3 And arrField(i) = "校对时间" Then
                '校对医嘱签名时,对校对时间特殊处理：由于是在执行过程之前取签名源文,这时还未写入数据库
                strLine = strLine & vbTab & colSomeTime("_" & NVL(rsTmp!相关ID, rsTmp!ID))
            ElseIf lng签名id = 0 And int操作类型 = 8 And arrField(i) = "执行终止时间" Then
                '停止医嘱签名时,对终止时间特殊处理：由于是在执行过程之前取签名源文,这时还未写入数据库
                strLine = strLine & vbTab & colSomeTime("_" & NVL(rsTmp!相关ID, rsTmp!ID))
            Else
                If IsNull(rsTmp.Fields(arrField(i)).value) Then
                    strLine = strLine & vbTab & ""
                Else
                    If Rec.IsType(rsTmp.Fields(arrField(i)).Type, adDBTimeStamp) Then
                        strLine = strLine & vbTab & Format(rsTmp.Fields(arrField(i)).value, "yyyy-MM-dd HH:mm:ss")
                    Else
                        strLine = strLine & vbTab & rsTmp.Fields(arrField(i)).value
                    End If
                End If
            End If
        Next
        strSource = strSource & vbCrLf & Mid(strLine, 2)
        strIDs = strIDs & "," & rsTmp!ID
        strSourceTmp = strSourceTmp & vbCrLf & Mid(strLine, 2)
        strIDsTmp = strIDsTmp & "," & rsTmp!ID
        strID = rsTmp!ID: str相关ID = rsTmp!相关ID & ""
        rsTmp.MoveNext
        If bln每组医嘱单独签名 Then
            '每组医嘱单独签名则返回集合
            If rsTmp.EOF = False Then
                If rsTmp!ID & "" <> str相关ID And (rsTmp!相关ID & "" <> str相关ID Or (str相关ID = "" And rsTmp!相关ID & "" = "")) And rsTmp!相关ID & "" <> strID Then
                    ColIDs.Add Mid(strIDsTmp, 2)
                    ColSource.Add Mid(strSourceTmp, 3)
                    strIDsTmp = "": strSourceTmp = ""
                End If
            ElseIf strSourceTmp <> "" Then
                ColIDs.Add Mid(strIDsTmp, 2)
                ColSource.Add Mid(strSourceTmp, 3)
                strIDsTmp = "": strSourceTmp = ""
            End If
        End If
    Loop
    
    strSource = Mid(strSource, 3)
    strIDs = Mid(strIDs, 2)
    If ColIDs.Count = 0 Then
        ColIDs.Add strIDs
        ColSource.Add strSource
    End If
    
    ReadAdviceSignSource = intRule
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceSign(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, ByVal str人员 As String, ByVal dat时间 As Date) As Long
'功能：获取指定医嘱操作的签名ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 签名ID From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[2] And 操作人员=[3] And 操作时间=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, int类型, str人员, dat时间)
    If Not rsTmp.EOF Then
        GetAdviceSign = NVL(rsTmp!签名ID, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceSigns(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, ByVal str人员 As String, ByVal dat时间 As Date) As String
'功能：获取指定医嘱操作的签名ID字符串(多病人签名的情况--确认停止)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strAdciseSigns As String
    
    On Error GoTo errH
    
    strSQL = "Select distinct 签名ID" & vbNewLine & _
            "From 病人医嘱状态 A" & vbNewLine & _
            "Where a.操作类型 = [2] And a.操作人员 = [3] And a.操作时间 = [4] And Exists" & vbNewLine & _
            "  (Select 1 From 病人医嘱状态 B" & vbNewLine & _
            "        Where a.操作类型 = b.操作类型 And a.操作人员 = b.操作人员 And a.操作时间 = b.操作时间 And b.医嘱id = [1])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, int类型, str人员, dat时间)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!签名ID, 0) <> 0 Then
            strAdciseSigns = strAdciseSigns & "," & rsTmp!签名ID
        End If
        rsTmp.MoveNext
    Loop
    GetAdviceSigns = Mid(strAdciseSigns, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicesSameSign(ByVal lng签名id As Long) As String
'功能：获取相同签名ID的多组医嘱IDs(多病人签名一起签名时，回退某个病人的多组医嘱的确认停止)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strAdciseSigns As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(b.相关ID,b.ID) as 医嘱ID" & vbNewLine & _
            " From 病人医嘱状态 A,病人医嘱记录 B" & vbNewLine & _
            " Where a.医嘱id=b.id" & vbNewLine & _
            " and a.签名id=[1]"


    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng签名id)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!医嘱ID, 0) <> 0 Then
            strAdciseSigns = strAdciseSigns & "," & rsTmp!医嘱ID
        End If
        rsTmp.MoveNext
    Loop
    GetAdvicesSameSign = Mid(strAdciseSigns, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetAdviceTime(ByVal lng医嘱ID As Long, ByVal int类型 As Integer) As Date
'功能：读取医嘱指定操作的时间
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select Max(操作时间) as 时间 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, int类型)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetAdviceTime = rsTmp!时间
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMergeIDs(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_相关ID As Long, ByVal COL_ID As Long) As String
'功能：获取指定一并给药的医嘱ID串(非一并给药返回当前医嘱ID)
'参数：lngRow=一并给药的开始药品行
    Dim lng相关ID As Long, i As Long
    Dim str医嘱ID As String
    
    With vsAdvice
        lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                str医嘱ID = str医嘱ID & "," & Val(.TextMatrix(i, COL_ID))
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeIDs = Mid(str医嘱ID, 2)
End Function

Public Function GetRXKey(ByRef rsRXKey As ADODB.Recordset, ByVal strKey As String, ByVal str医嘱ID As String) As String
'功能：返回药品处方条数限制关键字,用于处方NO分配
'参数：strKey=当前处方NO的Key,不包含处方条数限制Key部份
'      str医嘱ID=当前药品的医嘱ID串，一并给药包含多个ID，"ID1,ID2,..."
'                一并给药开始行或独立药品行才传入,一并给药中间行传入空
    Dim intNextCount As Integer
    Dim strNextID As String
    
    rsRXKey.Filter = "Key='" & strKey & "'"
    If rsRXKey.EOF Then
        strNextID = zlStr.ListMinus(str医嘱ID, "")
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey.AddNew
        rsRXKey!Key = strKey
        rsRXKey!医嘱ID = strNextID
        rsRXKey!条数 = intNextCount
        rsRXKey!张数 = 1
        rsRXKey.Update
    ElseIf str医嘱ID <> "" Then
        strNextID = zlStr.ListMinus(str医嘱ID, rsRXKey!医嘱ID)
        intNextCount = UBound(Split(strNextID, ",")) + 1
        
        rsRXKey!医嘱ID = rsRXKey!医嘱ID & "," & strNextID
        rsRXKey!条数 = rsRXKey!条数 + intNextCount
        rsRXKey.Update
    
        If rsRXKey!条数 > gintRXCount Then
            strNextID = zlStr.ListMinus(str医嘱ID, "")
            intNextCount = UBound(Split(strNextID, ",")) + 1
            
            rsRXKey!张数 = rsRXKey!张数 + 1
            rsRXKey!医嘱ID = strNextID
            rsRXKey!条数 = intNextCount
            rsRXKey.Update
        End If
    ElseIf str医嘱ID = "" Then
        '一并给药中间行,保持第一行的关键字
    End If

    GetRXKey = rsRXKey!张数
End Function

Public Function GetMergeCount(ByRef vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal COL_相关ID As Long, ByVal COL_收费细目ID As Long) As Long
'功能：获取指定一并给药的药品种数数量(非一并给药返回1行)
'参数：lngRow=一并给药的开始药品行
    Dim lng相关ID As Long, i As Long
    Dim str药品ID As String
    
    With vsAdvice
        lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                If InStr(str药品ID & ",", "," & Val(.TextMatrix(i, COL_收费细目ID)) & ",") = 0 Then
                    str药品ID = str药品ID & "," & Val(.TextMatrix(i, COL_收费细目ID))
                End If
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeCount = UBound(Split(Mid(str药品ID, 2), ",")) + 1
End Function

Public Function GetAdviceState(ByVal lng医嘱ID As Long, ByVal vDate As Date) As Integer
'功能：读取医嘱在指定时点的医嘱状态(主要用于暂停启用,因为暂停启用的操作时间现是发生时间)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    strSQL = "Select 操作类型 From 病人医嘱状态 Where 操作类型<>10 And 医嘱ID=[1] And 操作时间<=[2] Order by 操作时间 Desc"
    strSQL = "Select 操作类型 From (" & strSQL & ") Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lng医嘱ID, vDate)
    If Not rsTmp.EOF Then
        GetAdviceState = rsTmp!操作类型
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistsSpecAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：判断病人本次住院是否已经下达了确认的特殊医嘱(出院，转院，死亡)
'返回：如果存在，返回医嘱提示信息。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select A.姓名,A.婴儿,A.医嘱内容 From 病人医嘱记录 A,诊疗项目目录 B" & _
        " Where A.诊疗项目ID=B.ID And A.诊疗类别='Z' And B.操作类型 IN('5','6','11')" & _
        " And A.医嘱状态 Not IN(1,2,4) And A.病人ID=[1] And A.主页ID=[2]" & _
        " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        strMsg = strMsg & vbCrLf & "　　●" & rsTmp!医嘱内容 & IIF(NVL(rsTmp!婴儿, 0) <> 0, "(婴儿" & NVL(rsTmp!婴儿, 0) & ")", "")
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        rsTmp.MoveFirst
        strMsg = "提醒您，病人""" & rsTmp!姓名 & """已经确认下达以下特殊医嘱：" & vbCrLf & strMsg & vbCrLf & ""
    End If
    ExistsSpecAdvice = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistBalance(ByVal strNO As String) As Boolean
'功能：判断指定的收费划价单是否存在已经收费的内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From 门诊费用记录 Where Mod(记录性质,10)=1 And 记录状态 IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "BillExistBalance", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人医嘱附件(ByVal lng医嘱ID As Long) As String
'功能：返回指定医嘱的附件描述串
'参数：lng医嘱ID=可见行的医嘱ID(除药品外，是相关ID为空的医嘱ID)
'返回：格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 项目,必填,要素ID,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order by 排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get病人医嘱附件", lng医嘱ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "<Split1>" & rsTmp!项目 & "<Split2>" & NVL(rsTmp!必填, 0) & "<Split2>" & NVL(rsTmp!要素ID) & "<Split2>" & NVL(rsTmp!内容)
        rsTmp.MoveNext
    Loop
    Get病人医嘱附件 = Mid(strSQL, Len("<Split1>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医嘱项目附件(ByVal lng项目id As Long, Optional ByVal int场合 As Integer = 2) As String
'功能：返回指定诊疗项目要求的附件描述串
'参数：lng项目ID=诊疗项目ID,int场合=1门诊，=2住院，=4体检
'返回：格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select C.项目,C.内容,C.要素ID,C.必填" & _
        " From 病历单据应用 A,病历文件列表 B,病历单据附项 C" & _
        " Where A.诊疗项目ID=[1] And A.应用场合=[2]" & _
        " And A.病历文件ID=B.ID And B.种类=7 And B.ID=C.文件ID" & _
        " Order by C.排列"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get医嘱项目附件", lng项目id, int场合)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        '最后多一个<Split2>，表明是直接读取的项目附项描述，而不是编辑产生的，用于修改附项时识别，见frmAdviceEditEx的修改处理
        strSQL = strSQL & "<Split1>" & rsTmp!项目 & "<Split2>" & NVL(rsTmp!必填, 0) & "<Split2>" & NVL(rsTmp!要素ID) & "<Split2>" & NVL(rsTmp!内容) & "<Split2>1"
        rsTmp.MoveNext
    Loop
    Get医嘱项目附件 = Mid(strSQL, Len("<Split1>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMedicineSended(ByVal lngAdviceID As Long, ByVal DateLast As Date) As Boolean
'功能：检查指定医嘱的最近一次发送是否已发药
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1" & vbNewLine & _
            "From 病人医嘱发送 A, 住院费用记录 B, 药品收发记录 C" & vbNewLine & _
            "Where a.医嘱id = [1] And a.末次时间 = [2] And a.No = b.No And" & vbNewLine & _
            "      a.记录性质 = b.记录性质 And a.医嘱id = b.医嘱序号" & _
            " And b.Id = c.费用id And c.单据 In (9, 10) And Mod(c.记录状态, 3) = 1 And c.审核人 Is Null And Rownum = 1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckMedicineSended", lngAdviceID, DateLast)
    CheckMedicineSended = rsTmp.RecordCount = 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastSendMediCineID(ByVal lngAdviceID As Long, ByVal DateLast As Date, ByVal lng病人性质 As Long) As Long
'功能：根据药品医嘱ID获取最近一次发送的费用对应的收费项目ID(药品规格ID)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select B.收费细目ID" & _
            " From 病人医嘱发送 A," & IIF(lng病人性质 = 1, "门诊", "住院") & "费用记录 B" & _
            " Where A.NO=B.NO And A.记录性质=B.记录性质" & _
            " And B.记录状态 IN(0,1,3) And B.医嘱序号=A.医嘱ID And A.医嘱ID=[1] And A.末次时间=[2]"
    '可能该药品未计费(如自备药)
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetLastSendMediCineID", lngAdviceID, DateLast)
    If rsTmp.RecordCount > 0 Then
        GetLastSendMediCineID = Val("" & rsTmp!收费细目ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowRollNotify(ByVal strPatis As String)
'功能：检查一组病人是否存在已停止的需要超期收回的医嘱，并进行提示
'参数：strRollNotify=病人ID:主页ID,...
    Dim rsTmp As ADODB.Recordset, rsDrug As ADODB.Recordset
    Dim strSQL As String, strMsg As String, strSQLPati As String, strTemp As String
    Dim strThis As String, p As Long, n As Long, strUnRoll As String, blnDo As Boolean, lng药品ID As Long
    Dim varPar(0 To 10) As String

    On Error GoTo errH
    strUnRoll = zlDatabase.GetPara("发药后不收回", glngSys, p住院医嘱发送)
    strTemp = "Select C1 As 病人ID,C2 As 主页ID From Table(f_Num2list2([1]))"
    n = 0
    Do While True
        If Len(strPatis) < 4000 Then
            p = Len(strPatis) + 1
        Else
            p = InStrRev(Mid(strPatis, 1, 4000), ",")
        End If
        strThis = Mid(strPatis, 1, p - 1)
        
        If n > 10 Then
            strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 1) & "]")
        End If
        
        n = n + 1
        strPatis = Mid(strPatis, p + 1)
        If strPatis = "" Then Exit Do
    Loop
    
    '条件与超期收回中一致，但只包含当前状态为(自动)停止的。
    strSQL = "(A.执行时间方案 is NULL And (Nvl(A.频率次数,0)=0 Or Nvl(A.频率间隔,0)=0 Or A.频率间隔 is NULL))"
    strSQL = _
        " Select /*+ Rule*/ A.姓名,A.医嘱内容,A.ID,A.诊疗类别,A.上次执行时间,A.收费细目ID,b.病人性质" & _
        " From 病人医嘱记录 A,病案主页 B,诊疗项目目录 E,(" & strSQLPati & ") F" & _
        " Where A.诊疗项目ID=E.ID And a.病人id=b.病人id and a.主页id=b.主页id And A.病人ID = F.病人ID And A.主页ID = F.主页ID" & _
        " And Not(A.诊疗类别='H' And E.操作类型='1') And Not(A.诊疗类别='Z' And E.操作类型 In('4','14'))" & _
        " And Nvl(A.执行性质,0)<>0 And A.总给予量 is NULL And Nvl(A.医嘱期效,0)=0" & _
        " And ((Not " & strSQL & " And A.执行终止时间<A.上次执行时间)" & _
        " Or (" & strSQL & " And Trunc(A.执行终止时间)<Trunc(A.上次执行时间)+1))" & _
        " And A.医嘱状态=8 And (A.相关ID is Null Or A.诊疗类别 IN('5','6'))" & _
        " And A.开始执行时间 is Not NULL And A.病人来源<>3  And NVL(a.执行频次,'无')<>'必要时' And NVL(a.执行频次,'无')<>'需要时'" & _
        " And Not Exists(Select 1 From 病人医嘱记录 X Where 诊疗类别 IN('5','6') And X.相关ID=A.ID)" & _
        " Order by A.病人ID,A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "超期收回检查", varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
    Do While Not rsTmp.EOF
        blnDo = True
        
        If InStr(",5,6,", rsTmp!诊疗类别) > 0 And strUnRoll <> "" Then
            If Not IsNull(rsTmp!收费细目ID) Then
                lng药品ID = rsTmp!收费细目ID
            Else
                lng药品ID = GetLastSendMediCineID(Val(rsTmp!ID), CDate(rsTmp!上次执行时间), Val(rsTmp!病人性质 & ""))
            End If
            If lng药品ID <> 0 Then
                strSQL = "Select 发药类型 From 药品规格 Where 药品ID = [1] And 发药类型 is Not Null"
                Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "超期收回检查", lng药品ID)
                If rsDrug.RecordCount > 0 Then
                    If InStr("," & strUnRoll & ",", "," & rsDrug!发药类型 & ",") > 0 Then
                        If CheckMedicineSended(Val(rsTmp!ID), CDate(rsTmp!上次执行时间)) Then
                            blnDo = False
                        End If
                    End If
                End If
            Else '无需收回：医嘱未记费（如自备药）或相关费用被删除了（如划价单被删除）
                blnDo = False
            End If
        End If
        If blnDo Then
            strMsg = strMsg & vbCrLf & "●　病人：" & rsTmp!姓名 & "　医嘱：" & rsTmp!医嘱内容
        End If
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        MsgBox "下列已停止的医嘱被超期发送：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "该类医嘱可以使用""超期发送收回""进行处理。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckPathNotEvaluete(ByVal lng病人ID As Long, lng主页ID As Long, Optional ByRef blnIsSend As Boolean, Optional ByRef str日期 As String) As Boolean
'功能：检查路径病人当前是否未评估
'参数：blnIsSend  当天是否已经生成
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select b.日期,sysdate As 当前日期" & vbNewLine & _
            "From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And a.Id = b.路径记录id And a.当前阶段id = b.阶段id And a.当前天数 = b.天数 And Rownum = 1 And" & vbNewLine & _
            "      Not Exists" & vbNewLine & _
            "(Select 1 From 病人路径评估 C Where c.路径记录id = a.Id And c.阶段id = a.当前阶段id And c.天数 = a.当前天数)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取路径生成", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        str日期 = Format(rsTmp!日期 & "", "yyyy-MM-dd") & ""
        If Format(rsTmp!日期 & "", "yyyy-MM-dd") = Format(rsTmp!当前日期 & "", "yyyy-MM-dd") Then
            blnIsSend = True '
        Else
            blnIsSend = False
        End If
        CheckPathNotEvaluete = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathNotEvalueteOut(ByVal lng挂号ID As Long, Optional ByRef blnIsSend As Boolean, Optional ByRef str日期 As String) As Boolean
'功能：检查路径病人当前是否未评估
'参数：blnIsSend  当天是否已经生成
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select b.日期,sysdate As 当前日期" & vbNewLine & _
            "From 病人门诊路径 A, 病人门诊路径执行 B, 病人门诊路径记录 C" & vbNewLine & _
            "Where C.挂号ID=[1] And C.路径记录ID=A.ID And a.Id = b.路径记录id And a.当前阶段id = b.阶段id And a.当前天数 = b.天数 And Rownum = 1 And" & vbNewLine & _
            "      Not Exists" & vbNewLine & _
            "(Select 1 From 病人门诊路径评估 d Where d.路径记录id = a.Id And d.阶段id = a.当前阶段id And d.天数 = a.当前天数)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取路径生成", lng挂号ID)
    If rsTmp.RecordCount > 0 Then
        str日期 = Format(rsTmp!日期 & "", "yyyy-MM-dd") & ""
        If Format(rsTmp!日期 & "", "yyyy-MM-dd") = Format(rsTmp!当前日期 & "", "yyyy-MM-dd") Then
            blnIsSend = True '
        Else
            blnIsSend = False
        End If
        CheckPathNotEvalueteOut = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathItemIsMust(ByVal byt执行方式 As Byte, ByVal int天数 As Integer, ByVal lng路径记录Id As Long, _
                                    ByVal lng阶段Id As Long, ByVal lng项目id As Long, Optional ByVal intType As Integer) As Boolean
'功能:检查路径项目是否是必须生成的项目
    Dim blnMust As Boolean
    Dim strSQL As String
    Dim rsStep As ADODB.Recordset
    
    On Error GoTo errH:
    If byt执行方式 = 1 Then
        blnMust = True
    ElseIf byt执行方式 = 2 Or byt执行方式 = 4 Then  '至少一次或必须一次
        strSQL = "Select 开始天数,结束天数 From " & IIF(intType = 1, "门诊路径阶段", "临床路径阶段") & " Where ID = [1]"
        Set rsStep = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng阶段Id)
        If Not IsNull(rsStep!开始天数) Then
            If Not IsNull(rsStep!结束天数) Then
                blnMust = (int天数 = Val("" & rsStep!结束天数))
                If blnMust Then   '是否最后一天
                    '判断该项目之前有没有执行过(路径外项目除外)
                    strSQL = "Select 1 From " & IIF(intType = 1, "病人门诊路径执行", "病人路径执行") & " Where 路径记录ID = [1] And 阶段ID = [2] And 项目ID = [3] And 天数<[4] And rownum<2"
                    Set rsStep = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", lng路径记录Id, lng阶段Id, lng项目id, int天数)
                    If rsStep.RecordCount > 0 Then blnMust = False
                End If
            Else
                blnMust = True  '单天
            End If
        End If
    End If
    
    CheckPathItemIsMust = blnMust
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckAdviceAppend(ByVal strAppend As String) As String
'功能：对指定医嘱的申请附项填写情况进行检查
'参数：strAppend=当前申请附项的填写情况串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
'返回：必须要填写的申请附项内容，如"项目1,项目2..."
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    
    arrItem = Split(strAppend, "<Split1>")
    For i = 0 To UBound(arrItem)
        arrSub = Split(arrItem(i), "<Split2>")
        If Val(arrSub(1)) = 1 And Trim(arrSub(3)) = "" Then
            strItem = strItem & "," & arrSub(0)
        End If
    Next
    
    CheckAdviceAppend = Mid(strItem, 2)
End Function

Public Function ReplaceAppend(ByVal strTarget As String, ByVal strSource As String) As String
'功能：用已输入的申请附项内容，对指定医嘱的相同、空白附项进行缺省替换
'参数：strTarget=要被替换的医嘱附项描述串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
'      strSource=已输入的医嘱附项描述串,格式相同。
'返回：已被替换后医嘱附项描述串，格式相同。返回值可能与strTarget相同。
    Dim arrTarget As Variant, arrSub1 As Variant
    Dim arrSource As Variant, arrSub2 As Variant
    Dim i As Integer, j As Integer
    Dim strReturn As String, blnReplace As Boolean
    
    If strTarget = "" Or strSource = "" Then
        ReplaceAppend = strTarget: Exit Function
    End If
    
    arrTarget = Split(strTarget, "<Split1>")
    arrSource = Split(strSource, "<Split1>")
    
    For i = 0 To UBound(arrTarget)
        arrSub1 = Split(arrTarget(i), "<Split2>")
        
        blnReplace = False
        For j = 0 To UBound(arrSource)
            arrSub2 = Split(arrSource(j), "<Split2>")
            If arrSub1(0) = arrSub2(0) Then
                If arrSub1(3) = "" And arrSub2(3) <> "" Then
                    arrSub1(3) = arrSub2(3): blnReplace = True
                End If
                Exit For
            End If
        Next
        
        strReturn = strReturn & "<Split1>" & arrSub1(0) & "<Split2>" & arrSub1(1) & "<Split2>" & arrSub1(2) & "<Split2>" & arrSub1(3)
        If UBound(arrSub1) >= 4 Then
            '被自动替换了的，相当于取有缺省值，修改时不再识别，见"Get医嘱项目附件"函数
            If Not blnReplace Then strReturn = strReturn & "<Split2>" & arrSub1(4)
        End If
    Next
    
    ReplaceAppend = Mid(strReturn, Len("<Split1>") + 1)
End Function

Public Function GetMaxDate(lng病人ID As Long, lng主页ID As Long, Optional int原因 As Integer) As Date
'功能：获取转科病人最大的上次变动时间
'参数：int原因=返回上次变动的原因
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    int原因 = 0
    
    strSQL = "Select 开始时间,开始原因 From 病人变动记录" & _
        " Where 开始时间 is Not NULL And 终止时间 is NULL" & _
        " And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIF(IsNull(rsTmp!开始时间), GetMaxDate, rsTmp!开始时间)
        int原因 = NVL(rsTmp!开始原因, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBabyRegList(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, Optional ByRef rsBaby As ADODB.Recordset) As String
'功能：读取病人的婴儿姓名列表
'参数：lng就诊ID=住院病人为"主页ID",门诊病人为"挂号ID"
'返回："姓名1<Split>姓名2<Split>姓名3..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.序号,a.婴儿姓名,a.婴儿性别 as 性别 From 病人新生儿记录 a Where a.病人ID=[1] And a.主页ID=[2] Order by a.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetBabyRegList", lng病人ID, lng就诊ID)
    Set rsBaby = zlDatabase.CopyNewRec(rsTmp)
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "<Split>" & NVL(rsTmp!婴儿姓名)
        rsTmp.MoveNext
    Loop
    GetBabyRegList = Mid(strSQL, Len("<Split>") + 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdvicePrintPage(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, _
    ByVal int期效 As Integer, ByVal lng页号 As Long) As Long
'功能：根据指定的医嘱打印页号，获取与该页一起跨页打印的前面页号
'返回：与当前页一起跨页打印的起始页号，可能是前几页，但是最近次重整后的打印页
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "(Select Nvl(Max(页号),0) From 病人医嘱打印 Where 病人ID=[1] And 主页ID=[2] And Nvl(婴儿, 0)=[3] And Nvl(期效,0)=[4] And 医嘱ID Is Null)"
    strSQL = "Select D.页号 From 病人医嘱打印 A,病人医嘱记录 B,病人医嘱记录 C,病人医嘱打印 D" & _
        " Where A.医嘱ID=B.ID And B.相关ID=C.相关ID And C.ID=D.医嘱ID And D.页号=[5]-1 And D.页号>=" & strSQL & _
        " And A.病人ID=[1] And A.主页ID=[2] And Nvl(A.婴儿,0)=[3] And Nvl(A.期效,0)=[4] And A.页号=[5] And A.行号=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdvicePrintPage", lng病人ID, lng主页ID, int婴儿, int期效, lng页号)
    If Not rsTmp.EOF Then
        GetAdvicePrintPage = GetAdvicePrintPage(lng病人ID, lng主页ID, int婴儿, int期效, rsTmp!页号)
    Else
        GetAdvicePrintPage = lng页号
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Open_LIS_Report(ByVal frmParent As Object, ByVal lng医嘱ID As Long, ByVal lng病人ID As Long, ByVal blnCurrMoved As Boolean, ByVal blnPrint As Boolean, Optional ByVal bln禁预览打印 As Boolean) As Boolean
'调用LiwWork打印带图形的LIS报表
'参数：bln禁预览打印 在调用报表预览时是否显示预览界面的打印按钮，false 要显示，true 不显示
    Dim strChart(0 To 8) As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim intLoop As Integer
    Dim objLisWork As Object
    Dim lng发送号 As Long, lng标本id As Long
                    
    On Error GoTo errHandle
    Set objLisWork = CreateObject("zl9LisWork.clsLISImg")
    
    strSQL = "select 发送号 from 病人医嘱发送 a,病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lng医嘱ID)
    If Not rsTmp.EOF Then
        lng发送号 = NVL(rsTmp!发送号, 0)
    End If
    strSQL = "select max(标本ID) as ID from 检验项目分布 where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lng医嘱ID)
    If Not rsTmp.EOF Then
        lng标本id = NVL(rsTmp!ID, 0)
    End If
    If lng标本id = 0 Then
        strSQL = "select ID from 检验标本记录 b where b.医嘱id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lng医嘱ID)
        If Not rsTmp.EOF Then
            lng标本id = NVL(rsTmp!ID, 0)
        End If
    End If
    If lng发送号 = 0 Or lng标本id = 0 Then Exit Function
    
    strSQL = "select id from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng标本id)
    intLoop = 0
    Do Until rsTmp.EOF
        If Not objLisWork Is Nothing Then
            If objLisWork.Get_Chart2d_File(App.Path, rsTmp("ID")) Then
                strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
            End If
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If Not objLisWork Is Nothing Then
        If objLisWork.Get_ReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
            If Not blnPrint And bln禁预览打印 Then
                strTmp = "DisabledPrint=1"
            Else
                strTmp = "DisabledPrint=0"
            End If
            Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                            "病人ID=" & lng病人ID, "标本ID=" & lng标本id, _
                            "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                            "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                            "图形9=" & strChart(8), strTmp, IIF(blnPrint, 2, 1))
        End If
    End If
    '删除图形文件
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If Dir(strChart(intLoop)) <> "" Then Kill strChart(intLoop)
        End If
    Next
    
    Open_LIS_Report = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCuvetteNumber(rsNumber As ADODB.Recordset, ByVal str管码 As String, ByVal lng医嘱ID As Long, _
    ByVal lng相关ID As Long, ByVal str类别 As String, ByVal int操作类型 As Integer, ByVal lng执行科室ID As Long, _
    ByVal int婴儿 As Integer, ByVal lng诊疗项目ID As Long, ByVal int紧急 As Integer, ByVal str标本 As String, ByVal lng采集科室ID As Long) As String
    '功能：对检验医嘱生成样本条码
    '      1.一并采集的同一检验医嘱使用相同的样本条码
    '      2.相同管码的检验使用相同的样本条码
    '      3.校本条码规则:12位的"管码+医嘱ID"
    '参数：rsNumber=动态记录集，具有"管码、相关ID、样本条码"等字段
    Dim strTmp管码 As String, strTmp条码 As String
    
    If str类别 = "C" And str管码 <> "" Then '检验项目才有管码
        rsNumber.Filter = "相关ID=" & lng相关ID
        If rsNumber.EOF Then
            rsNumber.Filter = "诊疗项目id=" & lng诊疗项目ID
            If rsNumber.EOF Then
                rsNumber.Filter = "管码='" & str管码 & "' And 执行科室ID=" & lng执行科室ID & " And 婴儿=" & int婴儿 & _
                    " And 紧急标志=" & int紧急 & " And 标本='" & str标本 & "' And 采集科室ID=" & lng采集科室ID
                If rsNumber.EOF Then
                    '生成新的条码
                    rsNumber.AddNew
                    rsNumber!管码 = str管码
                    rsNumber!相关ID = lng相关ID
'                    rsNumber!样本条码 = str管码 & Format(lng医嘱ID, Replace(Space(12 - Len(str管码)), " ", "0"))
                    rsNumber!样本条码 = zlDatabase.GetNextNo(125, lng医嘱ID)
                    rsNumber!诊疗项目ID = lng诊疗项目ID
                    rsNumber!执行科室ID = lng执行科室ID
                    rsNumber!婴儿 = int婴儿
                    rsNumber!紧急标志 = int紧急
                    rsNumber!标本 = str标本
                    rsNumber!采集科室ID = lng采集科室ID
                    rsNumber.Update
                    
                    strTmp条码 = rsNumber!样本条码
                Else
                    '相同管码、执行科室、婴儿的检验使用相同的样本条码
                    strTmp管码 = NVL(rsNumber!管码)
                    strTmp条码 = NVL(rsNumber!样本条码)
                    
                    rsNumber.AddNew
                    rsNumber!管码 = strTmp管码
                    rsNumber!相关ID = lng相关ID
                    rsNumber!样本条码 = strTmp条码
                    rsNumber!诊疗项目ID = lng诊疗项目ID
                    rsNumber!执行科室ID = lng执行科室ID
                    rsNumber!婴儿 = int婴儿
                    rsNumber!紧急标志 = int紧急
                    rsNumber!标本 = str标本
                    rsNumber!采集科室ID = lng采集科室ID
                    rsNumber.Update
                End If
            Else
                '生成新的条码：相同检验的医嘱使用"不同的"条码
                rsNumber.AddNew
                rsNumber!管码 = str管码
                rsNumber!相关ID = lng相关ID
'                rsNumber!样本条码 = str管码 & Format(lng医嘱ID, Replace(Space(12 - Len(str管码)), " ", "0"))
                rsNumber!样本条码 = zlDatabase.GetNextNo(125, lng医嘱ID)
                rsNumber!诊疗项目ID = lng诊疗项目ID
                rsNumber!执行科室ID = lng执行科室ID
                rsNumber!婴儿 = int婴儿
                rsNumber!紧急标志 = int紧急
                rsNumber!标本 = str标本
                rsNumber!采集科室ID = lng采集科室ID
                rsNumber.Update
                
                strTmp条码 = rsNumber!样本条码
            End If
        Else
            '一并采集的检验项目使用相同的条码
            strTmp管码 = NVL(rsNumber!管码)
            strTmp条码 = NVL(rsNumber!样本条码)
            
            rsNumber.AddNew
            rsNumber!管码 = strTmp管码
            rsNumber!相关ID = lng相关ID
            rsNumber!样本条码 = strTmp条码
            rsNumber!诊疗项目ID = lng诊疗项目ID
            rsNumber!执行科室ID = lng执行科室ID
            rsNumber!婴儿 = int婴儿
            rsNumber!紧急标志 = int紧急
            rsNumber!标本 = str标本
            rsNumber!采集科室ID = lng采集科室ID
            rsNumber.Update
        End If
    ElseIf str类别 = "E" And int操作类型 = 6 Then
        '采集方式使用与医嘱相同(最近)的条码
        If Not rsNumber.EOF Then
            If NVL(rsNumber!相关ID, 0) = lng医嘱ID Then
                strTmp条码 = NVL(rsNumber!样本条码)
            End If
        End If
    End If
    
    GetCuvetteNumber = strTmp条码
End Function

Public Function GetExecDays(ByVal str分解时间 As String) As ADODB.Recordset
'功能：根据当前医嘱的执行时间串返回不重复的执行天数记录集
    Dim rsTmp As ADODB.Recordset
    Dim arrTmp As Variant, i As Long, strTmp As String
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "收费时间", adVarChar, 10
    rsTmp.Fields.Append "存在", adInteger '用于决定是否加入已存在的列表
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    arrTmp = Split(str分解时间, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = Format(arrTmp(i), "yyyy-MM-dd")
        rsTmp.Filter = "收费时间='" & strTmp & "'"
        If rsTmp.EOF Then
            rsTmp.AddNew
            rsTmp!收费时间 = strTmp
            rsTmp!存在 = 0
            rsTmp.Update
        End If
    Next
    rsTmp.Filter = ""
    Set GetExecDays = rsTmp
End Function


Public Function AdviceMoneyMake(ByVal lng病人ID As Long, ByVal lng主页ID As Long, rsMoneyNow As Recordset, rsMoneyDay As ADODB.Recordset, _
    ByVal lng医嘱ID As Long, ByVal lng诊疗项目ID, ByVal lng收费项目ID As Long, ByVal lng执行部门id As Long, ByVal str试管编码 As String, _
    ByVal str收费类别 As String, ByVal int收费方式 As Integer, ByVal str分解时间 As String, ByVal byt来源 As Byte, ByRef lng费用次数 As Long, ByVal dbl总量 As Double, _
    Optional ByVal lng当前医嘱ID As Long, Optional ByVal lng发送号 As Long, Optional ByVal dbl计价数量 As Double, Optional rsExec As Recordset, _
    Optional ByVal lng计算方式 As Long, Optional ByVal str频率 As String, Optional ByVal dbl单量 As Double, Optional ByVal int期效 As Integer = 1, _
    Optional ByVal int费用性质 As Integer, Optional ByVal str诊疗类别 As String, Optional ByVal str样本条码 As String, Optional ByVal str部位方法 As String, Optional ByRef dbl收费数量 As Double, Optional ByVal strMinDate As String) As Boolean
'功能：判断指定的医嘱费用是否应该产生
'参数：lng主页ID=住院病人才使用，门诊病人传入0不分具体挂号
'      rsMoneyNow=当前病人本次要发送的费用,动态记录集(收费方式=-1,表示首次不收时，一天只收一次的项目的记录)
'      rsMoneyDay=当前病人当天已发送的费用,静态记录集
'      lng医嘱ID=一组医嘱的ID
'      str分解时间=本次发送的执行时间串，以逗号分隔，并且排除了暂停的时间点
'      byt来源:1-门诊，2-住院
'      dbl计价数量=收费项目的计价数量
'      其他=当前行发送医嘱及费用信息
'      lng当前医嘱ID=当前行医嘱id
'      str样本条码=检验医嘱传入样本条码
'      str部位方法 检查项目医嘱子医嘱的检查部位和方法 固定格式，检查部位<sTab>检查方法，如："头部<sTab>平扫"
'      dbl收费数量  收费数次，外挂决定如果外挂不指定则为0
'      strMinDate   查询已发送的医嘱中收费情况时的最小时间
'以下是计量计时医嘱的数量组织规则
'1、长嘱可选频率、持续性、必要时和不定时以单量作为数次。
'2、临嘱一次性和需要时频率的医嘱取总量作为数次。
'3、临嘱可选频率取单量作为数次，最后一次取总量除以单量取末作为数次，例如红外照射治疗，总量80、单量25，每天4次，那么执行登记时，供执行四次，前三次本次数次为25，第四次为80除以25取模=5。
'4、批量执行登记页面医嘱清单单量后新增列：本次数次，用于显示本次数次。
'5、医嘱编辑时不允许录入首次用量。
'返回：
'      lng费用次数=一天只收一次时（3,4,5,6,7），返回本次发送要收取的次数
'      dbl总量=总的发送次数或数量
'      rsExec=医嘱执行计价的内容
    Dim lng材料ID As Long, blnMakeMoney As Boolean
    Dim rsDays As ADODB.Recordset, i As Long
    Dim arrTmp As Variant
    Dim dbl数量 As Double
    Dim strDate As String
    Dim dbl总量Tmp As Double
    Dim strSQL As String, rsTmp As Recordset, strTmp As String
    Dim str部位 As String, str方法 As String
    Dim blnTmp As Boolean
    Dim dblTmp收费数量 As Double
    Dim str最小日期 As String
    
    blnMakeMoney = True
    lng费用次数 = 1
    dbl收费数量 = 0
    
    If str部位方法 <> "" Then
        str部位 = Split(str部位方法, "<sTab>")(0)
        str方法 = Split(str部位方法, "<sTab>")(1)
    End If
    
    If int收费方式 = 9 Then
        '自定义
        '调用外挂接口，接口出错当作正常收取
        strTmp = lng医嘱ID & "<sTab>" & lng当前医嘱ID & "<sTab>" & lng诊疗项目ID & "<sTab>" & lng收费项目ID & "<sTab>" & int收费方式 & "<sTab>" & str部位 & "<sTab>" & str方法 & "<sTab>" & (dbl总量 * dbl计价数量)
        dblTmp收费数量 = -1
        On Error Resume Next
        Call CreatePlugInOK(IIF(byt来源 = 1, p门诊医嘱下达, p住院医嘱发送), -1)
        If Not gobjPlugIn Is Nothing Then
            blnTmp = gobjPlugIn.AdviceMakeFee(glngSys, IIF(byt来源 = 1, p门诊医嘱下达, p住院医嘱发送), strTmp, rsMoneyNow, dblTmp收费数量)
            Call zlPlugInErrH(err, "AdviceMakeFee")
            If err.Number <> 0 Then
                '接口出错了
                dblTmp收费数量 = -1
            Else
                If blnTmp Then
                    If dblTmp收费数量 = 0 Then
                        '如果收数量为0了则认为本次不用收费用
                        blnMakeMoney = False
                    End If
                    If dblTmp收费数量 > 0 Then
                        dbl收费数量 = dblTmp收费数量
                    End If
                End If
            End If
        End If
        err.Clear: On Error GoTo 0
        
        On Error GoTo errH
        
        If dblTmp收费数量 = -1 Then
            strSQL = "Select zl_fun_CustomExpenses([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17]) as 返回结果 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "AdviceMoneyMake", lng病人ID, lng主页ID, byt来源, lng当前医嘱ID, lng医嘱ID, int期效, str频率, lng诊疗项目ID, lng收费项目ID, _
                                                lng执行部门id, str诊疗类别, str收费类别, dbl总量, dbl单量, dbl计价数量, int费用性质, lng计算方式)
            If rsTmp.RecordCount > 0 Then
                strTmp = rsTmp!返回结果 & ""
                If Val(Split(strTmp, ":")(0)) = 0 Then
                    '不收取
                    blnMakeMoney = False
                Else
                    '要收取
                    If InStr(strTmp, ":") > 0 Then
                        If Val(Split(strTmp, ":")(1)) > 0 Then lng费用次数 = Val(Split(strTmp, ":")(1)): dbl收费数量 = lng费用次数
                    End If
                End If
            End If
        End If
    End If
    
    If int收费方式 = 0 Then
        '正常收费的，检查在本次发送中、本医嘱中是否被排斥
        rsMoneyNow.Filter = "(医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=5)" & _
            " Or (医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=6)" 'Or的使用
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf int收费方式 = 1 Then '检验试管费用(一次发送只收取一次)
        If str试管编码 <> "" Then
            '相同条码(试管)只收取一次
            rsMoneyNow.Filter = "试管编码='" & str试管编码 & "' And 样本条码='" & str样本条码 & "' And 收费项目ID=" & lng收费项目ID & " And 收费方式<>-1"
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
            
            '只收取试管对应的卫材费用
            If blnMakeMoney And str收费类别 = "4" Then
                lng材料ID = GetTubeMaterial(str试管编码)
                If lng材料ID <> 0 And lng收费项目ID <> lng材料ID Then blnMakeMoney = False
            End If
        End If
    ElseIf int收费方式 = 2 Then '一次发送只收取一次
        rsMoneyNow.Filter = "诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & " And 收费方式<>-1"
        If Not rsMoneyNow.EOF Then blnMakeMoney = False
    ElseIf InStr(",3,4,5,6,7,", int收费方式) > 0 Then
        '3-当天只收取一次；4-当天未执行收取一次；5-当天只收取一次，排斥其他项目；6-当天未执行收取一次，排斥其他项目
        
        '正常收费的，检查在本次发送中、本医嘱中是否被排斥
        If int收费方式 = 7 Then
            rsMoneyNow.Filter = "(医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=5)" & _
                " Or (医嘱ID=" & lng医嘱ID & " And 诊疗项目ID=" & lng诊疗项目ID & " And 收费方式=6)" 'Or的使用
            If Not rsMoneyNow.EOF Then blnMakeMoney = False
        End If
        
        If blnMakeMoney Then
            Set rsDays = GetExecDays(str分解时间)
                        
            '先从本次发送中的找(频率为一天一次且没有收的，判断时当成已收取,以便后续的其他医嘱"首次不收"时不再认为有首次)
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.Filter = "收费时间='" & rsDays!收费时间 & "' And 诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & _
                    IIF(int收费方式 = 7, "", " And 收费方式<>-1") & _
                    IIF((int收费方式 = 4 Or int收费方式 = 6) And lng执行部门id <> 0, " And 执行部门ID=" & lng执行部门id, "")
                If rsMoneyNow.RecordCount > 0 Then rsDays!存在 = 1
                rsDays.MoveNext
            Next
            '再从已发送中的找(当天及将来执行的)
            rsDays.Filter = "存在=0"
            For i = 1 To rsDays.RecordCount
                If i = 1 Then
                    If rsMoneyDay Is Nothing Then
                        If strMinDate = "" Then
                            str最小日期 = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd")
                        Else
                            str最小日期 = Format(strMinDate, "yyyy-MM-dd")
                        End If
                        Call GetPatiDayMoneyDetail(rsMoneyDay, lng病人ID, lng主页ID, byt来源, CDate(str最小日期))
                    End If
                End If
                rsMoneyDay.Filter = "收费时间='" & rsDays!收费时间 & "' And 诊疗项目ID=" & lng诊疗项目ID & " And 收费项目ID=" & lng收费项目ID & _
                    IIF(int收费方式 = 7, "", " And 收费方式<>-1") & _
                    IIF((int收费方式 = 4 Or int收费方式 = 6) And lng执行部门id <> 0, " And 执行否=0 And 执行部门ID=" & lng执行部门id, "")
                If rsMoneyDay.RecordCount > 0 Then rsDays!存在 = 1
                rsDays.MoveNext
            Next
        End If
    End If
                            
    '记录到本次发送明细项目记录中
    If InStr(",3,4,5,6,7,", int收费方式) > 0 Then
        If int收费方式 = 7 Then
            If blnMakeMoney Then
                rsDays.Filter = "存在=0"    '没收过的那些天(频率为一天一次但未收的当成收过了)，首次不收
                lng费用次数 = dbl总量 - rsDays.RecordCount
                blnMakeMoney = lng费用次数 > 0
            End If
        Else
            rsDays.Filter = "存在=0"
            blnMakeMoney = rsDays.RecordCount > 0
            lng费用次数 = rsDays.RecordCount    '一天一次，有多少天要收就有多少次
        End If
        If blnMakeMoney Or int收费方式 = 7 And lng费用次数 = 0 Then
            For i = 1 To rsDays.RecordCount
                rsMoneyNow.AddNew
                rsMoneyNow!医嘱ID = lng医嘱ID
                rsMoneyNow!诊疗项目ID = lng诊疗项目ID
                rsMoneyNow!收费项目ID = lng收费项目ID
                rsMoneyNow!试管编码 = str试管编码
                rsMoneyNow!样本条码 = str样本条码
                
                '首次不收时，如果频率为一天一次，则计算后的费用次数为0,为了让本次后续发送的其他医嘱正确计算首是否收取，需要产生记录，但收费方式特殊记录为-1
                rsMoneyNow!收费方式 = IIF(int收费方式 = 7 And lng费用次数 = 0, -1, int收费方式)
                rsMoneyNow!收费时间 = rsDays!收费时间
                rsMoneyNow!执行部门ID = lng执行部门id
                rsMoneyNow.Update
            
                rsDays.MoveNext
            Next
        End If
    ElseIf blnMakeMoney Then
        rsMoneyNow.AddNew
        rsMoneyNow!医嘱ID = lng医嘱ID
        rsMoneyNow!诊疗项目ID = lng诊疗项目ID
        rsMoneyNow!收费项目ID = lng收费项目ID
        rsMoneyNow!试管编码 = str试管编码
        rsMoneyNow!样本条码 = str样本条码
        rsMoneyNow!收费方式 = int收费方式
        '检查项目专用配合外挂使用
        rsMoneyNow!子医嘱ID = lng当前医嘱ID
        rsMoneyNow!检查部位 = str部位
        rsMoneyNow!检查方法 = str方法
        rsMoneyNow!数量 = IIF(dbl收费数量 > 0, dbl收费数量, dbl总量 * dbl计价数量)
        
        If str分解时间 <> "" Then
            rsMoneyNow!收费时间 = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd")  '此时间暂时没有用处
        Else
            rsMoneyNow!收费时间 = ""
        End If
        rsMoneyNow!执行部门ID = lng执行部门id
        rsMoneyNow.Update
    End If
    '读取医嘱执行计价(除药品卫材医嘱外的才存储)
    If InStr(",5,6,7,", "," & str诊疗类别 & ",") = 0 Then
        If str分解时间 <> "" And Not rsExec Is Nothing Then
            arrTmp = Split(str分解时间, ",")
            If dbl收费数量 = 0 Then
                dbl总量Tmp = dbl总量 * dbl计价数量
            Else
                '如果被外挂处理了，这里应该用新的总量来计算
                dbl总量Tmp = dbl收费数量
            End If
            If dbl单量 = 0 Then dbl单量 = 1
            For i = 0 To UBound(arrTmp)
                rsExec.AddNew
                rsExec!医嘱ID = lng当前医嘱ID
                rsExec!发送号 = lng发送号
                rsExec!要求时间 = Format(arrTmp(i), "yyyy-MM-dd HH:mm:ss")
                rsExec!收费细目ID = lng收费项目ID
                rsExec!费用性质 = int费用性质
                If blnMakeMoney Then
                    '卫材也可以输入单量总量
                    If str频率 <> "" And ((lng计算方式 = 0 Or lng计算方式 = 3) And dbl总量 > 0 Or lng计算方式 = 1 Or lng计算方式 = 2 Or str诊疗类别 = "4") Then
                        '计量和计时的需要乘以数次
                        If int期效 = 0 Then
                            '1、长嘱可选频率、持续性、必要时和不定时以单量作为数次。
                            dbl数量 = dbl计价数量 * dbl单量
                        Else
                            '3、临嘱可选频率取单量作为数次，最后一次剩余的数量，例如红外照射治疗，总量80、单量25，每天4次，那么执行登记时，供执行四次，前三次本次数次为25，第四次为80除以25取模=5。
                            '门诊有可能没有录入执行时间,分解时间就只有一个，按总量作为次数
                            If UBound(arrTmp) = 0 Then
                                If InStr(",1,2,3,4,5,6,7,9,", int收费方式) > 0 Then
                                    '特殊收费方式总量按  lng费用次数 来计算和医嘱发送窗口中的费用记录数量保持一致
                                    If dbl收费数量 > 0 Then
                                        dbl数量 = dbl收费数量
                                    Else
                                        dbl数量 = lng费用次数 * dbl计价数量
                                    End If
                                Else
                                    dbl数量 = dbl总量Tmp
                                End If
                            Else
                                If i = UBound(arrTmp) Then
                                    dbl数量 = dbl总量Tmp
                                Else
                                    If dbl总量Tmp >= dbl单量 Then
                                        dbl数量 = dbl计价数量 * dbl单量
                                    Else
                                        dbl数量 = dbl总量Tmp
                                    End If
                                    dbl总量Tmp = dbl总量Tmp - dbl数量
                                End If
                            End If
                        End If
                    Else
                        dbl数量 = dbl计价数量
                    End If
                    If i <> 0 Then
                        strDate = Format(arrTmp(i - 1), "yyyy-MM-dd")
                    End If
                    '一次发送收取一次，则只有第一次收取
                    If InStr(",1,2,", int收费方式) > 0 Then
                        If i <> 0 Then dbl数量 = 0
                    ElseIf InStr(",3,4,5,6,", int收费方式) > 0 Then
                        '3456当天只收取一次的，存在=0的收取，默认第一次有数量
                        rsDays.Filter = "存在=0 And 收费时间='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If Not (rsDays.RecordCount > 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate) Then
                            dbl数量 = 0
                        End If
                    ElseIf int收费方式 = 7 Then
                        '当天首次不收取的，存在=1就收取，存在=0的为首次
                        rsDays.Filter = "存在=1 And 收费时间='" & Format(arrTmp(i), "yyyy-MM-dd") & "'"
                        If rsDays.RecordCount = 0 And Format(arrTmp(i), "yyyy-MM-dd") <> strDate Then
                            dbl数量 = 0
                        End If
                    End If
                Else
                    '如果不收取，则设置为0
                    dbl数量 = 0
                End If
                rsExec!数量 = dbl数量
                rsExec.Update
            Next
        End If
    End If
    AdviceMoneyMake = blnMakeMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetTubeMaterial(ByVal str试管编码 As String) As Long
'功能：根据管码获取对应的试管材料ID
    Dim strSQL As String
    
    If grsTube Is Nothing Then
        On Error GoTo errH
        
        strSQL = "Select 编码,材料ID From 采血管类型 Where 材料ID is Not NULL"
        Set grsTube = New ADODB.Recordset
        Set grsTube = zlDatabase.OpenSQLRecord(strSQL, "GetTubeMaterial")
    End If
    
    grsTube.Filter = "编码='" & str试管编码 & "'"
    If Not grsTube.EOF Then GetTubeMaterial = NVL(grsTube!材料ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiDayMoneyDetail(rsMoneyDay As ADODB.Recordset, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt来源 As Byte, ByVal dat收费时间 As Date) As Boolean
'功能：获取指定病人某天(dat收费时间)之后医嘱产生的费用项目明细
'参数：lng主页ID=住院病人才使用
'      byt来源:1-门诊(含住院临嘱发送到门诊)，2-住院
'      dat收费时间 需要判断的最早的时间
'返回：rsMoneyDay，包含"诊疗项目ID,收费项目ID,执行部门ID,执行否,收费时间"字段
'      一次发某条临时嘱后产生的医嘱执行时间点分布在不同的天里并且天数日期不连续可能会漏算，例如下。
'        临嘱一次发送后执行时间点分别为：2014-11-1 XX:XX,2014-11-3 XX:XX,2014-11-5 XX:XX,2014-11-7 XX:XX
'        1.dat收费时间 = 2014-11-1则只算两天，2014-11-1，2014-11-7 正确应该算3天 2014-11-1，2014-11-5，2014-11-7
'        2.dat收费时间 = 2014-11-2则只算1天，2014-11-7             正确应该算2天 2014-11-5，2014-11-7
    Dim strSQL As String, str医嘱IDs As String
    Dim rsTmp As ADODB.Recordset
    Dim rs执行时间 As ADODB.Recordset
    Dim i As Long, j As Long
    Dim strToday As String, strDay As String
    Dim strTmp As String
    Dim n As Long, p As Long
    Dim strThis As String
    Dim strTabTmp As String
    Dim varPar(0 To 10) As String
        
    On Error GoTo errH
 
    '执行判断：
    '1.传入的是将填定到费用记录中的执行部门，因此也以费用记录中的执行部门为准判断。
    '2.除和跟踪卫材外，医嘱费用的执行科室与医嘱执行科室相同；以后如果不同了，该函数也可以适应
    '3.医嘱执行时，对应费用的执行状态也会同步标记。
    '4.首次不收的项目，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
    If byt来源 = 1 Then
        strSQL = "Select A.诊疗项目ID,C.收费细目ID as 收费项目ID,C.执行部门ID,Decode(Nvl(C.执行状态,0),0,0,1) as 执行否," & _
            " To_Char(b.首次时间,'yyyy-mm-dd') as 首次时间,Trunc(b.末次时间) - Trunc(b.首次时间) + 1 AS 天数,0 as 收费方式," & _
            " To_Char(b.末次时间,'yyyy-mm-dd') as 末次时间,A.频率间隔,A.间隔单位,a.医嘱期效,a.id,nvl(a.相关id,0) as 相关id" & _
            " From 病人医嘱记录 A,病人医嘱发送 B,门诊费用记录 C" & _
            " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1 And A.ID=B.医嘱ID And B.记录性质=C.记录性质 And B.NO=C.NO" & _
            " And B.医嘱ID=C.医嘱序号 And C.记录状态 IN(0,1) And b.末次时间>=[3]" & _
            " Union " & _
            " Select A.诊疗项目ID,D.收费细目id,D.执行科室ID as 执行部门ID,0 as 执行否," & _
            " To_Char(B.首次时间,'yyyy-mm-dd') as 首次时间,Trunc(b.末次时间) - Trunc(b.首次时间) + 1 AS 天数,-1 as 收费方式," & _
            " To_Char(b.末次时间,'yyyy-mm-dd') as 末次时间,A.频率间隔,A.间隔单位,a.医嘱期效,a.id,nvl(a.相关id,0) as 相关id" & _
            " From 病人医嘱记录 A,病人医嘱发送 B,病人医嘱计价 D" & _
            " Where A.病人ID=[1] And Nvl(A.主页ID,0) = [2] And a.医嘱期效 = 1" & _
            " And A.ID=B.医嘱ID And b.末次时间>=[3] And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
            " And Not Exists (Select 1 From 门诊费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
            " Order by 诊疗项目ID,收费项目ID"
    Else
        '长嘱，其他医嘱的相同费用，可能不同时间多次发送,Union去除了重复记录
        '首次不收的项目，则没有产生费用记录（但有医嘱发送记录）,需要读出来当成已生成的，以便其他首次不收的项目判断
        strSQL = "Select a.诊疗项目id, c.收费细目id As 收费项目id, c.执行部门id, Decode(Nvl(c.执行状态, 0), 0, 0, 1) As 执行否," & vbNewLine & _
            "     To_Char(b.首次时间,'yyyy-mm-dd') As 首次时间, Decode(b.首次时间,null, 1,Trunc(b.末次时间) - Trunc(b.首次时间) + 1) As 天数,0 as 收费方式," & vbNewLine & _
            " To_Char(b.末次时间,'yyyy-mm-dd') as 末次时间,A.频率间隔,A.间隔单位,a.医嘱期效,a.id,nvl(a.相关id,0) as 相关id" & _
            " From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C" & vbNewLine & _
            " Where a.病人id = [1] And a.主页id = [2] And a.Id = b.医嘱id And b.记录性质 = c.记录性质 And b.No = c.No And b.医嘱id = c.医嘱序号 And" & vbNewLine & _
            "      c.记录状态 In (0, 1) And b.末次时间>=[3]" & vbNewLine & _
            " Union " & vbNewLine & _
            " Select a.诊疗项目id, D.收费细目id, D.执行科室ID as 执行部门id, 0 As 执行否," & vbNewLine & _
            " To_Char(b.首次时间,'yyyy-mm-dd') As 首次时间,Decode(a.医嘱期效, 0, Trunc(b.末次时间) - Trunc(b.首次时间) + 1, 1) As 天数,-1 as 收费方式," & vbNewLine & _
            " To_Char(b.末次时间,'yyyy-mm-dd') as 末次时间,A.频率间隔,A.间隔单位,a.医嘱期效,a.id,nvl(a.相关id,0) as 相关id" & _
            " From 病人医嘱记录 A, 病人医嘱发送 B, 病人医嘱计价 D" & vbNewLine & _
            " Where a.病人id = [1] And a.主页id = [2] And a.Id = b.医嘱id And b.末次时间>=[3]" & vbNewLine & _
            " And A.ID=D.医嘱ID And D.收费方式=7" & vbNewLine & _
            " And Not Exists (Select 1 From 住院费用记录 C Where c.收费细目id=d.收费细目id  And b.记录性质 = c.记录性质 And b.No = c.No And a.Id = c.医嘱序号)" & vbNewLine & _
            " Order By 诊疗项目id, 收费项目id"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取当天及后续的医嘱", lng病人ID, lng主页ID, dat收费时间)
    
    If byt来源 = 2 Then
        rsTmp.Filter = "医嘱期效=0 and 相关id=0"
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                If InStr("," & str医嘱IDs & ",", "," & rsTmp!ID & ",") = 0 Then str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
            If str医嘱IDs <> "" Then
                strTmp = "Select /*+cardinality(x,10)*/ x.Column_Value as 医嘱ID From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)) X"
                n = 0
                Do While True
                    If Len(str医嘱IDs) < 4000 Then
                        p = Len(str医嘱IDs) + 1
                    Else
                        p = InStrRev(Mid(str医嘱IDs, 1, 4000), ",")
                    End If
                    strThis = Mid(str医嘱IDs, 1, p - 1)
                    
                    If n > 10 Then
                        strTabTmp = strTabTmp & vbNewLine & " Union All " & Replace(strTmp, "[2]", "'" & strThis & "'")
                    Else
                        varPar(n) = strThis
                        strTabTmp = IIF(strTabTmp = "", "", strTabTmp & vbNewLine & " Union All ") & Replace(strTmp, "[2]", "[" & (n + 2) & "]")
                    End If
                    n = n + 1
                    str医嘱IDs = Mid(str医嘱IDs, p + 1)
                    If str医嘱IDs = "" Then Exit Do
                Loop
                strSQL = "select a.医嘱id,To_Char(Trunc(a.要求时间), 'yyyy-mm-dd') As 执行时间 from 医嘱执行时间 a" & _
                    " where a.要求时间>=[1] and a.医嘱id in (" & strTabTmp & ")  group by a.医嘱id,Trunc(a.要求时间)"
                Set rs执行时间 = zlDatabase.OpenSQLRecord(strSQL, "读取当天及后续的医嘱", dat收费时间, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
            End If
        End If
        rsTmp.Filter = 0
    End If
    
    '将记录集按执行时间分成多条记录
    strToday = Format(dat收费时间, "yyyy-MM-dd")
    Set rsMoneyDay = New ADODB.Recordset '用于清除Filter属性
    Set rsMoneyDay = InitPatiExecDays
    
    For i = 1 To rsTmp.RecordCount
        If Val(rsTmp!医嘱期效 & "") = 1 Then
            If Val(rsTmp!频率间隔 & "") = 1 And rsTmp!间隔单位 & "" = "天" Or rsTmp!间隔单位 & "" = "小时" Or rsTmp!间隔单位 & "" = "分钟" Then
                For j = 1 To rsTmp!天数
                    If j = 1 Then
                        strDay = Format(rsTmp!首次时间, "yyyy-MM-dd")
                    Else
                        strDay = Format(DateAdd("d", j - 1, CDate(rsTmp!首次时间)), "yyyy-MM-dd")
                    End If
                    If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                Next
            Else
                '对于临嘱这里可能会漏算
                strDay = Format(rsTmp!首次时间, "yyyy-MM-dd")
                If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                strDay = Format(rsTmp!末次时间, "yyyy-MM-dd")
                If strDay >= strToday Then Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
            End If
        Else
            '长嘱不会错，通过 医嘱执行时间 可以确定是在那一天执行
            If Not rs执行时间 Is Nothing Then
                rs执行时间.Filter = "医嘱id=" & rsTmp!ID & " or 医嘱id=" & rsTmp!相关ID
                For j = 1 To rs执行时间.RecordCount
                    strDay = rs执行时间!执行时间 & ""
                    Call AddMoneyDayItem(rsTmp, rsMoneyDay, strDay)
                    rs执行时间.MoveNext
                Next
            End If
        End If
        
        rsTmp.MoveNext
    Next
    rsMoneyDay.Filter = ""
    
    GetPatiDayMoneyDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddMoneyDayItem(ByVal rsTmp As ADODB.Recordset, ByRef rsMoneyDay As ADODB.Recordset, ByVal strDay As String)
'功能：向 rsMoneyDay 添加数据
    rsMoneyDay.Filter = "诊疗项目ID=" & Val("" & rsTmp!诊疗项目ID) & " And 收费项目ID=" & Val("" & rsTmp!收费项目ID) & _
        " And 收费时间='" & strDay & "' And 执行否=" & Val("" & rsTmp!执行否) & " And 收费方式=" & Val("" & rsTmp!收费方式)
    If rsMoneyDay.RecordCount = 0 Then
        rsMoneyDay.AddNew
        rsMoneyDay!诊疗项目ID = Val("" & rsTmp!诊疗项目ID)
        rsMoneyDay!收费项目ID = Val("" & rsTmp!收费项目ID)
        rsMoneyDay!执行部门ID = Val("" & rsTmp!执行部门ID)
        rsMoneyDay!执行否 = Val("" & rsTmp!执行否)
        rsMoneyDay!收费方式 = Val("" & rsTmp!收费方式)
        rsMoneyDay!收费时间 = strDay
        rsMoneyDay.Update
    End If
End Sub

Private Function InitPatiExecDays() As ADODB.Recordset
'功能：初始化医嘱相关费用执行的记录集
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    rsTmp.Fields.Append "诊疗项目ID", adBigInt
    rsTmp.Fields.Append "收费项目ID", adBigInt
    rsTmp.Fields.Append "执行部门ID", adBigInt
    rsTmp.Fields.Append "收费方式", adInteger
    rsTmp.Fields.Append "执行否", adInteger
    rsTmp.Fields.Append "收费时间", adVarChar, 10
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set InitPatiExecDays = rsTmp
End Function

Public Function GetSkinTestResult(ByVal lng项目id As Long, ByVal str结果 As String) As Integer
'功能：根据皮试结果标注，返回阴阳性
'参数：str结果=皮试结果标注符号,如"(+)"
'返回：-1-阴性,1-阳性,0-无结果
    Dim arr阳性 As Variant, arr阴性 As Variant
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    If grsSkinTest Is Nothing Then
        strSQL = "Select ID,Nvl(标本部位,'阳性(+);阴性(-)') as 标注 From 诊疗项目目录 Where 类别='E' And 操作类型='1'"
        Set grsSkinTest = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSkinTest, strSQL, "GetSkinTestResult")
    End If
    
    grsSkinTest.Filter = "ID=" & lng项目id
    If grsSkinTest.EOF Then Exit Function
    
    arr阳性 = Split(Split(grsSkinTest!标注, ";")(0), ",")
    arr阴性 = Split(Split(grsSkinTest!标注, ";")(1), ",")
    
    For i = 0 To UBound(arr阳性)
        If Right(arr阳性(i), Len(str结果)) = str结果 Then
            GetSkinTestResult = 1: Exit Function
        End If
    Next
    For i = 0 To UBound(arr阴性)
        If Right(arr阴性(i), Len(str结果)) = str结果 Then
            GetSkinTestResult = -1: Exit Function
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查部门信息是否加载编码
    '返回:显示编码,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 13:11:01
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    If mlng部门编码平均长度 = 0 Then
        strSQL = "Select Avg(length(编码)) As 长度 From 部门表"
        On Error GoTo errH
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取部门编码的平均长度")
        mlng部门编码平均长度 = Val(NVL(rsTemp!长度))
    End If
    '由于编码长度可能过长,无法显示部门的名称,因此自动显示和不显示编码,当大于5时,不显示.小于5时,显示
   zlIsShowDeptCode = mlng部门编码平均长度 <= 5
      
   Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", _
    Optional blnSendKeys As Boolean = True, Optional blnAddItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim intIndex As Integer
    Dim strIDs As String, str简码 As String, strLike As String
    strLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "*", "")
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = strLike & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If NVL(!编码) = strSearch Then lngDeptID = NVL(!ID): iCount = 0:  Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp): Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(NVL(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If NVL(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(NVL(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(NVL(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(NVL(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(NVL(!简码)) Like strCompents Or UCase(Trim(NVL(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(NVL(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(NVL(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = NVL(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 And lngDeptID <= 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIF(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIF(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIF(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIF(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(NVL(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    If Cbo.Locate(cboDept, lngDeptID, True) = False Then
        If blnAddItem = True Then
            If rsTemp.RecordCount = 1 Then
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIF(zlIsShowDeptCode, rsTemp!编码 & "-", "") & rsTemp!名称
                cboDept.ItemData(cboDept.ListCount - 1) = Val(NVL(rsTemp!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "其他科室…"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            Else
                cboDept.RemoveItem cboDept.ListCount - 1
                cboDept.AddItem IIF(zlIsShowDeptCode, rsReturn!编码 & "-", "") & rsReturn!名称
                cboDept.ItemData(cboDept.ListCount - 1) = Val(NVL(rsReturn!ID))
                intIndex = cboDept.NewIndex
                cboDept.AddItem "其他科室…"
                cboDept.ItemData(cboDept.ListCount - 1) = 0
                cboDept.ListIndex = intIndex
            End If
            rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
            zlSelectDept = True
            Exit Function
        Else
            GoTo GoNotSel
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing

    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
    zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

'*********************************************************************************************************************
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Function CheckSpecialAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByRef strPaits As String) As Boolean
'功能：检查病人是否有需要只校对，不发送的特殊医嘱
'       持续护理等级,病重/危医嘱,术前术后医嘱不发送,记录入出量,,转科，出院，转院，死亡
'
'参数：strPaits=以逗号分隔的病人ID串,返回存在以上特殊医嘱的病人ID,主页ID;
'      lng病人ID,lng主页ID=单个病人时才传入
    Dim rsTmp As ADODB.Recordset, rsExists As ADODB.Recordset, strSQL As String, i As Long
    Dim strDepts As String, blnOnePati As Boolean, blnExists As Boolean
    
    strDepts = GetUser科室IDs   '当前操作人员的所属病区的所有科室
    On Error GoTo errH
    If strPaits = "" Then
        strSQL = "Select A.病人ID,A.主页ID,a.诊疗类别,e.操作类型,a.执行频次" & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 E" & vbNewLine & _
                "Where a.病人id = [1] And a.主页id = [2]"
    Else
        strSQL = "Select/*+ Rule*/ Distinct A.病人ID,A.主页ID,a.诊疗类别,e.操作类型,a.执行频次" & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 E,Table(f_Num2list([1])) B" & vbNewLine & _
                "Where a.病人id = B.Column_value"
    End If
    strSQL = strSQL & " And a.诊疗项目id = e.Id And a.医嘱状态 = 1" & vbNewLine & _
            " And (a.诊疗类别 = 'H' And e.操作类型 = '1' And e.执行频率 = 2 Or a.诊疗类别 = 'Z' And e.操作类型 In ('3', '4', '14', '5', '6', '9', '10', '11', '12') Or NVL(a.执行频次,'无')='必要时' Or NVL(a.执行频次,'无')='需要时')" & _
            " And Exists(" & _
            "Select M.姓名 From 人员表 M,执业类别 N" & _
            " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1)," & _
            "2,Substr(A.开嘱医生,1,Decode(Instr(A.开嘱医生,'/'),0,length(A.开嘱医生),Instr(A.开嘱医生,'/')-1))," & _
            "Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
            " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')" & _
            " )"
            
    If InStr(GetInsidePrivs(p住院医嘱发送), "全院医嘱校对") = 0 Then
        strSQL = strSQL & "  And (A.开嘱科室ID In (Select Column_Value From Table(f_Num2list([3]))) Or Exists(select 1 From 病人医嘱记录 Q,诊疗项目目录 O where Q.病人ID=A.病人ID AND Q.主页ID=A.主页ID AND Q.诊疗项目ID=O.ID AND Q.诊疗类别='Z'AND O.操作类型 ='7' AND Q.执行科室ID=A.开嘱科室ID And Q.医嘱状态=8)) "
    End If
    
    If strPaits = "" Then
        blnOnePati = True
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查特殊医嘱", lng病人ID, lng主页ID, strDepts)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查特殊医嘱", strPaits, 0, strDepts)
    End If
    
    strPaits = ""
    For i = 1 To rsTmp.RecordCount
        blnExists = False
        
        '5-出院;6-转院,11-死亡,校对时停止其他长嘱，发送时将病人标记为预出院
        '3-转科;4,14-术前术后医嘱
        If rsTmp!诊疗类别 = "Z" And InStr(",3,4,14,5,6,11,", "," & rsTmp!操作类型 & ",") > 0 Or rsTmp!执行频次 & "" = "必要时" Then
            blnExists = True
        Else
            strSQL = "Select 1" & vbNewLine & _
                    "From 病人医嘱记录 A, 诊疗项目目录 E" & vbNewLine & _
                    "Where a.病人id = [1] And a.主页id = [2] And a.诊疗项目id = e.Id And a.医嘱状态 In (3,5,6,7)"
            '护理等级医嘱
            If rsTmp!诊疗类别 = "H" And rsTmp!操作类型 = "1" Then
                strSQL = strSQL & " And a.诊疗类别 = 'H' And e.操作类型 = '1' And e.执行频率 = 2"
            ElseIf rsTmp!诊疗类别 = "Z" Then
                If rsTmp!操作类型 = "9" Then    '病重
                    strSQL = strSQL & " And a.诊疗类别 = 'Z' And e.操作类型 = '10'"
                ElseIf rsTmp!操作类型 = "10" Then '病危
                    strSQL = strSQL & " And a.诊疗类别 = 'Z' And e.操作类型 = '9'"
                ElseIf rsTmp!操作类型 = "12" Then '记录入出量
                    strSQL = strSQL & " And a.诊疗类别 = 'Z' And e.操作类型 = '12'"
                End If
            End If
                    
            Set rsExists = zlDatabase.OpenSQLRecord(strSQL, "检查特殊医嘱", Val(rsTmp!病人ID), Val("" & rsTmp!主页ID))
            blnExists = rsExists.RecordCount > 0
        End If
        
        If blnExists Then
            If blnOnePati Then
                Exit For
            Else
                If InStr(strPaits & ";", ";" & rsTmp!病人ID & "," & rsTmp!主页ID & ";") = 0 Then
                    strPaits = strPaits & ";" & rsTmp!病人ID & "," & rsTmp!主页ID
                End If
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If blnOnePati Then
        CheckSpecialAdvice = blnExists
    Else
        If strPaits = "" Then
            CheckSpecialAdvice = False
        Else
            strPaits = Mid(strPaits, 2)
            CheckSpecialAdvice = True
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'功能:检查收费项目是否未停用,并且服务于病人
'参数:bytFlag:服务对象:1-门诊,2-住院
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From 收费项目目录 Where ID = [1] And (撤档时间 is Null Or 撤档时间 > Sysdate) And 服务对象 In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
ErrHand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏", , , True)) = 1
    err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function


Public Function AssembleImage(AssembleViewer As Object, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As Object

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    On Error GoTo err
    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '计算横向和纵向图像数量
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '修正图像的宽高，不能大于最大值
    '如果大于intMaxWidth×intMaxHeight则，按照图像总长宽比，使用小于等于intMaxWidth×intMaxHeight作为新宽高,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '采集图像
    '将图像采集到临时图像集
    For i = 1 To AssembleViewer.Count
        '计算缩放比例 hj修改,解决多图合并时，放大的图象无法真正放大的问题
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    Rows = iRows: Cols = iCols
    
    If ImageCount <> 0 Then
        If Rows * Cols > ImageCount Then
            iBase = 6
            blnDoLoop = True
            
            While blnDoLoop
                iBase = iBase - 1
                
                If ImageCount Mod iBase = 0 Then
                    blnDoLoop = False
                End If
            Wend
        

            If RegionWidth > RegionHeight Then
                If ImageCount / iBase > iBase Then
                    Cols = ImageCount / iBase
                    Rows = iBase
                Else
                    Rows = ImageCount / iBase
                    Cols = iBase
                End If
            Else
                If ImageCount / iBase > iBase Then
                    Cols = iBase
                    Rows = ImageCount / iBase
                Else
                    Rows = iBase
                    Cols = ImageCount / iBase
                End If
            End If
        End If
    End If
err:
End Sub

Public Function GetRPTPicture(ByVal blnMoved As Boolean, ByVal lngReportID As Long, ByVal strRPTNO As String, ByRef aryPrintPara As Variant) As String
'功能：获取报告图片
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strPicPath As String
    Dim intPCount As Integer
    Dim oPicture As StdPicture
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryPara(19) As String, aryFlagPara(1) As String     '报告图中的图像记录
    Dim strFlagString As String '实际传给自定义报表的内容
    Dim int格式号 As Integer
    Dim intRows As Integer, intCols As Integer
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    
    Const cprET_单病历审核 = 3
    Const EPRMarkedPicture = 1
    
    Dim cTable As Object, objFile As Object
    
    
    On Error GoTo err
    
    Set cTable = CreateObject("zlRichEPR.cEPRTable")
    If cTable Is Nothing Then Exit Function
    Set objFile = CreateObject("Scripting.FileSystemObject")
        
    '获取图像
    strPicPath = App.Path & "\TmpImage\"
    If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
    
    '获取报告图像（包括标记图）生成本地文件
    '一个报告表格中可能排列多个报告图
    intPCount = 0
    strSQL = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
    If blnMoved = True Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取图像", lngReportID)
    Do While Not rsTmp.EOF
        If cTable.GetTableFromDB(cprET_单病历审核, lngReportID, Val("" & rsTmp!表格ID)) Then
            For i = 1 To cTable.Pictures.Count
                strPicFile = "PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                    Set oPicture = cTable.Pictures(i).DrawFinalPic
                Else
                    Set oPicture = cTable.Pictures(i).OrigPic
                End If
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '保存标记图和图象的路径
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        aryFlagPara(0) = strPicFile
                    Else
                        aryPara(intPCount) = strPicFile
                        dcmImages.AddNew
                        dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                        intPCount = intPCount + 1
                        If intPCount > UBound(aryPara) Then Exit Do
                    End If
                End If
            Next i
        End If
        rsTmp.MoveNext
    Loop
    
    '根据选择的自定义报表格式，组织图像
    '仅按一种报表格式处理
    int格式号 = 1
    strSQL = "Select b.名称,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
    "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2] And b.名称 not like '标记%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询是否需要组合图像", strRPTNO, int格式号)
    If rsTmp.RecordCount = 1 And intPCount >= 1 Then
        '组合图象
        ResizeRegion intPCount, rsTmp("W"), rsTmp("H"), intRows, intCols
        Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTmp("H"), rsTmp("W"))
        dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
    End If
    
    
    '获取图像，调用报表
    intPCount = 0       '记录图像的数量
    strSQL = "Select b.名称 From zlReports a, zlRptItems b" & vbNewLine & _
    "       Where a.Id = b.报表id And a.编号 = [1] And Nvl(b.下线, 0) = 1 And b.类型 = 11 And b.格式号 = [2]" & vbNewLine & _
    "       Order By b.名称"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取图象框", strRPTNO, int格式号)
    '装载图像数据
    intParaCount = 0
    Do While Not rsTmp.EOF
        
        '分别装在标记图和报告图
        If InStr(rsTmp!名称, "标记") <> 0 Then '标记图
            If aryFlagPara(0) <> "" Then strFlagString = rsTmp!名称 & "=" & aryFlagPara(0)
        Else    '报告图
            If intPCount > UBound(aryPara) Then Exit Do     '图像数量超过报告中的图像，退出
            If aryPara(intPCount) = "" Then Exit Do         '报表中的图象框比报告中的多，退出
            
            aryPrintPara(intParaCount) = rsTmp!名称 & "=" & aryPara(intPCount)
            intPCount = intPCount + 1
            intParaCount = intParaCount + 1
        End If
        rsTmp.MoveNext
    Loop
    
    '处理报表中图形比报告中少的情况
    For j = intParaCount To UBound(aryPrintPara)
        If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
    Next
    
    GetRPTPicture = strFlagString
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetAdviceRs() As ADODB.Recordset
'功能：产生一个包含医嘱记录所有字段的本地记录集对象
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select a.*,c.疾病ID,0 as EditState From 病人医嘱记录 A,病人诊断医嘱 B,病人诊断记录 C Where A.id=b.医嘱id and c.id=b.诊断id And a.ID=0 And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceRS")
    
    Set GetAdviceRs = zlDatabase.zlCopyDataStructure(rsTmp)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckUnExecutedAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngSpecialAdviceID As Long, ByVal intBabyNum As Integer) As String
'功能：检查未校对的医嘱，和未发送的临嘱。
'参数：strSpecialAdviceIDs=特殊医嘱ID(检查时要排除当前特殊医嘱ID)
'      intBabyNum=婴儿序号。
'返回：提示信息
'说明：该函数执行时只发送母亲特殊医嘱时会忽略对婴儿医嘱的检查，如果同时发送则都正常。
'      存储过程  Zl_病人变动记录_Change 和 Zl_病人变动记录_Preout 也有类似的检查，做了相应的调整。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    '未校对医嘱
    strSQL = "Select 1 From 病人医嘱记录 Where 病人id = [1] And 主页id = [2] And 医嘱状态 = 1 And ID<>[3] And Nvl(婴儿, 0) = [4] And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", lng病人ID, lng主页ID, lngSpecialAdviceID, intBabyNum)
    If rsTmp.RecordCount > 0 Then strMsg = "未校对的医嘱"
    
    '未发送的临嘱
    strSQL = "Select 1" & vbNewLine & _
        "From 病人医嘱记录" & vbNewLine & _
        "Where 病人id = [1] And 主页id = [2] And 医嘱期效 = 1 And 医嘱状态 In (2, 3) And Nvl(执行标记, 0) <> -1 And ID <> [3] And Nvl(婴儿, 0) = [4] And" & vbNewLine & _
        " Nvl(执行性质,0)<>0 and Rownum = 1"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", lng病人ID, lng主页ID, lngSpecialAdviceID, intBabyNum)
    If rsTmp.RecordCount > 0 Then
        If strMsg <> "" Then
            strMsg = strMsg & "、未发送的临嘱"
        Else
            strMsg = "未发送的临嘱"
        End If
    End If
        
    strSQL = "Select 1 From 病人医嘱记录 a Where a.病人id = [1] And a.主页id = [2] And a.医嘱期效 = 0 And a.医嘱状态=8 " & _
            "And Exists(Select 1 From 病人医嘱记录 b Where b.id = [3] And a.执行终止时间 > b.开始执行时间) And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", lng病人ID, lng主页ID, lngSpecialAdviceID)
    If rsTmp.RecordCount > 0 Then
        If strMsg <> "" Then
            strMsg = strMsg & "、未到停止时间的长嘱"
        Else
            strMsg = "未到停止时间的长嘱"
        End If
    End If
    
    
    CheckUnExecutedAdvice = strMsg
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function getChargeMode(ByVal intChargeMode As Integer) As String
    getChargeMode = Decode(intChargeMode, 0, "正常收取", 1, "检验试管费用", 2, "一次发送只收取一次", 3, "当天只收取一次", 4, _
        "当天未执行收取一次", 5, "当天只收取一次,排斥其他项目", 6, "当天未执行收取一次,排斥其他项目", 7, "每天首次不收取", 9, "自定义")
End Function

Public Sub Load输液滴速(ByRef cbo滴速 As ComboBox, ByRef lbl滴速单位 As Label, ByVal blnTurn As Boolean, Optional ByVal bln输血 As Boolean = False)
'功能：加载滴速下拉列表
'参数：blnTurn=是否转换输液单位
    Dim i As Long
    Dim arrTmp() As String
    
    If bln输血 = False Then
        If blnTurn Then
            If lbl滴速单位.Caption = "毫升/小时" Then
                lbl滴速单位.Caption = "滴/分钟"
            Else
                lbl滴速单位.Caption = "毫升/小时"
            End If
        End If
        If lbl滴速单位.Caption = "滴/分钟" Then
            arrTmp = Split("20,30,40,50,60,70,80,20-40,40-60", ",")
        Else
            arrTmp = Split("60,120,180,240,300,600,120-240", ",")
        End If
    Else
        arrTmp = Split("15,30,快速,加压", ",")
    End If
    
    cbo滴速.Clear
    For i = 0 To UBound(arrTmp)
        cbo滴速.AddItem arrTmp(i)
    Next
    
End Sub

Public Function GetKSSAuditQuestion(ByVal lng医嘱ID As Long) As String
'功能：获取抗菌用药审核未通过的反馈信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(操作说明,'无') as 操作说明 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=12 Order by 操作时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
    If Not rsTmp.EOF Then
        GetKSSAuditQuestion = rsTmp!操作说明
    Else
        If gbln血库系统 Then
            '血库将输血医嘱的设为审核不通过后，操作类型为16
            strSQL = "Select Nvl(操作说明,'无') as 操作说明 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=16 Order by 操作时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
            If Not rsTmp.EOF Then GetKSSAuditQuestion = rsTmp!操作说明
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetMaxAdviceNO(ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal lng婴儿 As Long, Optional ByVal str挂号单 As String) As Long
'功能：获取当前病人的最大医嘱序号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If str挂号单 <> "" Then
        strSQL = "Select nvl(Max(序号),1) as 序号 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[4] And Nvl(婴儿,0)=[3]"
    Else
        If lng主页ID = 0 Then
            strSQL = "Select Nvl(Max(序号),1) as 序号 From 病人医嘱记录 Where 病人ID=[1] And 主页ID Is Null"
        Else
            strSQL = "Select Nvl(Max(序号),1) as 序号 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(婴儿,0)=[3]"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID, lng婴儿, str挂号单)
    If Not rsTmp.EOF Then GetMaxAdviceNO = rsTmp!序号

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlIsCheckMedicinePayMode(ByVal str医疗付款名称 As String, _
    Optional ByRef bln医保 As Boolean, Optional ByRef bln公费 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医疗付款方式是否公费或医保
    '入参:str医疗付款名称-医疗付款名称
    '出参:bln医保-true,表示医保
    '        bln公费-true,表示是公费
    '返回:是医保或公费医疗,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-17 16:25:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "": bln医保 = False: bln公费 = False
    If grs医疗付款方式 Is Nothing Then
        strSQL = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    ElseIf grs医疗付款方式.State <> 1 Then
        strSQL = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    End If
    If strSQL <> "" Then
        Set grs医疗付款方式 = zlDatabase.OpenSQLRecord(strSQL, "获取医疗付款方式")
    End If
    grs医疗付款方式.Find "名称='" & str医疗付款名称 & "'", , adSearchForward, 1
    If grs医疗付款方式.EOF Then Exit Function
    bln医保 = Val(NVL(grs医疗付款方式!是否医保)) = 1
    bln公费 = Val(NVL(grs医疗付款方式!是否公费)) = 1
    zlIsCheckMedicinePayMode = bln医保 Or bln公费
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check检查部位Enable(ByVal lng项目id As Long, ByVal str部位 As String, ByVal str性别 As String, Optional ByVal str方法 As String, Optional ByRef blnExists As Boolean) As Boolean
'功能：检查部位是否适用于指定的性别
'参数：blnExists 判断是否存在这个检查部位或方法
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(b.适用性别, 0) as 适用性别" & vbNewLine & _
            "From 诊疗项目部位 A, 诊疗检查部位 B, 诊疗项目目录 C" & vbNewLine & _
            "Where a.类型 = b.类型 And a.部位 = b.名称 And a.类型 = c.操作类型 And a.项目id = c.Id And" & vbNewLine & _
            "       c.Id = [1] And a.部位 = [2] And Replace(A.方法,chr(9),'')=[3]"
    blnExists = True
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id, str部位, str方法)
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "适用性别=" & IIF(str性别 = "男", 1, IIF(str性别 = "女", 2, 0)) & " Or 适用性别=0"
        Check检查部位Enable = rsTmp.RecordCount > 0
    Else
        blnExists = False
    End If
        
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckStopedUnAffirm(ByVal strPatis As String, ByRef strPatisName As String) As Boolean
'功能：检查指定的病人是否存在已停止但未确认停止的医嘱
'参数：strPatis=病人ID1:主页ID,病人ID2:主页ID,......
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    If InStr(strPatis, ",") = 0 Then
        strSQL = "Select a.姓名" & vbNewLine & _
            "From 病人医嘱记录 A" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And Exists (Select 1 From 病人医嘱状态 B Where a.Id = b.医嘱id And b.操作类型 = 8 And b.签名id Is Not Null)" & _
            " And a.医嘱状态 = 8 And a.医嘱期效 = 0 And Nvl(a.婴儿, 0) = 0 And Nvl(A.执行标记,0)<>-1 And Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", Val(Split(strPatis, ":")(0)), Val(Split(strPatis, ":")(1)))
    Else
        strSQL = "Select /*+ leading(b) use_nl(a)*/Distinct a.姓名" & vbNewLine & _
                "From 病人医嘱记录 A, Table(f_Num2list2([1])) B" & vbNewLine & _
                "Where a.病人id = b.C1 And a.主页id = b.C2 And a.医嘱状态 = 8 And Nvl(A.执行标记,0)<>-1 " & _
                " And Exists (Select 1 From 病人医嘱状态 C Where a.Id = C.医嘱id And C.操作类型 = 8 And C.签名id Is Not Null) " & _
                " And a.医嘱期效 = 0 And Nvl(a.婴儿, 0) = 0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strPatis)
    End If

    CheckStopedUnAffirm = rsTmp.RecordCount > 0
    For i = 1 To rsTmp.RecordCount
        strPatisName = strPatisName & "," & rsTmp!姓名
        rsTmp.MoveNext
    Next
    If strPatisName <> "" Then strPatisName = Mid(strPatisName, 2)

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetRsRedoDate(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Date
'功能：获取病人医嘱最近重整时间
    Dim strSQL As String, rsTmp As ADODB.Recordset
 
    strSQL = "Select 医嘱重整时间 as 时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取重整时间", lng病人ID, lng主页ID)
    
    GetRsRedoDate = NVL(rsTmp!时间, CDate("1900-01-01"))
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check过敏试验(frmParent As Object, ByVal txt医嘱内容 As Object, ByVal lng病人ID As Long, ByVal lng药名ID As Long, ByVal str名称 As String, ByVal bln自动皮试 As Boolean, Optional ByRef lng皮试ID As Long, _
    Optional ByVal lng主页ID As Long, Optional ByVal lng药品ID As Long, Optional ByRef bln连续用药 As Boolean, Optional ByVal str开始时间 As String, Optional ByRef bln阳性禁示 As Boolean) As String
'功能：检查西成药，中成药的过敏试验，门诊和住院医嘱下达公用
'参数：frmParent当前调用的窗体
'      txt医嘱内容 控件，医嘱编辑界面的文本输入框控件对象
'      lng药名ID=药品诊疗项目ID
'      str名称=药品名称,用于提示
'      bln连续用药 bln判断连续用药 =true 表明要判断并将结果以 bln连续用药 返回（仅用于住院）
'      str开始时间 医嘱的开始执行时间
'      bln阳性禁示 药品品种或者规格对应的过敏实验项目设为 不允许脱敏使用，返回（用于界面判断）
'返回：为空表示通过
'      lng皮试ID=要自动添加的皮试项目ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dat开始时间 As Date
    Dim lngTmp皮试ID As Long
    Dim str皮试名称 As String
    
    Dim bln规格 As Boolean
    Dim rs皮试 As ADODB.Recordset
    Dim blnTmp阳性禁示 As Boolean
    
    On Error GoTo errH
    
    lng皮试ID = 0
    bln连续用药 = False
    bln阳性禁示 = False
    
    '判断当前药品是不是绑定了皮试项目，先判断规格，再判断品种
    If lng药品ID <> 0 Then
        strSQL = "Select A.用法ID,B.名称,B.执行标记 From 药品用法用量 A,诊疗项目目录 B Where A.用法ID=B.ID And A.性质=0 And A.药品ID=[1] And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) order by b.编码"
        Set rs皮试 = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng药品ID)
        If Not rs皮试.EOF Then bln规格 = True
    End If
    
    If Not bln规格 Then
        strSQL = "Select A.用法ID,B.名称,B.执行标记 From 诊疗用法用量 A,诊疗项目目录 B Where A.用法ID=B.ID And A.性质=0 And A.项目ID=[1] And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) order by b.编码"
        Set rs皮试 = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng药名ID)
    End If
    
    
    '需要皮试
    If Not rs皮试.EOF Then
        '如果只有一个皮试项目将其记录下来
        If rs皮试.RecordCount = 1 Then
            lngTmp皮试ID = rs皮试!用法ID
            str皮试名称 = rs皮试!名称 & ""
        End If
        
        rs皮试.Filter = "执行标记=2"
        blnTmp阳性禁示 = Not rs皮试.EOF
        
        If str开始时间 <> "" Then
            dat开始时间 = CDate(str开始时间)
        Else
            dat开始时间 = zlDatabase.Currentdate
        End If
        '取有效时间内的最后一次过敏结果登记
        strSQL = "Select 药物名,结果,Nvl(过敏时间,记录时间) as 过敏时间 From 病人过敏记录" & _
            " Where 病人ID=[1] And 药物ID=[2] And [3]+Nvl(过敏时间,记录时间)>=[4]" & _
            " Order by Nvl(过敏时间,记录时间) Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng病人ID, lng药名ID, gint过敏登记有效天数, dat开始时间)
        If Not rsTmp.EOF Then
            '有过敏结果登记记录,根据是否阳性决定是否提示
            If NVL(rsTmp!结果, 0) = 1 Then
                If blnTmp阳性禁示 Then
                    bln阳性禁示 = True
                    strMsg = "该病人在" & Format(rsTmp!过敏时间, "M月d日") & "的过敏实验中对""" & NVL(rsTmp!药物名, str名称) & """过敏(+)。" & _
                        vbCrLf & vbCrLf & "该项目已设置了不能进行脱敏使用，故禁止使用。"
                Else
                    strMsg = "该病人在" & Format(rsTmp!过敏时间, "M月d日") & "的过敏实验中对""" & NVL(rsTmp!药物名, str名称) & """过敏(+)。" & _
                        vbCrLf & vbCrLf & "是否仍然使用该药品？"
                End If
            Else
                strMsg = "" '为阴性,通过
            End If
        Else
            If lng主页ID <> 0 Then
                '如果按规格下达则判断 药品id,否则按品种判断75874
                strSQL = "select 1 from 病人医嘱记录 a where a.病人id=[1] and a.主页id=[2]" & _
                    IIF(lng药品ID <> 0, " and a.收费细目id=[3]", " and a.诊疗项目id=[3]") & _
                    " And [4]+a.上次执行时间>=[5] And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lng病人ID, lng主页ID, IIF(0 <> lng药品ID, lng药品ID, lng药名ID), gint过敏登记有效天数, dat开始时间)
                bln连续用药 = Not rsTmp.EOF
            End If
            
            If Not bln连续用药 Then
                If bln自动皮试 Then
                    '问题：31144,如果是多个皮试项目，则弹出选择器，供用户选择一个皮试。
                    If lngTmp皮试ID = 0 Then
                        vRect = zlControl.GetControlRect(txt医嘱内容.hwnd)
                        
                        If bln规格 Then
                            strSQL = "Select A.用法ID as ID,B.名称 as 请选择一种皮试" & _
                                    " From 药品用法用量 A,诊疗项目目录 B" & _
                                    " Where A.用法ID=B.ID And A.性质=0 And A.药品ID=[1]" & _
                                    " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) order by b.编码"
                            Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "皮试医嘱选择", False, "", "", False, False, True, _
                                vRect.Left, vRect.Top, txt医嘱内容.Height, blnCancel, False, True, lng药品ID)
                        Else
                            strSQL = "Select A.用法ID as ID,B.名称 as 请选择一种皮试" & _
                                    " From 诊疗用法用量 A,诊疗项目目录 B" & _
                                    " Where A.用法ID=B.ID And A.性质=0 And A.项目ID=[1]" & _
                                    " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) order by b.编码"
                            Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "皮试医嘱选择", False, "", "", False, False, True, _
                                vRect.Left, vRect.Top, txt医嘱内容.Height, blnCancel, False, True, lng药名ID)
                        End If
                            
                        If Not blnCancel Then
                            lng皮试ID = rsTmp!ID
                            strMsg = "" '自动添加,不提示
                        Else
                            strMsg = "在对病人使用""" & str名称 & """前，要求先进行过敏测试，" & vbCrLf & _
                                    "但您刚才没有选择对应过敏测试，是否仍然使用该药品？"
                        End If
                    Else
                        lng皮试ID = lngTmp皮试ID
                        strMsg = "" '自动添加,不提示
                    End If
                Else
                    '要求皮试,则提示皮试
                    strMsg = "在对病人使用""" & str名称 & """前，要求先进行""" & str皮试名称 & """，" & vbCrLf & _
                        "但没有发现有效的过敏试验结果，是否仍然使用该药品？"
                End If
            End If
        End If
    End If
    Check过敏试验 = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNoneSendID(ByVal lng病人ID As Long, ByVal str标识 As String, ByVal bytType As Byte, Optional ByVal bln发送输液药品 As Boolean, Optional ByVal lng挂号ID As Long, Optional ByRef strAdviceDrugIDs As String) As String
'功能：获取因为皮试过敏(阳性(+))或无皮试结果而不发送的医嘱ID
'参数：lng病人 病人ID
'      str标识 如果是门诊则是挂号单，住院则是主页id
'      bytType 1门诊，2住院
'      bln发送输液药品 =输液药品发送页面调用
'      strAdviceDrugIDs  出参数   受限制的药品行的医嘱IDs
    Dim rsTmp As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim str用法IDs As String
    Dim strPatiFilter As String
    Dim strPatiFilterAnd As String
    Dim bln自动皮试 As Boolean
    Dim strOK医嘱IDs As String
    Dim strNO医嘱IDs As String
    Dim lng主页ID As Long
    Dim strDrugIDs As String '要被排开的药品医嘱ID
    
    strAdviceDrugIDs = ""
    
    bln自动皮试 = Val(zlDatabase.GetPara("医嘱发送皮试限制", glngSys, IIF(bytType = 1, p门诊医嘱下达, p住院医嘱下达))) <> 0
    
    If Not bln自动皮试 Then Exit Function
    
    If bytType = 2 And bln发送输液药品 = False Then
        If Val(zlDatabase.GetPara("根据皮试结果限制医嘱发送类型", glngSys, p住院医嘱下达)) = 1 Then Exit Function
    End If
    
    strPatiFilter = IIF(bytType = 1, " And A.病人ID+0=[1] And  A.挂号单=[2]", " And A.病人ID=[1] And  A.主页ID=[2]")
    strPatiFilterAnd = IIF(bytType = 1, " And A.病人ID+0=B.病人ID And  A.挂号单=B.挂号单", " And A.病人ID=B.病人ID And  A.主页ID=B.主页ID")
    
    On Error GoTo errH
    
    '取出需要皮试的药品医嘱
    strSQL = "Select a.相关id,a.id,b.用法id,a.诊疗项目ID,a.开始执行时间,a.开始执行时间 as 过敏时间 From 病人医嘱记录 A,诊疗用法用量 B" & _
        " Where a.诊疗项目id = b.项目id And b.性质=0 and A.诊疗类别 IN('5','6') And a.医嘱状态<>4 and a.皮试阳性说明 is null and nvl(a.皮试结果,'空')<>'连续用药'" & strPatiFilter & _
        " union all " & _
        " Select a.相关id,a.id,b.用法id,a.诊疗项目ID,a.开始执行时间,a.开始执行时间 as 过敏时间 From 病人医嘱记录 A,药品用法用量 B" & _
        " Where a.收费细目id = b.药品id And b.性质=0 and A.诊疗类别 IN('5','6') And a.医嘱状态<>4 and a.皮试阳性说明 is null and nvl(a.皮试结果,'空')<>'连续用药'" & strPatiFilter
    strSQL = "Select a.相关id,a.id,a.用法id,a.诊疗项目ID,a.开始执行时间,a.过敏时间,1 as 结果 from (" & strSQL & ") a group by a.相关id,a.id,a.用法id,a.诊疗项目ID,a.开始执行时间,a.过敏时间 order by a.相关id"
    Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, IIF(bytType = 1, str标识, Val(str标识)))
    
    If Not rsDrug.EOF Then
        If bytType = 1 Then
            lng主页ID = lng挂号ID
        Else
            lng主页ID = Val(str标识)
        End If
        strSQL = "Select a.药物id,[3]+a.过敏时间 as 过敏时间, Nvl(a.结果,0) as 结果 From 病人过敏记录 A " & _
            " Where a.记录来源 = 2 And a.病人ID=[1] And a.主页ID=[2] And a.过敏时间 = (Select Max(x.过敏时间) From 病人过敏记录 X  Where x.病人id = a.病人id And x.主页id = a.主页id And x.药物id = a.药物id)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID, gint过敏登记有效天数)
        If Not rsTmp.EOF Then
            '有数据进行筛选
            Set rsDrug = zlDatabase.CopyNewRec(rsDrug)
            For i = 1 To rsDrug.RecordCount
                rsTmp.Filter = "药物id=" & rsDrug!诊疗项目ID
                If Not rsTmp.EOF Then
                    rsDrug!过敏时间 = rsTmp!过敏时间
                    rsDrug!结果 = rsTmp!结果
                End If
                rsDrug.MoveNext
            Next
            rsDrug.MoveFirst
            For i = 1 To rsDrug.RecordCount
                If rsDrug!过敏时间 < rsDrug!开始执行时间 Or Val(rsDrug!结果 & "") = 1 Then
                    If InStr("," & strNO医嘱IDs & ",", "," & rsDrug!相关ID & ",") = 0 Then
                        strNO医嘱IDs = strNO医嘱IDs & "," & rsDrug!相关ID
                    End If
                    strDrugIDs = strDrugIDs & "," & rsDrug!ID
                End If
                rsDrug.MoveNext
            Next
        Else
            '无过敏结果，则全部排除掉
            For i = 1 To rsDrug.RecordCount
                If InStr("," & strNO医嘱IDs & ",", "," & rsDrug!相关ID & ",") = 0 Then
                    strNO医嘱IDs = strNO医嘱IDs & "," & rsDrug!相关ID
                End If
                strDrugIDs = strDrugIDs & "," & rsDrug!ID
                rsDrug.MoveNext
            Next
        End If
    End If
    
    If strNO医嘱IDs <> "" Then
        'strAdviceDrugIDs    药品行的医嘱IDs
        strAdviceDrugIDs = Mid(strDrugIDs, 2)
        strSQL = "Select a.ID From 病人医嘱记录 a Where a.诊疗类别 IN ('5','6','E') And  Instr([3],','||Nvl(a.相关ID,ID)||',')>0 " & strPatiFilter
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, IIF(bytType = 1, str标识, Val(str标识)), "," & strNO医嘱IDs & ",")
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = strSQL & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        GetNoneSendID = Mid(strSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckApplication(ByVal lngID As Long, ByVal int场合 As Integer) As Boolean
'功能：检查诊疗项目是否对应了申请附项
'参数：lngID诊疗项目ID，int场合 1-门诊，2-住院
'返回：true=对应了申请附项
    Dim strSQL As String, rsTmp As Recordset
    
    On Error GoTo errH
    strSQL = "Select 1 From 病历单据应用 A, 病历单据附项 B Where a.病历文件id = b.文件id And a.诊疗项目id = [1] And 应用场合 = [2] And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckApplication", lngID, int场合)
    CheckApplication = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPrice(ByVal lngPatId As Long, ByVal lngPageId As Long, _
    ByVal lng诊疗项目ID As Long, ByVal str部分方法组合 As String, ByVal lng执行类型 As Long, _
    ByVal lng病人来源 As Long, ByVal lng执行科室ID As Long) As Double
'功能：获取检查项目的费用合计
'参数：部分方法组合格式：部位;方法1,方法2|部位2;方法1,方法3
'     执行类型：1常规，2床旁，3术中
'    病人来源：1-门诊，2住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim strSQL部位 As String
    Dim str部分方法 As String
    Dim str部位 As String
    Dim str方法 As String
    Dim objExp As Object
    Dim strGrad As String
    Dim strTmp As String
    
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = GetObject("", "zlPublicExpense.clsPublicExpense")
        If objExp Is Nothing Then Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
    End If
    
    '获取费用等级
    If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, lngPatId, lngPageId, "", strTmp, strTmp, strGrad) = False Then
        strGrad = ""
    End If
    
    strSQL部位 = " Select '' as 标本部位, '' as 检查方法 From Dual "
    
    If Not Trim(str部分方法组合) = "" Then
        For i = 0 To UBound(Split(str部分方法组合, "|"))
            str部分方法 = Split(str部分方法组合, "|")(i)
            str部位 = Split(str部分方法, ";")(0)
            str方法 = Split(str部分方法, ";")(1)
            For j = 0 To UBound(Split(str方法, ","))
                strSQL部位 = strSQL部位 & " Union All " & _
                    "Select '" & str部位 & "','" & Split(str方法, ",")(j) & "' From Dual "
            Next
        Next
    End If
    
    strSQL = "Select Sum(b.收费数量 * d.现价) As 合计" & vbNewLine & _
            "From (Select *" & vbNewLine & _
            "       From (Select c.诊疗项目id, c.收费项目id, c.检查部位, c.检查方法, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id," & vbNewLine & _
            "                     Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & vbNewLine & _
            "              From 诊疗收费关系 C," & vbNewLine & _
            "                   (" & strSQL部位 & ") A, 诊疗项目目录 d " & vbNewLine & _
            "              Where c.诊疗项目id = D.ID And d.计价性质 =0 and c.诊疗项目id = [1] And (a.标本部位 Is Null And [2] In (1, 2) And c.费用性质 = 1 Or" & vbNewLine & _
            "                    a.标本部位 = c.检查部位 And a.检查方法 = c.检查方法 And Nvl(c.费用性质, 0) = 0 Or" & vbNewLine & _
            "                    a.检查方法 Is Null And Nvl(c.费用性质, 0) = 0 And c.检查部位 Is Null And c.检查方法 Is Null) And" & vbNewLine & _
            "                    (c.适用科室id Is Null Or c.适用科室id = [3] And c.病人来源 = [4]))" & vbNewLine & _
            "       Where Nvl(适用科室id, 0) = Top) B, 收费项目目录 C, 收费价目 D" & vbNewLine & _
            "Where b.收费项目id = c.Id And b.收费项目id = d.收费细目id And" & vbNewLine & _
           IIF(strGrad = "", _
            "       D.价格等级 Is Null ", _
            "      ((instr( ';4;5;6;7;', ';' || C.类别 || ';')=0 And D.价格等级=[5]) " & _
                    " Or (D.价格等级 Is Null And Not Exists(Select 1 From 收费价目 " & _
                                    " Where b.收费项目Id=收费细目ID And Sysdate Between d.执行日期 And d.终止日期 And" & _
                                    " (instr( ';4;5;6;7;', ';' || C.类别 || ';')=0 And 价格等级=[5]) )))") & " And " & vbNewLine & _
            "      ((Sysdate Between d.执行日期 And d.终止日期) Or (Sysdate >= d.执行日期 And d.终止日期 Is Null)) And" & vbNewLine & _
            "      (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.服务对象 In ([4], 3) And" & vbNewLine & _
            "      (c.站点 = '0' Or c.站点 Is Null)"

    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPrice", lng诊疗项目ID, lng执行类型, lng执行科室ID, lng病人来源, strGrad)
    If rsTmp.RecordCount > 0 Then GetPrice = Val(rsTmp!合计 & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceStopTime(ByVal lng医嘱ID As Long) As String
'功能：对于填写了执行登记的医嘱，确保停止时间大于最近一次的要求执行时间
'返回：最早的停止时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select To_Char(Max(要求时间), 'YYYY-MM-DD HH24:MI') As 执行时间 From 病人医嘱执行 Where 医嘱id = [1]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
    If Not rsTmp.EOF Then GetAdviceStopTime = "" & rsTmp!执行时间
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CreateScript(Optional ByRef objVBA As Object, Optional ByRef objScript As clsScript) As Boolean
'功能：创建Script和VBA对象
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    err.Clear: On Error GoTo 0
    
    If Not objVBA Is Nothing Then
        objVBA.Language = "VBScript"
        Set objScript = New clsScript
        objVBA.AddObject "clsScript", objScript, True
        CreateScript = True
    End If
End Function

Public Function GetAdviceDiag(ByVal lng医嘱ID As Long, Optional ByRef str诊断 As String) As String
'功能：获得医嘱对应的诊断信息
'参数：str诊断=关联诊断的诊断名称字符串
'返回：关联诊断的ID，逗号分隔
    Dim rsTmp As Recordset, strSQL As String
    Dim strReturn As String
    
    strSQL = "Select  A.ID,a.诊断描述" & vbNewLine & _
            "From 病人诊断记录 A, 病人诊断医嘱 B" & vbNewLine & _
            "Where b.诊断id=a.id And  b.医嘱ID=[1]" & vbNewLine & _
            "Order By b.rowID"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱相关诊断", lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str诊断 = str诊断 & "," & rsTmp!诊断描述
            strReturn = strReturn & "," & rsTmp!ID
            rsTmp.MoveNext
        Loop
        str诊断 = Mid(str诊断, 2)
        strReturn = Mid(strReturn, 2)
    End If
    GetAdviceDiag = strReturn
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CanUnExec(ByVal datExec As Date, Optional ByVal datNow As Date) As Boolean
'功能：根据执行记录的执行时间判断能否取消执行或取消完成
'参数：datExec=执行记录的执行时间
'      datNow =当前时间
'返回：CanUnExec=true-可以取消执行，false-不可以取消执行

    Dim lngDatDiff As Long
    If datExec <> CDate(Format("0", "yyyy-MM-dd HH:mm")) Then
        If datNow = CDate(0) Then
            datNow = zlDatabase.Currentdate
        End If
        lngDatDiff = DateDiff("D", datExec, datNow)
        CanUnExec = lngDatDiff <= gint医嘱执行有效天数
    Else
        CanUnExec = True
    End If
    
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Function InitObjRecipeAudit(ByVal lngProgram As Long) As Boolean
    If gobjRecipeAudit Is Nothing Then
        On Error Resume Next
        Set gobjRecipeAudit = CreateObject("zl9RecipeAudit.clsBusiness")
        If Not gobjRecipeAudit Is Nothing Then
            Call gobjRecipeAudit.Init(gcnOracle, lngProgram = 1252)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjRecipeAudit = Not gobjRecipeAudit Is Nothing
End Function

Public Function Check中药存储库房(ByVal lngBegin As Long, ByVal lngEnd As Long, str中药名称 As String, ByRef vsAdvice As VSFlexGrid, ByVal bytMode As Byte, _
ByVal lng病人科室ID As Long, ByVal COL_类别 As Long, ByVal col_医嘱内容 As Long, ByVal COL_收费细目ID As Long, ByVal COL_执行科室ID As Long) As Boolean
'功能：检查指定的中药配方的所有药品是否设置了存储库房
'参数：str中药名称=返回未设置存储库房的中药名称串
'      bytMode 1-门诊场合,2-住院场合
    Dim lng中药房 As Long, strIDs As String
    Dim i As Long, lng执行科室ID As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colID As New Collection
    
    Check中药存储库房 = True
   
    With vsAdvice
         For i = lngBegin To lngEnd
             If .TextMatrix(i, COL_类别) = "7" Then
                 colID.Add .TextMatrix(i, col_医嘱内容), "C" & .TextMatrix(i, COL_收费细目ID)
                 strIDs = strIDs & "," & .TextMatrix(i, COL_收费细目ID)
                 If lng执行科室ID = 0 Then lng执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
             End If
         Next
    End With
    If strIDs <> "" Then
        strIDs = Mid(strIDs, 2)
        strSQL = "Select /*+ rule*/Column_Value as ID" & vbNewLine & _
                "From Table(f_Num2list([1])) B" & vbNewLine & _
                "Where Not Exists (Select 1 From 收费执行科室 A Where a.收费细目id = b.Column_Value" & vbNewLine & _
                " And Nvl(a.病人来源,[4]) = [4] And 执行科室id = [2] And Nvl(开单科室id, [3]) = [3])"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strIDs, lng执行科室ID, lng病人科室ID, bytMode)
        If rsTmp.RecordCount > 0 Then
            For i = 1 To rsTmp.RecordCount
                str中药名称 = str中药名称 & "," & colID("C" & rsTmp!ID)
                rsTmp.MoveNext
            Next
            str中药名称 = Mid(str中药名称, 2)
            Check中药存储库房 = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathAdviceIsExe(ByVal lng医嘱ID As Long) As Boolean
'功能:检查当前路径医嘱对应的路径项目是否已经执行登记
'参数:医嘱ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, p临床路径应用, 1)) Then
        '92225 补录医嘱未校对，不管项目是否执行登记都允许删除。
        strSQL = "Select Count(1) as 行数 From 病人医嘱记录 Where ID = [1] And 紧急标志 = 2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
        If rsTmp.RecordCount > 0 Then
            If NVL(rsTmp!行数, 0) = 1 Then Exit Function
        End If
        
        strSQL = "Select a.执行时间" & vbNewLine & _
                 "From 病人路径执行 A, (Select Min(a.路径执行id) As 路径执行id From 病人路径医嘱 A Where a.病人医嘱id = [1]) B" & vbNewLine & _
                 "Where a.Id = b.路径执行id"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!执行时间) Then
                CheckPathAdviceIsExe = True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathAdviceIsExeOut(ByVal lng医嘱ID As Long) As Boolean
'功能:检查当前路径医嘱对应的路径项目是否已经执行登记
'参数:医嘱ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, P门诊路径应用, 1)) Then
        strSQL = "Select a.执行时间" & vbNewLine & _
                 "From 病人门诊路径执行 A, (Select Min(a.路径执行id) As 路径执行id From 病人门诊路径医嘱 A Where a.病人医嘱id = [1]) B" & vbNewLine & _
                 "Where a.Id = b.路径执行id"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!执行时间) Then
                CheckPathAdviceIsExeOut = True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceInsure(ByVal int险类 As Integer, ByVal bln提醒对码 As Boolean, ByVal lng病人ID As Long, ByVal lng病人性质 As Long, _
   ByVal strIDs1 As String, ByVal strIDs2 As String, ByVal str医嘱内容 As String, Optional ByVal lng病人病区ID As Long) As String
'功能：医保病人下达医嘱时，医嘱录入后，对医嘱涉及的计价项目的保险对码情况进行检查
'参数：strIDs1:药品卫材的收费细目ID字符串（一组医嘱例如：青霉素+葡萄糖）:收费细目ID1,收费细目ID2,···
'      strIDs2 ：其他诊疗项目的诊疗项目ID（一组医嘱例如：输血项目+输血途径）:执行科室字符串 诊疗项目ID1:执行科室1,诊疗项目ID2:执行科室2,···
'      lng病人性质=1门诊，=2住院
'      str医嘱内容：用户提示时显示的医嘱内容
'      bln提醒对码=False 表示当前不继续检查，=True 继续检查
'返回：提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If gint医保对码 = 0 Or int险类 = 0 Or Not bln提醒对码 Then Exit Function
    If gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类) Then Exit Function
    
    
    If strIDs1 = "" And strIDs2 = "" Then Exit Function
    
    If strIDs1 <> "" Then
        If Mid(strIDs1, 1, 1) = "," Then strIDs1 = Mid(strIDs1, 2)
        strSQL = "Select Column_Value as 收费项目ID From Table(f_Num2list([1]))"
    End If
    If strIDs2 <> "" Then
        If Mid(strIDs2, 1, 1) = "," Then strIDs2 = Mid(strIDs2, 2)
        If strIDs1 <> "" Then strSQL = strSQL & " Union All "
        '由于没有加部位等条件，所以要用Distinct
        strSQL = strSQL & "Select 收费项目ID From (" & _
                "Select Distinct C.收费项目ID,C.适用科室id" & _
                " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                " From 诊疗收费关系 C,Table(f_Num2list2([2])) D Where C.诊疗项目ID=D.c1" & _
                "      And (C.适用科室ID is Null or C.适用科室ID = Nvl(D.c2,[4]) And C.病人来源 = " & IIF(lng病人性质 = 1, 1, 2) & ")" & _
                " ) Where Nvl(适用科室id, 0) = Top"
    End If
    
    strSQL = "Select /*+ RULE */ Distinct C.名称,B.收费细目ID" & _
        " From (" & strSQL & ") A,保险支付项目 B,收费项目目录 C" & _
        " Where A.收费项目ID=B.收费细目ID(+) And A.收费项目ID=C.ID" & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        " And B.险类(+)=[3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceInsure", strIDs1, strIDs2, int险类, lng病人病区ID)
    strSQL = "": i = 0
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!收费细目ID) Then
            If i = 8 Then
                strSQL = strSQL & vbCrLf & "… …"
                Exit Do
            End If
            strSQL = strSQL & vbCrLf & "●" & rsTmp!名称
            i = i + 1
        End If
        rsTmp.MoveNext
    Loop
    If strSQL <> "" Then
        CheckAdviceInsure = "当前病人是医保病人，但医嘱的以下计价项目没有设置对应的保险项目！" & vbCrLf & vbCrLf & _
            "医嘱内容：" & vbCrLf & str医嘱内容 & vbCrLf & vbCrLf & "计价项目：" & strSQL
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckWaitQuittance(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：发送转科医嘱时候 检查是否存在未审核销帐的单，据根据参数做相应提示
'返回：真假
    
    Dim strSQL As String
    Dim strDrug As String
    Dim rsTmp As ADODB.Recordset
    
    If gbyt转科时未审核销帐单据检查 = 0 Then Exit Function
    
    On Error GoTo ErrHand
    
    strSQL = " Select Distinct a.No, d.名称 项目, c.名称 As 部门" & _
        " From 住院费用记录 a, 病人费用销帐 b, 部门表 c, 收费项目目录 d" & _
        " Where a.Id = b.费用id And a.收费细目id = d.Id And b.审核部门id = c.Id(+)" & _
        " And b.审核时间 Is Null And a.病人id = [1] And Nvl(a.主页id, 0) = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckWaitQuittance", lng病人ID, lng主页ID)
    strDrug = ""
    Do While Not rsTmp.EOF
        If strDrug = "" Then
            strDrug = "单据[" & NVL(rsTmp!NO) & "]中的" & NVL(rsTmp!项目) & "：在" & NVL(rsTmp!部门, "[未知部门]") & "未审核"
        Else
            If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & "单据[" & NVL(rsTmp!NO) & "]中的" & NVL(rsTmp!项目) & "：在" & NVL(rsTmp!部门, "[未知部门]") & "未审核" & vbCrLf) = 0 Then
                If LenB(StrConv(strDrug & vbCrLf & "单据[" & NVL(rsTmp!NO) & "]中的" & NVL(rsTmp!项目) & "：在" & NVL(rsTmp!部门, "[未知部门]") & "未审核", vbFromUnicode)) <= 1000 Then
                    strDrug = strDrug & vbCrLf & "单据[" & NVL(rsTmp!NO) & "]中的" & NVL(rsTmp!项目) & "：在" & NVL(rsTmp!部门, "[未知部门]") & "未审核"
                Else
                    strDrug = strDrug & vbCrLf & "... ..."
                End If
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If strDrug <> "" Then
        If gbyt转科时未审核销帐单据检查 = 1 Then
            If MsgBox("该病人存在未审核销帐的单据：" & vbCrLf & vbCrLf & strDrug & vbCrLf & vbCrLf & "确定要发送转科医嘱吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckWaitQuittance = True
                Exit Function
            End If
        Else
            MsgBox "该病人存在未审核销帐的单据：" & vbCrLf & vbCrLf & strDrug & vbCrLf & vbCrLf & "不允发送转科医嘱。", vbInformation, gstrSysName
            CheckWaitQuittance = True
            Exit Function
        End If
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get医嘱附项内容(ByVal lng医嘱ID As Long, ByVal str中文名 As String) As String
'功能:根据医嘱ID，元素名称、返回医嘱的对应元素的申请附项内容
'参数:str中文名  诊治所见项目.中文名

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
     
    strSQL = "Select a.内容 From 病人医嘱附件 A, 诊治所见项目 B" & _
        " Where a.要素id = b.Id And a.医嘱id = [1] And b.中文名 = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID, str中文名)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = IIF(strTmp = "", "", strTmp & ",") & rsTmp!内容
            rsTmp.MoveNext
        Next
    End If
    
    Get医嘱附项内容 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get医技科室医嘱IDs(ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal lng执行科室ID As Long, ByVal bln住院 As Boolean, ByVal lng前提ID As Long) As String
'功能：在当前科室执行的所有医嘱
'参数：str就诊ID 如果是住院是主页id ，门诊是挂号单；bln住院=true 表示住院医嘱
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    strSQL = "Select ID From 病人医嘱记录 A Where a.病人id = [1] And a.执行科室id = [2]" & _
        IIF(bln住院, " And a.主页id = [3]", " And a.挂号单 = [3] ") & " Order By 开嘱时间 Desc"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng执行科室ID, IIF(bln住院, Val(str就诊ID), str就诊ID))
    Do While Not rsTmp.EOF
        strTmp = IIF(strTmp = "", "", strTmp & ",") & rsTmp!ID
        rsTmp.MoveNext
    Loop
    If Len(strTmp) > 4000 Then
        '如果不截取4000，后续所有SQL处理就比较麻烦了，从业务上看，只有门诊多次就诊使用相同的挂号单到医技站进行治疗时会产生这么多医技医嘱，例如血透等；所以只提取最近4000长度的ID即可，以前的医嘱没有什么作用
        strTmp = Mid(strTmp, 1, 3980)
        strTmp = Mid(strTmp, 1, InStrRev(strTmp, ",") - 1)
    End If
    
    If strTmp <> "" Then
        If InStr("," & strTmp & ",", "," & lng前提ID & ",") = 0 Then
            strTmp = strTmp & "," & lng前提ID
        End If
    Else
        strTmp = lng前提ID
    End If
    
    Get医技科室医嘱IDs = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncTraReaction(ByVal lng医嘱ID As Long, ByVal lngMoudle As Long, ByVal blnMoved As Boolean) As Boolean
'功能：输血反应
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long, lng病人来源 As Long
    

    If InitObjBlood(True) = False Then Exit Function
    
    On Error GoTo errH
    strSQL = "Select B.病人ID,B.主页ID,B.病人来源,A.ID 就诊ID From 病人挂号记录 A,病人医嘱记录 B where B.挂号单=A.NO(+) And  B.id=[1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "输出反应", lng医嘱ID)
    lng病人ID = Val("" & rsTemp!病人ID)
    If IsNull(rsTemp!主页ID) Then
        lng主页ID = Val("" & rsTemp!就诊ID)
    Else
        lng主页ID = Val("" & rsTemp!主页ID)
    End If
    lng病人来源 = Val("" & rsTemp!病人来源)
    Call gobjPublicBlood.zlShowBloodReaction(Nothing, glngSys, lngMoudle, 1, lng病人ID, lng主页ID, lng病人来源)
    
    FuncTraReaction = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Read发药窗口(lngID As Long) As ADODB.Recordset
'功能：获取指定药房的发药窗口
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 发药窗口 Where 药房ID=[1] Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngID)
    Set Read发药窗口 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人打印记录DelSQL(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal intBaby As Integer, Optional ByVal int期效 As Integer, _
    Optional ByVal lng医嘱ID As Long, Optional ByVal str医嘱IDs As String, Optional ByVal blnHaveBaby As Boolean, Optional ByRef strMsg As String) As String
'功能：获取病人将要删除  病人医嘱打印  表中数据的过程SQL，包括删除预打印的记录 病人医嘱打印.打印时间 is null
'      如："Zl_病人医嘱打印_Delete(718,1,0,0,1)|Zl_病人医嘱打印_Delete(718,1,0,1,1)|Zl_病人医嘱打印_Delete(718,1,1,0,1)";
'参数：lng病人ID，lng主页ID，intBaby，int期效，lng医嘱ID，str医嘱IDs
'      blnHaveBaby 当前病人是否有婴儿
'intType 调用时机：
'      intType=2通用界面作废医嘱时，必须传的参数有 intType，lng病人ID，lng主页ID，str医嘱IDs，blnHaveBaby
'      intType=3医技站作废医嘱时(只能单条作废)，必须传的参数有 intType，lng病人ID，lng主页ID，intBaby，int期效，lng医嘱ID，blnHaveBaby
'      intType=4工作站主界面删除医嘱时，必须传的参数有 intType，lng病人ID，lng主页ID，intBaby，int期效，str医嘱IDs，blnHaveBaby
'          intType=5屏蔽打印时，必须传的参数有 intType，lng病人ID，lng主页ID，intBaby，int期效，lng医嘱ID，blnHaveBaby
'          intType=5的情况暂时没使用，在以前的程序中已经完全控制了。
'      strMsg 返回参数，例“张三的长期医嘱单打印记录将从第3页起被清除。”
'说明：intType in (3,4,5)时返回的过程SQL只有一条，外面可不做处理直接用。

    Dim rsTmp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim strSQL As String, strMsgTmp As String
    Dim str婴儿姓名 As String
    Dim i As Integer
    Dim intNoP As Integer '1－清除未打印但已生成的记录，0－清除已经打过的记录及之后的
    
    On Error GoTo errH
    
    If blnHaveBaby Then
        If intType = 1 Or intType = 2 Then
            strSQL = "Select 序号,婴儿姓名 as 姓名 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2]"
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "Get病人预打印记录DelSQL", lng病人ID, lng主页ID)
        Else
            strSQL = "Select 婴儿姓名 as 姓名 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] and 序号=[3]"
            Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "Get病人预打印记录DelSQL", lng病人ID, lng主页ID, intBaby)
            If Not rsBaby.EOF Then str婴儿姓名 = rsBaby!姓名 & ""
        End If
    End If
    
    'SQL中取出的  位置  是一组医嘱中的最小位置，对于一组医嘱多条的情况不用单独处理，在删除医嘱预打印记录时一定是以一组医嘱为单位。
    '在删除已经打印的记录时，按页删除，药品医嘱可能会有特殊情况（药品医嘱占用两行），但不影响正确性。
    Select Case intType
    Case 1, 2
        strSQL = "Select a.婴儿,a.期效, Min(LPad(页号,4,'0')||LPad(行号,3,'0')) As 位置,Min(a.页号) As 页号,min(a.打印时间) as 打印时间" & _
            " From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
            " Where a.病人id = b.病人id and b.id=a.医嘱id And a.主页id = b.主页id and a.病人id=[1] and a.主页id=[2]" & _
            " And b.id In (Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))" & _
            " Group By a.婴儿,a.期效 having Min(a.页号)>0"
    Case 3, 5
        strSQL = "Select Min(LPad(页号,4,'0')||LPad(行号,3,'0')) As 位置,Min(a.页号) As 页号,min(a.打印时间) as 打印时间" & _
            " From 病人医嘱打印 A, 病人医嘱记录 B" & _
            " Where a.病人id = b.病人id And a.主页id = b.主页id and b.id=a.医嘱id and a.病人id=[1] and a.主页id=[2]" & _
            " And a.婴儿=[3] and a.期效=[4] and (b.id =[5] or b.相关id=[5]) having Min(a.页号)>0"
    Case 4
        strSQL = "Select Min(LPad(页号,4,'0')||LPad(行号,3,'0')) As 位置,Min(a.页号) As 页号,min(a.打印时间) as 打印时间" & _
            " From 病人医嘱打印 A, 病人医嘱记录 B" & _
            " Where a.病人id = b.病人id And a.主页id = b.主页id and b.id=a.医嘱id and a.病人id=[1] and a.主页id=[2]" & _
            " And a.婴儿=[3] and a.期效=[4] and (b.id =(Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))" & _
            " or b.相关id=(Select Column_Value From Table(Cast(f_Num2list([6]) As zlTools.t_Numlist)))) having Min(a.页号)>0"
    End Select
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get病人预打印记录DelSQL", lng病人ID, lng主页ID, intBaby, int期效, lng医嘱ID, str医嘱IDs)
    
    ' IsNull(rsTmp!打印时间) 则只需要清除记录，在界面上不用提示
    strSQL = ""
    If Not rsTmp.EOF Then
        Select Case intType
        Case 1, 2
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!婴儿 & "") <> 0 Then
                    rsBaby.Filter = "序号=" & Val(rsTmp!婴儿 & "")
                    If Not rsBaby.EOF Then str婴儿姓名 = rsBaby!姓名 & ""
                End If
                intNoP = 1
                If Not IsNull(rsTmp!打印时间) Then
                    intNoP = 0
                    strMsgTmp = IIF(strMsgTmp = "", "", strMsgTmp & vbCrLf) & _
                        "该病人" & IIF(str婴儿姓名 = "", "", "婴儿-" & str婴儿姓名) & "的" & IIF(Val(rsTmp!期效 & "") = 0, "长期", "临时") & "医嘱单的打印记录将从第" & _
                        Val(rsTmp!页号 & "") & "页开始被清除。"
                    str婴儿姓名 = ""
                End If
                strSQL = strSQL & "|" & "Zl_病人医嘱打印_Delete(" & lng病人ID & "," & lng主页ID & "," & Val(rsTmp!婴儿 & "") & "," & Val(rsTmp!期效 & "") & "," & Val(rsTmp!页号 & "") & ",'" & rsTmp!位置 & "')"
                rsTmp.MoveNext
            Next
            strSQL = Mid(strSQL, 2)
        Case 3, 4, 5
            intNoP = 1
            If Not IsNull(rsTmp!打印时间) Then
                intNoP = 0
                strMsgTmp = "该病人" & IIF(str婴儿姓名 = "", "", "婴儿-" & str婴儿姓名) & "的" & IIF(int期效 = 0, "长期", "临时") & "医嘱单的打印记录将从第" & _
                        Val(rsTmp!页号 & "") & "页开始被清除。"
            End If
            strSQL = "Zl_病人医嘱打印_Delete(" & lng病人ID & "," & lng主页ID & "," & intBaby & "," & int期效 & "," & Val(rsTmp!页号 & "") & ",'" & rsTmp!位置 & "')"
        End Select
    End If
    
    strMsg = ""
    
    strMsg = strMsgTmp
    
    Get病人打印记录DelSQL = strSQL
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal bytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    
    strSQL = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B," & IIF(bytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIF(bytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetStockCheck")
    For i = 1 To rsTmp.RecordCount
        colStock.Add NVL(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function ItemExistInsure(ByVal lng病人ID As Long, ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As Boolean
'功能：判断收费项目是否设置了保险支付项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    If gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类) Then
        ItemExistInsure = True: Exit Function
    End If
    
    strSQL = "Select 1 From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng收费细目ID, int险类)
    ItemExistInsure = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng病人ID As Long, ByVal byt来源 As Byte) As Currency
'功能:获取指定病人的记帐划价单金额合计
'参数:byt来源:1-门诊，2-住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTab As String
    strTab = IIF(byt来源 = 1, "门诊费用记录", "住院费用记录")
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(Sum(实收金额),0) As 划价费用合计 From " & strTab & " Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]"
    
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 项目Id,使用限量,已用数量,使用限量-已用数量 可用数量 From 病人审批项目 Where 病人ID=[1] And 主页ID=[2]"
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRevoke(ByVal lng医嘱ID As Long) As Boolean
'功能：(门诊)对要作废的医嘱对应的费用的结帐情况进行检查
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng发送号 As Long
    
    If gbytBillOpt = 0 Then
        CheckAdviceBalanceRevoke = True
        Exit Function
    End If
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSQL = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    '部份条件见"ZL_病人医嘱记录_作废"
    strSQL = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1 And A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID" & _
        " And B.记录性质=2 And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where 系统=[3] And 模块 is NULL And Nvl(私有,0)=0 And 参数号=68 And Nvl(参数值,'0')='0')" & _
            " )" & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID, glngSys)
    If Not rsTmp.EOF Then
        If gbytBillOpt = 1 Then
            If MsgBox("要作废医嘱的对应费用中存在已结帐的费用，确实要作废吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "要作废医嘱的对应费用中存在已结帐的费用，不能作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckAdviceBalanceRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBillingRevoke(ByVal lng医嘱ID As Long) As Boolean
'功能：(门诊)对要作废的医嘱对应的记帐费用的审核情况进行检查
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng发送号 As Long
    
    On Error GoTo errH
    
    '医嘱ID为传入值的这条医嘱不一定发送了的,甚至无发送。
    strSQL = "Select Distinct 发送号 From 病人医嘱发送" & _
        " Where 医嘱ID IN(Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
    If rsTmp.EOF Then Exit Function
    lng发送号 = rsTmp!发送号
    
    '部份条件见"ZL_病人医嘱记录_作废"
    strSQL = "Select A.NO,A.序号" & _
        " From 门诊费用记录 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 I" & _
        " Where A.NO=B.NO And A.记录性质 IN(2,12) And A.记录状态=1" & _
        " And A.划价人 Is Not NULL And A.划价人<>A.操作员姓名" & _
        " And A.医嘱序号=B.医嘱ID And B.医嘱ID=C.ID And B.记录性质=2" & _
        " And C.诊疗项目ID=I.ID And B.发送号=[1] And (C.ID=[2] Or C.相关ID=[2])" & _
        " And (" & _
            " A.收费类别 Not In ('5','6','7','E')" & _
            " Or A.收费类别='E' And I.操作类型 Not In ('2','3','4')" & _
            " Or A.收费类别 In ('5','6','7') And Nvl(A.执行状态,0)=0" & _
            " Or Exists(Select 1 From zlParameters Where 系统=[3] And 模块 is NULL And Nvl(私有,0)=0 And 参数号=68 And Nvl(参数值,'0')='0')" & _
            " )"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID, glngSys)
    If Not rsTmp.EOF Then Exit Function
    
    CheckAdviceBillingRevoke = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'功能：获取指定病人的剩余额
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(费用余额,0) as 费用余额,Nvl(预交余额,0) as 预交余额" & _
            " From 病人余额 Where 性质=1 And 类型 = " & IIF(lng主页ID = 0, 1, 2) & " And 病人ID= [1] "
    
    If curModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSQL = strSQL & " Union All  Select -1* " & curModiMoney & " as 费用余额,0 as 预交余额 From Dual"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
            
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If lng主页ID <> 0 Then
        strSQL = strSQL & " Union All " & _
            " Select -1*Nvl(Sum(金额),0) as 费用余额,0 as 预交余额" & _
            " From 保险模拟结算 Where 病人ID=[1] And 主页ID=[2]"
        strSQL = "Select Sum(费用余额) as 费用余额,Sum(预交余额) as 预交余额 From (" & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Set GetMoneyInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceBalanceRoll(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, Optional ByVal blnBat As Boolean) As Boolean
'功能：(住院)对要回退的医嘱对应的费用的结帐情况进行检查(一个病人一次住院的)
'参数：blnBat=是否要进行批量回退
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intInsure As Integer
    
    On Error GoTo errH
        
    '取要回退的记帐NO
    If blnBat Then
        strSQL = "Select Distinct 医嘱ID,NO From 病人医嘱发送 Where 记录性质=2 And 发送号=[1]"
    Else
        strSQL = "Select Distinct A.医嘱ID,A.NO From 病人医嘱发送 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID And A.记录性质=2 And A.发送号=[1] And (B.ID=[2] Or B.相关ID=[2])"
    End If
    '取这些NO的结帐情况(非划价未销帐)
    strSQL = "Select A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 住院费用记录 A,(" & strSQL & ") B Where A.NO=B.NO And A.医嘱序号=B.医嘱ID And A.记录性质 IN(2,12) " & _
        " Group by A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID)
    If Not rsTmp.EOF Then
        strSQL = "Select A.病人ID,A.险类 From 病案主页 A,病人医嘱记录 B" & _
            " Where Rownum=1 And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng医嘱ID)
        If Not rsTmp.EOF Then intInsure = NVL(rsTmp!险类, 0)
        If intInsure <> 0 Then '先对医保的限制进行检查
            If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, rsTmp!病人ID, intInsure) Then
                MsgBox "该病人为医保病人，要回退医嘱的发送费用中存在已结帐的费用，不能回退。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If gbytBillOpt <> 0 Then
            If gbytBillOpt = 1 Then
                If MsgBox("要回退医嘱的发送费用中存在已结帐的费用，确实要回退吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gbytBillOpt = 2 Then
                MsgBox "要回退医嘱的发送费用中存在已结帐的费用，不能回退。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    CheckAdviceBalanceRoll = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAdviceDrugSurplus(ByVal lng发送号 As Long, Optional ByVal lng医嘱ID As Long) As String
'功能：检查待回退药品医嘱的数量是否大于当前留存的数量
'参数：lng发送号=要回退的发送号
'      lng医嘱ID=要回退的一组药品医嘱的ID，如果不指定由表示批量回退多条医嘱
'返回：提示信息
'说明：护士不能回退医生的操作，所以只涉及住院费用记录(医生才可能发送临嘱为门诊费用)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select C.医嘱内容 as 药品,A.收费细目ID as 药品ID,A.病人病区ID as 病区ID,A.执行部门ID as 库房ID,Sum(A.数次) as 回退数量" & _
        " From 住院费用记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
        " Where A.医嘱序号=B.医嘱ID And A.NO=B.NO And A.记录性质=B.记录性质" & _
        " And B.医嘱ID=C.ID And A.收费类别 In('5','6') And A.价格父号 Is Null" & _
        " And B.发送号=[1] And C.诊疗类别 IN('5','6') And (C.相关ID=[2] Or [2]=0)" & _
        " Group by C.医嘱内容,A.收费细目ID,A.病人病区ID,A.执行部门ID"
    strSQL = _
        " Select A.药品,D.名称 as 库房,C.住院包装,C.住院单位,A.回退数量,B.留存数量" & _
        " From (" & strSQL & ") A,药品留存计划 B,药品规格 C,部门表 D" & _
        " Where A.库房ID=D.ID And A.药品ID=C.药品ID" & _
        " And A.病区ID=B.部门ID(+) And A.库房ID=B.库房ID(+) And A.药品ID=B.药品ID(+) And B.状态(+)=0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckAdviceDrugSurplus", lng发送号, lng医嘱ID)
    Do While Not rsTmp.EOF
        If NVL(rsTmp!回退数量, 0) > NVL(rsTmp!留存数量, 0) And NVL(rsTmp!留存数量, 0) <> 0 Then
            strMsg = strMsg & vbCrLf & "●[" & rsTmp!药品 & "]从""" & rsTmp!库房 & """的回退数量 " & _
                FormatEx(NVL(rsTmp!回退数量, 0) / NVL(rsTmp!住院包装, 1), 5) & rsTmp!住院单位 & "，当前留存数量 " & _
                FormatEx(NVL(rsTmp!留存数量, 0) / NVL(rsTmp!住院包装, 1), 5) & rsTmp!住院单位
        End If
        rsTmp.MoveNext
    Loop
    
    If strMsg <> "" Then strMsg = "下列药品的回退数量大于留存数量：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "要继续吗？"
    CheckAdviceDrugSurplus = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CanEditBloodAdvice(ByVal lngID As Long, ByVal int审核状态 As Integer, ByVal bln急 As Boolean, Optional ByVal bln用血 As Boolean = False, Optional ByVal blnMsg As Boolean = True) As Boolean
'功能：输血医嘱可否编辑 int审核状态 取值有：0，1，2，4，5；bln急 是否是紧急医嘱。（只检查备血医嘱，这里主要是以前老备血流程的检查）
    Dim strMsg As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    CanEditBloodAdvice = True
    
    If int审核状态 = 0 Or int审核状态 = 1 Then Exit Function
    On Error GoTo ErrHand
    strSQL = "Select 医嘱ID from 病人医嘱状态 where 医嘱ID=[1] and 操作类型=[2]"
    If gbln血库系统 Then
        If bln用血 = True Then Exit Function '用血医嘱允许修改
        If int审核状态 = 5 Or int审核状态 = 2 Then
            strMsg = "该输血申请已经被血库接收" & IIF(int审核状态 = 5, "正在配血", "并且已完成配血") & "，不允许修改，若要修改请与输血科联系。"
        ElseIf int审核状态 = 6 Then
            strMsg = "该输血申请已经被血库接收并且已停止配血，不允许修改，若要修改请与输血科联系。"
        ElseIf int审核状态 = 4 Or int审核状态 = 7 Then
            If Not bln急 And gbln输血分级管理 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "输血申请", lngID, IIF(int审核状态 = 4, 11, 18))
                If Not rsTmp.EOF Then
                    If int审核状态 = 4 Then
                        strMsg = "该输血申请已经完成审核，不允许修改，若要修改请在输血审核管理中回退审核操作。"
                    Else
                        strMsg = "该输血申请已经开始审核，不允许修改，若要修改请在输血审核管理中回退审核操作。"
                    End If
                End If
            End If
        End If
    Else
        If int审核状态 = 2 Then
            strMsg = "该输血申请已经审核，不允许修改，若要修改请在输血审核管理中回退审核操作。"
        End If
    End If
    
    If strMsg <> "" Then
        If blnMsg = True Then
            MsgBox strMsg, vbInformation, "输血申请"
        End If
        CanEditBloodAdvice = False
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckCHLimited(ByVal lngRow As Long, ByVal int付数 As Integer, ByRef blnOutOfRange As Boolean, ByRef vsAdvice As VSFlexGrid, ByVal COL_相关ID As Long, ByVal COL_诊疗项目ID As Long, ByVal COL_类别 As Long, ByVal COL_单量 As Long) As Boolean
'功能：检查中药配方每味药的处方限量
'参数：blnMsg 是否弹出消息提示框
'      blnOutOfRange 传出参数，调用程序可以根据该参数做进一步处理，例如将 .TextMatrix(i, COL_是否超量)="1"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lng中药数量 As Long
    Dim colAmount As New Collection, strIDs As String

    CheckCHLimited = True
    
    On Error GoTo errH
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_类别) = "7" Then
                
                    '草药可以按规格下达后，计算用量时要特殊处理
                    If InStr("," & strIDs & ",", "," & .TextMatrix(i, COL_诊疗项目ID) & ",") = 0 Then
                        strIDs = strIDs & "," & .TextMatrix(i, COL_诊疗项目ID)
                        Call colAmount.Add(Val(.TextMatrix(i, COL_单量)), "_" & .TextMatrix(i, COL_诊疗项目ID))
                    Else
                        lng中药数量 = colAmount("_" & .TextMatrix(i, COL_诊疗项目ID))
                        lng中药数量 = lng中药数量 + Val(.TextMatrix(i, COL_单量))
                        
                        Call colAmount.Remove("_" & .TextMatrix(i, COL_诊疗项目ID))
                        Call colAmount.Add(lng中药数量, "_" & .TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            Else
                Exit For
            End If
        Next
    End With
    If strIDs = "" Then Exit Function
    strIDs = Mid(strIDs, 2)
        
    strSQL = "Select /*+ rule*/ A.ID,A.名称,A.计算单位,B.处方限量 From 诊疗项目目录 A,药品特性 B Where A.ID=B.药名ID And Nvl(B.处方限量,0)<>0" & _
            " And A.ID IN (Select Column_Value From Table(f_Num2list([1])))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckCHLimited", strIDs)
             
    For i = 1 To rsTmp.RecordCount
        If int付数 * colAmount("_" & rsTmp!ID) > rsTmp!处方限量 Then
            blnOutOfRange = True: Exit Function
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitCardRs(ByRef rsCard As ADODB.Recordset)
'功能：初始化记录集，医嘱编辑界面下方卡片控件信息
'说明：门诊住院通用，修改此方法时注意考虑门诊住院两种情况
    Set rsCard = New ADODB.Recordset
    
    With rsCard.Fields
        .Append "是否新增", adInteger '0-新增,1-修改
        .Append "是否保存", adInteger '默认为 0 申请单界面保存按钮不可用 0－不可用，1－可用
        .Append "紧急医嘱", adInteger '默认为 0
        .Append "发送为不收费记帐单", adInteger '默认为 0
        .Append "停止所有长嘱", adInteger '默认为 0
        .Append "生效时间", adVarChar, 20 '即医嘱开始执行时间
        .Append "手术时间", adVarChar, 20
        .Append "手术情况", adInteger '默认为 0
        .Append "医生嘱托", adVarChar, 500
        .Append "执行科室ID", adBigInt '医嘱执行科室的 部门ID
        .Append "附加执行ID", adBigInt '麻醉医嘱执行科室的 部门ID，输血申请时为输血途径执行科室id
        .Append "附加执行科室ID", adBigInt '麻醉医嘱执行科室的 部门ID，输血申请时为输血途径执行科室id
        .Append "项目IDs", adVarChar, 500 '一组医嘱中的项目ID串，顺序固定，即为医嘱行的顺序，如：3542,2532,5478,......
        .Append "会诊科室IDs", adVarChar, 500 '参加会诊的科室id串
        .Append "安排时间", adVarChar, 20 '手术医嘱则是手术时间，输血医嘱则是输血时间，会诊会诊医嘱的开始时间
        .Append "用药理由", adVarChar, 500 '输血申请时为 输血原因
        .Append "医嘱标志", adInteger '0-完全新增,1-确定主项目新增，2-医嘱复制新增，3-修改已经改医嘱，4-修改原始医嘱
        .Append "主项目ID", adBigInt   '主医嘱的诊疗项目ID
        .Append "总量", adVarChar, 10 '输血申请时为 输血总量
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub DeleteRsExec(ByRef rsExec As Recordset, ByRef lng医嘱ID As Long)
'功能：根据传入的医嘱ID，删除医嘱执行计价中的该医嘱对应的数据
    rsExec.Filter = "医嘱ID=" & lng医嘱ID
    Do While Not rsExec.EOF
        rsExec.Delete
        rsExec.MoveNext
    Loop
End Sub

Public Function ApplyInPacs(frmParent As Object, ByRef lng申请序号 As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng婴儿序号 As Long, ByVal lng病人性质 As Long, ByVal lng医嘱ID As Long, _
    ByVal lng医护科室ID As Long, ByVal lng科室id As Long, ByVal lng病区ID As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, Optional ByRef clsMipModule As Object, Optional ByVal lng项目id As Long, Optional ByVal lng前提ID As Long) As Long
'功能：调用检查申请单
'参数： lng医嘱ID=修改申请单时当前行的医嘱ID,lng申请序号 =当前修改行的申请序号
'       lng病人性质 0-普通住院病人,1-门诊留观病人,2-住院留观病人
'       lng医护科室ID 医护科室ID
'       lng科室ID 如果是转出病人，则为原科室ID；
'       lng病区ID 如果是转出病人，则为原病区ID；
'       objVBA objScript rsDefine VB对像和记录集用于产生医嘱内容文本；blnMoved 数据是否已经转出；clsMipModule 消息对象
'返回：申请序号
    Dim objPacspplication As New clsPacsApplication
    Dim objAppPages()  As clsApplicationData
    Dim rsPati As Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim lngAdviceID As Long
    Dim strMsg As String
    Dim strExtra As String
    Dim strTmp As String
    Dim str部位 As String, str方法 As String
    Dim strTmp方法 As String
    Dim objTmp As New clsApplicationData
    Dim str检查入院诊断 As String
    Dim bln中医 As Boolean
    Dim str类型 As String
    Dim str摘要 As String
    Dim strItems As String
    Dim strRISDel As String
    Dim strRISAdd As String 'RIS参数，格式：医嘱ID:诊疗项目ID,....
    Dim arrSQL() As String
    Dim lng假医嘱ID As Long  '避免医嘱ID序列值的浪费，在最后提交事务时产生真的医嘱ID
    Dim blnDo As Long
    Dim strTabAdvice As String '医嘱信息临时表
    Dim strTxtAdvice As String '医嘱内容
    Dim bln提醒对码 As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim rsPrice As ADODB.Recordset
    Dim lng医嘱序号 As Long
    
    On Error GoTo errH
    
    str检查入院诊断 = zlDatabase.GetPara("要求输入入院诊断", glngSys, p住院医嘱下达)
    '诊断检查
    If InStr(str检查入院诊断, "D") > 0 Then
        bln中医 = Sys.DeptHaveProperty(lng科室id, "中医科")
        str类型 = IIF(bln中医, "2,12", "2")
        If Not ExistsDiagNoses(lng病人ID, lng主页ID, str类型) Then
            strMsg = "病人的入院诊断还没有输入，请先输入病人的入院诊断再下达检查申请。"
        End If
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select A.住院号, A.当前床号, A.出生日期, Nvl(B.姓名, A.姓名) as  姓名, Nvl(B.性别, A.性别) as  性别, Nvl(B.年龄, A.年龄) as 年龄, A.门诊号, A.健康号,b.险类,b.费别,b.病人性质" & vbNewLine & _
            "From 病人信息 A, 病案主页 B" & vbNewLine & _
            "Where A.病人id = B.病人id And A.病人id = [1] And B.主页id = [2]"
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ApplyInPacs", lng病人ID, lng主页ID)
    
    If rsPati.RecordCount = 0 Then
        MsgBox "未能正确读取病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    '初始化检查申请单对象
    Call objTmp.MakePacsData(lng申请序号, objAppPages())
    Call objPacspplication.InitComponents(Get开嘱科室ID(UserInfo.ID, 0, lng科室id, IIF(lng病人性质 = 1, 1, 2)), frmParent)
    If objPacspplication.ShowApplicationForm(lng病人ID, IIF(lng病人性质 = 1, 1, 2), 0, lng主页ID, IIF(lng申请序号 = 0, lng医嘱ID, lng申请序号), objAppPages(), lng婴儿序号, , lng项目id) Then
        On Error GoTo errH
        If lng申请序号 = 0 Then
            strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ApplyInPacs")
            lng申请序号 = Val(rsTmp!申请序号)
        End If
        On Error Resume Next
        err.Clear
        If UBound(objAppPages) >= 0 Then
            If err.Number = 0 Then
                On Error GoTo errH
                ReDim Preserve arrSQL(0)
                lng医嘱序号 = GetMaxAdviceNO(lng病人ID, lng主页ID, lng婴儿序号)     '病人医嘱记录.序号，递增
                For i = 0 To UBound(objAppPages)
                    If lng医嘱ID = 0 Or objAppPages(i).blnIsModify = True Then
                        '调用 zl_AdviceCheck 函数检查
                        str摘要 = ""
                        str摘要 = gclsInsure.GetItemInfo(Val(rsPati!险类 & ""), lng病人ID, 0, "", 0, "", CStr(objAppPages(i).lngProjectId))
                        objAppPages(i).strAbstract = str摘要
                        strTmp = objAppPages(i).strPartMethod
                        If strTmp <> "" Then
                            For k = 0 To UBound(Split(strTmp, "|"))
                                str部位 = Split(Split(strTmp, "|")(k), ";")(0)
                                strTmp方法 = Split(Split(strTmp, "|")(k), ";")(1)
                                For j = 0 To UBound(Split(strTmp方法, ","))
                                    str方法 = Split(strTmp方法, ",")(j)
                                    strExtra = strExtra & "," & str部位 & ":" & str方法
                                Next
                            Next
                            strExtra = "||0||0||" & Mid(strExtra, 2) & "||0"
                        End If
                        strExtra = str摘要 & strExtra
                        strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 2, lng病人ID, lng主页ID, Val(rsPati!险类 & ""), 1, "D", objAppPages(i).lngProjectId, _
                           objAppPages(i).lngRequestRoomId, UserInfo.姓名, IIF(objAppPages(i).lngExeRoomId <= 0, 0, objAppPages(i).lngExeRoomId), IIF(objAppPages(i).lngExeRoomId <= 0, "5", objAppPages(i).lngExeRoomType), 0, 0, strExtra)
                        
                        If Not rsTmp.EOF Then
                            strMsg = NVL(rsTmp!结果)
                            If strMsg <> "" Then
                                Select Case Val(Split(strMsg, "|")(0))
                                Case 1 '提示
                                    If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strMsg = "": Exit Function
                                    End If
                                Case 2 '禁止
                                    MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                    strMsg = "": Exit Function
                                End Select
                                strMsg = ""
                            End If
                        End If
                        
                        blnDo = GetPacsAdviceSQLData(objAppPages(i), 0, lng申请序号, lng病人ID, lng主页ID, lng婴儿序号, "", lng科室id, objVBA, objScript, rsDefine, _
                            strRISDel, strRISAdd, arrSQL, lng假医嘱ID, strTabAdvice, strTxtAdvice, lng前提ID, lng医嘱序号)
                        If Not blnDo Then Exit Function
                        
                        strItems = objAppPages(i).lngProjectId & ":" & objAppPages(i).lngExeRoomId
                        '医保对码检查
                        If gint医保对码 = 2 Then bln提醒对码 = True
                        strMsg = CheckAdviceInsure(Val(rsPati!险类 & ""), bln提醒对码, lng病人ID, Val(rsPati!病人性质 & ""), "", strItems, Left(strTxtAdvice, 50))
                        If strMsg <> "" Then
                            If gint医保对码 = 1 Then
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                If vMsg = vbIgnore Then bln提醒对码 = False
                            ElseIf gint医保对码 = 2 Then
                                MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                                Exit Function
                            End If
                            strMsg = ""
                        End If
                        If Val(rsPati!险类 & "") <> 0 Then
                            If gclsInsure.GetCapability(support实时监控, lng病人ID, Val(rsPati!险类 & "")) Then
                                If MakePriceRecord申请单("22", lng病人ID, lng主页ID, strTabAdvice, strItems, rsPati!费别 & "", lng科室id, rsPrice) Then
                                    If Not gclsInsure.CheckItem(Val(rsPati!险类 & ""), 1, 0, rsPrice) Then
                                        MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的PACS申请单不能保存。", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        strTabAdvice = ""
                    End If
                Next
                
                If Not SavePacsData(0, Mid(strRISDel, 2), Mid(strRISAdd, 2), arrSQL, lng假医嘱ID, lngAdviceID) Then
                     Exit Function
                End If
                ApplyInPacs = lng申请序号
                Call ZLHIS_CIS_001(clsMipModule, lng病人ID, rsPati!姓名 & "", rsPati!住院号 & "", , IIF(lng病人性质 = 1, 1, 2), lng主页ID, lng病区ID, , lng科室id, "", , rsPati!当前床号 & "", _
                    lngAdviceID, 0, 1, "D", "", UserInfo.姓名, Format(objAppPages(0).strRequestTime, "yyyy-MM-dd HH:mm:ss"), lng科室id, "", , , "")
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ApplyOutPacs(frmParent As Object, ByRef lng申请序号 As Long, ByVal lng病人ID As Long, ByVal str挂号单 As String, ByVal lng医嘱ID As Long, ByVal lng挂号科室ID As Long, _
    ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, ByVal blnMoved As Boolean, Optional ByVal lng项目id As Long, Optional ByVal lng前提ID As Long) As Long
'功能：调用检查申请单
'参数：lng医嘱ID=修改申请单时当前行的医嘱ID,lng申请序号 =当前修改行的申请序号
'       lng挂号科室ID 挂号执行科室ID
'       objVBA objScript rsDefine VB对像和记录集用于产生医嘱内容文本；blnMoved 数据是否已经转出；
    Dim objPacspplication As New clsPacsApplication
    Dim objAppPages()  As clsApplicationData
    Dim rsPati As Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim strMsg As String
    Dim strExtra As String
    Dim strTmp As String
    Dim str部位 As String, str方法 As String
    Dim strTmp方法 As String
    Dim objTmp As New clsApplicationData
    Dim str摘要 As String
    Dim strItems As String
    Dim strRISDel As String
    Dim strRISAdd As String 'RIS参数，格式：医嘱ID:诊疗项目ID,....
    Dim arrSQL() As String
    Dim lng假医嘱ID As Long  '避免医嘱ID序列值的浪费，在最后提交事务时产生真的医嘱ID
    Dim blnDo As Long
    Dim strTabAdvice As String '医嘱信息临时表
    Dim strTxtAdvice As String '医嘱内容
    Dim bln提醒对码 As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim rsPrice As ADODB.Recordset
    Dim lng医嘱序号 As Long

    On Error GoTo errH
    '执行部门(号别科室)即病人科室
    strSQL = "Select A.姓名,A.性别,A.年龄,B.门诊号,B.住院号,B.健康号,a.ID as 挂号ID," & _
        " B.险类,B.就诊诊室,C.名称 as 执行部门,A.登记时间,b.费别" & _
        " From 病人挂号记录 A,病人信息 B,部门表 C" & _
        " Where A.NO(+)=[2] And a.记录性质(+)=1 And a.记录状态(+)=1 And B.病人ID=[1]" & _
        " And A.病人ID(+)=B.病人ID And A.执行部门ID=C.ID(+)"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If
    
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ApplyOutPacs", lng病人ID, str挂号单)

    If rsPati.RecordCount = 0 Then
        MsgBox "未能正确读取病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    '初始化检查申请单对象
    Call objPacspplication.InitComponents(lng挂号科室ID, frmParent)
    Call objTmp.MakePacsData(lng申请序号, objAppPages())
    If objPacspplication.ShowApplicationForm(lng病人ID, 1, Val(rsPati!挂号ID & ""), 0, IIF(lng申请序号 = 0, lng医嘱ID, lng申请序号), objAppPages(), , , lng项目id) Then
        On Error GoTo errH
        If lng申请序号 = 0 Then
            strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ApplyOutPacs")
            lng申请序号 = Val(rsTmp!申请序号)
        End If
        On Error Resume Next
        err.Clear
        If UBound(objAppPages) >= 0 Then
            If err.Number = 0 Then
                On Error GoTo errH
                ReDim Preserve arrSQL(0)
                lng医嘱序号 = GetMaxAdviceNO(lng病人ID, 0, 0)
                For i = 0 To UBound(objAppPages)
                    If lng医嘱ID = 0 Or objAppPages(i).blnIsModify = True Then
                        '调用 zl_AdviceCheck 函数检查
                        str摘要 = ""
                        str摘要 = gclsInsure.GetItemInfo(Val(rsPati!险类 & ""), lng病人ID, 0, "", 0, "", CStr(objAppPages(i).lngProjectId))
                        objAppPages(i).strAbstract = str摘要
                        strTmp = objAppPages(i).strPartMethod
                        If strTmp <> "" Then
                            For k = 0 To UBound(Split(strTmp, "|"))
                                str部位 = Split(Split(strTmp, "|")(k), ";")(0)
                                strTmp方法 = Split(Split(strTmp, "|")(k), ";")(1)
                                For j = 0 To UBound(Split(strTmp方法, ","))
                                    str方法 = Split(strTmp方法, ",")(j)
                                    strExtra = strExtra & "," & str部位 & ":" & str方法
                                Next
                            Next
                            strExtra = "||0||0||" & Mid(strExtra, 2) & "||0"
                        End If
                        strExtra = str摘要 & strExtra
                        strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 1, lng病人ID, Val(rsPati!挂号ID & ""), Val(rsPati!险类 & ""), 1, "D", objAppPages(i).lngProjectId, _
                           objAppPages(i).lngRequestRoomId, UserInfo.姓名, objAppPages(i).lngExeRoomId, IIF(objAppPages(i).lngExeRoomId <= 0, "5", objAppPages(i).lngExeRoomType), 0, 0, strExtra)
                        
                        If Not rsTmp.EOF Then
                            strMsg = NVL(rsTmp!结果)
                            If strMsg <> "" Then
                                Select Case Val(Split(strMsg, "|")(0))
                                Case 1 '提示
                                    If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strMsg = "": Exit Function
                                    End If
                                Case 2 '禁止
                                    MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                                    strMsg = "": Exit Function
                                End Select
                                strMsg = ""
                            End If
                        End If
                        
                        blnDo = GetPacsAdviceSQLData(objAppPages(i), 1, lng申请序号, lng病人ID, 0, 0, str挂号单, lng挂号科室ID, objVBA, objScript, rsDefine, _
                            strRISDel, strRISAdd, arrSQL, lng假医嘱ID, strTabAdvice, strTxtAdvice, lng前提ID, lng医嘱序号)
                        If Not blnDo Then Exit Function
                        
                        strItems = objAppPages(i).lngProjectId & ":" & objAppPages(i).lngExeRoomId
                        '医保对码检查
                        If gint医保对码 = 2 Then bln提醒对码 = True
                        strMsg = CheckAdviceInsure(Val(rsPati!险类 & ""), bln提醒对码, lng病人ID, 1, "", strItems, Left(strTxtAdvice, 50))
                        If strMsg <> "" Then
                            If gint医保对码 = 1 Then
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                If vMsg = vbIgnore Then bln提醒对码 = False
                            ElseIf gint医保对码 = 2 Then
                                MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                                Exit Function
                            End If
                            strMsg = ""
                        End If
                        
                        If Val(rsPati!险类 & "") <> 0 Then
                            If gclsInsure.GetCapability(support实时监控, lng病人ID, Val(rsPati!险类 & "")) Then
                                If MakePriceRecord申请单("21", lng病人ID, Val(rsPati!挂号ID & ""), strTabAdvice, strItems, rsPati!费别 & "", lng挂号科室ID, rsPrice) Then
                                    If Not gclsInsure.CheckItem(Val(rsPati!险类 & ""), 0, 0, rsPrice) Then
                                        MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的PACS申请单不能保存。", vbInformation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        strTabAdvice = ""
                    End If
                Next
                If Not SavePacsData(1, Mid(strRISDel, 2), Mid(strRISAdd, 2), arrSQL, lng假医嘱ID) Then
                     Exit Function
                End If
                ApplyOutPacs = lng申请序号
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePacsData(ByVal intType As Integer, ByVal strRISDel As String, ByVal strRISAdd As String, ByRef arrSQL() As String, ByVal lng假医嘱ID As Long, Optional ByRef lng医嘱ID As Long) As Boolean
'功能：提交数据
'参数：  intType 0－住院医嘱，1－门诊医嘱,lng医嘱ID－返回一条医嘱的ID
    Dim varTmp As Variant
    Dim strTmp As String
    Dim i As Long, j As Long
    Dim varID As Variant
    Dim blnStartTran As Boolean
    Dim lngTmp As Long
    
    On Error GoTo errH
    
    '产生真实的医嘱ID值
    For i = 1 To lng假医嘱ID
        j = zlDatabase.GetNextID("病人医嘱记录")
        If i = 1 Then
            strTmp = j
        Else
            strTmp = strTmp & "," & j
        End If
    Next

    varID = Split(strTmp, ",")

    For i = 1 To UBound(arrSQL)
        strTmp = arrSQL(i)
        If InStr(strTmp, "<FAKEID>") > 0 Then
            j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
            strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
            If InStr(strTmp, "<FAKEID>") > 0 Then '最多替换两次
                j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
            End If
            arrSQL(i) = strTmp
        End If
    Next
    
    varTmp = Split(strRISDel, ",")
    If strRISDel <> "" Then
        On Error Resume Next
        For i = 0 To UBound(varTmp)
            strTmp = varTmp(i)
            If 0 <> gobjRis.HISSchedulingEx(Val(Split(strTmp, ":")(0)), Val(Split(strTmp, ":")(1))) Then
                MsgBox "当前启用了影像信息系统接口，本次操作删除或修改了已经预约医嘱，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系！", vbInformation, gstrSysName
            End If
        Next
        err.Clear: On Error GoTo 0
    End If
    On Error GoTo errH
    Call gcnOracle.BeginTrans: blnStartTran = True
        For i = 1 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "PACS申请单保存医嘱")
        Next
    Call gcnOracle.CommitTrans: blnStartTran = False
    
    SavePacsData = True
    
    lng医嘱ID = Val(varID(0))
    
    If strRISAdd <> "" Then
        varTmp = Split(strRISAdd, ",")
        j = IIF(intType = 1, 1, 2)
        On Error Resume Next
        For i = 0 To UBound(varTmp)
            strTmp = varTmp(i)
            
            lngTmp = Val(Split(strTmp, ":")(0))
            lngTmp = Val(varID(lngTmp - 1)) '换为真实的医嘱
            
            Call gobjRis.HISScheduling(j, lngTmp, Val(Split(strTmp, ":")(1)))
        Next
        err.Clear: On Error GoTo 0
    End If
    Exit Function
errH:
    If blnStartTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPacsAdviceSQLData(ByRef adviceInf As clsApplicationData, ByVal intType As Integer, ByVal lng申请序号 As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng婴儿序号 As Long, ByVal str挂号单 As String, ByVal lng科室id As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset, _
    ByRef strRISDel As String, ByRef strRISAdd As String, ByRef arySql() As String, ByRef lng假医嘱ID As Long, ByRef strTabAdvice As String, ByRef strTxtAdvice As String, ByVal lng前提ID As Long, ByRef lng医嘱序号 As Long) As Boolean
'------------------------------------------------
'功能：获取保存医嘱的SQL，注意真实的医嘱ID是在执行事物时才产生的
'参数： intType 0－住院医嘱，1－门诊医嘱；arySql过程SQL，strTabAdvice 医嘱信息生成的临表SQL查询
'       lng科室ID 住院调用 如果是转出病人，则为原科室ID；门诊调用 挂号执行科ID
'       strRISDel 需要取消RIS预约的信息，strRISAdd 需要重新预约的RIS信息，"21341,12343,..."
'返回：true则继续，false退出
'------------------------------------------------
    Dim i As Long, j As Long
    Dim arrSQL() As String
    Dim int急 As Integer '紧急标志
    
    Dim str医嘱内容 As String
    Dim str项目名称 As String
    
    Dim lng医嘱ID As Long
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim lng计价特性 As Long
    Dim blnRIS预约 As Boolean
    Dim str医嘱IDs As String
    Dim strTmp As String
    Dim varID As Variant
    Dim rsData As ADODB.Recordset
    Dim str部位 As String
    Dim strTmp方法  As String
    Dim str方法 As String
    Dim lngTmpID As Long
    
    Dim arrAppend As Variant
    Dim blnIsDel As Boolean
    Dim lng必填 As Long
    Dim lng排列 As Long
    Dim lng要素ID As Long
    Dim str附项内容 As String
    Dim strDiag As String

    On Error GoTo errH


    '获取医嘱所需的相关数据
    If adviceInf Is Nothing Then Exit Function
    
    strSQL = "Select 名称,计价性质 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", adviceInf.lngProjectId)
    If Not rsTmp.EOF Then
        str项目名称 = rsTmp!名称 & ""
        lng计价特性 = Val(rsTmp!计价性质 & "")
    End If
    
    lng假医嘱ID = lng假医嘱ID + 1
    lng医嘱ID = lng假医嘱ID ' zlDatabase.GetNextID("病人医嘱记录")        '获取医嘱ID
    
    ReDim Preserve arrSQL(1)
    
    If adviceInf.lngUpdateAdviceId <> 0 Then
        '修改医嘱，删除后重新插入
        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & adviceInf.lngUpdateAdviceId & ",1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    End If
    
    If intType = 0 Then
        '组织主医嘱插入语句
        int急 = IIF(adviceInf.blnIsAdditionalRec = True, 2, IIF(adviceInf.blnIsPriority, 1, 0))
    Else
        '组织主医嘱插入语句
        int急 = IIF(adviceInf.blnIsPriority, 1, 0)
    End If
    
    str医嘱内容 = FormatAdviceContext(str项目名称, adviceInf.strPartMethod, adviceInf.lngExeType, objVBA, objScript, rsDefine)
    strTxtAdvice = str医嘱内容
    lng医嘱序号 = lng医嘱序号 + 1
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(<FAKEID>" & lng医嘱ID & "</FAKEID>,NULL," & lng医嘱序号 & "," & IIF(intType = 0, 2, 1) & _
                    "," & lng病人ID & "," & IIF(intType = 0, lng主页ID & "," & lng婴儿序号, "Null,0") & ",1,1,'D'," & adviceInf.lngProjectId & _
                    ",NULL,NULL,NULL,1,'" & str医嘱内容 & "',Null,Null,'一次性',NULL,NULL,NULL,NULL," & lng计价特性 & "," & _
                    IIF(adviceInf.lngExeRoomId <= 0, "Null", adviceInf.lngExeRoomId) & _
                    "," & IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & "," & int急 & _
                    ",to_date('" & Format(adviceInf.strStartExeTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')" & _
                    ",NULL," & lng科室id & "," & adviceInf.lngRequestRoomId & ",'" & UserInfo.姓名 & "'," & _
                    "to_date('" & Format(adviceInf.strRequestTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                    IIF(intType = 0, "null", "'" & str挂号单 & "'") & "," & ZVal(lng前提ID) & ",Null," & adviceInf.lngExeType & _
                    ",NULL," & IIF(adviceInf.strAbstract = "", "Null", "'" & adviceInf.strAbstract & "'") & ",'" & UserInfo.姓名 & "',Null,NULL,NULL,NULL," & lng申请序号 & ")"
    
    strTabAdvice = "Select " & lng医嘱ID & " As ID," & lng医嘱序号 & " As 序号,-null As 相关id, 'D' As 诊疗类别," & adviceInf.lngProjectId & " As 管码项目id," & adviceInf.lngProjectId & " As 诊疗项目id," & vbNewLine & _
        " 1 As 总量, 0 As 单量, null As 标本部位,null As 检查方法," & adviceInf.lngExeType & " As 执行标记," & lng计价特性 & " As 计价特性,null As 附加手术," & _
        IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & " As 执行性质," & adviceInf.lngExeRoomId & " As 执行科室id From Dual"
    
    strDiag = adviceInf.strDiagnoseId
    If intType = 1 Then
        If strDiag = "" Then
            '门诊病人用申请单时取一条诊断来进行默认关联
            strSQL = "Select a.ID From 病人诊断记录 A,病人挂号记录 b Where a.病人id=b.病人id and a.主页id =b.id and b.no=[1] and  a.记录来源 = 3  order by a.诊断类型,a.诊断次序"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str挂号单)
            If Not rsTmp.EOF Then
                strDiag = rsTmp!ID & ""
            End If
        End If
    End If
    
    '诊断关联信息
    If strDiag <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(<FAKEID>" & lng医嘱ID & "</FAKEID>,'" & strDiag & "')"
    End If
    
    '组织部位插入语句
    For i = 0 To UBound(Split(adviceInf.strPartMethod, "|")) '部位1;方法1,方法2,方法3|部位n;方法1,方法2,方法3---
        str部位 = Split(Split(adviceInf.strPartMethod, "|")(i), ";")(0)
        strTmp方法 = Split(Split(adviceInf.strPartMethod, "|")(i), ";")(1)
        For j = 0 To UBound(Split(strTmp方法, ","))
            lng医嘱序号 = lng医嘱序号 + 1     '病人医嘱记录.序号，递增
            str方法 = Split(strTmp方法, ",")(j)
            lng假医嘱ID = lng假医嘱ID + 1
            lngTmpID = lng假医嘱ID
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(<FAKEID>" & lngTmpID & "</FAKEID>,<FAKEID>" & lng医嘱ID & "</FAKEID>," & lng医嘱序号 & _
                 IIF(intType = 0, ",2," & lng病人ID & "," & lng主页ID & "," & lng婴儿序号 & ",", ",1," & lng病人ID & ",NULL,0,") & _
                 "1,1,'D'," & adviceInf.lngProjectId & ",NULL,NULL,NULL,1," & _
                 "'" & str项目名称 & "',NULL," & _
                 "'" & str部位 & "','一次性',NULL,NULL,NULL,NULL," & lng计价特性 & "," & _
                 IIF(adviceInf.lngExeRoomId <= 0, "null", adviceInf.lngExeRoomId) & "," & IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & "," & int急 & ",to_date('" & Format(adviceInf.strStartExeTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),NULL," & _
                 lng科室id & "," & adviceInf.lngRequestRoomId & _
                 ",'" & UserInfo.姓名 & "',to_date('" & Format(adviceInf.strRequestTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')," & _
                 IIF(intType = 0, "NULL", "'" & str挂号单 & "'") & "," & ZVal(lng前提ID) & ",'" & str方法 & "'," & adviceInf.lngExeType & ",NULL,NULL,'" & UserInfo.姓名 & "',NULL,NULL,NULL,NULL," & lng申请序号 & ")"
            
            strTabAdvice = strTabAdvice & " Union All " & _
                "Select " & lngTmpID & " As ID," & lng医嘱序号 & " As 序号," & lng医嘱ID & " As 相关id, 'D' As 诊疗类别," & adviceInf.lngProjectId & " As 管码项目id," & adviceInf.lngProjectId & " As 诊疗项目id," & vbNewLine & _
                " 1 As 总量, 0 As 单量,'" & str部位 & "' As 标本部位,'" & str方法 & "' As 检查方法," & adviceInf.lngExeType & " As 执行标记," & lng计价特性 & " As 计价特性,null As 附加手术," & _
                IIF(adviceInf.lngExeRoomId <= 0, "5", adviceInf.lngExeRoomType) & " As 执行性质," & adviceInf.lngExeRoomId & " As 执行科室id From Dual"
            
        Next
    Next
    lng医嘱序号 = lng医嘱序号 + 1
    
    '组织申请附项插入语句
    blnIsDel = False
    If adviceInf.strRequestAffix <> "" Then
        arrAppend = Split(adviceInf.strRequestAffix, "|")
        For j = 0 To UBound(arrAppend)
            strTmp = "": lng必填 = 0: lng排列 = 0: lng要素ID = 0
            If InStr(adviceInf.strRequestAffixCfg, Split(arrAppend(j), ":")(0) & ":") > 0 Then
                strTmp = Mid(adviceInf.strRequestAffixCfg, InStr(adviceInf.strRequestAffixCfg, Split(arrAppend(j), ":")(0) & ":") + Len(Split(arrAppend(j), ":")(0) & ":"))
                If InStr(strTmp, "|") > 0 Then
                    strTmp = Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                End If
                If strTmp <> "" Then
                    lng必填 = Split(strTmp, ",")(0)
                    lng排列 = Split(strTmp, ",")(1)
                    lng要素ID = Val(Split(strTmp, ",")(2))
                End If
            End If
            strTmp = arrAppend(j)
            str附项内容 = Replace(strTmp, Mid(strTmp, 1, InStr(strTmp, ":")), "")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(<FAKEID>" & lng医嘱ID & "</FAKEID>,'" & Split(arrAppend(j), ":")(0) & "'," & lng必填 & "," & lng排列 & "," & ZVal(lng要素ID) & ",'" & str附项内容 & "'," & IIF(Not blnIsDel, 1, 0) & ")"
            blnIsDel = True
        Next
    End If
    
    If HaveRIS And gbln启用影像信息系统预约 Then
        blnRIS预约 = True
        strSQL = "select a.ID,b.预约id from 病人医嘱记录 a,RIS检查预约 b where a.id=b.医嘱id and a.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", adviceInf.lngUpdateAdviceId)
    End If
    If blnRIS预约 Then
        If Not rsTmp.EOF Then
            strRISDel = strRISDel & "," & Val(rsTmp!ID & "") & ":" & Val(rsTmp!预约id & "")
        End If
    End If
     
    For i = 1 To UBound(arrSQL)
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = arrSQL(i)
    Next
    
    If blnRIS预约 Then
        strRISAdd = strRISAdd & "," & lng医嘱ID & ":" & adviceInf.lngProjectId
    End If
    
    GetPacsAdviceSQLData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatAdviceContext(ByVal strAdvicePro As String, _
    ByVal strAdvicePart As String, ByVal lngExeType As Long, ByRef objVBA As Object, ByRef objScript As Object, ByRef rsDefine As ADODB.Recordset) As String
'根据系统基本参数，格式化医嘱内容

    Dim strReturn As String
    Dim i As Long
    Dim Arr部位 As Variant
    Dim str部位方法 As String
    Dim strTmp As String
    
    If objVBA Is Nothing Then
        On Error Resume Next
        Set objVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not objVBA Is Nothing Then
            objVBA.Language = "VBScript"
            Set objScript = New clsScript
            objVBA.AddObject "clsScript", objScript, True
        End If
    End If
    On Error GoTo errH
    
    rsDefine.Filter = "诊疗类别='D'"
    If rsDefine.RecordCount > 0 Then
        strReturn = rsDefine!医嘱内容 & ""
    End If
    strTmp = ""
    '获取部位方法
    '前:部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中
    '后:部位名1(方法名1,方法名2),部位名2(方法名1,方法名2)-----
    If strAdvicePart <> "" Then
        strTmp = strAdvicePart
        Arr部位 = Split(Split(strTmp, Chr(9))(0), "|")
        strTmp = ""
        For i = 0 To UBound(Arr部位)
            strTmp = strTmp & "," & Split(Arr部位(i), ";")(0) & "(" & Split(Arr部位(i), ";")(1) & ")"
        Next
        strTmp = Mid(strTmp, 2)
    End If
    str部位方法 = strTmp
    
    If strReturn = "" Then
        strReturn = strAdvicePro & "," & _
                            Decode(lngExeType, 1, ",床旁执行", 2, ",术中执行", "") & IIF(strAdvicePart <> "", ":" & str部位方法, "")
    Else
        If InStr(strReturn, "[检查项目]") > 0 Then
            strReturn = Replace(strReturn, "[检查项目]", _
                                            """" & strAdvicePro & Decode(lngExeType, 1, ",床旁执行", 2, ",术中执行", "") & _
                                            """")
        End If

        '替换部位方法
        If InStr(strReturn, "[检查部位]") > 0 Then
            strReturn = Replace(strReturn, "[检查部位]", _
                                            """" & str部位方法 & """")
        End If
        strReturn = objVBA.Eval(strReturn)
    End If
    FormatAdviceContext = strReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBloodState(ByVal int紧急 As Integer, ByVal int备血 As Integer) As String
'功能：获取新下达的输血医嘱的审核状态
'参数：int紧急 0－非紧急，1－紧急；int备血 0－备血 要血库发血，1－用血 起一个通知作用
    Dim strTmp As String
    Dim str审核状态 As String
    
    strTmp = IIF(gbln输血分级管理, "1", "0")
    strTmp = strTmp & IIF(gbln血库系统, "1", "0")
    strTmp = strTmp & int紧急 & int备血
    Select Case strTmp
    Case "1000", "1001", "1100"
        str审核状态 = "1"
    Case "0100", "0110", "1110"
        str审核状态 = "4"
    End Select
    
    GetBloodState = str审核状态
End Function

Public Function CanAutoExeItem(ByVal lng科室id As Long, ByVal str类别 As String, ByVal str操作类型 As String, ByVal int执行分类 As Integer) As Boolean
'功能：判断当前项目是不是可以自动完成
    Dim varArr As Variant
    Dim i As Long
    Dim blnResult As Boolean
    Dim lngResult As Long
    Dim str医嘱类别 As String
    Dim strPar As String
    
    strPar = zlDatabase.GetPara("本科执行自动完成医嘱类别", glngSys, p住院医嘱发送, , , , , lng科室id)
    
    str医嘱类别 = str类别 & IIF("" = str操作类型, 0, str操作类型) & int执行分类
    Select Case str医嘱类别
        Case "E21"
            lngResult = 0 '输液
        Case "E22"
            lngResult = 1 '注射
        Case "E24"
            lngResult = 2 '口服
        Case "E60"
            lngResult = 3 '采集
        Case "E13", "E15"
            lngResult = 4 '过敏试验
        Case "E00"
            lngResult = 5 '普通治疗
        Case "E50"
            lngResult = 6 '特殊治疗
        Case "E20"
            lngResult = 7 '其他给药途径
        Case Else
            lngResult = 8 '其他医嘱
    End Select
    
    If InStr("," & strPar & ",", "," & lngResult & ",") > 0 Or strPar = "*" Then
        blnResult = True
    End If
    
    If blnResult Then
        If (str医嘱类别 = "E13" Or str医嘱类别 = "E15") And Mid(gstr医嘱核对, 2, 1) = "1" Then
            blnResult = False
        ElseIf (str类别 = "K" Or str类别 = "E" And str操作类型 = "8") And Mid(gstr医嘱核对, 1, 1) = "1" Then
            blnResult = False
        End If
    End If
    
    CanAutoExeItem = blnResult
End Function

Public Function RevokeOutAdvice(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal str挂号单 As String, ByVal str姓名 As String, ByVal str门诊号 As String, ByVal lng挂号科室ID As Long, ByVal lng医嘱ID As Long, _
    ByVal lng医嘱状态 As Long, ByVal str诊疗类别 As String, ByVal str操作类型 As String, ByVal lng审核状态 As Long, ByVal str发送时间 As String, _
    ByVal lng签名 As Long, ByVal lngType As Long, ByVal str医嘱内容 As String, ByVal blnMoved As Boolean, ByRef clsMip As Object, ByVal int场合 As Integer) As Boolean
'功能：门诊医嘱作废操作
'参数：lngType 1－表示一并给药，2－检验行；
'      blnMoved 数据是否转出
'      clsMip 消息对象

    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lng证书ID As Long, lng签名id As Long
    Dim strSign As String, intRule As Integer
    Dim strSource As String, strIDs As String, blnDo As Boolean
    Dim strTimeStamp As String, blnTran As Boolean, strErr As String, strTimeStampCode As String
    Dim lngRIS医嘱ID As Long
    Dim strAdvice输血 As String
    Dim arrSQL As Variant
    Dim i As Long
    
    '检查是否可以作废
    If lng医嘱ID = 0 Then
        MsgBox "该病人没有医嘱可以作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lng医嘱状态 <> 8 Then
        MsgBox "当前选择的医嘱尚未发送或已经作废。", vbInformation, gstrSysName
        Exit Function
    End If
    If str诊疗类别 = "K" And gbln血库系统 Then
        If InitObjBlood(True) Then
            strAdvice输血 = lng医嘱ID
        End If
    End If
    
    If str诊疗类别 = "K" And gbln血库系统 And lng审核状态 = 2 Then
        On Error GoTo errH
        strSQL = "Select Nvl(执行分类,0) as 执行分类 from 病人医嘱记录 A, 诊疗项目目录 B  where A.相关ID  = [1] and A.诊疗项目ID = B.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询诊疗项目的执行分类", lng医嘱ID)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp!执行分类) = 0 Then
                MsgBox "本次作废的输血医嘱已经完成配血，不能直接作废医嘱，若要作废请与输血科联系。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        On Error GoTo 0
    End If
    
    '已有费用转出不允许作废
    If zlDatabase.DateMoved(str发送时间) Then
        If MovedBySend(lng医嘱ID, 0, 1) Then
            MsgBox "该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。" & vbCrLf & _
                "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '电子签名检查和提示
    If lng签名 = 1 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
            Else
                MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
    End If
    
    '检查作废医嘱对应的费用结帐情况
    If Not CheckAdviceBalanceRevoke(lng医嘱ID) Then Exit Function
    
    '已审核记帐费用检查
    If InStr(GetInsidePrivs(p门诊医嘱下达), "作废已审核记帐医嘱") = 0 Then
        If Not CheckAdviceBillingRevoke(lng医嘱ID) Then
            MsgBox "要作废医嘱的对应记帐划价费用已经审核，不能作废。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If lngType = 1 Then
        If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("确实要作废医嘱""" & str医嘱内容 & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    arrSQL = Array()
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人诊断医嘱_delete(" & lng医嘱ID & ")"   '门诊医嘱作废时删除病人诊断医嘱 中对应的记录
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人危急值医嘱_Update(3,null," & lng医嘱ID & ")"   '删除危急值对应关系
    
    '作废时进行电子签名
    If strSign <> "" Then
        If gobjESign.CertificateStoped(UserInfo.姓名) = False Then
            '获取签名医嘱源文
            strIDs = lng医嘱ID
            intRule = ReadAdviceSignSource(4, lng病人ID, str挂号单, strIDs, 0, blnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & strIDs & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSign
            Else
                Exit Function
            End If
        End If
    End If
    
    'RIS回退，回退失败则退出
    If InStr(",D,F,", str诊疗类别) > 0 Or str诊疗类别 = "E" And lngType <> 2 Then '检查、手术、治疗
        If HaveRIS(True) Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lng医嘱ID) <> 1 Then
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISRollAdvice)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            lngRIS医嘱ID = lng医嘱ID
            err.Clear: On Error GoTo 0
        End If
    End If
    
    Call CreatePlugInOK(p门诊医嘱下达, int场合)
    
    '调用作废前外挂接口
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        strErr = ""
        blnDo = gobjPlugIn.AdviceRevokedBefore(glngSys, p门诊医嘱下达, lng病人ID, lng挂号ID, lng医嘱ID, int场合, strErr)
        Call zlPlugInErrH(err, "AdviceRevokedBefore")
        If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
            If Not blnDo Then
                MsgBox strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
        blnDo = False
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "RevokeOutAdvice"
    Next
    If strAdvice输血 <> "" Then
        If gobjPublicBlood.AdviceOperation(p门诊医生站, lng医嘱ID, 4, False, strErr) = False Then
            gcnOracle.RollbackTrans: blnTran = False
            MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    
    If Not (clsMip Is Nothing) Then
        If clsMip.IsConnect Then
            Call ZLHIS_CIS_024(clsMip, lng病人ID, str姓名, , str门诊号, 1, lng挂号ID, lng挂号科室ID, "", lng医嘱ID, str诊疗类别, str操作类型)
        End If
    End If
    '调用作废后外挂接口
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.AdviceRevoked(glngSys, p门诊医嘱下达, lng病人ID, lng挂号ID, lng医嘱ID, int场合)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    Call InitObjLis(p门诊医生站)
    '调用LIS作废申请单
    If Not gobjLIS Is Nothing Then
        If gobjLIS.DelLisApplicationForm(CStr(lng医嘱ID), strErr) = False Then
            MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
        End If
    End If
    '调用数据交换平台，向LIS,PACS取消申请单
    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    If Not gobjExchange Is Nothing Then
        If str诊疗类别 = "D" Then
            blnDo = True
        ElseIf str诊疗类别 = "E" Then
            blnDo = lngType = 2
        End If
        If blnDo Then
            Call gobjExchange.SendMsg(IIF(str诊疗类别 = "D", 2, 1), "病人ID::" & lng病人ID & "||主页ID::0||医嘱ID::" & lng医嘱ID & "||操作类型::0")
        End If
    End If
    '调用预约中心服务
    If str诊疗类别 = "Z" And str操作类型 = "2" Then
        Call Svr预约入院取消服务(lng挂号ID)
    End If
    RevokeOutAdvice = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        If blnTran Then
        blnTran = False
        'HIS事务回滚再调用RIS发送 lngRIS医嘱ID
        If lngRIS医嘱ID <> 0 And HaveRIS(True) Then
            strSQL = "Select a.病人id, a.主页id, a.挂号单, a.开嘱科室id, a.执行科室id,a.诊疗项目ID, a.诊疗类别 As 类别, b.发送号, a.Id As 医嘱id, Decode(a.挂号单, Null, 2, 1) As 病人来源" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B Where a.Id = b.医嘱id And a.Id =[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "RevokeOutAdvice", lngRIS医嘱ID)
            If Not rsTmp.EOF Then
                Call gobjRis.HISSendAdvice(rsTmp, 1, Val(rsTmp!病人ID & ""), 0, rsTmp!挂号单 & "", Val(rsTmp!发送号 & ""))
            End If
        End If
    End If
End Function

Public Function CheckLISAppAdvice(ByVal int场合 As Integer, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int险类 As Integer, ByVal str类别 As String, _
    ByVal lng诊疗项目ID As Long, ByVal lng开嘱科室ID As Long, ByVal str开嘱医生 As String, ByVal lng执行科室ID As Long, ByVal lng执行性质 As Long, ByVal str摘要 As String) As Boolean
'功能：医嘱数据库端检查 Zl_Advicecheck
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    On Error GoTo errH
    
    strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", int场合, lng病人ID, lng就诊ID, int险类, 1, str类别, lng诊疗项目ID, _
         lng开嘱科室ID, str开嘱医生, lng执行科室ID, lng执行性质, 0, 0, str摘要)
    
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!结果)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '提示
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '禁止
                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                strMsg = "": Exit Function
            End Select
            strMsg = ""
        End If
    End If
    CheckLISAppAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub DefCommandPlugInPopup(ByRef objBar As Object, ByRef rsBar As ADODB.Recordset)
'功能：在医嘱卡右键弹出菜单
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
            objControl.IconId = rsBar!图标ID
            objControl.Parameter = rsBar!功能名
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
End Sub

Public Sub InitCardRsBlood(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "用血安排", adInteger '0-普通，1－紧急
        .Append "临床诊断IDs", adVarChar, 2000 '诊断ID 串逗号分割   病人诊断记录.ID
        .Append "待诊", adInteger '0/1
        .Append "输血类型", adVarChar, 500
        .Append "输血目的", adVarChar, 500
        .Append "输血性质", adInteger
        .Append "即往输血史", adInteger
        .Append "既往输血反应史", adInteger
        .Append "输血禁忌及过敏史", adInteger
        .Append "孕产情况", adVarChar, 10 '1/1 表示:1孕1产
        .Append "受血者属地", adInteger
        .Append "是否签订同意书", adInteger, 1, adFldIsNullable
        .Append "是否已评估", adInteger, 1, adFldIsNullable
        .Append "预定输血日期", adVarChar, 500
        .Append "血型", adInteger
        .Append "RHD", adInteger
        .Append "输血项目ID", adBigInt
        .Append "输血执行科室ID", adBigInt
        .Append "预定输血量", adDouble
        .Append "输血途径项目ID", adBigInt
        .Append "输血途径执行科室ID", adBigInt
        .Append "滴速", adVarChar, 2000  '用血申请录入滴速存放在E类医嘱嘱托中
        .Append "备注", adVarChar, 2000
        .Append "输血申请日期", adVarChar, 500
        .Append "申请科室ID", adBigInt
        .Append "临床诊断描述", adVarChar, 4000 '界面显示诊断文字
        .Append "检查结果", adVarChar, 4000 '项目检查结果
        .Append "申请项目", adVarChar, 4000 '输血申请项目 格式：项目ID,申请量,血型,RH;项目ID,申请量,血型,RH...... (新申请单使用)
        .Append "申请其他项目SQL", adVarChar, 2000
        .Append "检验项目SQL", adVarChar, 2000
        .Append "诊断关联信息SQL", adVarChar, 2000
        .Append "申请项目SQL", adVarChar, 2000
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub InitCardRsOperate(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "临床诊断IDs", adVarChar, 2000
        .Append "临床诊断描述", adVarChar, 4000 '界面显示诊断文字
        .Append "手术情况", adInteger '住院用到
        .Append "主手术项目ID", adBigInt
        .Append "附手术项目IDs", adVarChar, 2000
        .Append "麻醉项目ID", adBigInt
        .Append "手术执行科室ID", adBigInt
        .Append "麻醉执行科室ID", adBigInt
        .Append "生效时间", adVarChar, 100 '开始时间
        .Append "手术时间", adVarChar, 100 '安排时间
        .Append "申请附项", adVarChar, 4000
        .Append "申请科室ID", adBigInt
        .Append "附手术必要时", adVarChar, 2000 '附加手术必要时
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Sub InitCardRsLIS(ByRef rsCard As ADODB.Recordset)
    Set rsCard = New ADODB.Recordset
    With rsCard.Fields
        .Append "临床诊断IDs", adVarChar, 2000
        .Append "申请信息", adVarChar, 4000
    End With
    rsCard.CursorLocation = adUseClient
    rsCard.LockType = adLockOptimistic
    rsCard.CursorType = adOpenStatic
    rsCard.Open
End Sub

Public Function ShowApply检查(frmParent As Object, ByVal lngNo As Long, Optional ByVal lng医嘱ID As Long) As Boolean
'功能：查看检查申请单
    Dim rsTmp As ADODB.Recordset
    Dim objAppPages()  As clsApplicationData
    Dim objTmp As New clsApplicationData
    Dim strSQL As String
    Dim int婴儿 As Integer
    
    On Error GoTo errH
    If lngNo = 0 Then
        strSQL = "Select a.申请序号,a.婴儿 From 病人医嘱记录 A where a.Id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowPacsApplication", lng医嘱ID)
        If rsTmp.RecordCount = 0 Then
            MsgBox "没有找到您指定的检查医嘱。", vbInformation, gstrSysName
            Exit Function
        Else
            lngNo = Val(rsTmp!申请序号 & "")
            int婴儿 = Val(rsTmp!婴儿 & "")
        End If
    Else
        strSQL = "Select a.婴儿 From 病人医嘱记录 A where a.申请序号 =[1] and rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowPacsApplication", lngNo)
        int婴儿 = Val(rsTmp!婴儿 & "")
    End If
    If lngNo = 0 Then
        MsgBox "该医嘱没有对应申请单。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set rsTmp = objTmp.MakePacsData(lngNo, objAppPages(), True)
    Call frmPacsApplication.InitComponents(Val(rsTmp!开嘱科室id & ""), frmParent)
    ShowApply检查 = frmPacsApplication.ShowApplicationForm(Val(rsTmp!病人ID & ""), Val(rsTmp!病人性质 & ""), Val(rsTmp!就诊ID & ""), Val(rsTmp!主页ID & ""), lngNo, objAppPages(), int婴儿, False)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get申请序号() As Long
'功能：获取申请序号
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel")
    Get申请序号 = Val(rsTmp!申请序号 & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub InitPlugInRs(ByRef rsDataPlugIn As ADODB.Recordset)
'功能：初化记录集，用于 clsPlugIn.AdviceBeforeSend方法的入参
    Set rsDataPlugIn = New ADODB.Recordset
    rsDataPlugIn.Fields.Append "病人ID", adBigInt
    rsDataPlugIn.Fields.Append "就诊ID", adBigInt
    rsDataPlugIn.Fields.Append "挂号单", adVarChar, 30
    rsDataPlugIn.Fields.Append "医嘱ID", adBigInt
    rsDataPlugIn.Fields.Append "相关ID", adBigInt
    rsDataPlugIn.Fields.Append "收费细目ID", adBigInt
    rsDataPlugIn.Fields.Append "分解时间", adVarChar, 40000
    rsDataPlugIn.Fields.Append "次数", adInteger
    rsDataPlugIn.Fields.Append "单量", adDouble
    rsDataPlugIn.Fields.Append "单量单位", adVarChar, 100
    rsDataPlugIn.Fields.Append "总量", adDouble
    rsDataPlugIn.Fields.Append "总量单位", adVarChar, 100
    rsDataPlugIn.Fields.Append "场合", adInteger
    rsDataPlugIn.CursorLocation = adUseClient
    rsDataPlugIn.LockType = adLockOptimistic
    rsDataPlugIn.CursorType = adOpenStatic
    rsDataPlugIn.Open
End Sub

Public Function CheckDocEmpowerEx(ByVal lng诊疗项目ID As Long, ByVal strAppend As String) As Boolean
'功能：检查操作员是否具有手术项目的执行权
'参数：strAppend=当前申请附项的填写情况串,格式为"项目名1<Split2>0/1(必填否)<Split2>要素ID<Split2>内容<Split1>..."
    Dim strSQL As String, rsTmp As Recordset
    Dim arrItem As Variant, arrSub As Variant
    Dim strItem As String, i As Integer
    Dim lngID As Long
    Dim strDoc As String
    
    On Error GoTo errH
    If strAppend <> "" Then
        strSQL = "select A.ID from 诊治所见项目 A,诊治所见分类 B where a.分类id=b.id and b.编码='06' and A.中文名='主刀医生'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckDocEmpowerEx")
        If rsTmp.RecordCount > 0 Then
            lngID = rsTmp!ID
            arrItem = Split(strAppend, "<Split1>")
            For i = 0 To UBound(arrItem)
                arrSub = Split(arrItem(i), "<Split2>")
                If Val(arrSub(2)) = lngID Then
                    If Trim(arrSub(3)) <> "" Then
                        strDoc = Trim(arrSub(3))
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    If strDoc = "" Then strDoc = UserInfo.姓名
    strSQL = "Select Count(*) as 权限 From 人员手术权限 A,人员表 B Where A.人员id = B.ID And B.姓名=[1] And A.诊疗项目id = [2] And A.记录性质 = 2"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckUserEmpower", strDoc, lng诊疗项目ID)
    CheckDocEmpowerEx = Val(rsTmp!权限 & "") > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPlugInBar(ByVal lng模块 As Long, ByVal int场合 As Integer, rsBar As ADODB.Recordset) As String
'功能：组织外挂部件的菜单样按钮
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugInOK(lng模块, int场合)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lng模块, int场合, strXML)
    Call zlPlugInErrH(err, "GetFuncNames")
    err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'功能：组织菜单到本地记录集中，注意对老版本的兼容处理
'参数：strFunc 老版本功能列串，strXML含配置信息的功能串
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    If strXML = "" And strFunc = "" Then Exit Sub
    If strXML = "" And strFunc <> "" Then
        '兼容以前老版本的方式
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '暂定为200个扩展功能插件，防止死循环
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'功能：将功能串转换为记录集方式
'参数：strFunc 功能串，intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '第一个独立按钮显示分割线
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !功能名 = strFuncName
            !菜单名 = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'功能：分配功能ID，加菜单快键
'参数：lngV 版本，1-老版，2-新版
'返回：字符串，以前低版本方式的功能串
    Dim i As Long
    '分配功能ID，图标ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !序号 = i
            !功能ID = conMenu_Tool_PlugIn_Item + i
            !图标ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'功能：设定快键
'参数：lngV 版本，1-老版，2-新版 intType 功能按钮属于那一栏 1-菜单栏，2-工具栏，3-左键栏
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '如果只有一个，也归为独立按钮
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !菜单名 = !菜单名 & "(&" & i & ")"
                    Else
                        !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !菜单名 = !菜单名 & "(&" & i & ")"
                Else
                    !菜单名 = !菜单名 & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "序号", adBigInt '用于排序
    rsBar.Fields.Append "功能ID", adBigInt '菜单按钮 Control.ID
    rsBar.Fields.Append "图标ID", adBigInt
    rsBar.Fields.Append "功能名", adVarChar, 1000 '去掉关键字之后的 名称 即工具栏上的按钮名称
    rsBar.Fields.Append "菜单名", adVarChar, 1000 '菜单栏/右键菜单 名称
    rsBar.Fields.Append "IsAuto", adInteger '是否自动执行功能
    rsBar.Fields.Append "IsGroup", adInteger '是否分割线
    rsBar.Fields.Append "IsInTool", adInteger '是否独立显示
    rsBar.Fields.Append "BarType", adInteger '1-菜单栏，2－工具栏，3－弹出栏
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub

Public Sub Make待执行消息(ByVal strSendDate As String)
'功能：医嘱发送后产生待执行消息提醒
'参数：strSendDate 发送时间
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL As Variant
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    strSQL = "select b.病人id,b.主页id,b.病人来源,c.当前病区id as 病区ID,c.出院科室id as 科室ID,a.执行部门id" & vbNewLine & _
        "from (select max(a.医嘱id) as 医嘱ID,a.执行部门id from  病人医嘱发送 A where a.执行状态 = 0 And Exists" & vbNewLine & _
        "(Select 1 From 部门性质说明 X Where x.部门id = a.执行部门id And x.工作性质='护理') And a.发送时间 =[1] group by a.执行部门id) a," & vbNewLine & _
        " 病人医嘱记录 B,病案主页 c Where a.医嘱id = b.Id and b.病人id=c.病人id and b.主页id=c.主页id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", CDate(strSendDate))
    
    If Not rsTmp.EOF Then
        arrSQL = Array()
        For i = 1 To rsTmp.RecordCount
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_业务消息清单_insert(" & rsTmp!病人ID & "," & rsTmp!主页ID & "," & rsTmp!科室ID & "," & rsTmp!病区ID & "," & rsTmp!病人来源 & _
                ",'有待执行的医嘱。','0010','ZLHIS_CIS_034','" & strSendDate & "',1,0,null,'" & rsTmp!执行部门ID & "')"
            rsTmp.MoveNext
        Next
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "mdlCISKernel"
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Get申请单相关参数()
'功能：获取申请单相关参数值
    Dim str启用 As String
    Dim str必用 As String
    Dim varTmp As Variant
    str启用 = zlDatabase.GetPara(238, glngSys, , "11|11|11|11|1")
    str必用 = zlDatabase.GetPara(260, glngSys, , "11")
    varTmp = Split(str启用, "|")
    gstrOutUseApp = Mid(varTmp(0), 1, 1) & Mid(varTmp(1), 1, 1) & Mid(varTmp(2), 1, 1) & Mid(varTmp(3), 1, 1)
    gstrInUseApp = Mid(varTmp(0), 2, 1) & Mid(varTmp(1), 2, 1) & Mid(varTmp(2), 2, 1) & Mid(varTmp(3), 2, 1) & Mid(varTmp(4), 1, 1)
    gblnOut必用 = Mid(str必用, 1, 1) = "1"
    gblnIn必用 = Mid(str必用, 2, 1) = "1"
End Sub

Public Function GetDiag诊断描述(ByVal str诊断IDs As String) As String
'功能：获得指定诊断的诊断描述
'返回：str诊断=关联诊断的诊断名称字符串
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str诊断 As String
    
    strSQL = "Select  A.ID,a.诊断描述 From 病人诊断记录 A " & _
        " Where NVL(A.编码序号,1) = 1 And a.id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetDiag诊断描术", str诊断IDs)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            str诊断 = str诊断 & "," & rsTmp!诊断描述
            rsTmp.MoveNext
        Loop
        str诊断 = Mid(str诊断, 2)
    End If
    GetDiag诊断描述 = str诊断
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDataRIS预约(ByVal strIDs As String) As ADODB.Recordset
'功能：获取指定范围医嘱在RIS中产生预约信息的数据
'参数：strIDs 主医嘱ID串
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "select b.医嘱id as ID,b.预约id from RIS检查预约 b where b.医嘱id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)"
    Set GetDataRIS预约 = zlDatabase.OpenSQLRecord(strSQL, "GetDataRIS预约", strIDs)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckPathInItem(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str诊疗项目IDs As String, ByRef str分类 As String, ByVal lng阶段Id As Long, ByVal bln中药配方 As Boolean, ByVal byt期效 As Byte, Optional ByVal bln西药 As Boolean) As Long
'功能：检查临床路径病人，当前输入的医嘱（一组诊疗项目）是否是当前阶段的路径内项目，如果是，则返回项目ID
'      必须且仅执行一次的项目，生成时必定已产生，再添加就当成路径外项目。
'参数：str诊疗项目IDs= '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
'      lng阶段ID=用于检查合并路径的项目是否匹配
'      bln中药配方=中药配方单独处理，根据参数设置的允许修改的中药比例来算
'      byt期效 =医嘱期效
'返回：路径项目ID和分类名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str证候IDs As String
    Dim lng改动中药味数 As Long, dbl中药味数 As Double, lng中药味数 As Long
    Dim i As Long
    Dim arrTmp As Variant, blnTmp As Boolean
    Dim bln匹配期效 As Boolean
    Dim bln不匹配 As Boolean
        Dim bln药品分类相同不算路径外 As Boolean
    Dim str药品分类ids As String
    
    str分类 = ""
    If str诊疗项目IDs = "0" Then
    '自由录入的医嘱固定当成路径外项目
        CheckPathInItem = 0
    Else
        'Wm_Concat函数在分组中的排序有问题,且在10.2.0.5中返回值类型变化了
'        strSQL = "Select 分类,路径项目id" & vbNewLine & _
'                "From (Select 分类,路径项目id, 组id, Wmsys.Wm_Concat(诊疗项目id) 诊疗项目ids" & vbNewLine & _
'                "       From (Select Rownum, c.路径项目id, d.诊疗项目id, b.分类, Decode(内容要求,1,Nvl(d.相关id,d.id),0) 组id" & vbNewLine & _
'                "              From 病人临床路径 A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
'                "              Where a.病人id = [1] And a.主页id = [2] And b.阶段id = a.当前阶段id And b.Id = c.路径项目id And c.医嘱内容id = d.Id And b.执行方式<>4" & vbNewLine & _
'                "              Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
'                "       Group By 分类,路径项目id,组id)" & vbNewLine & _
'                "Where 诊疗项目ids = [3]"
        bln匹配期效 = CBool(zlDatabase.GetPara("匹配时期效不同算路径外项目", glngSys, p临床路径应用, "0"))
        lng中药味数 = Val(zlDatabase.GetPara("中药配方允许修改的中药味数上限", glngSys, p临床路径应用, "30"))
        bln不匹配 = Val(zlDatabase.GetPara("药品医嘱不匹配为路径外项目", glngSys, p临床路径应用, "0")) = 1
                bln药品分类相同不算路径外 = Val(zlDatabase.GetPara("药品医嘱相同分类不算路径外医嘱", glngSys, p临床路径应用, "0")) = 1
        If bln药品分类相同不算路径外 And bln西药 Then
            strSQL = "Select f_List2str(Cast(Collect(To_Char(分类id)) As t_Strlist)) As 药品分类ids" & vbNewLine & _
                    "From 诊疗项目目录" & vbNewLine & _
                    "Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", str诊疗项目IDs)
            If rsTmp.RecordCount > 0 Then
                str药品分类ids = rsTmp!药品分类IDs & ""
            End If
        End If
        '给药途径可能因为收取不同费用的原因，实际使用时不是定义的给药途径，所以只判断药品相同即可
        If Not bln中药配方 Then
            strSQL = "Select 分类, 路径项目id,期效,诊疗项目ids,药品分类IDs" & vbNewLine & _
                    "From (Select 分类, 路径项目id, 组id,期效,f_List2str(Cast(Collect(To_Char(诊疗项目id)) As t_Strlist)) As 诊疗项目ids,f_List2str(Cast(Collect(To_Char(药品分类ID)) As t_Strlist)) As 药品分类IDs" & vbNewLine & _
                    "       From (Select c.路径项目id, b.分类, d.诊疗项目id, Nvl(d.相关id, d.Id) 组id,d.期效,d.序号,e.分类ID as 药品分类ID" & vbNewLine & _
                    "              From 病人临床路径 A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D,诊疗项目目录 E" & vbNewLine & _
                    "              Where a.病人id = [1] And a.主页id = [2] And b.阶段id = [4] And b.Id = c.路径项目id And c.医嘱内容id = d.Id And D.诊疗项目ID = E.ID And " & vbNewLine & _
                    "                    (b.执行方式 <> 4 or b.执行方式 = 4 And Not Exists(Select 1 From 病人路径执行 E Where e.路径记录ID=a.id And a.当前阶段id=e.阶段id and b.id=e.项目id))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From 诊疗项目目录 E Where D.诊疗项目ID = E.ID And E.类别 = 'E' And  E.操作类型 In('2','3','4','6'))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From 诊疗项目目录 E Where D.诊疗项目ID = E.ID And E.类别 In('G','F','D') And D.相关ID<>0 )" & vbNewLine & _
                    "              Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
                    "       Group By 分类, 路径项目id, 组id,期效)" & vbNewLine & _
                    IIF(bln药品分类相同不算路径外 And bln西药, IIF(InStr(str药品分类ids, ",") > 0, " Where instr(药品分类IDs,',')>0 ", "Where (药品分类IDs = [6] or instr(','||药品分类IDs||',',','||[6]||',')>0)"), _
                        IIF(InStr(str诊疗项目IDs, ",") > 0, " Where instr(诊疗项目ids,',')>0 ", "Where (诊疗项目ids = [3] or instr(','||诊疗项目ids||',',','||[3]||',')>0)")) & IIF(bln匹配期效, " And 期效 =[5]", "")
        
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng病人ID, lng主页ID, str诊疗项目IDs, lng阶段Id, byt期效, str药品分类ids)
            
            If bln药品分类相同不算路径外 And bln西药 Then
                If InStr(str药品分类ids, ",") > 0 Then
                    '多个项目判断时，如检验项目，忽略顺序，如果其中有一个是路径外的那么一组就是路径外的
                    arrTmp = Split(str药品分类ids, ",")
                    Do While Not rsTmp.EOF
                        blnTmp = True
                        For i = 0 To UBound(arrTmp)
                            If InStr("," & rsTmp!药品分类IDs & ",", "," & arrTmp(i) & ",") = 0 Then
                                blnTmp = False
                                Exit For
                            End If
                        Next
                        If blnTmp Then
                            CheckPathInItem = rsTmp!路径项目ID
                            str分类 = rsTmp!分类
                            GoTo FuncEnd
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                Else
                    '如果有多个路径项目，则只取第一个
                    If rsTmp.RecordCount > 0 Then
                        CheckPathInItem = rsTmp!路径项目ID
                        str分类 = rsTmp!分类
                        If Not bln匹配期效 Then
                            '存在多个情况下根据期效进行匹配,优先匹配诊疗项目ID和期效都相同的（为了保证在同一阶段,同一药品,不同项目,期效不一致的情况下,能正常匹配）
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!期效 & "" = byt期效 & "" Then
                                    CheckPathInItem = rsTmp!路径项目ID
                                    str分类 = rsTmp!分类
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            Else
                If InStr(str诊疗项目IDs, ",") > 0 Then
                    '多个项目判断时，如检验项目，忽略顺序，如果其中有一个是路径外的那么一组就是路径外的
                    arrTmp = Split(str诊疗项目IDs, ",")
                    Do While Not rsTmp.EOF
                        blnTmp = True
                        For i = 0 To UBound(arrTmp)
                            If InStr("," & rsTmp!诊疗项目ids & ",", "," & arrTmp(i) & ",") = 0 Then
                                blnTmp = False
                                Exit For
                            End If
                        Next
                        If blnTmp Then
                            CheckPathInItem = rsTmp!路径项目ID
                            str分类 = rsTmp!分类
                            GoTo FuncEnd
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                Else
                    '如果有多个路径项目，则只取第一个
                    If rsTmp.RecordCount > 0 Then
                        CheckPathInItem = rsTmp!路径项目ID
                        str分类 = rsTmp!分类
                        If Not bln匹配期效 Then
                            '存在多个情况下根据期效进行匹配,优先匹配诊疗项目ID和期效都相同的（为了保证在同一阶段,同一药品,不同项目,期效不一致的情况下,能正常匹配）
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!期效 & "" = byt期效 & "" Then
                                    CheckPathInItem = rsTmp!路径项目ID
                                    str分类 = rsTmp!分类
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            End If
        Else
            '匹配时排开不符合的证候
            str证候IDs = Get证候IDs(lng病人ID, lng主页ID)
            strSQL = "Select  分类, 路径项目id, 组id, f_List2str(Cast(Collect(To_Char(诊疗项目id)) As t_Strlist)) As 诊疗项目ids" & vbNewLine & _
                    "From (Select c.路径项目id, b.分类, d.诊疗项目id, Nvl(d.相关id, d.Id) 组id, d.序号" & vbNewLine & _
                    "       From 病人临床路径 A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
                    "       Where a.病人id = [1] And a.主页id = [2] And b.阶段id = [3] And b.Id = c.路径项目id And c.医嘱内容id = d.Id And" & vbNewLine & _
                    "             (b.执行方式 <> 4 Or b.执行方式 = 4 And Not Exists" & vbNewLine & _
                    "              (Select 1 From 病人路径执行 E Where e.路径记录id = a.Id And a.当前阶段id = e.阶段id And b.Id = e.项目id)) And Exists" & vbNewLine & _
                    "        (Select 1 From 诊疗项目目录 E Where d.诊疗项目id = e.Id And e.类别 = '7') " & vbNewLine & _
                    IIF(str证候IDs <> "", " And (Instr(',' || [4] || ',', ',' || d.组合项目id || ',') > 0 Or d.组合项目id Is Null)", "") & vbNewLine & _
                    "       Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
                    "Group By 分类, 路径项目id, 组id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng病人ID, lng主页ID, lng阶段Id, str证候IDs)
            Do While Not rsTmp.EOF
                If rsTmp!诊疗项目ids & "" <> "" Then
                    '允许改动的中药
                    dbl中药味数 = (UBound(Split(rsTmp!诊疗项目ids & "", ",")) + 1) * lng中药味数 / 100
                    lng改动中药味数 = 0
                    '先找，配方外的中药
                    For i = 0 To UBound(Split(str诊疗项目IDs, ","))
                        If InStr("," & rsTmp!诊疗项目ids & ",", "," & Split(str诊疗项目IDs, ",")(i) & ",") = 0 Then
                            lng改动中药味数 = lng改动中药味数 + 1
                        End If
                    Next
                    '再找配方中的中药，当且缺少的
                    If rsTmp!诊疗项目ids & "" <> "" Then
                        For i = 0 To UBound(Split(rsTmp!诊疗项目ids & "", ","))
                            If InStr("," & str诊疗项目IDs & ",", "," & Split(rsTmp!诊疗项目ids & "", ",")(i) & ",") = 0 Then
                                lng改动中药味数 = lng改动中药味数 + 1
                            End If
                        Next
                    End If
                    '如果在允许的范围之内，则匹配成功，否则继续匹配
                    If lng改动中药味数 <= dbl中药味数 Then
                        CheckPathInItem = rsTmp!路径项目ID
                        str分类 = rsTmp!分类
                        Exit Do
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
FuncEnd:
    If CheckPathInItem = 0 And bln不匹配 = True Then
        
        strSQL = "Select 分类, 路径项目id,期效,诊疗项目ids,ID" & vbNewLine & _
                    "From (Select 分类, 路径项目id, 组id,期效,f_List2str(Cast(Collect(To_Char(诊疗项目id)) As t_Strlist)) As 诊疗项目ids,ID" & vbNewLine & _
                    "       From (Select c.路径项目id, b.分类, d.诊疗项目id, Nvl(d.相关id, d.Id) 组id,d.期效,d.序号,A.ID" & vbNewLine & _
                    "              From 病人临床路径 A, 临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
                    "              Where a.病人id = [1] And a.主页id = [2] And b.阶段id = [4] And b.Id = c.路径项目id And c.医嘱内容id = d.Id And" & vbNewLine & _
                    "                    (b.执行方式 <> 4 or b.执行方式 = 4 And Not Exists(Select 1 From 病人路径执行 E Where e.路径记录ID=a.id And a.当前阶段id=e.阶段id and b.id=e.项目id))" & vbNewLine & _
                    "                    And Exists(Select 1 From 诊疗项目目录 E Where D.诊疗项目ID = E.ID And E.类别 In('5','6') And D.相关ID<>0 )" & vbNewLine & _
                    "              Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
                    "       Group By 分类, 路径项目id, 组id,期效,ID) A Where exists (select 1 from 病人路径执行 E where E.路径记录ID=A.ID And E.阶段ID=[4] and E.项目ID=A.路径项目ID)" & _
                    " and exists(select 1 from 诊疗项目目录 where 类别 in('5','6') and ID In (select /*+cardinality(A,10)*/ column_value from Table(f_Str2list([3]) ) A))"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckPathInItem", lng病人ID, lng主页ID, str诊疗项目IDs, lng阶段Id, byt期效)
        Do While Not rsTmp.EOF
            CheckPathInItem = rsTmp!路径项目ID
            str分类 = rsTmp!分类
            
            Exit Function
        Loop
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get证候IDs(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取该病人证候IDs，逗号分割
    Dim strSQL As String, rsTmp As Recordset
    Dim str证候IDs As String
    
    strSQL = "Select 证候ID From 病人诊断记录 Where 病人id = [1] And 主页id = [2] And NVL(编码序号,1) = 1 And 证候id Is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get证候IDs", lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        str证候IDs = str证候IDs & "," & rsTmp!证候id
        rsTmp.MoveNext
    Loop
    Get证候IDs = Mid(str证候IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByRef str路径项目分类 As String = "-1") As ADODB.Recordset
'功能：获取路径病人当前路径信息
'返回：str分类=当前天数最后一个路径项目所属的分类
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str路径项目分类 <> "-1"
    str路径项目分类 = ""
    strSQL = "Select a.路径记录id, a.当前阶段id, a.当前天数, a.开始日期, b.日期, b.分类" & vbNewLine & _
            "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, Max(b.Id) 执行id, Min(c.日期) As 开始日期" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B, 病人路径执行 C" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.Id = c.路径记录id And b.阶段id + 0 = a.当前阶段id And b.天数 = a.当前天数 And a.状态 = 1 And" & vbNewLine & _
            "             a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
            "       Group By a.Id, a.当前阶段id, a.当前天数) A, 病人路径执行 B" & vbNewLine & _
            "Where a.执行id = b.Id"

    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSQL, "病人当前路径信息", lng病人ID, lng主页ID)
    If rsRet.RecordCount > 0 And blnDo Then
        str路径项目分类 = "" & rsRet!分类
        
        '如果当天生成了医嘱类项目，则取医嘱类项目的分类
        strSQL = "Select 分类" & vbNewLine & _
                "From 病人路径执行" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From 病人路径执行 A" & vbNewLine & _
                "            Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3] And Exists (Select 1 From 病人路径医嘱 B Where a.Id = b.路径执行id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人当前路径信息", Val(rsRet!路径记录id), Val(rsRet!当前阶段ID), CDate(rsRet!日期))
        If rsTmp.RecordCount > 0 Then
            str路径项目分类 = "" & rsTmp!分类
        End If
    End If
    Set GetPatiPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiPathAppend(ByVal lng记录ID As Long, ByVal dat日期 As Date) As ADODB.Recordset
'功能:根据日期获取有效阶段和天数
    Dim strSQL      As String
    
    strSQL = "Select 阶段id,天数" & vbNewLine & _
            "From (Select a.阶段id, a.天数, a.登记时间" & vbNewLine & _
            "       From 病人路径执行 A" & vbNewLine & _
            "       Where a.路径记录id = [1] And a.日期 = [2]" & vbNewLine & _
            "       Order By a.登记时间 Desc)" & vbNewLine & _
            "Where Rownum < 2"
        
    On Error GoTo errH
    Set GetPatiPathAppend = zlDatabase.OpenSQLRecord(strSQL, "GetPatiPathAppend", lng记录ID, dat日期)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPathOutItemID(ByVal lng路径记录Id As Long, ByVal DatAddDate As Date, Optional ByVal bytFunc As Byte) As Long
'功能：获取刚才添加的路径外项目的执行ID
'参数:bytFunc=0 住院临床路径 1-门诊临床路径
'返回：
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If bytFunc = 1 Then
        strSQL = "Select Max(b.Id) 执行id" & vbNewLine & _
                    "       From 病人门诊路径执行 B" & vbNewLine & _
                    "       Where b.路径记录id = [1] And b.登记时间 = [2] And Nvl(b.项目ID,0) = 0"
    Else
        strSQL = "Select Max(b.Id) 执行id" & vbNewLine & _
                "       From 病人路径执行 B" & vbNewLine & _
                "       Where b.路径记录id = [1] And b.登记时间 = [2] And Nvl(b.项目ID,0) = 0"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPathOutItemID", lng路径记录Id, DatAddDate)
    If rsTmp.RecordCount > 0 Then
        GetPathOutItemID = Val("" & rsTmp!执行Id)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MakePriceRecord申请单(ByVal strType As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strAdvice As String, ByVal str项目科室 As String, ByVal str费别 As String, ByVal lng开单科室ID As Long, ByRef rsPrice As ADODB.Recordset) As Boolean
'功能：工作站主界面单独使用申请单保存申请单时，生成对应的用于医保的费用明细记录集，判断逻辑同门诊/住院医嘱编辑界面MakePriceRecord方法
'参数：strAdvice 医嘱临时表SQL，str项目科室：诊疗项目ID:执行科室ID,...，lng主页ID 如果是门诊病人则传入的 挂号ID
'      strType 调用方 两位数字表示第一位表示申请单 1－检验，2－检查，3－输血，4－手术，5－会诊；第二位表示门诊或者住院 1－门诊，2－住院
'返回：有计价数据记录集内容才返回True
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str诊疗收费 As String
    Dim lng执行科室ID As Long, blnLoad As Boolean
    Dim dbl单价 As Double, dbl金额 As Double, dbl实收 As Double
    Dim str项目 As String, blnDo As Boolean
    Dim lng诊断ID As Long, lng疾病ID As Long
    Dim int场合 As String
    
    int场合 = Mid(strType, 2, 1)
    
    On Error GoTo errH
    
    str诊疗收费 = "Select c.诊疗项目id, c.收费项目id, c.检查部位, c.检查方法, c.费用性质, c.收费数量, c.固有对照, c.从属项目, c.收费方式, c.适用科室id,c.top From (" & _
            "Select /*+cardinality(D,10)*/ Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,C.适用科室id" & _
            " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
            " From 诊疗收费关系 C,Table(f_Num2list2([1])) D Where C.诊疗项目ID=D.c1" & _
            "      And (C.适用科室ID is Null or C.适用科室ID = D.c2 And C.病人来源 = 1)" & _
            " ) c Where Nvl(c.适用科室id, 0) = c.Top"
                
    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
        " Select A.序号,A.诊疗类别,C.类别 as 收费类别,B.收费项目ID as 收费细目ID,D.收入项目ID," & _
        " Decode(A.总量,0,1,A.总量)*B.收费数量 as 数量,Decode(C.是否变价,1,D.缺省价格,D.现价) as 单价," & _
        " C.是否变价,C.屏蔽费别,A.执行科室ID, a.附加手术,d.附术收费率" & _
        " From (" & strAdvice & ") A,(" & str诊疗收费 & ") B,收费项目目录 C,收费价目 D,诊疗项目目录 E,采血管类型 F" & _
        " Where a.诊疗项目id = b.诊疗项目id And (a.相关id Is Null And a.执行标记 In (1, 2) And b.费用性质 = 1 Or" & _
        "            a.标本部位 = b.检查部位 And a.检查方法 = b.检查方法 And Nvl(b.费用性质, 0) = 0 Or" & _
        "            a.检查方法 Is Null And Nvl(b.费用性质, 0) = 0 And b.检查部位 Is Null And b.检查方法 Is Null) And a.管码项目id = e.Id And" & vbNewLine & _
        "            e.试管编码 = f.编码(+) And (Nvl(b.收费方式, 0) = 1 And c.类别 = '4' And b.收费项目id = f.材料id Or" & vbNewLine & _
        "            Not (Nvl(b.收费方式, 0) = 1 And c.类别 = '4' And f.材料id Is Not Null)) And Nvl(a.计价特性, 0) = 0 And" & vbNewLine & _
        "            Nvl(a.执行性质, 0) Not In (0, 5) And b.收费项目id = c.Id And b.收费项目id = d.收费细目id And" & vbNewLine & _
        "            ((Sysdate Between d.执行日期 And d.终止日期) Or (Sysdate >= d.执行日期 And d.终止日期 Is Null)) And" & vbNewLine & _
        "            (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.服务对象 In (1, 3) And" & vbNewLine & _
        "            (c.站点 = '-' Or c.站点 Is Null)"

    strSQL = "Select a.序号,a.诊疗类别,a.收费类别,a.收费细目id,a.收入项目id,a.数量,a.单价,a.是否变价,a.屏蔽费别,a.执行科室id,a.附加手术,a.附术收费率 From (" & strSQL & ") A Order by a.序号,a.收入项目ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MakePriceRecord申请单", str项目科室)
    If Not rsTmp.EOF Then
        '初始化记录集
        Set rsPrice = New ADODB.Recordset
        With rsPrice
            .Fields.Append "病人ID", adBigInt
            .Fields.Append "主页ID", adBigInt, , adFldIsNullable
            .Fields.Append "收费类别", adVarChar, 1
            .Fields.Append "收费细目ID", adBigInt
            .Fields.Append "数量", adDouble
            .Fields.Append "单价", adDouble
            .Fields.Append "实收金额", adDouble
            .Fields.Append "开单人", adVarChar, 100, adFldIsNullable
            .Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
            
            If int场合 = 1 Then
                .Fields.Append "疾病ID", adBigInt, , adFldIsNullable
                .Fields.Append "诊断ID", adBigInt, , adFldIsNullable
            End If
            
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenStatic
            .Open
        End With
        '加入费用明细
        dbl实收 = 0: blnDo = True
        Do While Not rsTmp.EOF
            '执行科室
            If blnDo Then
                lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
            End If
            
            '单价
            dbl单价 = Format(NVL(rsTmp!单价, 0), gstrDecPrice) '其他变价取的缺省价格
   
            '金额
            dbl金额 = CCur(NVL(rsTmp!数量, 0) * dbl单价)
            If NVL(rsTmp!附加手术, 0) = 1 Then
                dbl金额 = dbl金额 * NVL(rsTmp!附术收费率, 100) / 100
            End If
            dbl金额 = Format(dbl金额, gstrDec)
            
            If NVL(rsTmp!屏蔽费别, 0) = 0 And str费别 <> "" Then
                dbl金额 = ActualMoney(str费别, rsTmp!收入项目ID, dbl金额, rsTmp!收费细目ID, lng执行科室ID, NVL(rsTmp!数量, 0))
            End If
            
            dbl实收 = dbl实收 + dbl金额
            
            '项目变化时加入
            str项目 = rsTmp!序号 & "," & rsTmp!收费细目ID
            blnDo = False: rsTmp.MoveNext
            If Not rsTmp.EOF Then
                If rsTmp!序号 & "," & rsTmp!收费细目ID <> str项目 Then blnDo = True
            Else
                blnDo = True
            End If
            rsTmp.MovePrevious
            
            If blnDo Then
                rsPrice.AddNew
                rsPrice!病人ID = lng病人ID
                rsPrice!主页ID = lng主页ID
                rsPrice!收费类别 = rsTmp!收费类别
                rsPrice!收费细目ID = rsTmp!收费细目ID
                rsPrice!数量 = NVL(rsTmp!数量, 0)
                rsPrice!单价 = dbl单价
                rsPrice!实收金额 = dbl金额
                rsPrice!开单人 = UserInfo.姓名
                rsPrice!开单科室 = CStr(Sys.RowValue("部门表", lng开单科室ID, "名称"))
'                If lng疾病ID <> 0 Then mrsPrice!疾病id = lng疾病ID
'                If lng诊断ID <> 0 Then mrsPrice!诊断id = lng诊断ID
                rsPrice.Update
                dbl实收 = 0
            End If
            
            rsTmp.MoveNext
        Loop
        
        rsPrice.MoveFirst
        MakePriceRecord申请单 = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FuncClinicPay(frmMain As Object, ByVal lng病人ID As Long, ByVal strNO As String)
'功能：诊间支付
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str结算医嘱IDs As String
    Dim bln使用预交 As Boolean
    
    On Error GoTo errH
    
    If gobjSquareCard Is Nothing Then
        On Error Resume Next
        Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        err.Clear: On Error GoTo 0
        If Not gobjSquareCard Is Nothing Then
            If gobjSquareCard.zlInitComponents(frmMain, p门诊医嘱下达, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set gobjSquareCard = Nothing
            End If
        End If
    End If
    
    If Not gobjSquareCard Is Nothing Then
        strSQL = "Select f_List2str(Cast(Collect(a.医嘱序号 || '') As t_Strlist)) As 医嘱ids" & vbNewLine & _
            "From 病人医嘱记录 B, 门诊费用记录 A" & vbNewLine & _
            "Where a.病人id =[1] And a.记录性质=1 And a.记录状态 = 0 And a.医嘱序号=b.Id And b.医嘱状态=8 And b.挂号单 =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, strNO)
        If Not IsNull(rsTmp!医嘱ids) Then
            bln使用预交 = Val(zlDatabase.GetPara("诊间支付允许使用预交款", glngSys, p门诊医嘱下达, "1")) = 1
            str结算医嘱IDs = rsTmp!医嘱ids & ""
            Call gobjSquareCard.zlSquareAffirm(frmMain, p门诊医嘱下达, GetInsidePrivs(p门诊医嘱下达), lng病人ID, 0, False, 1, , str结算医嘱IDs, , , bln使用预交)
        Else
            MsgBox "当前病人没有可结算的医嘱费用。", vbInformation, gstrSysName
        End If
    Else
        MsgBox "医疗卡部件（zl9CardSquare）初始化失败!", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function InitObjPublicDrug() As Boolean
    If gobjPublicDrug Is Nothing Then
        On Error Resume Next
        Set gobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
        If Not gobjPublicDrug Is Nothing Then
            Call gobjPublicDrug.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjPublicDrug = Not gobjPublicDrug Is Nothing
End Function

Public Function FuncLisRptFileView(frmMain As Object, ByVal lng医嘱ID As Long) As Boolean
'功能：打开LIS报告文件查看
'返回：成功打开－true
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String
    Dim objFile As New FileSystemObject
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng报告ID As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    strSQL = "select a.id,a.类型,a.报告名 from 医嘱报告内容 a,病人医嘱报告 b where a.id=b.报告id and b.医嘱id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
    
    If rsTmp.EOF Then
        MsgBox "该医嘱没有产生报告文件！", vbInformation, gstrSysName:
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If rsTmp.RecordCount = 1 Then
        lng报告ID = Val(rsTmp!ID & "") '如果只有一份则直接打开
    Else
        strSQL = "select a.id,b.查阅状态 as 记录状态,null as 处理人,c.姓名,c.年龄,c.性别,a.报告名 as 内容,d.名称 as 执行科室," & vbNewLine & _
            " a.创建时间 as 记录时间,a.类型,a.报告名" & vbNewLine & _
            " from 医嘱报告内容 a,病人医嘱报告 b,病人医嘱记录 c,部门表 d" & vbNewLine & _
            " where a.id=b.报告id and b.医嘱id=c.id and c.执行科室id=d.id and c.id=[1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng医嘱ID)
        Screen.MousePointer = 0
        lng报告ID = frmCardSel.ShowMe(rsTmp, frmMain)
        rsTmp.Filter = "ID=" & lng报告ID
        Screen.MousePointer = 11
    End If
    
    If lng报告ID <> 0 Then
        strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!报告名 & "." & IIF(Val(rsTmp!类型 & "") = 0, "pdf", "html")
        If objFile.FileExists(strFile) Then objFile.DeleteFile strFile, True
        
        strFile = Sys.ReadLob(glngSys, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
        
        lngRetu = ShellExecute(frmMain.hwnd, "open", strFile, "", "", SW_SHOWNORMAL)
        If lngRetu <= 32 Then
            Select Case lngRetu
            Case 2: strInfo = "错误的关联"
            Case 29: strInfo = "关联失败"
            Case 30: strInfo = "关联应用程式忙碌中..."
            Case 31: strInfo = "没有关联任何应用程式"
            Case Else: strInfo = "无法识别的错误"
            End Select
            MsgBox "文件打开时出错：" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
        Else
            '成功打开后标记为已阅
            strSQL = "Zl_报告查阅记录_Insert(" & lng医嘱ID & ",null," & lng报告ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISKernel")
            FuncLisRptFileView = True
        End If
    End If
    
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FlexScroll(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能：支持滚轮的滚动
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
        Case -7864320  '向下滚
            zlCommFun.PressKey vbKeyPageDown
        Case 7864320   '向上滚
            zlCommFun.PressKey vbKeyPageUp
        End Select
    End Select
    FlexScroll = CallWindowProc(glngPreHWnd, hwnd, wMsg, wParam, lParam)
End Function




Public Function GetTsPrivs(ByVal lngMdl As Long) As String
    '读取特殊医嘱权限
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim i As Integer
    On Error GoTo errH

    
    If lngMdl = p门诊医嘱下达 Then
        If gstrTsPrivsMZ <> "" Then
            GetTsPrivs = gstrTsPrivsMZ
            Exit Function
        Else
            strSQL = "Select 门诊特殊医嘱权限 as 权限 From 人员表 Where ID=[1]"
        End If
    Else
        If gstrTsPrivsZY <> "" Then
            GetTsPrivs = gstrTsPrivsZY
            Exit Function
        Else
            strSQL = "Select 住院特殊医嘱权限 as 权限 From 人员表 Where ID=[1]"
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get部门性质", UserInfo.ID)
    If rsTmp.RecordCount <> 0 Then
        strTmp = IIF(rsTmp!权限 & "" = "", "0000", rsTmp!权限 & "")
    End If
    
    '组织权限字符串
    If strTmp = "0000" Then
        GetTsPrivs = ";"
    Else
        GetTsPrivs = IIF(Mid(strTmp, 1, 1) = 1, ";下达毒性药嘱", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 2, 1) = 1, ";下达麻醉药嘱", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 3, 1) = 1, ";下达精神药嘱", "")
        GetTsPrivs = GetTsPrivs & IIF(Mid(strTmp, 4, 1) = 1, ";下达贵重药嘱", "")
        GetTsPrivs = GetTsPrivs & ";"
    End If
    
    '缓存门诊住院权限数据
    If lngMdl = p门诊医嘱下达 Then
        gstrTsPrivsMZ = GetTsPrivs
    Else
        gstrTsPrivsZY = GetTsPrivs
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get自定义申请单(ByVal int场合 As Integer, ByRef str自定义申请单IDs As String) As Boolean
'功能：获取设置好的自定义申请单
    Dim strSQL As String, rsTmp As Recordset
    Dim strReturn As String
    
    If str自定义申请单IDs <> "" Then Exit Function
    
    strSQL = "Select a.Id, a.名称" & vbNewLine & _
                "From 病历文件列表 A, 自定义申请单文件 B" & vbNewLine & _
                "Where a.Id = b.文件id And a.格式 = 1 and exists(select 1 from 病历单据应用 C Where A.Id=C.病历文件ID And C.应用场合=[1])" & vbNewLine & _
                "Group By a.Id, a.名称" & vbNewLine & _
                "Having Count(1) = 3" '3个文件都有的时候才显示，否则可能无法使用
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get自定义申请单", 2)
    Do While Not rsTmp.EOF
        strReturn = strReturn & "|" & rsTmp!ID & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    str自定义申请单IDs = Mid(strReturn, 2)
    Get自定义申请单 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatiOutPathInfo(ByVal lng病人ID As Long, ByVal lng科室id As Long, Optional ByRef str路径项目分类 As String = "-1") As ADODB.Recordset
'功能：获取路径病人当前路径信息
'返回：str分类=当前天数最后一个路径项目所属的分类
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str路径项目分类 <> "-1"
    str路径项目分类 = ""
    strSQL = "Select a.路径记录id, a.当前阶段id, a.当前天数, b.日期, b.分类" & vbNewLine & _
            "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, Max(b.Id) 执行id" & vbNewLine & _
            "       From 病人门诊路径 A, 病人门诊路径执行 B, 病人门诊路径记录 C" & vbNewLine & _
            "       Where C.路径记录ID=A.ID And a.Id = b.路径记录id And b.阶段id = a.当前阶段id And b.天数 = a.当前天数 And a.状态 =1 And A.病人ID =[1] and A.科室ID=[2]" & vbNewLine & _
            "       Group By a.Id, a.当前阶段id, a.当前天数) A, 病人门诊路径执行 B" & vbNewLine & _
            "Where a.执行id = b.Id"
    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSQL, "病人当前路径信息", lng病人ID, lng科室id)
    If rsRet.RecordCount > 0 And blnDo Then
        str路径项目分类 = "" & rsRet!分类
        
        '如果当天生成了医嘱类项目，则取医嘱类项目的分类
        strSQL = "Select 分类" & vbNewLine & _
                "From 病人门诊路径执行" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From 病人门诊路径执行 A" & vbNewLine & _
                "            Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3] And Exists (Select 1 From 病人门诊路径医嘱 B Where a.Id = b.路径执行id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人当前路径信息", Val(rsRet!路径记录id), Val(rsRet!当前阶段ID), CDate(rsRet!日期))
        If rsTmp.RecordCount > 0 Then
            str路径项目分类 = "" & rsTmp!分类
        End If
    End If
    Set GetPatiOutPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FuncViewDrugExplain(ByVal lng药品ID As Long, objParent As Object)
'功能：查看药品说明书
    Dim objDrugExplain As New frmDrugExplain
    
    If lng药品ID = 0 Then
        MsgBox "当前药品未按规格下达，不能查看说明书。", vbInformation, gstrSysName
        Exit Sub
    End If
    objDrugExplain.ShowMe lng药品ID, objParent
End Sub

Public Function InitObjBlood(Optional ByVal blnMsg As Boolean = True) As Boolean
'判断如果血库部件为空就初始化
    If gobjPublicBlood Is Nothing Then
        On Error Resume Next
        Set gobjPublicBlood = CreateObject("zlPublicBlood.clsPublicBlood")
        If Not gobjPublicBlood Is Nothing Then
            If gobjPublicBlood.zlInitCommon(gcnOracle, gstrDBUser) = False Then
                Set gobjPublicBlood = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    InitObjBlood = Not gobjPublicBlood Is Nothing
    If InitObjBlood = False And blnMsg = True Then
        MsgBox "血库公共部件[zlPublicBlood]创建失败，将会影响输血流程及功能使用，请检查此部件是否存在以及是否正确注册！", vbInformation, gstrSysName
    End If
End Function

Public Function GetBloodCapacity(ByVal int场合 As Integer, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal strDate As String, ByVal bln计算24H量 As Boolean, _
    Optional ByVal intBaby As Integer, Optional ByVal lngActiveID As Long = 0) As Double
'功能：计算病人24小时输血量或本次就诊输血总量
'参数：int场合--1 门诊;2-住院
'         lng病人ID：病人身份标识ID
'         lng就诊ID---int场合=1为挂号ID；int场合=2为主页ID
'         strDate---时间
'         bln计算24H量---true计算24总量,false本次就诊总量
'         intBaby---婴儿序号
'         lngActiveID--医嘱ID，不为0则计算的量剔除改医嘱
    Dim rsTemp As New Recordset
    Dim dblNum As Double
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    If int场合 = 2 Then
        strSQL = _
            " Select Id, 开嘱时间, 申请量, 输血时间" & vbNewLine & _
            " From (With 医嘱记录 As (Select Decode(Nvl(c.医嘱id, 0), 0, b.诊疗项目id, c.诊疗项目id) 诊疗项目id, b.Id, b.开嘱时间," & vbNewLine & _
            "                            Decode(Nvl(c.医嘱id, 0), 0, b.总给予量, c.申请量) 申请量," & vbNewLine & _
            "                            Nvl(To_Char(b.手术时间, 'YYYY-MM-DD HH24:MI'), b.标本部位) As 输血时间" & vbNewLine & _
            "                     From 输血申请项目 c, 诊疗项目目录 p, 病人医嘱记录 q, 病人医嘱记录 b" & vbNewLine & _
            "                     Where c.医嘱id(+) = b.Id And p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And" & vbNewLine & _
            "                           q.相关id = b.Id And q.诊疗类别 = 'E' And b.病人id = [1] And b.主页id = [2] And Nvl(b.婴儿, 0) = [3] And" & vbNewLine & _
            "                           b.诊疗类别 = 'K' And b.医嘱状态 Not In (-1, 2, 4))" & vbNewLine & _
            "        Select b.Id, b.开嘱时间, b.申请量 * Decode(Upper(a.计算单位), 'ML', 1, Nvl(a.计算系数, 1)) 申请量, b.输血时间" & vbNewLine & _
            "        From 诊疗项目目录 a, 医嘱记录 b" & vbNewLine & _
            "        Where a.Id = b.诊疗项目id)"
    ElseIf int场合 = 1 Then
        strSQL = _
            " Select Id, 开嘱时间, 申请量, 输血时间" & vbNewLine & _
            " From (With 医嘱记录 As (Select Decode(Nvl(c.医嘱id, 0), 0, b.诊疗项目id, c.诊疗项目id) 诊疗项目id, b.Id, b.开嘱时间," & vbNewLine & _
            "                            Decode(Nvl(c.医嘱id, 0), 0, b.总给予量, c.申请量) 申请量," & vbNewLine & _
            "                            Nvl(To_Char(b.手术时间, 'YYYY-MM-DD HH24:MI'), b.标本部位) As 输血时间" & vbNewLine & _
            "                     From 输血申请项目 c, 诊疗项目目录 p, 病人医嘱记录 q, 病人医嘱记录 b, 病人挂号记录 d" & vbNewLine & _
            "                     Where c.医嘱id(+) = b.Id And p.Id = q.诊疗项目id And (p.操作类型 = 9 Or p.操作类型 = 8 And p.执行分类 = 0) And" & vbNewLine & _
            "                           q.相关id = b.Id And q.诊疗类别 = 'E' And b.诊疗类别 = 'K' And b.医嘱状态 Not In (-1, 2, 4) And b.挂号单 = d.No And" & vbNewLine & _
            "                           d.病人id = [1] And d.Id = [2])" & vbNewLine & _
            "        Select b.Id, b.开嘱时间, b.申请量 * Decode(Upper(a.计算单位), 'ML', 1, Nvl(a.计算系数, 1)) 申请量, b.输血时间" & vbNewLine & _
            "        From 诊疗项目目录 a, 医嘱记录 b" & vbNewLine & _
            "        Where a.Id = b.诊疗项目id)"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取输血量", lng病人ID, lng就诊ID, intBaby)
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        dblNum = 0
        Do While Not rsTemp.EOF
            If lngActiveID <> Val(rsTemp!ID & "") Then
                If bln计算24H量 = False Then
                    dblNum = dblNum + Val("" & rsTemp("申请量"))
                Else
                    If rsTemp("输血时间") & "" <> "" Then
                        If CDate(rsTemp("输血时间")) > CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) - 1 And CDate(rsTemp("输血时间")) <= CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) Then dblNum = dblNum + Val("" & rsTemp("申请量"))
                    Else
                        If CDate(rsTemp("开嘱时间")) > CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) - 1 And CDate(rsTemp("开嘱时间")) <= CDate(Format(strDate, "yyyy-MM-dd HH:mm:ss")) Then dblNum = dblNum + Val("" & rsTemp("申请量"))
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    GetBloodCapacity = dblNum
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetBloodVerifyState(ByVal int病人来源 As Integer, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str申请时间 As String, ByVal dblTotalByML As Double, _
    Optional ByVal int紧急 As Integer, Optional ByVal int备血 As Integer, Optional ByVal intBaby As Integer, Optional ByVal lngActiveID As Long) As String
'功能：根据输血审核模式，重新计算备血医嘱的审核状态(输血三级别审核)：目前只支持启用了血库系统，并且是申请单方式下达申请
'参数：
'   int病人来源：1-门诊；2-住院
'   lng病人ID：病人ID
'   lng就诊ID：int病人来源=1为挂号ID；int病人来源=2为主页ID
'   str申请时间：输血申请的预定输血时间，用于计算24小时总量使用
'   dblTotalByML：本次输血申请总量，注意：总量为换算后的ML单位总量
'   int紧急：是否是紧急医嘱，紧急医嘱不用审核-0－非紧急，1－紧急；
'   int备血：是否是备血医嘱，只有备血医嘱才审核；0－备血 ，1－用血
'   intBaby：婴儿序号
'   lngActiveID：医嘱ID。注意：不为0时，本次计算的24小时总量将不包含该医嘱对应的申请量。主要是针对申请单修改时剔除之前的总量
    Dim strPrivs As String
    Dim dbl24h量 As Double
    Dim str审核状态 As String
    
    On Error GoTo ErrHand
    str审核状态 = GetBloodState(int紧急, int备血)
    If str审核状态 = "1" And gbln血库系统 = True Then '审核状态已经决定了是备血申请
        ' 具有医务科权限的，则直接通过;具有医生权限的<1600则直接通过,否则按照审核规则
        strPrivs = GetInsidePrivs(p输血审核管理)
        If gbln输血申请三级审核 = True Then
            If InStr(";" & strPrivs & ";", ";医务科;") = 0 Then
                '修改时则需要剔除此医嘱本身（GetBloodCapacity中排除）
                dbl24h量 = GetBloodCapacity(int病人来源, lng病人ID, lng就诊ID, str申请时间, True, intBaby, lngActiveID)
                
                If InStr(";" & strPrivs & ";", ";科主任;") <> 0 Then   '只具有科主任权限
                     If dbl24h量 < 1600 - dblTotalByML Then '小于1600直接通过
                         str审核状态 = 4
                     Else
                         str审核状态 = 7 '大于1600，进入医务科审核环节
                     End If
                Else
                    '即没有医务科也没有科主任权限
                    If dbl24h量 < 800 - dblTotalByML Then '小于800，开单人是上级医师则直接过,否则进入初审环节
                        If UserInfo.专业技术职务 = "主任医师" Or UserInfo.专业技术职务 = "副主任医师" Or UserInfo.专业技术职务 = "主治医师" Then
                            str审核状态 = 4
                        End If
                    Else
                        '小于1600的，开单人是上级医师则进入二级审核发环节,否则进入初审环节
                        If dbl24h量 < 1600 - dblTotalByML And (UserInfo.专业技术职务 = "主任医师" Or UserInfo.专业技术职务 = "副主任医师" Or UserInfo.专业技术职务 = "主治医师") Then
                            str审核状态 = 7
                        End If
                    End If
                End If
            Else
                str审核状态 = 4
            End If
        Else
            str审核状态 = IIF(strPrivs <> "", 4, 1)
        End If
    End If
    GetBloodVerifyState = str审核状态
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetBloodTotalByML(ByVal strItems As String) As Double
'功能：根据申请项目和量，获取输血医嘱的总量(ML)
'参数：strItems:申请项目信息,格式：诊疗项目ID,申请量;诊疗项目ID,申请量
    Dim i As Integer
    Dim arrItem
    Dim strIDs As String
    Dim objItem As New Collection
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim dblNum As Double, dblTotal As Double
    
    If strItems = "" Then Exit Function
    On Error GoTo ErrHand
    arrItem = Split(strItems, ";")
    For i = 0 To UBound(arrItem)
        If InStr(1, "," & strIDs & ",", "," & Split(arrItem(i), ",")(0) & ",") = 0 Then
            strIDs = strIDs & "," & Split(arrItem(i), ",")(0)
            objItem.Add Val(Split(arrItem(i), ",")(1)), "_" & Split(arrItem(i), ",")(0)
        End If
    Next
    If Left(strIDs, 1) = "," Then strIDs = Mid(strIDs, 2)
    strSQL = "Select /*+ CARDINALITY(b 10) */" & vbNewLine & _
        "  a.id,a.名称, a.计算单位, a.计算系数" & vbNewLine & _
        " From 诊疗项目目录 a, Table(f_Num2list([1])) b" & vbNewLine & _
        " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetBloodTotalByML", strIDs)
    Do While Not rsTemp.EOF
        dblNum = Val(objItem("_" & rsTemp!ID))
        If UCase(rsTemp!计算单位 & "") <> "ML" Then
            dblNum = dblNum * Val(NVL(rsTemp!计算系数, 1))
        End If
        dblTotal = dblTotal + dblNum
        rsTemp.MoveNext
    Loop
    GetBloodTotalByML = dblTotal
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Get原液皮试(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String) As ADODB.Recordset
'功能：按病人获取原液皮试的医嘱项目
'参数：门诊病人传 str挂号单；               lng病人ID, lng主页ID 传为0
'      住院病人传入 lng病人ID, lng主页ID ； str挂号单 传空串
'返回：记录集形式
'说明：用  病人医嘱发送.标本发送批号 来关联药品行和过敏实验医嘱行，使用用 药品行的医嘱ID。
'      规则 一个原液过敏实验项目，收费对照中对照了药品费用(药品A)，病人下达此皮试并且标记结果为阴性，当病人再次发送药品医嘱A时总量-1
'      一但减-1后则打上标记后续发送不用再减。
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim rs皮试 As ADODB.Recordset
    Dim str标号 As String
    
    On Error GoTo errH
    
    '原液皮试，结果为阴性，标本发送批号 来关联
    If str挂号单 = "" Then
        strSQL = "Select b.医嘱Id as 皮试医嘱ID,nvl(b.标本发送批号,0) as 标号,max(e.id) as 药品ID" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 诊疗收费关系 D, 收费项目目录 E" & vbNewLine & _
            "Where a.皮试结果 = '(-)' And a.Id = b.医嘱id And a.诊疗项目id = c.Id And c.类别 = 'E' And c.操作类型 = '1' And c.执行分类 = 5" & vbNewLine & _
            "And  c.Id=d.诊疗项目id And d.收费项目id = e.Id And e.类别 In ('5', '6') And a.病人ID =[1] and a.主页ID=[2]" & vbNewLine & _
            "group by b.医嘱id,b.标本发送批号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    Else
        strSQL = "Select b.医嘱Id as 皮试医嘱ID,nvl(b.标本发送批号,0) as 标号,max(e.id) as 药品ID" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 诊疗收费关系 D, 收费项目目录 E" & vbNewLine & _
            "Where a.皮试结果 = '(-)' And a.Id = b.医嘱id And a.诊疗项目id = c.Id And c.类别 = 'E' And c.操作类型 = '1' And c.执行分类 = 5" & vbNewLine & _
            "And  c.Id=d.诊疗项目id And d.收费项目id = e.Id And e.类别 In ('5', '6') And a.挂号单 =[1]" & vbNewLine & _
            "group by b.医嘱id,b.标本发送批号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", str挂号单)
    End If
    Set rs皮试 = zlDatabase.CopyNewRec(rsTmp)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!标号 & "") <> 0 Then
                str标号 = str标号 & "," & Val(rsTmp!标号 & "")
            End If
            rsTmp.MoveNext
        Next
        
        If str标号 <> "" Then
            strSQL = "select a.id from 病人医嘱记录 a,病人医嘱发送 b where a.id=b.医嘱ID and a.id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X) group by a.id"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", Mid(str标号, 2))
            str标号 = ""
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    str标号 = str标号 & "," & rsTmp!ID '已经关联处理了的
                    rsTmp.MoveNext
                Next
            End If
            For i = 1 To rs皮试.RecordCount
                If Val(rs皮试!标号 & "") <> 0 Then
                    If InStr(str标号 & ",", "," & rs皮试!标号 & ",") = 0 Then
                        rs皮试!标号 = 0
                    End If
                End If
                rs皮试.MoveNext
            Next
            rs皮试.MoveFirst
        End If
    End If
    Set Get原液皮试 = rs皮试
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get原液皮试药品(ByVal lng项目id As Long) As Long
'功能：获取原液皮试对应的药品信息
'参数：lng项目ID 原液皮试项目ID
'返回：收费对照中的 药品收费项目ID（规格ID）
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Max(b.Id) As 药品id" & vbNewLine & _
        "From 诊疗收费关系 A, 收费项目目录 B" & vbNewLine & _
        "Where a.收费项目id = b.Id And b.类别 In ('5', '6') And a.诊疗项目id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id)
    
    If Not rsTmp.EOF Then
        Get原液皮试药品 = Val(rsTmp!药品ID & "")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BloodApplyPrintCheck(ByVal lng医嘱ID As Long, ByVal int场合 As Integer, ByVal int类型 As Integer, ByVal int模式 As Integer) As Boolean
'功能说明:申请单预览打印时对医嘱的相关内容进行检查，并返回提示及处理结果。
' --入参说明：
' ----调用场合_in=1-门诊,2-住院
' ----申请类型_In=1-输血申请单;2-取血通知单(便于医院根据申请类型控制)
' ----用血安排_In=0-普通输血;1-紧急输血(便于医院根据输血紧急程度控制)
' ----模式_in:0=预览是调用；1-打印是调用
' --函数返回："处理结果|提示信息",处理结果=0-正常,1-询问提示,2-禁止；处理结果为0时，无需返回提示信息及分隔符。
'返回：TRUE允许,False禁止
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strMsg As String
    
    On Error GoTo ErrHand
    strSQL = "Select Zl1_Fun_BloodApplyPrint([1],[2],[3],[4]) as 结果 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl1_Fun_BloodApplyPrint", lng医嘱ID, int场合, int类型, int模式)
    If Not rsTmp.EOF Then
        strMsg = NVL(rsTmp!结果)
        If strMsg <> "" Then
            Select Case Val(Split(strMsg, "|")(0))
            Case 1 '提示
                If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    strMsg = "": Exit Function
                End If
            Case 2 '禁止
                MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                strMsg = "": Exit Function
            End Select
            strMsg = ""
        End If
    End If
                
    BloodApplyPrintCheck = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetApplyCustom(ByVal lng项目id As Long) As Long
'功能：获取自定义申请单的文件ID
'参数：lng项目ID 诊疗项目ID
'返回：自定义申请单的文件ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.病历文件ID" & vbNewLine & _
        "From 病历单据应用 A, 病历文件列表 B" & vbNewLine & _
        "Where a.病历文件id = b.Id And b.种类 = 7 And Nvl(b.格式, 0) = 1 And a.诊疗项目id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng项目id)
    
    If Not rsTmp.EOF Then
        GetApplyCustom = Val(rsTmp!病历文件ID & "")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get会诊医嘱IDs(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strDepts As String) As String
'功能：获取病人指定科室的会诊医嘱
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH

    strSQL = "Select a.id From 病人医嘱记录 A, 诊疗项目目录 B Where a.诊疗项目id = b.Id And a.诊疗类别 = 'Z' And b.操作类型 = '7' And a.医嘱状态 = 8 And a.病人id =[1] And a.主页id =[2] And a.开嘱科室id In (Select /*+cardinality(x,10)*/ x.Column_Value  From Table(f_Num2list([3])) X )"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID, strDepts)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        Get会诊医嘱IDs = Mid(strTmp, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckCanSendAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngSpecialAdviceID As Long, ByVal intBabyNum As Integer) As Boolean
'功能：截止到当前特殊的开始执行时间内是否还有可以发送的长期医嘱
'参数：strSpecialAdviceIDs=特殊医嘱ID转科医嘱
'      intBabyNum=婴儿序号。
'返回：true 存在，false 不存在
'说明：该函数执行时只发送母亲特殊医嘱时会忽略对婴儿医嘱的检查，如果同时发送则都正常。
'      忽略  备用医嘱，输液医嘱只发两天的情况

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strEnd As String
    Dim rsSend As ADODB.Recordset
    Dim i As Long
    Dim strPause As String
    Dim datBegin As Date
    Dim datEnd As Date
    Dim str首次时间 As String
    Dim str末次时间 As String
    Dim str分解时间 As String
    Dim lng次数 As Long
    Dim lng组ID As Long
    
    
    On Error GoTo errH
    
    If Val(zlDatabase.GetPara("存在未发送医嘱时禁止处理转科医嘱", glngSys, p住院医嘱发送, 0)) = 0 Then Exit Function
    
    '不用判断长嘱是不是已经停止，校对/发送 转科医嘱时会自动停，且停嘱时间是转科医嘱的[开始执行时间]
    '计算医嘱是不是可以发送，截止时间为转科医嘱的开始执行时间，医嘱期效都是长嘱，已校对状态
    strSQL = "Select to_char(b.开始执行时间,'yyyy-mm-dd hh24:mi:ss') as 日期 From 病人医嘱记录 b Where b.id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", lngSpecialAdviceID)
    If Not rsTmp.EOF Then
        strEnd = rsTmp!日期 & ""
        '其它非药长嘱
        strSQL = "select a.id, a.相关ID,Nvl(A.相关ID,A.ID) as 组ID,a.诊疗类别,a.医嘱状态,a.开始执行时间,a.上次执行时间,a.执行时间方案,a.频率次数,a.频率间隔,a.间隔单位,a.执行终止时间" & _
            " from 病人医嘱记录 a,诊疗项目目录 E where a.诊疗项目ID=e.id and a.病人id=[2] and a.主页id=[3] and nvl(a.婴儿,0)=[4]" & _
            " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4) And A.开始执行时间<=[1] And (A.上次执行时间<[1] Or A.上次执行时间 is NULL)" & _
            " And (A.执行终止时间>A.上次执行时间 Or A.执行终止时间 is NULL Or A.上次执行时间 Is NULL)" & _
            " And (A.执行终止时间>A.开始执行时间 Or A.执行终止时间 is NULL) And A.医嘱期效=0" & _
            " And Not(Nvl(a.诊疗类别,'自由')='H' And E.操作类型='1' And E.执行频率=2)" & _
            " And Not(Nvl(a.诊疗类别,'自由')='Z' And E.操作类型 IN ('4','14','9','10','12'))" & _
            " And Not (a.诊疗类别 = 'E' And a.相关id Is Not Null And e.操作类型 = '3') And Nvl(a.诊疗类别, '自由') Not In ('5', '6', '7') And Not Exists (Select ID From 病人医嘱记录 X Where 诊疗类别 In ('5', '6', '7') And x.相关id = a.id)" & _
            " And a.开始执行时间 Is Not Null And Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3 And Nvl(a.执行频次, '无') <> '必要时' And Nvl(a.执行频次, '无') <> '需要时'" & _
            " order by a.序号"
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", CDate(strEnd), lng病人ID, lng主页ID, intBabyNum)
        For i = 1 To rsSend.RecordCount
            If IsNull(rsSend!相关ID) Then
                strPause = GetAdvicePause(rsSend!ID)
                '当前医嘱的发送计算时间段
                datBegin = rsSend!开始执行时间
                If Not IsNull(rsSend!上次执行时间) Then
                    If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                        datBegin = DateAdd("s", 1, rsSend!上次执行时间)
                    Else
                        datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, rsSend!频率间隔, rsSend!间隔单位)
                        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
                        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                    End If
                End If
                datEnd = CDate(strEnd)
                If Not IsNull(rsSend!执行终止时间) Then
                    If rsSend!执行终止时间 < CDate(strEnd) Then
                        datEnd = rsSend!执行终止时间
                    End If
                End If
                '计算分解时间及次数
                If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                    lng次数 = Calc持续性长嘱次数(datBegin, datEnd, Format(NVL(rsSend!上次执行时间), "yyyy-MM-dd HH:mm:ss"), Format(NVL(rsSend!执行终止时间), "yyyy-MM-dd HH:mm:ss"), strPause)
                    If lng次数 <> 0 Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                Else
                    '执行频率为"可选频率"的项目
                    str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位, rsSend!开始执行时间)
                    If str分解时间 <> "" Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                End If
            End If
            rsSend.MoveNext
        Next
                    
        '药品长嘱
        strSQL = "select a.id,a.相关ID,Nvl(A.相关ID,A.ID) as 组ID,a.诊疗类别,a.医嘱状态,a.开始执行时间,a.上次执行时间,a.执行时间方案,a.频率次数,a.频率间隔,a.间隔单位,a.执行终止时间" & _
            " from 病人医嘱记录 a,诊疗项目目录 E where a.诊疗项目ID=e.id and a.病人id=[2] and a.主页id=[3] and nvl(a.婴儿,0)=[4]" & _
            " And Nvl(A.医嘱状态,0) Not IN(-1,1,2,4) And A.开始执行时间<=[1] And (A.上次执行时间<[1] Or A.上次执行时间 is NULL)" & _
            " And (A.执行终止时间>A.上次执行时间 Or A.执行终止时间 is NULL Or A.上次执行时间 Is NULL)" & _
            " And (A.执行终止时间>A.开始执行时间 Or A.执行终止时间 is NULL) And A.医嘱期效=0" & _
            " And a.诊疗类别 In ('5', '6', '7') And a.开始执行时间 Is Not Null And Nvl(a.执行标记, 0) <> -1 And a.病人来源 <> 3 And Nvl(a.执行频次, '无') <> '必要时' And Nvl(a.执行频次, '无') <> '需要时'" & _
            " order by a.序号"
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱", CDate(strEnd), lng病人ID, lng主页ID, intBabyNum)
        For i = 1 To rsSend.RecordCount
            If lng组ID <> Val(rsSend!组ID & "") Then
                lng组ID = Val(rsSend!组ID & "")
                strPause = GetAdvicePause(rsSend!ID)
                '当前医嘱的发送计算时间段
                datBegin = rsSend!开始执行时间
                If Not IsNull(rsSend!上次执行时间) Then
                    If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                        datBegin = DateAdd("s", 1, rsSend!上次执行时间) '"持续性"的项目
                    Else
                        datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, rsSend!频率间隔, rsSend!间隔单位)
                        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
                        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                    End If
                End If
                datEnd = CDate(strEnd)
                If Not IsNull(rsSend!执行终止时间) Then
                    If rsSend!执行终止时间 < CDate(strEnd) Then
                        datEnd = rsSend!执行终止时间
                    End If
                End If
                '计算分解时间及次数
                If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                    lng次数 = Calc持续性长嘱次数(datBegin, datEnd, Format(NVL(rsSend!上次执行时间), "yyyy-MM-dd HH:mm:ss"), Format(NVL(rsSend!执行终止时间), "yyyy-MM-dd HH:mm:ss"), strPause)
                    If lng次数 <> 0 Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                Else
                    '执行频率为"可选频率"的项目
                    str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位, rsSend!开始执行时间)
                    If str分解时间 <> "" Then
                        CheckCanSendAdvice = True
                        Exit Function
                    End If
                End If
            End If
            rsSend.MoveNext
        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Svr预约入院取消服务(ByVal lng挂号ID As Long)
'功能：调用预约中心服务取消住院申请
    Dim strJsIn As String
    Dim strJsOut As String
    Dim strErr As String
   
    strJsIn = "{""input_in"":{""rgst_id"": """ & lng挂号ID & """}}"
    Call Sys.NewSystemSvr("预约中心", "住院申请取消", strJsIn, strJsOut, strErr)
    
    If strErr <> "" Then
        MsgBox "预约入院取消服务:" & strErr, vbInformation, gstrSysName
    End If
End Sub

Private Function GetParaURL(ByVal strSysName As String, ByVal strServiceName As String) As String
'功能:
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strUrl As String
    
    On Error GoTo errH
    strSQL = "Select 服务地址 From 三方服务配置目录 Where 系统标识 = [1] And 服务名称 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", strSysName, strServiceName)
    If Not rsTmp.EOF Then strUrl = Trim(rsTmp!服务地址 & "")
    GetParaURL = strUrl
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get输液类医嘱(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intType As Integer) As String
'功能：获取指定病人的某类特殊输液医嘱
'参数：intType 调用场合，0 - 一般发送窗口，1 - 输液发送窗口
'说明：营养、自备药/胰岛素、不取药、离院带药
'返回：需要从发送的医嘱中排除掉的  医嘱
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strIDs As String
    Dim str期效 As String
    Dim lngPar期效 As Long
    Dim strPar给药 As String
    Dim str自备 As String
    Dim str自备输液 As String
    Dim str不取 As String
    Dim str离院 As String
    Dim str营养 As String
    Dim str常规 As String '普通的正常输液医嘱
    Dim strAll As String
    Dim blnOK As Boolean
    Dim strNoInIDs As String
    Dim strTmpALL As String
    Dim varTmp As Variant
    Dim str性质 As String
    
    On Error GoTo errH
    
    lngPar期效 = Val(zlDatabase.GetPara("医嘱类型", glngSys, p输液配置中心, "1")) - 1
    If lngPar期效 = -1 Then
        str期效 = ",0,1,"
    Else
        str期效 = lngPar期效
    End If
     
    strPar给药 = zlDatabase.GetPara("输液给药途径", glngSys, p输液配置中心)
    
    '一般发送窗口和输液发送窗口两边过滤医嘱恰好相反，能在静配发的就不在一般发送中显示，在取参数值判断时更据入参intType来
    str性质 = ""
    blnOK = Val(zlDatabase.GetPara("自备药允许发往静配中心", glngSys, p输液配置中心, "0")) = intType
    str性质 = str性质 & IIF(blnOK, "1", "0")
    blnOK = Val(zlDatabase.GetPara("不取药允许发往静配中心", glngSys, p输液配置中心, "0")) = intType
    str性质 = str性质 & IIF(blnOK, "1", "0")
    blnOK = Val(zlDatabase.GetPara("离院带药允许发往静配中心", glngSys, p输液配置中心, "0")) = intType
    str性质 = str性质 & IIF(blnOK, "1", "0")
         
    strSQL = "Select a.id,a.相关id,a.执行性质 As 药品执行性质, a.执行标记 As 药品执行标记, b.执行性质 As 给药执行性质, b.执行标记 As 给药执行标记," & vbNewLine & _
        " c.操作类型 as 给药操作类型,c.执行分类 as 给药执行分类,c.执行标记 as 给药执行标记,nvl(a.执行科室id,0) as 药品执行科室id,c.id as 给药项目id,d.药品id as 输液自备,a.医嘱期效" & vbNewLine & _
        " From 病人医嘱记录 A, 病人医嘱记录 B,诊疗项目目录 c,输液自备药清单 d" & vbNewLine & _
        " Where a.相关id = b.Id and b.诊疗项目id=c.id and a.收费细目id=d.药品id(+) And a.诊疗类别 In ('5', '6')" & vbNewLine & _
        " And A.开始执行时间 is Not NULL And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 and nvl(a.医嘱状态,0)<>4 and a.病人id=[1] and a.主页id=[2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get输液类医嘱", lng病人ID, lng主页ID)
    
    For i = 1 To rsTmp.RecordCount
        blnOK = True
        
        If blnOK Then
            '给药途径输液类
            If Val(rsTmp!给药操作类型 & "") <> 2 Or Val(rsTmp!给药执行分类 & "") <> 1 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            '满足给药途径参数
            If strPar给药 <> "" Then
                If InStr("," & strPar给药 & ",", "," & rsTmp!给药项目id & ",") = 0 Then
                    blnOK = False
                End If
            End If
        End If
        
        If blnOK Then
            '满足期效参数
            If InStr("," & str期效 & ",", "," & rsTmp!医嘱期效 & ",") = 0 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            '满足科室参数
            If InStr("," & gstr输液配置中心 & ",0,", "," & rsTmp!药品执行科室id & ",") = 0 Then
                blnOK = False
            End If
        End If
        
        If blnOK Then
            If intType = 0 Then
                '一般发送窗口
                '常规的静配医嘱，后续会再处理
                If Val(rsTmp!药品执行科室id & "") > 0 Then
                    If InStr("," & strTmpALL & ",", "," & rsTmp!相关ID & ",") = 0 Then
                        strTmpALL = strTmpALL & "," & rsTmp!相关ID
                    End If
                End If
            End If
            
            
            If Val(rsTmp!给药执行性质 & "") <> 5 And Val(rsTmp!药品执行性质 & "") = 5 And Val(rsTmp!药品执行标记 & "") = 2 Then
                '不取药
                If InStr("," & strAll & ",", "," & rsTmp!相关ID & ",") = 0 Then
                    str不取 = str不取 & "," & rsTmp!相关ID
                    strAll = strAll & "," & rsTmp!相关ID
                End If
            ElseIf Val(rsTmp!给药执行性质 & "") <> 5 And Val(rsTmp!药品执行性质 & "") = 5 And Val(rsTmp!药品执行标记 & "") <> 2 Then
                '自备药--需进一步细分：一般自备药，用静配自备药
                '输液自备
                If Val(rsTmp!输液自备 & "") = 0 Then
                    If InStr("," & strAll & ",", "," & rsTmp!相关ID & ",") = 0 Then
                        str自备 = str自备 & "," & rsTmp!相关ID
                        strAll = strAll & "," & rsTmp!相关ID
                    End If
                Else
                    If InStr("," & strAll & ",", "," & rsTmp!相关ID & ",") = 0 Then
                        str自备输液 = str自备输液 & "," & rsTmp!相关ID
                        strAll = strAll & "," & rsTmp!相关ID
                    End If
                End If
            ElseIf Val(rsTmp!给药执行性质 & "") = 5 And Val(rsTmp!药品执行性质 & "") <> 5 Then
                '离院带药
                If InStr("," & strAll & ",", "," & rsTmp!相关ID & ",") = 0 Then
                    str离院 = str离院 & "," & rsTmp!相关ID
                    strAll = strAll & "," & rsTmp!相关ID
                End If
            ElseIf Val(rsTmp!给药执行标记 & "") = 2 Then
                '营养
                If InStr("," & strAll & ",", "," & rsTmp!相关ID & ",") = 0 Then
                    str营养 = str营养 & "," & rsTmp!相关ID
                    strAll = strAll & "," & rsTmp!相关ID
                End If
            End If
        End If
        rsTmp.MoveNext
    Next
    
    
    If strTmpALL <> "" Then
        varTmp = Split(strTmpALL, ",")
        strTmpALL = ""
        For i = 0 To UBound(varTmp)
            If InStr("," & str自备 & str自备输液 & str不取 & str离院 & str营养 & ",", "," & varTmp(i) & ",") = 0 Then
                strTmpALL = strTmpALL & "," & varTmp(i)
            End If
        Next
    End If
    
    '营养类固定排除
    '当参数启用不接收自备药时，  str自备输液 类要保留下不进行排除
    '依次为：自备药、不取药、离院带药
    Select Case str性质
    Case "000"
        If str自备 <> "" Then
            strNoInIDs = strNoInIDs & str自备
        End If
        If str不取 <> "" Then
            strNoInIDs = strNoInIDs & str不取
        End If
        If str离院 <> "" Then
            strNoInIDs = strNoInIDs & str离院
        End If
    Case "001"
        If str自备 <> "" Then
            strNoInIDs = strNoInIDs & str自备
        End If
        If str不取 <> "" Then
            strNoInIDs = strNoInIDs & str不取
        End If
'        If str离院 <> "" Then
'            strNoInIDs = strNoInIDs & str离院
'        End If
    Case "010"
        If str自备 <> "" Then
            strNoInIDs = strNoInIDs & str自备
        End If
'        If str不取 <> "" Then
'            strNoInIDs = strNoInIDs & str不取
'        End If
        If str离院 <> "" Then
            strNoInIDs = strNoInIDs & str离院
        End If
    Case "011"
        If str自备 <> "" Then
            strNoInIDs = strNoInIDs & str自备
        End If
'        If str不取 <> "" Then
'            strNoInIDs = strNoInIDs & str不取
'        End If
'        If str离院 <> "" Then
'            strNoInIDs = strNoInIDs & str离院
'        End If
    Case "100"
'        If str自备 <> "" Then
'            strNoInIDs = strNoInIDs & str自备
'        End If
        If str不取 <> "" Then
            strNoInIDs = strNoInIDs & str不取
        End If
        If str离院 <> "" Then
            strNoInIDs = strNoInIDs & str离院
        End If
    Case "101"
'        If str自备 <> "" Then
'            strNoInIDs = strNoInIDs & str自备
'        End If
        If str不取 <> "" Then
            strNoInIDs = strNoInIDs & str不取
        End If
'        If str离院 <> "" Then
'            strNoInIDs = strNoInIDs & str离院
'        End If
    Case "110"
'        If str自备 <> "" Then
'            strNoInIDs = strNoInIDs & str自备
'        End If
'        If str不取 <> "" Then
'            strNoInIDs = strNoInIDs & str不取
'        End If
        If str离院 <> "" Then
            strNoInIDs = strNoInIDs & str离院
        End If
    Case "111"
'        If str自备 <> "" Then
'            strNoInIDs = strNoInIDs & str自备
'        End If
'        If str不取 <> "" Then
'            strNoInIDs = strNoInIDs & str不取
'        End If
'        If str离院 <> "" Then
'            strNoInIDs = strNoInIDs & str离院
'        End If
    End Select
    
    
    '一般发送窗口固定排除  营养 、自备输液【输液自备药清单】、常规
    If intType = 0 Then
        If str营养 <> "" Then
            strNoInIDs = strNoInIDs & str营养
        End If
        
        If str自备输液 <> "" Then
            strNoInIDs = strNoInIDs & str自备输液
        End If
        
        If strTmpALL <> "" Then
            strNoInIDs = strNoInIDs & strTmpALL
        End If
    End If
    
    strNoInIDs = Mid(strNoInIDs, 2)
    Get输液类医嘱 = strNoInIDs
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowSQLSelectCIS(frmParent As Object, ByVal strSQL As String, ByVal strDetail As String, bytStyle As Byte, _
    ByVal strTitle As String, ByVal bln末级 As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：多功能选择器,使用ADO.Command打开,允许使用[x]参数
'参数：
'     frmParent=显示的父窗体
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     strDetail  明细级SQL
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型。其中，
'               格式为："bytSize=?"表示设置字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体。
'               格式为：ColSet:...时表示列宽设置,ColSet格式:列宽设置|列名1,宽度1;列名2,宽度2.....|悬浮提示|列名。
'               格式为：HeadCap=SQL列名1,列表展示列名1;SQL列名2,列表展示列名2；该项目用来手工指定SQL列在列表中展示名称，一般用于编码名称列，但是不改变列的Key
'               格式为：MultiCheckReturn=0,1：多选时只返回勾选行，由于多选点确定默认返回当前行所以增加该参数控制，该控制启用后，不支持默认行的返回，但是仍旧支持双击行自动返回。
'               格式为：HideNullCols=0,1;是否隐藏SQl中的null as 写法的列
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    arrPar = arrInput
    Set ShowSQLSelectCIS = frmNew.ShowSelect(frmParent, strSQL, strDetail, bytStyle, strTitle, bln末级, strSeek, strNote, blnShowSub, _
                                        blnShowRoot, blnNoneWin, X, Y, txtH, Cancel, blnMultiOne, blnSearch, False, arrPar)
End Function

Public Sub LisInfoTrans(ByVal strLIS As String, ByRef rsData As ADODB.Recordset, ByRef rsSub As ADODB.Recordset)
'功能:LIS检验申请单信息转换,将LIS申请单传回来的字符串信息转换成记录集
'参数:strLis 申请单信息
'     rsData 出参,记录集方式
'     rsSub 次级信息,主要为诊疗项目目录相关信息
    Dim arrTmp As Variant
    Dim arrTmp1 As Variant
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim str耐受 As String
    Dim lng序号 As Long
    Dim varRow医嘱 As Variant
    Dim str医嘱 As String
    Dim strALL项目IDs As String
    Dim str附项 As String
    
    On Error GoTo errH
    
    '记录集,一行记录代表一组医嘱
    Set rsData = New ADODB.Recordset
    rsData.Fields.Append "序号", adBigInt
    rsData.Fields.Append "医嘱", adVarChar, 4000
    rsData.Fields.Append "组号", adBigInt
    rsData.Fields.Append "是否耐受", adInteger
    rsData.Fields.Append "时间ID", adVarChar, 18
    rsData.Fields.Append "时间内容", adVarChar, 400
    
    rsData.Fields.Append "采集科室ID", adBigInt
    rsData.Fields.Append "采集项目ID", adBigInt
    rsData.Fields.Append "执行科室ID", adBigInt
    rsData.Fields.Append "检验项目ID", adBigInt
    rsData.Fields.Append "开始执行时间", adVarChar, 40
    rsData.Fields.Append "标本", adVarChar, 400
    rsData.Fields.Append "附项", adVarChar, 4000
    rsData.Fields.Append "嘱托", adVarChar, 4000
    rsData.Fields.Append "紧急", adBigInt
 
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    arrTmp = Split(strLIS, "<Split B>")
    For i = 0 To UBound(arrTmp)
        str医嘱 = arrTmp(i)
        varRow医嘱 = Split(str医嘱, "<Split A>")
        '正常申请单开出来的的检验医嘱只有一个检验项目ID
        'varRow医嘱(8)普通检查申请，下标只到8，耐受试验信息在varRow医嘱(9)中
        If UBound(varRow医嘱) > 8 Then
            str耐受 = varRow医嘱(9)
            arrTmp1 = Split(str耐受, "<split2>")
            If Val(arrTmp1(0)) = 1 Then '判断是否耐受重复增加医嘱行
                For j = 1 To UBound(arrTmp1)
                    varRow医嘱 = Split(arrTmp1(j), "<split3>")
                    lng序号 = lng序号 + 1
                    rsData.AddNew
                    rsData!序号 = lng序号
                    rsData!组号 = i + 1
                    rsData!医嘱 = str医嘱
                    rsData!是否耐受 = 1
                    rsData!时间ID = Val(varRow医嘱(0))
                    rsData!时间内容 = varRow医嘱(1)
                    Call LisInfoTransToRs(rsData!医嘱 & "", rsData, strALL项目IDs)
                    rsData.Update
                Next
            Else
                lng序号 = lng序号 + 1
                rsData.AddNew
                rsData!序号 = lng序号
                rsData!医嘱 = str医嘱
                Call LisInfoTransToRs(rsData!医嘱 & "", rsData, strALL项目IDs)
                rsData.Update
            End If
        Else
            lng序号 = lng序号 + 1
            rsData.AddNew
            rsData!序号 = lng序号
            rsData!医嘱 = str医嘱
            Call LisInfoTransToRs(rsData!医嘱 & "", rsData, strALL项目IDs)
            rsData.Update
        End If
    Next
    rsData.MoveFirst
    Set rsSub = Get诊疗项目记录(0, Mid(strALL项目IDs, 2))
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LisInfoTransToRs(ByVal strLIS As String, ByRef rsData As ADODB.Recordset, ByRef strALL项目IDs As String)
'功能:LIS申请信息添加到记录集中
'说明:仅在本模块的 LisInfoTrans 方法中调用
    Dim varRow医嘱 As Variant
 
    varRow医嘱 = Split(strLIS, "<Split A>")
    
    rsData!采集科室ID = Val(varRow医嘱(0))
    rsData!执行科室ID = Val(varRow医嘱(1))
    rsData!开始执行时间 = varRow医嘱(2)
    rsData!标本 = varRow医嘱(3)
    rsData!附项 = varRow医嘱(4)
    rsData!嘱托 = varRow医嘱(5)
    rsData!紧急 = Val(varRow医嘱(6))
    rsData!采集项目ID = Val(varRow医嘱(7))
    rsData!检验项目ID = Val(varRow医嘱(8)) '申请单开出来的的检验医嘱只有一个检验项目ID，不考虑一并采集的情况
    
    If InStr("," & strALL项目IDs & ",", "," & rsData!采集项目ID & ",") = 0 Then
        strALL项目IDs = strALL项目IDs & "," & rsData!采集项目ID
    End If
    
    If InStr("," & strALL项目IDs & ",", "," & rsData!检验项目ID & ",") = 0 Then
        strALL项目IDs = strALL项目IDs & "," & rsData!检验项目ID
    End If
      
End Sub

Public Function IsLis耐受项目(ByVal lngMod As Long, ByVal lng项目id As Long, ByRef strErr As String) As Boolean
'功能：检查当前检验项目是否是耐受项目
'参数：lngMod 模块号，lng项目ID 诊疗项目id
'      strErr 出参，错误信息
'说明：要考虑LIS部件中接口不存在的情况
    Dim blnTmp As Boolean
 
    strErr = ""
    Call InitObjLis(lngMod)
    If Not gobjLIS Is Nothing Then
        On Error Resume Next
        blnTmp = gobjLIS.IsToleranceItem(lng项目id, strErr)
        If Not blnTmp And strErr <> "" Then
            blnTmp = False
        End If
        If 438 = err.Number Then
            blnTmp = False
        End If
    End If
    IsLis耐受项目 = blnTmp
End Function

Public Function GetAdviceFeeKind(lngAdviceID As Long) As Byte
'功能：根据医嘱ID获取临嘱发送的费用单据的性质，1=门诊费用，2=住院费用
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    GetAdviceFeeKind = 2
    strSQL = "Select 记录性质,门诊记帐 From 病人医嘱发送 Where 医嘱ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetAdviceFeeKind", lngAdviceID)
    If rsTmp.RecordCount > 0 Then
        If rsTmp!记录性质 = 1 Or rsTmp!记录性质 = 2 And Val("" & rsTmp!门诊记帐) = 1 Then
            GetAdviceFeeKind = 1
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
