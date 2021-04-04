Attribute VB_Name = "mdlOutExse"
Option Explicit '要求变量声明
'=======系统控制相关变量============
Public Enum 身份验证Enum
    id门诊收费 = 0
    id入院登记 = 1
    id帐户管理 = 2
    id挂号 = 3
    id结帐 = 4
    id门诊确认 = 5
End Enum

Public Enum 医院业务
    support门诊预算 = 0
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    'support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤销出院 = 17            '允许撤消病人出院
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
    support出院无实际交易 = 29      '出院接口中是否要与接口商进行交易
    support多单据收费 = 30          '是否支持多单据收费
    
    support门诊收费存为划价单 = 31  '将门诊收费单转为划价单保存，修改以前固定判断某个医保的方式
    
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support多单据收费必须全退 = 39  '多单据收费必须全退
    
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support多单据一次结算 = 47      '多单据预结算时，医保接口仅在最后一次调用时返回结算结果，HIS中再分摊到每张单据上
    
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support医生确定处方类型 = 48
    support实时监控 = 60             '是否启用费用实时监控
    '刘兴洪:27536 20100119
    support不提醒缴款金额不足 = 64            '在收费时,如果收费参数的"不进行缴款输入和累计控制"为true时,同时是医保病人时没有输入缴款金额时不提醒用户
    support退费后打印回单 = 65   '医保病人是否退费后打印回单:问题
    support门诊_不分单据结算 = 80               '预结算、结算都只调用一次医保交易:一卡通同步更改
    
    support挂号不收取病历费 = 81    '在挂号时，不使用医保收取病历费

    support按单据全退 = 82 '门诊退费时，按单据进行退费，86176
    support多单据分单据结算 = 83 '多单据一次结算按单据进行医保报销，86321
    support一次结算分单据退费 = 85 '按一次结算调用医保接口，但按单据退费,91602
End Enum
Public Type Ty_InsurePatiPara
    允许不设置医保项目 As Boolean
    门诊收费存为划价单 As Boolean
    不提醒缴款金额不足 As Boolean
    门诊必须传递明细 As Boolean
    医保接口打印票据 As Boolean
    医生确定处方类型 As Boolean
    多单据一次结算 As Boolean
    门诊连续收费 As Boolean
    门诊预结算 As Boolean
    多单据收费 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    blnOnlyBjYb As Boolean '本地仅支持北京医保:刘兴洪
    退费后打印回单 As Boolean '
End Type
Public Enum 交易Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum
Public gobjCustBill As Object               '自定义记帐单部件
Public gobjRegist As Object                 '挂号部件
Public gobjPatient As Object                 '病人管理系统部件
Public gclsInsure As New clsInsure
Public gbytWarn As Byte '记帐报警返回值
Public gstrModiNO As String '修改后产生的新单据号
'============费用系统参数=====================

Public gstr医保费用类型 As String '医保病人允许的费用类型
Public gstr公费费用类型 As String '公费病人允许的费用类型
Public gstrCustomerAppellation As String    '对消费者的称呼:病人,客户
Public gbln退费申请模式 As Boolean '退费是否使用申请审核模式
Public gbln简码切换 As Boolean '35242
'病人输入方式
Public gblnInputName As Boolean '允许姓名输入
Public gblnInputID As Boolean '允许病人ID输入
Public gblnInputCard As Boolean '允许刷卡输入
Public gblnInputNO As Boolean '允许挂号单输入
Public gblnUnPopPriceBill As Boolean '不弹出划价单选择
Public gobjSquare As SquareCard  '卡结算部件  42301
'就诊卡
'Public gbytCardNOLen As Byte '就诊卡号长度
'Public gblnShowCard As Boolean '是否就诊卡号显示为正常符号
Public gobjPublicExpense As Object
Public gintPriceGradeStartType As Integer
Public gstr药品价格等级 As String
Public gstr卫材价格等级 As String
Public gstr普通价格等级 As String

Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbyt预存款退费验卡 As Byte '预存款退费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbln消费卡退费验卡 As Boolean '消费卡退费时是否刷卡验证

'票据控制
Public gobjBillPrint As Object '第三方票据打印部件
Public gblnBillPrint As Boolean '第三方票据打印部件是否可用

Public gblnStrictCtrl As Boolean '是否严格票据管理
Public gbytFactLength As Byte '票据号码长度
Public gblnSharedInvoice As Boolean
Public glngFactNormal As Long       '普通发票格式
Public glngFactMediCare As Long     '医保发票格式

'药房、窗口相关控制
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gbytAssign As Byte '发药窗口动态分配方式(0,1)
Public gbln收费后自动发药 As Boolean '是否自动发药退药
Public gbln门诊自动发料 As Boolean '门诊收费或记帐,记帐划价单审核后自动发料
Public gbytMediOutMode As Byte '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
Public gbln划价立即缴款 As Boolean '提取划价单后是否立即缴款:39253
Public gbln不显示无库存卫材 As Boolean
'单据输入控制
Public gstrLike As String  '项目匹配方法,%或空
Public gblnMyStyle As Boolean '使用个性化风格

Public gbln收费类别 As Boolean '是否首先输入类别
Public gblnFeeKindCode As Boolean '不输类别时,首位当作收费类别简码
Public gstrMatchMode As String  '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
Public gbln处方限量 As Boolean '处方限量检查,如果输入时选择了允许,则保存时,不再检查
Public gblnPrePayPriority As Boolean '优先使用预交款
Public gbytAutoSplitBill As Byte '单据按类别或执行科室自动分组
Public gbyt医保对码检查 As Byte '0-不进行检查、1-检查并提醒未对码项目、2-检查并禁止未对码项目
Public gcurMaxMoney As Currency '单笔费用最大提醒金额

'中药输入快捷
Public grsABCNum As ADODB.Recordset
Public gstrABC As String '输入允许的快捷字母
'刘兴洪 问题:????    日期:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000

'费用计算控制
Public gBytMoney As Byte '收费分币处理方法
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gblnMultiBalance As Boolean  '多张单据使用多种结算方式模式
Public glngAddedItem As Long    '自动加收挂号费的收费项目ID

'操作控制
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用
 

'==============费用本机参数===============
'票据控制
Public gobjTax As Object '税控打印接口对象
Public gblnTax As Boolean '本机是否使用税控打印
Public gstrTax As String
Public gint收费清单 As Integer      '0-不打印,1-要打印,2-选择是否打印
Public gint划价通知单 As Integer    '0-不打印,1-要打印,2-选择是否打印

'药房、窗口控制
Public glng西药房 As Long '指定的西药房,0为动态分配
Public glng中药房 As Long '指定的中药房,0为动态分配
Public glng成药房 As Long '指定的成药房,0为动态分配
Public glng发料部门 As Long '指定的卫材发料部门,0为动态分配

Public gstr西窗 As String  '指定的西药房发药窗口,空为动态分配
Public gstr中窗 As String '指定的中药房发药窗口,空为动态分配
Public gstr成窗 As String  '指定的成药房发药窗口,空为动态分配
Public gbln药房上班安排 As Boolean     '是否启用了药房上班安排

Public glng误差细目ID As Long           '是否设置了误差项
Public gstr误差费名称 As String         '误差费名称(结算方式中的误差费名称)

Public gstr误差收据费目 As String
'分离发药时要检查库存的药房
Public gstr西药房 As String
Public gstr成药房 As String
Public gstr中药房 As String

Public gbln其它药房 As Boolean '是否显示其它药房库存
Public gbln其它药库 As Boolean '是否显示其它药库库存
Public gbyt库存显示方式 As Byte   '问题:31936: 其他药房或药库(非操作员所属科室)的库存显示方式:0-直接显示库存;1-显示有无

'输入控制
Public gstrCardPass As String '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
Public gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
Public gstr收费类别 As String '可输入的收费类别
Public gblnPay As Boolean '中药是否输入付数
Public gblnTime As Boolean '变价项目是否输入数次
Public gbyt科室医生 As Byte '0-开单人确定开单科室,1-开单科室确定开单人,2-开单人和开单科室相互独立
Public gbln不缺省开单人 As Boolean
Public gbln缺省科室优先 As Boolean
Public gbln必须输开单人 As Boolean
Public gcurMax As Currency '单据允许输入的最大金额

Public gstr费别 As String '缺省的病人费别
Public gstr结算方式 As String  '缺省使用的结算方式
Public gstr所属部门ID As String '存储当前操作员的所属部门ID,
Public gbln性别 As Boolean '光标是否经过该项目
Public gbln年龄 As Boolean
Public gbln费别 As Boolean
Public gbln医疗付款 As Boolean
Public gbln加班 As Boolean
Public gbln开单日期 As Boolean
Public gbln开单人 As Boolean
Public gbyt开单人显示 As Byte

Public gblnSeekName As Boolean '是否通过姓名进行模糊查找
Public gblnOnlyUnitPatient As Boolean '门诊记帐时,输入姓名时只查合约单位病人
Public gintNameDays As Integer '通过姓名模糊查找天数
Public gblnSeekBill As Boolean '是否自动搜寻划价单据
Public gintSeekDays As Integer '自动搜录单据的天数
Public gblnCheckRegeventDept As Boolean '检查病人挂号科室
Public gbytUnRegevent As Byte      '未挂号病人收费,0-允许,1-提醒,2-禁止
Public gbln允许录入特殊使用的抗生素 As Boolean '92727

'LED语音报价控制
Public gblnLED As Boolean '是否使用Led显示
Public gbln手工报价 As Boolean '使用Led后,是否手工报价
Public gblnLedDispDetail As Boolean '使用Led后,每增加一行单据是否在设备上显示收费明细
Public gblnLedWelcome As Boolean    '使用Led后,在收费时,输入新病人或导入划价单时,是否显示欢迎信息并发声

'其它控制
Public gint病人来源 As Integer '收费，划价时的病人来源(1-门诊,2-住院)
Public gbln病人来源受权限控制 As Boolean '是否允许更改病人来源

Public gbln药房单位 As Boolean '划价,记帐,收费时是否按照门诊单位进行显示；划价,收费也可能按住院单位
Public gstr药房单位 As String '根据病人来源决定如"门诊单位"或"住院单位"
Public gstr药房包装 As String '根据病人来源决定如"门诊包装"或"住院包装"

Public gblnCheckTest As Boolean '检查皮试结果
Public gbln累计 As Boolean '收费是否显示累计
Public gbln护士 As Boolean '收费划价是否显示护士
Public gint分类合计 As Integer '0-按收据费目,1-按收入项目
Public gblnMulti As Boolean '是否支持多单据收费
Public gblnShowErr As Boolean '查看收费单据时是否显示误差费用
Public gintDelPrice As Integer '进入划价管理时,删除n天内的划价单
Public gbln记帐打印 As Boolean
Public gbln划价打印 As Boolean
Public gbln审核打印 As Boolean

Public Type TY_Reg_Para  '挂号相关参数
    bytNODaysGeneral As Byte    '普通挂号有效天数
    bytNoDayseMergency As Byte '急诊挂号有效天数
End Type
'------------------------------------------------------------------------------------------
'卡支付相关
Public Type gTY_PayMoney
    lng医疗卡类别ID As Long
    bln消费卡 As Boolean
    str结算方式 As String
    str名称 As String
    str刷卡卡号 As String
    str刷卡密码 As String
    lng消费卡ID As Long
    str限制类别 As String   '上次刷卡限制类别
    dbl已刷金额 As Double '已经刷的消费卡金额
    str交易流水号 As String
    str交易说明 As String
    bln读卡 As Boolean
    bln卡号密文  As Boolean
    int医疗卡长度 As Integer
    bln支票 As Boolean
    bln自制卡 As Boolean
    blnOneCard As Boolean '是否一卡通结算
    int性质 As Integer '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算;<0 表示第三方支付
    strNo As String
    lngID As Long '预交ID
    lng结帐ID As Long
    dbl帐户余额 As Double
End Type
Public gtyPrePatiPay As gTY_PayMoney '上次病人的支付方式
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--定义系统参数
'问题:27990
Private Type Ty_System_Para
     byt药品名称显示 As Byte   '药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
     byt输入药品显示 As Byte  '输入药品显示（通过输入简码方式进入选择器时药品名称的显示）：0-按输入匹配显示，1-固定显示通用名和商品名
     byt条码卫材识别控制 As Byte   '是否仅条码识别::1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
     Sy_Reg  As TY_Reg_Para     '挂号相关:34717
End Type
Public gTy_System_Para As Ty_System_Para

'-------------------------------------------------------------------------------------------------------------------------------------------------
'刘兴洪:模块参数定义
Private Type Ty_Module_Para
    byt缴款控制 As Byte    '缴款控制:0-代表不进行缴款输入和累计控制,1-代表输入缴款后才结束病人累计(改变病人除外)，2-收费时必须要输入缴款金额
    int提醒剩余票据张数 As Integer      '收费时,票据在剩余X张后开始提醒收费员:-1代表不提醒
    byt票据分配规则 As Byte   '25187:票据分配规则:0-根据实际打印分配票号;1-根据系统预定规则分配;2-根据用户自定义规则分配
    byt票据汇总条件 As Byte            'byt票据分配规则>=1时:按票据分配规则汇总的条件:25187
    bln票据分单据 As Boolean    '分单据 :25187
    int执行科室 As Integer        'N个执行科室分页:25187
    int收据费目 As Integer        'N个收据费目分页:25187
    int收费细目 As Integer        'N个收费细目分页:25187
    bln分别打印 As Boolean      '多张单据收费分别打印:25187,合并
    byt票据生成方式 As Byte     '0-按费目打印,1-按细目打印,10-先按执行科室分别打,再按费目打印,11-先按执行科室分别打,再按细目打印:25187,合并
    byt门诊收据行次 As Byte     '收费收据总行次:25187,合并
    bln一张票据 As Boolean '收费一次只用一张票据:25187,合并
    bln工本费 As Boolean '是否收取工本费:25187,合并
    bln误差占用票据 As Boolean  '25187,合并
    bln使用加减切换  As Boolean '47457
    byt药品摆药退费方式 As Byte '47400
    byt退费缺省选择方式 As Byte '87489
    byt刷卡缺省金额操作 As Byte '86853
    bln只对医保结算成功单据收费 As Boolean '91665
    str本机收费执行科室 As String '96357，格式：科室ID1,科室ID2,科室ID3,...
    str已设置收费执行科室 As String '已设置了的本机收费执行科室，格式：科室ID1,科室ID2,科室ID3,...
    bln医保结算光标缺省定位 As Boolean
    bln现金退款缺省方式 As Boolean
    bln体检分别打印 As Boolean '104983
    str缺省退现 As String '112753
End Type

Public gTy_Module_Para As Ty_Module_Para
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public gstrMatchMethod As String
Private mlng部门编码平均长度 As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public grsTotal As ADODB.Recordset  '统计汇总的累计金额
Public grs收费类别 As ADODB.Recordset
Public gobjPlugIn As Object
Public gobjPublicDrug As Object '药品公共部件,105872
Public gblnUserIsClinic As Boolean

Public Function zlGet收费类别() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '返回:返回收费类别集
    '编制:刘兴洪
    '日期:2013-02-21 17:08:51
    '问题:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    On Error GoTo errHandle
    gstrSQL = "Select  编码,名称 From 收费项目类别"
    If grs收费类别 Is Nothing Then
        Set grs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    ElseIf grs收费类别.State <> 1 Then
        Set grs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    End If
    Set zlGet收费类别 = grs收费类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetDefaultWindow(ByVal str类别 As String, ByVal lng药房ID As Long) As String
'功能:获取缺省的药房窗口设置
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str类别
        Case "5"
            If InStr(gstr西窗, ":") > 0 Then '旧数据没有存药房ID
                 strTmp = gstr西窗
            ElseIf glng西药房 > 0 And gstr西窗 <> "" Then
                strTmp = glng西药房 & ":" & gstr西窗
            End If
        Case "6"
            If InStr(gstr中窗, ":") > 0 Then
                 strTmp = gstr中窗
            ElseIf glng中药房 > 0 And gstr中窗 <> "" Then
                 strTmp = glng中药房 & ":" & gstr中窗
            End If
        Case "7"
            If InStr(gstr中窗, ":") > 0 Then
                 strTmp = gstr成窗
            ElseIf glng成药房 > 0 And gstr成窗 <> "" Then
                 strTmp = glng成药房 & ":" & gstr成窗
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str类别
                Case "5"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lng药房ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function


Public Function Get发药窗口(ByVal Curdate As Date, ByVal lng药房ID As Long, ByVal str类别 As String, _
    str西窗 As String, str成窗 As String, str中窗 As String) As String
'功能：获取药品对应的发药窗口
'参数：lng药房ID=执行部门ID,curDate=当前时间
'说明：在同一材质类药房的发药窗口内平均分配
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '指定时固定分配(指定是指没有对应药房上班时指定)
    Select Case str类别
        Case "5"
            If str西窗 <> "" Then
                Get发药窗口 = str西窗
            ElseIf glng西药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str西窗 = Get发药窗口
            End If
        Case "6"
            If str成窗 <> "" Then
                Get发药窗口 = str成窗
            ElseIf glng成药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str成窗 = Get发药窗口
            End If
        Case "7"
            If str中窗 <> "" Then
                Get发药窗口 = str中窗
            ElseIf glng中药房 > 0 Then
                Get发药窗口 = GetDefaultWindow(str类别, lng药房ID)
                str中窗 = Get发药窗口
            End If
    End Select
    
    
    If Get发药窗口 <> "" Then
        strSQL = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药房ID, Get发药窗口)
        If rsTmp.EOF Then Get发药窗口 = ""
        Exit Function
    End If
    
    '动态分配上班的非专家窗口,98876
    strSQL = "Select Zl_Get发药窗口([1],[2],[3]) As 窗口 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取发药窗口", lng药房ID, gbytAssign, Curdate)
    If Not rsTmp.EOF Then
        Get发药窗口 = Nvl(rsTmp!窗口)
    End If
    
    If Get发药窗口 <> "" Then
        Select Case str类别
            Case "5"
                str西窗 = Get发药窗口
            Case "6"
                str成窗 = Get发药窗口
            Case "7"
                str中窗 = Get发药窗口
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Public Function Get工本费() As Detail
'功能：获取工本费的收费细目ID
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.*,C.名称 as 类别名称 " & _
        " From 收费项目目录 A,收费特定项目 B,收费项目类别 C " & _
        " Where B.收费细目ID=A.ID And C.编码=A.类别 " & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.特定项目='工本费'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    If Not rsTmp.EOF Then
        Set Get工本费 = New Detail
        With Get工本费
            .ID = rsTmp!ID
            .类别 = rsTmp!类别
            .类别名称 = rsTmp!类别名称
            .编码 = rsTmp!编码
            .名称 = rsTmp!名称
            .规格 = Nvl(rsTmp!规格)
            .计算单位 = Nvl(rsTmp!计算单位)
            .变价 = False '不可变价
            .加班加价 = False '不加班加价
            .屏蔽费别 = True '与费别无关
            .说明 = Nvl(rsTmp!说明)
            .执行科室 = Nvl(rsTmp!执行科室, 3) '缺省为操作员所在科室
        End With
    Else
        Set Get工本费 = Nothing
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set Get工本费 = Nothing
End Function

Public Function isSimple(strNo As String) As Boolean
'功能：判断单据是否为简单收费产生的
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct 发药窗口,收费类别,Nvl(数次,1) as 数次" & _
        " From 门诊费用记录 Where 记录状态 IN(1,3)" & _
        " And 记录性质=1 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!数次 = 1 And rsTmp!收费类别 = "Z" And Nvl(rsTmp!发药窗口) = "Z" Then isSimple = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(strPrivs As String, str姓名 As String, str适用病人 As String, _
    rsWarn As ADODB.Recordset, cur余额 As Currency, cur当日额 As Currency, _
    cur单据金额 As Currency, cur担保 As Currency, str类别 As String, _
    ByVal str类别名 As String, ByRef str已报类别 As String, Optional bln多病人 As Boolean) As Integer
'功能:对病人记帐进行报警提示
'参数:
'     str姓名=病人姓名,用于提示
'     str适用病人=根据病人身份返回的记帐报警适用方案
'     rsWarn=当前病区记帐报警设置记录
'     cur余额=病人余额,用于累计报警
'     cur当日额=病人当日发生的费用额,用于每日报警
'     cur单据金额=病人单据中输入的费用
'     cur担保=病人担保费用额,用于累计报警
'     str类别=当前要检查的类别,用于分类报警
'     str类别名=类别名称,用于提示
'返回:0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
'     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    Dim i As Integer, byt标志 As Byte
    Dim bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '报警参数检查
    rsWarn.Filter = "病区ID=0 And 适用病人='" & str适用病人 & "'"
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    If bln多病人 Then
        '示例：",周:-,张:DEF,李:567,张567"
        '报警标志2示例：",周:-①,张:DEF①,李:567①,张567②"
        bln已报警 = str已报类别 & "," Like "*," & str姓名 & ":-*,*" _
            Or str已报类别 & "," Like "*," & str姓名 & ":*" & str类别 & "*,*"
    Else
        '示例："-" 或 ",ABC,567,DEF"
        '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
        bln已报警 = InStr(str已报类别, str类别) > 0 Or str已报类别 Like "-*"
    End If
    
    If bln已报警 Then
        If byt标志 = 2 Then
            If bln多病人 Then
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str姓名 & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str姓名 & ":*" & str类别 & "*,*" Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For  '说明见住院模块
                    End If
                Next
            Else
                If str已报类别 Like "-*" Then
                    byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
                Else
                    arrTmp = Split(str已报类别, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str类别) > 0 Then
                            byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                            'Exit For '说明见住院模块
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名 <> "" Then str类别名 = """" & str类别名 & """费用"
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        '先只有两种:1.强制记帐,无权限时,禁止记帐
                        Call MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐", vbInformation + vbOKOnly, gstrSysName)
                        BillingWarn = 3
                        '--26349
'                        If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                            BillingWarn = 2
'                        Else
'                            BillingWarn = 1
'                        End If

                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur余额 + cur担保 - cur单据金额 < 0 Then
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                            BillingWarn = 3
                        Else
                            MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    ElseIf cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur余额 + cur担保 - cur单据金额 < 0 Then
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                                MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                BillingWarn = 3
                            Else
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                                BillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        Call MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐.", vbOKOnly + vbInformation, gstrSysName)
                        BillingWarn = 3
'                        If MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                            BillingWarn = 2
'                        Else
'                            BillingWarn = 1
'                        End If
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        BillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStock(ByVal lng药品ID As Long, ByVal lng药房ID As Long, Optional ByVal lng批次 As Long = -1) As Double
'功能：获取指定药房指定药品库存(以零售单位)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
    If lng批次 = -1 Or lng批次 = 0 Then
        strSQL = _
            " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
            " Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.性质=1 And A.药品ID=[1] And A.库房ID=[2]"
    Else
        strSQL = _
            " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
            " Where Nvl(A.批次,0)= [3] And (A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
            " And A.性质=1 And A.药品ID=[1] And (A.库房ID = [2] Or A.库房ID In (Select 虚拟库房id From 虚拟库房对照 Where 科室id = [2] And Rownum < 2))  "
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng药品ID, lng药房ID, lng批次)
    If Not rsTmp.EOF Then GetStock = rsTmp!库存
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMultiStock(ByVal lng药品ID As Long, ByVal str药房IDs As String) As Double
'功能：获取指定药房指定药品库存(以零售单位)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
        " Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.性质=1 And A.药品ID=[1] And instr([2],','|| A.库房ID ||',')>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药品ID, "," & str药房IDs & ",")
    If Not rsTmp.EOF Then GetMultiStock = rsTmp!库存
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPlace(ByVal lng药品ID As Long, ByVal lng药房ID As Long, _
    Optional bln卫材 As Boolean = False) As String
'功能：获取指定药品在指定药房的库房货位
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    If bln卫材 Then
        strSQL = "Select 库房货位 From 材料储备限额 Where 材料ID=[1] And 库房ID=[2]"
    Else
        strSQL = "Select 库房货位 From 药品储备限额 Where 药品ID=[1] And 库房ID=[2]"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药品ID, lng药房ID)
    If Not rsTmp.EOF Then GetPlace = Nvl(rsTmp!库房货位)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub GetDoctor(lng科室ID As Long, ByRef rsTmp As ADODB.Recordset, _
    Optional ByVal bln仅操作员部门 As Boolean, Optional ByVal str部门性质 As String = "临床")
    '功能：获取指定开单科室的医生或护士,如果未指定开单科室，则获取所有医生或护士
    '入参：
    '     bln仅操作员部门-操作员的所属部门下的人员
    Dim strSQL As String, bln护士 As Boolean
    Dim strWhere As String
    
    On Error GoTo errH
    If lng科室ID = 0 And bln仅操作员部门 Then
        strWhere = " And Exists (Select 1 From 部门人员 M, 部门人员 N" & _
                    " Where m.部门id = n.部门id And m.人员id = a.Id And n.人员id = [1])"
    End If
    
    str部门性质 = Replace(str部门性质, "'", "")
    If str部门性质 <> "" Then
        If InStr(1, str部门性质, ",") > 0 Then
            strWhere = strWhere & " And Instr(','||[2]||',',','||d.工作性质||',')>0"
        Else
            strWhere = strWhere & " And d.工作性质 = [2]"
        End If
    End If
    
    '允许部门工作性质为非临床的医生或护士,因为可能该人所属的一个部门是非末级部门.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务,A.专业技术职务,B.缺省" & _
            " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
            " Where A.ID = B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID And C.人员性质 IN('医生','护士') " & _
            " And D.服务对象 IN(" & gint病人来源 & ",3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & strWhere & vbNewLine & _
            " Order by " & IIf(gbyt开单人显示 = 1, "简码", "编号") & ",缺省 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.ID, str部门性质)
    End If
   
    '开单人允许有护士，并且收费项目允许治疗及材料时才读取护士
    bln护士 = gbln护士 And (gstr收费类别 = "" Or gstr收费类别 Like "*'E'*" Or gstr收费类别 Like "*'M'*" Or gstr收费类别 Like "*'4'*")
    If lng科室ID = 0 Then
        rsTmp.Filter = IIf(bln护士, "", "人员性质='医生'")
    Else
        rsTmp.Filter = "部门ID=" & lng科室ID & IIf(bln护士, "", " And 人员性质='医生'")
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub GetDoctorDept(ByRef rsTmp As ADODB.Recordset, _
    Optional ByVal bln仅操作员部门 As Boolean, _
    Optional ByVal str性质 As String = "'临床','产科'", _
    Optional ByVal lngDeptID As Long)
    '功能：获取所有开单科室
    '入参：
    '     bln仅操作员部门-操作员的所属部门
    '     str性质='临床','护理','中药房',...,允许为空
    '      lngDeptID-当前病区ID或科室ID
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
    If bln仅操作员部门 Then
        strWhere = " And Exists (Select 1 From 部门人员 Where 部门id = a.Id And 人员id = [1])"
    End If
    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strWhere = strWhere & " And Instr(','||[2]||',',','||B.工作性质||',')>0"
        Else
            strWhere = strWhere & " And B.工作性质 = [2]"
        End If
    End If
    If lngDeptID <> 0 Then strWhere = strWhere & " And A.ID=[3]"
    
    strSQL = _
        "Select A.ID, A.编码, A.名称, A.简码, 0 As 缺省, B.工作性质, Decode(D.优先级, 1, (Decode(C.优先级, 1, 1, 2)), 3) 优先级" & vbNewLine & _
        "From 部门表 A, 部门性质说明 B," & vbNewLine & _
        "     (Select 部门id, Max(Decode(Instr('检查,检验,手术,治疗,营养,体检', 工作性质), 0, 1, 2)) As 优先级" & vbNewLine & _
        "       From 部门性质说明 Where 服务对象 <> 0" & vbNewLine & _
        "       Group By 部门id) C, (Select 部门id, Max(Decode(服务对象, 1, 1, 2)) As 优先级 From 部门性质说明 Where 服务对象 <> 0 Group By 部门id) D" & vbNewLine & _
        "Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And A.ID = B.部门id And" & vbNewLine & _
        "      B.部门id = C.部门id And B.部门id = D.部门id And B.服务对象 In (" & gint病人来源 & ", 3)" & vbNewLine & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & strWhere & vbNewLine & _
        "Order By 优先级,编码"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.ID, str性质, lngDeptID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

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

Public Function ExistWindow(ByVal lng药房ID As Long, ByRef rs发药窗口 As ADODB.Recordset) As Boolean
'功能：确定医院指定药房是否使用发药窗口
'说明：因为专家窗仅指定,不动态分配,所以排开
    Dim strSQL As String
    
    On Error GoTo errH
    
    If rs发药窗口 Is Nothing Then
        strSQL = "Select 名称,药房ID From 发药窗口 Where Nvl(专家,0)=0"
        Set rs发药窗口 = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rs发药窗口, strSQL, "mdlOutExse")
    End If
    
    If lng药房ID = 0 Then
        rs发药窗口.Filter = ""
    Else
        rs发药窗口.Filter = "药房ID=" & lng药房ID
    End If
    If Not rs发药窗口.EOF Then ExistWindow = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 编码, 名称, 简码 From 收费项目类别"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "获取收费类别")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(intType As Integer, Optional lngBalance As Long, Optional strNos As String, Optional lng打印ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在二次结算
    '返回:True-存在二次结算数据 False-不存在二次结算数据
    '入参:intType:0-收费数据，使用lngBalance为结算序号
    '     intType:1-收费数据，使用strNos为单据号
    '     intType:2-根据打印ID来判断是否使用补结算
    '     lng打印ID-打印ID>0时，从临时表中取数
    '编制:刘尔旋
    '日期:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, cllPro As Collection
    Dim strValue(0 To 0) As String
    Dim varData() As Variant
    
    On Error GoTo errHandle
        
    If lng打印ID > 0 And intType = 2 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From 费用补充记录 A, (Select Distinct 结帐id From 门诊费用记录 B,临时票据打印内容 C Where B.NO=C.NO and mod(B.记录性质,10)=1 and C.性质=1 and C.ID=[2]) B" & vbNewLine & _
        " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", lngBalance, lng打印ID)
    ElseIf intType = 0 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From 费用补充记录 A, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
        " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
        strSQL = strSQL & " Union " & _
        " Select 1 From 费用补充记录 Where 结算序号 = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", lngBalance, lng打印ID)
    Else
        If Len(strNos) <= 4000 Then
            strSQL = "" & _
            " Select /*+ Rule */ 1" & vbNewLine & _
            " From 费用补充记录 A," & vbNewLine & _
            "      (Select Distinct 结帐id" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
            " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", strNos)
        Else
            If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
            If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
            strSQL = "With 单据信息 as (" & strSQL & ")" & vbCrLf
            strSQL = strSQL & _
            "      Select Distinct A1.结帐id" & vbNewLine & _
            "       From 门诊费用记录 A1,单据信息 A2" & vbNewLine & _
            "       Where Mod(A1.记录性质, 10) = 1 And A1.NO=A2.NO " & vbNewLine
            strSQL = "" & _
            "   Select 1" & vbNewLine & _
            "   From 费用补充记录 A," & vbNewLine & _
            "        (" & strSQL & ") B " & vbNewLine & _
            "   Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecordByArray(strSQL, "检查二次结算", varData)
        End If
    End If
    
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function InitSysPar() As Boolean
    '功能：初始化系统参数
    '返回：真-处理成功
    Dim strValue As String
    On Error Resume Next
    
    '费用金额小数点位数
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    '问题:35242
    gbln简码切换 = IIf(Val(zlDatabase.GetPara("简码匹配方式切换", , , 1)) = 1, 1, 0) = 1
    
    '卡号显示方式
    'gblnShowCard = zlDatabase.GetPara(12, glngSys) = "0"
    
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '收费分币处理方式
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    gBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 2, 1)))

    '病人输入方式
    strValue = zlDatabase.GetPara(17, glngSys, , "1111")
    gblnInputName = (Mid(strValue, 1, 1) = "1")
    gblnInputCard = (Mid(strValue, 2, 1) = "1")
    gblnInputNO = (Mid(strValue, 3, 1) = "1")
    gblnInputID = (Mid(strValue, 4, 1) = "1")
    
    '指定药房时限制库存
    gblnStock = zlDatabase.GetPara(18, glngSys) = "1"
    
    '窗口分配方式
    gbytAssign = Val(zlDatabase.GetPara(19, glngSys, , 0))
        
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(0))
    'gbytCardNOLen = Val(Split(strValue, "|")(4))
    
    '挂号有效天数
    '刘兴洪:34717
    '两位:前一位普能挂号;后一位急诊挂号
    strValue = zlDatabase.GetPara(21, glngSys, , "01") & "1"
    gTy_System_Para.Sy_Reg.bytNODaysGeneral = Val(Left(strValue, 1))
    gTy_System_Para.Sy_Reg.bytNoDayseMergency = Val(Mid(strValue, 2, 1))
    'If gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 Then gTy_System_Para.Sy_Reg.bytNODaysGeneral = 1
    ' If gTy_System_Para.Sy_Reg.bytNoDayseMergency = 0 Then gTy_System_Para.Sy_Reg.bytNoDayseMergency = 1
    
    '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys, , 0))
    
    '票号严格控制
    gblnStrictCtrl = Mid(zlDatabase.GetPara(24, glngSys, , "00000"), 1, 1) = "1"
        
    '一卡通消费刷卡控制
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdbl预存款消费验卡 = Val(Split(strValue, "|")(0))
    gbyt预存款退费验卡 = Val(Split(strValue, "|")(1))
    gbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"
    
    '医保费用类型
    gstr医保费用类型 = "'" & Replace(zlDatabase.GetPara(41, glngSys), "|", "','") & "'"

    '公费费用类型
    gstr公费费用类型 = "'" & Replace(zlDatabase.GetPara(42, glngSys), "|", "','") & "'"
    
    '收费项目输入简码匹配方式
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    
    '自动发药退药
    gbln收费后自动发药 = zlDatabase.GetPara(45, glngSys) = "1"
            
    '刷卡要求输入密码
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
        
      
    '保险对码检查
    gbyt医保对码检查 = Val(zlDatabase.GetPara(59, glngSys, , 1))
    
    
    
    
    '单笔费用最大提醒金额
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    
    '是否要求首先输入类别
    gbln收费类别 = zlDatabase.GetPara(72, glngSys) = "1"
    
    '西成药是否按商品名显示
    'gbln商品名 = zlDatabase.GetPara(74, glngSys) = "1"
        
    
    '多张单据使用多种结算方式模式
    gblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
        
      
    '自动发料
    gbln门诊自动发料 = zlDatabase.GetPara(92, glngSys) = "1"
    
    '从属项目汇总计算折扣
    gbln从项汇总折扣 = zlDatabase.GetPara(93, glngSys) = "1"
    

    '记帐报警包含划价费用
    gbln报警包含划价费用 = zlDatabase.GetPara(98, glngSys) = "1"
        
    '当不输类别时,输入费用项目时,首位当作类别简码
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1" And Not gbln收费类别
    
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    gbln退费申请模式 = IIf(zlDatabase.GetPara(151, glngSys) = "1", True, False)
    gbln不显示无库存卫材 = zlDatabase.GetPara(316, glngSys) = "1"
    'e.个人全局参数
    '-------------------------------------------------------------------------------------------------
    '问题:27990
    With gTy_System_Para
        .byt输入药品显示 = Val(zlDatabase.GetPara("输入药品显示")) '0-按输入匹配显示，1-固定显示通用名和商品名
        .byt药品名称显示 = Val(zlDatabase.GetPara("药品名称显示"))  '：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        .byt条码卫材识别控制 = Val(zlDatabase.GetPara(320, glngSys, , "0"))      '1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
    End With
    
    '当前操作员是否为临床部门人员
    gblnUserIsClinic = UserIsClinic(UserInfo.ID)
    InitSysPar = True
End Function

Public Sub InitLocPar(lngModul As Long)
'功能：初始化费用模块参数
    Dim strValue As String, intType As Integer
    Dim arrTmp As Variant
    
    On Error Resume Next
    
    'a.本机注册表存储的模块参数
    '----------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
    End If
    'b.数据库存储的公共全局参数
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
     'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gstr结算方式 = zlDatabase.GetPara("缺省结算方式", glngSys, lngModul)
        gstr费别 = zlDatabase.GetPara("缺省费别", glngSys, lngModul)
        
        If glngSys Like "8??" Then
            gstr收费类别 = "'5','6','7'"
        Else
            gstr收费类别 = zlDatabase.GetPara("收费类别", glngSys, lngModul)
        End If
        
        gbln手工报价 = zlDatabase.GetPara("手工报价", glngSys, lngModul) = "1"
        gblnLedDispDetail = zlDatabase.GetPara("LED显示收费明细", glngSys, lngModul, "1") = "1"
        gblnLedWelcome = zlDatabase.GetPara("LED显示欢迎信息", glngSys, lngModul, "1") = "1"
        
        gblnSharedInvoice = zlDatabase.GetPara("挂号共用收费票据", glngSys, lngModul) = "1"
        gblnPay = zlDatabase.GetPara("中药付数", glngSys, lngModul) = "1"
        gblnTime = zlDatabase.GetPara("变价数次", glngSys, lngModul) = "1"
        gbln护士 = zlDatabase.GetPara("显示护士", glngSys, lngModul) = "1"
    
        gbln药房单位 = zlDatabase.GetPara("药品单位", glngSys, lngModul) = "1"
    
        glng中药房 = zlDatabase.GetPara("缺省中药房", glngSys, lngModul)
        glng西药房 = zlDatabase.GetPara("缺省西药房", glngSys, lngModul)
        glng成药房 = zlDatabase.GetPara("缺省成药房", glngSys, lngModul)
        glng发料部门 = zlDatabase.GetPara("缺省发料部门", glngSys, lngModul)
        gbln其它药房 = zlDatabase.GetPara("显示其它药房库存", glngSys, lngModul) = "1"
        gbln其它药库 = zlDatabase.GetPara("显示其它药库库存", glngSys, lngModul) = "1"
        If lngModul = 1121 Then
            gbln划价立即缴款 = zlDatabase.GetPara("提取划价后立即缴款", glngSys, lngModul) = "1"
        End If

        
        '31936:其他药房或药库(非操作员所属科室)的库存显示方式:0-直接显示库存;1-显示有无
        '         此参数需要结合:显示其它药房库存和显示其它药库库存的参数设置来决定,如果勾选了其中一项,该参数才起作用
        If lngModul = 1121 Then
            gbyt库存显示方式 = 0: '门诊收费不变
        Else
            gbyt库存显示方式 = Val(zlDatabase.GetPara("库存显示方式", glngSys, lngModul))
        End If
        
        '分离发药时的检查
        gstr西药房 = zlDatabase.GetPara("西药房选择", glngSys, lngModul)
        gstr成药房 = zlDatabase.GetPara("成药房选择", glngSys, lngModul)
        gstr中药房 = zlDatabase.GetPara("中药房选择", glngSys, lngModul)
        
        gbyt科室医生 = Val(zlDatabase.GetPara("科室医生", glngSys, lngModul, 0))
        gbyt开单人显示 = Val(zlDatabase.GetPara("开单人显示方式", glngSys, lngModul, 1))
        gblnSeekName = zlDatabase.GetPara("姓名模糊查找", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("姓名查找天数", glngSys, lngModul))
        gbln允许录入特殊使用的抗生素 = Val(zlDatabase.GetPara("允许录入特殊使用的抗生素", glngSys, lngModul)) = 1
    End If
    
    gbln病人来源受权限控制 = False
    
    If lngModul = 1120 Or lngModul = 1121 Then
        gcurMax = Val(zlDatabase.GetPara("最大金额", glngSys, lngModul))
        gstr中窗 = zlDatabase.GetPara("中药房窗口", glngSys, lngModul)
        gstr西窗 = zlDatabase.GetPara("西药房窗口", glngSys, lngModul)
        gstr成窗 = zlDatabase.GetPara("成药房窗口", glngSys, lngModul)
        
        '光标经过项目
        gbln性别 = zlDatabase.GetPara("性别", glngSys, lngModul) = "1"
        gbln年龄 = zlDatabase.GetPara("年龄", glngSys, lngModul) = "1"
        gbln费别 = zlDatabase.GetPara("费别", glngSys, lngModul) = "1"
        gbln医疗付款 = zlDatabase.GetPara("医疗付款", glngSys, lngModul) = "1"
        gbln加班 = zlDatabase.GetPara("加班", glngSys, lngModul) = "1"
        gbln开单日期 = zlDatabase.GetPara("开单日期", glngSys, lngModul) = "1"
        gbln开单人 = zlDatabase.GetPara("开单人", glngSys, lngModul) = "1"
        
        gbln不缺省开单人 = zlDatabase.GetPara("不使用缺省开单人", glngSys, lngModul) = "1"
        gbln必须输开单人 = zlDatabase.GetPara("必须要输入开单人", glngSys, lngModul) = "1"
        gbln缺省科室优先 = zlDatabase.GetPara("缺省科室优先", glngSys, lngModul) = "1"
        gint分类合计 = Val(zlDatabase.GetPara("分类合计方式", glngSys, lngModul))
        If lngModul = 1120 Then
            gint划价通知单 = Val(zlDatabase.GetPara("划价通知单打印方式", glngSys, lngModul))
            gintDelPrice = Val(zlDatabase.GetPara("取消划价单", glngSys, lngModul))
        ElseIf lngModul = 1121 Then
            '刘兴洪:单独处理模块参数,其他参数,陆续在进行处理
            Dim strTmp As String
            With gTy_Module_Para
                'gbln缴款结束 = zlDatabase.GetPara("收费缴款输入控制", glngSys, lngModul) = "1"
                .byt缴款控制 = Val(zlDatabase.GetPara("收费缴款输入控制", glngSys, lngModul))   '问题:22343;51670
                strTmp = Trim(zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, lngModul, "0|10"))
                If Val(Split(strTmp & "|", "|")(0)) = 0 Then
                    .int提醒剩余票据张数 = -1
                Else
                    .int提醒剩余票据张数 = Val(Split(strTmp & "|", "|")(1))     '问题:26948
                End If
            End With
            '25187
            '启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
            strTmp = Trim(zlDatabase.GetPara("票据分配规则", glngSys, lngModul, "0||0;0;0;0;0"))
            arrTmp = Split(strTmp & "||", "||")
            With gTy_Module_Para
                .byt票据分配规则 = Val(arrTmp(0))
                arrTmp = Split(arrTmp(1) & ";;;;;", ";")
                .bln票据分单据 = Val(arrTmp(0)) = 1
                .int执行科室 = Val(arrTmp(1))
                .int收据费目 = Val(arrTmp(2))
                .int收费细目 = Val(arrTmp(3))
                .byt票据汇总条件 = Val(arrTmp(4))
                .bln分别打印 = zlDatabase.GetPara("多张单据收费分别打印", glngSys, lngModul) = "1"
                .byt票据生成方式 = Val(zlDatabase.GetPara("收费票据生成方式", glngSys, lngModul))
                .byt门诊收据行次 = Val(zlDatabase.GetPara("收费收据总行次", glngSys, lngModul, 3))
                .bln一张票据 = zlDatabase.GetPara("收费每次只用一张票据", glngSys, lngModul) = "1"
                .bln工本费 = Val(zlDatabase.GetPara("收据加收工本费", glngSys, lngModul, "0")) = 1
                .bln误差占用票据 = Val(zlDatabase.GetPara("误差项不使用票据", glngSys, lngModul, "0")) = "0"
                .bln使用加减切换 = Val(zlDatabase.GetPara("使用加减切换支付方式", glngSys, lngModul, "1")) = "1"
                .byt药品摆药退费方式 = Val(zlDatabase.GetPara("药品摆药退费方式", glngSys, lngModul, "0"))  '47400
                .byt退费缺省选择方式 = Val(zlDatabase.GetPara("退费缺省选择方式", glngSys, lngModul, "0")) '87489
                .byt刷卡缺省金额操作 = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0")) '86853
                .bln只对医保结算成功单据收费 = Val(zlDatabase.GetPara("只对医保结算成功单据收费", glngSys, lngModul, "0")) '91665
                .str本机收费执行科室 = zlDatabase.GetPara("本机收费执行科室", glngSys, lngModul)
                .str已设置收费执行科室 = zlGet已设置收费执行科室(lngModul)
                .bln医保结算光标缺省定位 = Val(zlDatabase.GetPara("医保结算光标缺省定位", glngSys, lngModul, "0")) = "1"
                .bln现金退款缺省方式 = Val(zlDatabase.GetPara("现金退款缺省方式", glngSys, lngModul, "0")) = "1"
                .bln体检分别打印 = zlDatabase.GetPara("体检病人分单据打印", glngSys, lngModul) = "1"
                .str缺省退现 = zlDatabase.GetPara("非三方卡退费缺省方式", glngSys, lngModul, "")
            End With
            
            gbln累计 = zlDatabase.GetPara("显示累计", glngSys, lngModul) = "1"
            gblnCheckTest = zlDatabase.GetPara("检查皮试结果", glngSys, lngModul) = "1"
            gblnPrePayPriority = zlDatabase.GetPara("优先使用预交款", glngSys, lngModul) = "1"
            
            
            
            glngAddedItem = Val(zlDatabase.GetPara("自动加收挂号费", glngSys, lngModul))
            gblnShowErr = zlDatabase.GetPara("显示误差费用", glngSys, lngModul) = "1"
            gblnMulti = zlDatabase.GetPara("多单据收费", glngSys, lngModul) = "1"
            
            gblnSeekBill = zlDatabase.GetPara("搜寻划价单据", glngSys, lngModul) = "1"
            gintSeekDays = Val(zlDatabase.GetPara("搜寻单据天数", glngSys, lngModul))
            gblnUnPopPriceBill = zlDatabase.GetPara("不弹出划价单选择", glngSys, lngModul) = "1"
        
            gblnCheckRegeventDept = zlDatabase.GetPara("检查病人挂号科室", glngSys, lngModul) = "1"
            gbytUnRegevent = Val(zlDatabase.GetPara("未挂号病人收费", glngSys, lngModul))
            gbytAutoSplitBill = Val(zlDatabase.GetPara("自动组合单据", glngSys, lngModul))
            gint收费清单 = Val(zlDatabase.GetPara("收费清单打印方式", glngSys, lngModul))
            glngFactNormal = Val(zlDatabase.GetPara("普通发票格式", glngSys, lngModul))
            glngFactMediCare = Val(zlDatabase.GetPara("医保发票格式", glngSys, lngModul))
        End If
    ElseIf lngModul = 1122 Then
        gint分类合计 = 0        '刘兴洪:由于门诊记帐没有该参数设置,因此,只能按收据费目来统计
        gblnOnlyUnitPatient = zlDatabase.GetPara("只查找合约单位病人", glngSys, lngModul) = "1"
        gbln记帐打印 = zlDatabase.GetPara("记帐打印", glngSys, lngModul) = "1"
        gbln划价打印 = zlDatabase.GetPara("划价打印", glngSys, lngModul) = "1"
        gbln审核打印 = zlDatabase.GetPara("审核打印", glngSys, lngModul) = "1"
    End If
    
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gint病人来源 = Val(zlDatabase.GetPara("病人来源", glngSys, lngModul, , , , intType))
        '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        gbln病人来源受权限控制 = IIf(intType = 1 Or intType = 3 Or intType = 15, True, False)
        
        If gint病人来源 = 1 Then
            gstr药房单位 = "门诊单位": gstr药房包装 = "门诊包装"
        Else
            gstr药房单位 = "住院单位": gstr药房包装 = "住院包装"
        End If
    End If
    
    'd.数据环境参数
    '-------------------------------------------------------------------------------------------------
    If lngModul = 1120 Or lngModul = 1121 Or lngModul = 1122 Then
        gbln药房上班安排 = Check药房上班安排
    End If
    
    If lngModul = 1121 Then
        Call GetErrorItem(glng误差细目ID, gstr误差收据费目)
        gstr误差费名称 = zlGet误差费名称
    End If
    
    If lngModul = 1124 Then
        gstr误差费名称 = zlGet误差费名称
        gbln药房单位 = zlDatabase.GetPara("药品单位显示", glngSys, lngModul) = "1"
        gstr药房单位 = "门诊单位": gstr药房包装 = "门诊包装"
        gblnSeekName = Split(zlDatabase.GetPara("姓名模糊查找方式", glngSys, lngModul, "0|0"), "|")(0) = "1"
        gintNameDays = Val(Split(zlDatabase.GetPara("姓名模糊查找方式", glngSys, lngModul, "0|0"), "|")(1))
        With gTy_Module_Para
            .bln使用加减切换 = Val(zlDatabase.GetPara("使用加减切换支付方式", glngSys, lngModul, "1")) = "1"
            .byt药品摆药退费方式 = Val(zlDatabase.GetPara("药品摆药退费方式", glngSys, lngModul, "0"))  '47400
        End With
    End If
End Sub

Private Function zlGet已设置收费执行科室(ByVal lngModul As Long) As String
    '获取已设置的所有收费执行科室
    '返回：
    '   格式:科室ID1,科室ID2,科室ID3,...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select b.参数值" & vbNewLine & _
            " From zlParameters A, zlUserParas B" & vbNewLine & _
            " Where a.Id = b.参数id And Nvl(a.系统, 0) = [1] And Nvl(a.模块, 0) = [2]" & vbNewLine & _
            "       And a.参数名 = [3] And b.参数值 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取本机收费执行科室参数", glngSys, lngModul, "本机收费执行科室")
    If rsTemp Is Nothing Then Exit Function
    Do While Not rsTemp.EOF
        strTemp = strTemp & "," & Nvl(rsTemp!参数值)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlGet已设置收费执行科室 = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBalanceSet() As ADODB.Recordset
    '功能：返回一个结算记录集对象
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "结算方式", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "结算金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "结算性质", adBigInt, , adFldIsNullable '1-医保;2-消费卡;3-医疗卡;0-其他
    rsTmp.Fields.Append "卡类别ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "卡号", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "交易流水号", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "交易说明", adVarChar, 500, adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str开单人 As String, ByVal str开单科室 As String, _
    ByVal intMode As Integer, ByVal intPrice As Integer, Optional ByVal intPage As Integer, Optional ByVal lngRow As Long) As ADODB.Recordset
'功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
'字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室,执行科室ID,
'         单据性质（1-收费单,2-记帐单),是否划价(1-划价;0-正常的收费及记帐单)
'参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行
'         intMode:单据性质（1-收费单,2-记帐单)
'         intPrice:是否划价(1-划价;0-正常的收费及记帐单)
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset, rsPrice As ADODB.Recordset
    Dim intStartPage  As Integer, intPages As Integer
    
    On Error GoTo errHandle
    
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    '79420,李南春,2014/11/10:调整记录集字段大小
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    '131048,焦博,2018-9-5,“CheckChargeItem”接口中的rsDetail 记录集中增加字段:
    '                                  执行科室ID、单据性质（1-收费单,2-记帐单)、是否划价(1-划价,0-正常的收费及记帐单)
    rsTmp.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "单据性质", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "是否划价", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    intStartPage = IIf(intPage <= 0, 1, intPage)
    intPages = IIf(intPage <= 0, objBill.Pages.Count, intPage)
    For p = intStartPage To intPages
        If objBill.Pages(p).NO <> "" Then '提取划价单
            strSQL = "Select 病人ID, NULL as 主页ID, 收费类别, 收费细目id, Avg(数次 * Nvl(付数, 1)) 数量," & vbNewLine & _
                    "        Sum(标准单价) As 单价, Sum(实收金额) 实收金额,Max(执行部门id) as 执行部门id" & vbNewLine & _
                    " From 门诊费用记录" & vbNewLine & _
                    " Where NO = [1] And 记录性质 = 1 And 记录状态 = 0" & vbNewLine & _
                    " Group By 收费细目id, 病人id, 收费类别"
            Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, "读取划价单", objBill.Pages(p).NO)
            With rsPrice
                For i = 1 To .RecordCount
                    rsTmp.Filter = "收费细目ID=" & !收费细目ID
                    If rsTmp.RecordCount = 0 Then
                        rsTmp.AddNew
                        
                        rsTmp!病人ID = Nvl(!病人ID, objBill.病人ID)
                        rsTmp!主页ID = Nvl(!主页ID, objBill.主页ID)
                        rsTmp!收费类别 = !收费类别
                        rsTmp!收费细目ID = !收费细目ID
                        
                        rsTmp!数量 = !数量
                        rsTmp!单价 = !单价
                        rsTmp!实收金额 = !实收金额
                        
                        rsTmp!开单人 = str开单人
                        rsTmp!开单科室 = str开单科室
                        rsTmp!执行科室ID = !执行部门ID
                        rsTmp!单据性质 = intMode
                        rsTmp!是否划价 = intPrice
                    Else
                        rsTmp!数量 = rsTmp!数量 + !数量
                        rsTmp!单价 = (rsTmp!单价 + !单价) / 2
                        rsTmp!实收金额 = rsTmp!实收金额 + !实收金额
                    End If
                    
                    rsTmp.Update
                    .MoveNext
                Next
            End With
        Else
            If lngRow = 0 Then
                intB = 1
                intE = objBill.Pages(p).Details.Count
            Else
                intB = lngRow
                intE = lngRow
            End If
            
            For i = intB To intE
                dbl单价 = 0: cur实收 = 0
                With objBill.Pages(p).Details(i)
                    If lngRow = 0 Then
                        rsTmp.Filter = "收费细目ID=" & .收费细目ID
                        blnNew = rsTmp.RecordCount = 0
                    Else
                        blnNew = True
                    End If
                                    
                    If blnNew Then
                        rsTmp.AddNew
                        
                        rsTmp!病人ID = objBill.病人ID
                        rsTmp!主页ID = objBill.主页ID
                        
                        rsTmp!收费类别 = .收费类别
                        rsTmp!收费细目ID = .收费细目ID
                        rsTmp!执行科室ID = .执行部门ID
                        rsTmp!单据性质 = intMode
                        rsTmp!是否划价 = intPrice
                        
                        For j = 1 To .InComes.Count
                            dbl单价 = dbl单价 + .InComes(j).标准单价
                            cur实收 = cur实收 + .InComes(j).实收金额
                        Next
                        If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                            '从药房单位转换为售价单位
                            rsTmp!数量 = IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                            rsTmp!单价 = Format(dbl单价 / .Detail.药房包装, gstrFeePrecisionFmt)
                        Else
                            rsTmp!数量 = IIf(.付数 = 0, 1, .付数) * .数次
                            rsTmp!单价 = Format(dbl单价, gstrFeePrecisionFmt)
                        End If
                        rsTmp!实收金额 = Format(cur实收, gstrDec)
                        
                        rsTmp!开单人 = str开单人
                        rsTmp!开单科室 = str开单科室
                    Else
                        For j = 1 To .InComes.Count
                            dbl单价 = dbl单价 + .InComes(j).标准单价
                            cur实收 = cur实收 + .InComes(j).实收金额
                        Next
                        If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                            '从药房单位转换为售价单位
                            rsTmp!数量 = rsTmp!数量 + IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                            rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价 / .Detail.药房包装, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                        Else
                            rsTmp!数量 = rsTmp!数量 + IIf(.付数 = 0, 1, .付数) * .数次
                            rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                        End If
                        rsTmp!实收金额 = rsTmp!实收金额 + Format(cur实收, gstrDec)
                    End If
                    
                    rsTmp.Update
                End With
            Next
        End If
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetApply(ByVal strNos As String, ByVal bytFlag As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退费申请数据
    '入参:strNOs-单据号,多个用逗号分离
    '     bytFlag-记录性质
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-05 11:59:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    strSQL = "" & _
    "   Select NO,申请人,申请时间,申请原因,审核人,审核原因,Nvl(状态,0) As 状态" & _
    "   From 病人退费申请 " & _
    "   Where NO IN  (Select Column_value From Table(f_str2List([1]))) And 记录性质 = [2]"
    On Error GoTo errH
    Set GetApply = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos, bytFlag)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ImportBill(strNo As String, blnModi As Boolean, bytFlag As Byte, _
    Optional ByVal int险类 As Integer, Optional ByVal bln工本费 As Boolean = True, _
    Optional ByVal bln不读煎法 As Boolean, Optional ByVal str药品价格等级 As String, _
    Optional ByVal str卫材价格等级 As String, Optional ByVal str普通价格等级 As String) As ExpenseBill
'功能：读取费用单据到单据对象中(目前忽略从属项目,当作独立项目),用于修改或导入时用
'参数：
'      strNO=单据号
'      bytFlag=记录性质,'0-收费,1-划价,2-门诊记帐
'      int险类=是否导入(修改)医保收费单据(相关验证已通过)
'      bln不读煎法   简单记帐等修改单据时不用读煎法
'返回：存放单据信息的单据对象
'说明：因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
'      该过程仅用于修改读入，修改时要排开误差处理费用(仅门诊收费有)
'      不管是导入还是修改单据,都不应包含已停用收费细目
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim intCurNo As Integer, strInfo As String, strAdvance As String
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    Dim dblAllTime As Double, dblCurTime As Double
    Dim dbl加班加价率 As Double
    Dim dblStock As Double, str药房IDs As String, str停用项目序号 As String
    Dim colSerial As New Collection
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim str摘要 As String, strWherePriceGrade
     
    On Error GoTo errH
    '价格等级
    If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And d.价格等级 = [4])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And d.价格等级 = [5])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And d.价格等级 = [6])" & vbNewLine & _
            "            Or (d.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From 收费价目" & vbNewLine & _
            "                                Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And 价格等级 = [4])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And 价格等级 = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And 价格等级 = [6])))))"
    Else
        strWherePriceGrade = " And d.价格等级 Is Null"
    End If
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
    strSQL = _
        " Select X.药品ID,W.材料ID,W.跟踪在用," & _
        "       A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态,0 as  多病人单,A.婴儿费,A.费别,A.姓名,A.性别,A.年龄," & _
        "       A.付款方式 ,A.标识号,A.病人ID, 0 as 主页ID,0 as 病人病区ID,A.病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
        "       A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        "       A.标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(F.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
        "       B.屏蔽费别,B.说明,B.执行科室,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
        "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
        "       Decode(A.收费类别,'4',1,X." & gstr药房包装 & ") as 药房包装," & _
        "       Decode(A.收费类别,'4',B.计算单位,X." & gstr药房单位 & ") as 药房单位," & _
        "       Decode(A.收费类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,Nvl(A.医嘱序号,0) 医嘱序号,B.录入限量,A.结论,M1.名称 as 诊疗名称" & _
        " From 门诊费用记录 A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E," & _
        "          收费项目别名 F,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
        " Where Nvl(A.附加标志,0)<>9 And A.记录性质=[2] And Instr([3],','||A.记录状态||',')>0" & _
                IIf(Not bln工本费, " And Nvl(A.附加标志,0)<>8", "") & " And A.NO=[1]" & _
        "       And A.价格父号 Is Null And A.收费细目ID=B.ID And A.收费细目ID=D.收费细目ID" & _
        "       And A.收费类别=C.编码 And D.收入项目ID=E.ID And A.收费细目ID=W.材料ID(+) " & _
        "       And A.收费细目ID=X.药品ID(+) And X.药名ID=M1.ID(+)" & _
        "       And A.收费细目ID=F.收费细目ID(+) And F.码类(+)=1 And F.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.门诊标志 " & IIf(gint病人来源 = 1, " IN(1,4)", "= 2") & _
        "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
            strWherePriceGrade
        
    strSQL = "Select * From (" & strSQL & ") Order by 序号"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, IIf(bytFlag = 0, 1, bytFlag), _
        IIf(blnModi, ",0,1,", ",0,1,3,"), str药品价格等级, str卫材价格等级, str普通价格等级)
    
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Pages(1).Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
NextRecord: Do While Not .EOF
                '检查收费项目是否停用或服务于门诊病人
                '主项停用时,不导从项
                If Not blnModi And InStr(",5,6,7,", !收费类别) = 0 Then
                    If InStr(1, str停用项目序号 & ",", "," & !从属父号 & ",") > 0 Then
                        .MoveNext
                        GoTo NextRecord
                    Else
                        If Not CheckFeeItemAvailable(!收费细目ID, 1) Then
                            str停用项目序号 = str停用项目序号 & "," & !序号
                            MsgBox "单据[" & strNo & "]中第" & !序号 & "行收费项目:" & !名称 & "" & vbCrLf & _
                                "已停用或不再服务于病人,将不会被导入." & IIf(IsNull(!从属父号), "如果有从属项目,也不会被导入.", ""), vbInformation, gstrSysName
                            .MoveNext
                            GoTo NextRecord
                        End If
                    End If
                End If
            
                '处理单据主体=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.Pages(1).NO = "" '要清空以便修改时表明是直接输入的费用
                    objBill.Pages(1).开单部门ID = IIf(IsNull(!开单部门ID), 0, !开单部门ID)
                    objBill.Pages(1).开单人 = IIf(IsNull(!开单人), "", !开单人)
                    objBill.Pages(1).医嘱序号 = !医嘱序号
                    
                    objBill.病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
                    objBill.主页ID = IIf(IsNull(!主页ID), 0, !主页ID)
                    objBill.病区ID = IIf(IsNull(!病人病区ID), 0, !病人病区ID)
                    objBill.科室ID = IIf(IsNull(!病人科室id), 0, !病人科室id)
                    objBill.姓名 = IIf(IsNull(!姓名), "", !姓名)
                    objBill.性别 = IIf(IsNull(!性别), "", !性别)
                    objBill.年龄 = IIf(IsNull(!年龄), "", !年龄)
                    objBill.标识号 = IIf(IsNull(!标识号), 0, !标识号)
                    objBill.床号 = "" & !付款方式
                    objBill.费别 = IIf(IsNull(!费别), "", !费别)
                    objBill.门诊标志 = IIf(IsNull(!门诊标志), 0, !门诊标志)
                    objBill.加班标志 = IIf(IsNull(!加班标志), 0, !加班标志)
                    objBill.婴儿费 = IIf(IsNull(!婴儿费), 0, !婴儿费)
                    objBill.划价人 = IIf(IsNull(!划价人), "", !划价人)
                    objBill.操作员编号 = IIf(IsNull(!操作员编号), "", !操作员编号)
                    objBill.操作员姓名 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
                    objBill.发生时间 = !发生时间
                    objBill.登记时间 = !登记时间
                    objBill.多病人单 = (IIf(IsNull(!多病人单), 0, !多病人单) = 1)
                End If
                
                '处理收费细目=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                            
                '处理序号,从属父号
                intCurNo = intCurNo + 1
                objBillDetail.序号 = intCurNo '实际是行号
                colSerial.Add intCurNo, "_" & !序号 '记录原序号现在的行号
                If Not IsNull(!从属父号) Then
                    objBillDetail.从属父号 = colSerial("_" & !从属父号)
                End If
                                                                    
                '使用原定的动态费别
                objBillDetail.费别 = IIf(IsNull(!费别), "", !费别)
                objBillDetail.收费类别 = !收费类别
                objBillDetail.收费细目ID = !收费细目ID
                objBillDetail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.付数 = Nvl(!付数, 1)
                If InStr(",5,6,7,", !收费类别) > 0 And gbln药房单位 Then
                    objBillDetail.数次 = Nvl(!数次, 0) / Nvl(!药房包装, 1)
                Else
                    objBillDetail.数次 = Nvl(!数次, 0)
                End If
                objBillDetail.原始数量 = objBillDetail.付数 * objBillDetail.数次
                
                objBillDetail.发药窗口 = IIf(IsNull(!发药窗口), "", !发药窗口)
                
                objBillDetail.附加标志 = IIf(IsNull(!附加标志), 0, !附加标志)
                
                objBillDetail.摘要 = IIf(IsNull(!摘要), "", !摘要)
                
                objBillDetail.执行部门ID = IIf(IsNull(!执行部门ID), 0, !执行部门ID)
                
                objBillDetail.原始执行部门ID = objBillDetail.执行部门ID     '用于修改时快速判断库存
                
                objBillDetail.Detail.ID = !收费细目ID
                objBillDetail.Detail.编码 = !编码
                objBillDetail.Detail.变价 = (IIf(IsNull(!是否变价), 0, !是否变价) = 1)
                objBillDetail.Detail.从项数次 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.固有从属 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.规格 = IIf(IsNull(!规格), "", !规格)
                objBillDetail.Detail.计算单位 = Nvl(!计算单位)
                
                objBillDetail.Detail.药房单位 = Nvl(!药房单位)
                objBillDetail.Detail.药房包装 = Nvl(!药房包装, 1)
                
                If InStr(",4,5,6,7,", !收费类别) > 0 Then
                    dblStock = GetStock(!收费细目ID, !执行部门ID)
                Else
                    dblStock = 0
                End If

                If InStr(",5,6,7,", !收费类别) > 0 And gbln药房单位 Then dblStock = dblStock / Nvl(!药房包装, 1)
                If blnModi Then
                    If InStr(",5,6,7,", !收费类别) > 0 Or !收费类别 = "4" And Nvl(!跟踪在用, 0) = 1 Then dblStock = dblStock + objBillDetail.原始数量
                End If
                objBillDetail.Detail.库存 = dblStock
                
                objBillDetail.Detail.加班加价 = (IIf(IsNull(!加班加价), 0, !加班加价) = 1)
                objBillDetail.Detail.类别 = IIf(IsNull(!类别), "", !类别)
                objBillDetail.Detail.类别名称 = IIf(IsNull(!类别名称), "", !类别名称)
                objBillDetail.Detail.名称 = IIf(IsNull(!名称), "", !名称)
                objBillDetail.Detail.商品名 = Nvl(!商品名)
                objBillDetail.Detail.屏蔽费别 = (IIf(IsNull(!屏蔽费别), 0, !屏蔽费别) = 1)
                objBillDetail.Detail.说明 = IIf(IsNull(!说明), "", !说明)
                objBillDetail.Detail.执行科室 = IIf(IsNull(!执行科室), 0, !执行科室)
                objBillDetail.Detail.类型 = IIf(IsNull(!费用类型), "", !费用类型)
                objBillDetail.Detail.诊疗名称 = Nvl(!诊疗名称)
                objBillDetail.Detail.中药形态 = Val(Nvl(!结论))
                
                If InStr(",5,6,7,", !收费类别) > 0 Then
                    objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                    objBillDetail.Detail.处方限量 = Get处方限量(objBillDetail.Detail.ID)
                End If
                objBillDetail.Detail.录入限量 = Val("" & !录入限量)
                
                objBillDetail.Detail.药名ID = IIf(IsNull(!药名ID), 0, !药名ID)
                objBillDetail.Detail.变价 = IIf(IsNull(!是否变价), 0, !是否变价) = 1
                objBillDetail.Detail.分批 = IIf(IsNull(!分批), 0, !分批) = 1
                objBillDetail.Detail.跟踪在用 = Nvl(!跟踪在用, 0) = 1
                
                '问题:41136
                str摘要 = objBillDetail.摘要
                If Not blnModi Then '90304
                    str摘要 = gclsInsure.GetItemInfo(int险类, objBill.病人ID, objBillDetail.收费细目ID, str摘要, 1, , "|1")
                    objBillDetail.摘要 = str摘要
                End If
                
                '处理价格部份=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '按照现有的价格设置重新计算
                    If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        '时价药品计算价格(分批可不分批)
                        dblAllTime = !付数 * !数次 '这里是售价数量
                        If dblAllTime <> 0 Or Nvl(!是否变价, 0) = 0 Then
                            Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                        "获取药品当前售价", CLng(!收费细目ID), objBillDetail.执行部门ID, dblAllTime)
                            If rsPrice.EOF Then
                                '获取价格失败
                                If !收费类别 = "4" Then
                                    MsgBox "卫生材料""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
                                Else
                                    MsgBox "药品""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
                                End If
                                objBillIncome.标准单价 = 0
                            Else
                                strPrice = Nvl(rsPrice!Price) & "|||"
                                varPrice = Split(strPrice, "|")
                                objBillIncome.标准单价 = Val(varPrice(0))
                                dbl剩余数量 = Val(varPrice(2))
                                
                                If dbl剩余数量 <> 0 And Nvl(!是否变价, 0) = 1 Then
                                    '数量未分解完毕
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                End If
                            End If
                        Else
                            objBillIncome.标准单价 = 0
                        End If
                    ElseIf Nvl(!是否变价, 0) = 1 Then
                        If Abs(!标准单价) > Abs(Val(Nvl(!现价))) Then
                            objBillIncome.标准单价 = Val(Nvl(!缺省价格))
                        Else
                            objBillIncome.标准单价 = !标准单价
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                                        
                    If InStr(",5,6,7,", !收费类别) > 0 And gbln药房单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!药房包装, 1), gstrFeePrecisionFmt)
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrFeePrecisionFmt)
                    End If
                    objBillIncome.现价 = IIf(IsNull(!现价), 0, !现价) '现价原价对药品变价无用
                    objBillIncome.原价 = IIf(IsNull(!原价), 0, !原价)
                    objBillIncome.收入项目ID = IIf(IsNull(!现收入ID), 0, !现收入ID)
                    objBillIncome.收入项目 = IIf(IsNull(!收入项目), "", !收入项目)
                    objBillIncome.收据费目 = IIf(IsNull(!现费目), "", !现费目)
                    
                    '应收金额=单价*付次*数次
                    objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    
                    '附加手术费率用计算(所有收入项目)
                    If IIf(IsNull(!附加标志), 0, !附加标志) = 1 And IIf(IsNull(!收费类别), "", !收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * IIf(IsNull(!附术收费率), 1, !附术收费率 / 100)
                    End If
                    
                    '加班费用率计算
                    dbl加班加价率 = 0
                    If IIf(IsNull(!加班标志), 0, !加班标志) = 1 And IIf(IsNull(!加班加价), 0, !加班加价) = 1 Then
                        dbl加班加价率 = IIf(IsNull(!加班加价率), 0, !加班加价率 / 100)
                        objBillIncome.应收金额 = objBillIncome.应收金额 + objBillIncome.应收金额 * dbl加班加价率
                    End If
                    objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                    
                    '计算实收金额
                    If IIf(IsNull(!屏蔽费别), 0, !屏蔽费别) = 1 Then
                        objBillIncome.实收金额 = objBillIncome.应收金额
                    Else
                        '使用原定的动态费别
                        objBillIncome.实收金额 = ActualMoney(objBillDetail.费别, !现收入ID, objBillIncome.应收金额, _
                            objBillDetail.收费细目ID, objBillDetail.执行部门ID, objBillDetail.原始数量, dbl加班加价率)
                    End If
                    
                    With objBillIncome
                        '获取项目保险信息,仅医保病人才算
                        If int险类 <> 0 And bytFlag = 0 Then
                            strAdvance = objBillDetail.摘要 & "||" & objBillDetail.原始数量
                            strInfo = gclsInsure.GetItemInsure(objBill.病人ID, objBillDetail.收费细目ID, .实收金额, True, int险类, strAdvance)
                            If strInfo <> "" Then
                                objBillDetail.保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                                objBillDetail.保险大类ID = Val(Split(strInfo, ";")(1))
                                .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                                objBillDetail.保险编码 = CStr(Split(strInfo, ";")(3))
                                
                                If UBound(Split(strInfo, ";")) >= 4 Then
                                    If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.摘要 = CStr(Split(strInfo, ";")(4))
                                    If UBound(Split(strInfo, ";")) >= 5 Then
                                        If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.类型 = Split(strInfo, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                    
                        objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
                    End With
                    
                    '判断下一条记录是否属于当前行
                    blnDo = False
                    int序号 = !序号
                    .MoveNext
                    If Not .EOF Then blnDo = (int序号 = !序号)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Pages(1).Details.Add .费别, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .发药窗口, _
                        .付数, .数次, .附加标志, .执行部门ID, .InComes, , .保险项目否, .保险大类ID, .保险编码, .摘要, .原始数量, .原始执行部门ID
                    
                    '设置工本费
                    If objBill.Pages(1).Details(objBill.Pages(1).Details.Count).附加标志 = 8 Then
                        objBill.Pages(1).Details(objBill.Pages(1).Details.Count).工本费 = True
                    End If
                End With
            Loop
        End With
        
        '读取中药煎法
        If Not bln不读煎法 Then
            strSQL = "Select 外观 From 药品收发记录 Where NO=[1] And 单据=[2]"  '8-收费处方发药；9-记帐单处方发药；
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, IIf(bytFlag = 0 Or bytFlag = 1, 8, 9))
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!外观) Then
                    objBill.Pages(1).煎法 = rsTmp!外观
                End If
            End If
        End If
        
    End If
    
    Set ImportBill = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillOnlyFactMoney(ByVal strNo As String) As Boolean
'功能：判断一张收费单据是否仅有工本费
'参数：strNO=F0000001,不带'号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(ID) as 所有,Sum(Decode(Nvl(附加标志,0),8,1,0)) as 工本费 From 门诊费用记录" & _
        " Where 记录性质=1 And 记录状态=1 And 价格父号 is Null And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "OutExse", strNo)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!工本费, 0) = 1 And Nvl(rsTmp!所有, 0) = 1 Then
            BillOnlyFactMoney = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxFact(ByVal strNo As String) As String
'功能：获取指定收费单据(可能为多张中的一张)发出的最大票据号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    '应取最后一次打印的最大号码
    strSQL = "Select Max(ID) From 票据打印内容 Where 数据性质=1 And NO=[1]"
    strSQL = "Select Max(A.号码) as 号码 " & _
        " From 票据使用明细 A,票据打印内容 B" & _
        " Where B.数据性质=1 And B.ID=(" & strSQL & ")" & _
        " And A.打印ID=B.ID And A.票种=1 And A.性质=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetMaxFact = Nvl(rsTmp!号码)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function RePrintCharge(ByVal bytType As Byte, strNos As String, frmParent As Object, _
                            ByRef lng领用ID As Long, ByVal strReclaimInvoice As String, _
                            Optional blnDelOpt As Boolean, Optional DateDel As Date, _
                            Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                            Optional ByVal blnDelRecord As Boolean, _
                            Optional lngShareUseID As Long, Optional strUseType As String = "", _
                            Optional blnOnePatiPrint As Boolean = False, _
                            Optional strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前收款记录重新打印一张票据
    '入参:1-重打;2-补打
    '       strNOs -指定要重打的单据号，带引号，可能是多个单据号，为"'AAA','BBB',..."的形式
    '       lng领用ID-上次使用的领用ID,第一次使用或在清单管理界面调用时没有
    '       strReclaimInvoice-实际收回的发票数(只有票据分配规则为1和2才传入)
    '       blnDelOpt-退费重打操作调用
    '       DateDel-退费时间
    '       intPrintFormat-打印格式序号
    '       intPrintOldFormat-老版票据打印格式
    '       blnVirtualPrint-医保接口打印票据，HIS不调打印只走票号
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '       lngShareUseID-共享票号领用ID(27559)
    '       blnOnePatiPrint-按病人一次打印
    '       strPriceGrade-价格等级，用于计算工本费
    '返回:打印成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 11:50:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim strRptName As String, int发票张数 As Integer
    Dim bytPrintType As Byte, lng打印ID As Long
    Dim bln分别打印 As Boolean
    
    blnHaveInvoice = lng领用ID <> 0     '主要是退费用,如果存在领用的,则必须回收发票,然后重打发票:30386
    
    If blnHaveInvoice = False And blnDelOpt Then
        blnHaveInvoice = zlCheckIsPrintInvoice(strNos)
    End If
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_1"
   
    '如果严格控制票据使用
    If gblnStrictCtrl Then
        '此时只判断是否有,打印之前再根据张数判断是否够用
        int发票张数 = 1
        If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 And blnOnePatiPrint = False Then
            int发票张数 = UBound(Split(strNos, ",")) + 1
            If int发票张数 = -1 Then int发票张数 = 1
        End If
        If zlCheckInvoiceValied(lng领用ID, int发票张数, strInvoice, lngShareUseID, strUseType) = False Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If intPrintFormat = 0 Then   '以缺省票据格式显示
                intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", intPrintFormat
            '由于没有格式的传入,因此,需要强制缺省到指定格式
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            '取出选择的格式
            intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        '取下一个票据号码
        If Not gblnStrictCtrl Then
            
            '有可能是第一次使用
            Do
                blnInput = False
                '非严格控制时直接从本地读取
                strInvoice = UCase(zlDatabase.GetPara("当前收费票据号", glngSys, 1121, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlstr.Increase(strInvoice)
                    strInvoice = UCase(InputBox("请确认" & IIf(bytType = 1, "重打", "补打") & "使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '用户取消输入,允许打印
                If strInvoice = "" Then
                    If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '检查输入有效性
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                            MsgBox "输入的票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '根据票据领用读取
                blnInput = False
                strInvoice = GetNextBill(lng领用ID)
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    '30386:打印了发票的,必需重打再发出
                    If frmInputBox.InputBox(frmParent, "开始发票号", "无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "开始发票号", "请确认" & IIf(bytType = 1, "重打", "补打") & "使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnHaveInvoice And blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function
                
                '检查输入有效性
                If blnInput Then
                    If zlCheckInvoiceValied(lng领用ID, 1, strInvoice, lngShareUseID, strUseType) Then blnValid = True
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        bytPrintType = IIf(blnDelOpt, 3, 2)
        If bytType = 2 And gTy_Module_Para.byt票据分配规则 <> 0 Then
            bytPrintType = 4  '补打
        End If
        
        
        If blnOnePatiPrint Then
            '保存临进数据
            If zlSaveTempPrintData(Replace(strNos, "'", ""), lng领用ID, strInvoice, lng打印ID) = False Then Exit Function
        End If
        
        bln分别打印 = gTy_Module_Para.bln分别打印 And Check体检单据(Replace(strNos, "'", "")) = False And blnOnePatiPrint = False
        
        '1-新单打印,2-重打,3-退费打印,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入)
        If blnDelOpt Then
            Call frmPrint.ReportPrint(bytPrintType, strNos, "", strReclaimInvoice, lng领用ID, lngShareUseID, strInvoice, _
                DateDel, , , bln分别打印, intPrintFormat, blnVirtualPrint, , strUseType, , blnOnePatiPrint, lng打印ID, strPriceGrade)
        Else
            Call frmPrint.ReportPrint(bytPrintType, strNos, "", strReclaimInvoice, lng领用ID, lngShareUseID, strInvoice, _
                zlDatabase.Currentdate, , , bln分别打印, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType, , _
                blnOnePatiPrint, lng打印ID, strPriceGrade)
        End If
        RePrintCharge = True
    End If
End Function

Public Function Check体检单据(ByVal strNos As String) As Boolean
    '只要有一张是体检,则认为全部是体检单据
    'strNOs -指定要重打的单据号，可能是多个单据号，为"AAA,BBB,..."的形式
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errH:
    If strNos = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
            " From 门诊费用记录 A, Table(Cast(f_Str2list([1]) As t_Strlist)) B" & vbNewLine & _
            " Where a.No = b.Column_Value And Mod(a.记录性质, 10) = 1 And a.门诊标志 = 4 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否包含体检费用", strNos)
    Check体检单据 = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintDelCharge(ByVal lng结算序号 As Long, frmParent As Object, _
                            ByRef lng领用ID As Long, Optional ByVal blnDelOpt As Boolean, Optional DateDel As Date, _
                            Optional intPrintFormat As Integer, Optional blnVirtualPrint As Boolean, _
                            Optional ByVal blnDelRecord As Boolean, _
                            Optional lngShareUseID As Long, Optional strUseType As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印退费发票(红票)
    '入参:
    '       lng冲销ID -指定要重打的单据的结算序号
    '       lng领用ID-上次使用的领用ID,第一次使用或在清单管理界面调用时没有
    '       DateDel-退费时间
    '       blnDelOpt-是否是重打或补打调用
    '       intPrintFormat-打印格式序号
    '       blnVirtualPrint-医保接口打印票据，HIS不调打印只走票号
    '       blnDelRecord-重打时，是否是对退费记录进行重打(目前只有北京医保(医保接口打印票据)才允许)
    '       lngShareUseID-共享票号领用ID(27559)
    '返回:打印成功,返回true,否则返回False
    '编制:
    '日期:2016-05-27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInvoice As String, strInfo As String
    Dim j As Integer, blnValid As Boolean, blnInput As Boolean
    Dim blnDo As Boolean, blnHaveInvoice As Boolean
    Dim lng领用IDTemp As Long
    Dim strRptName As String
    Dim bytPrintType As Byte, lng打印ID As Long
    Dim bln分别打印 As Boolean
    
    strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_7"
   
    '如果严格控制票据使用
    If gblnStrictCtrl Then
        '此时只判断是否有,打印之前再根据张数判断是否够用
        lng领用ID = GetInvoiceGroupID(1, 1, lng领用ID, lngShareUseID, , strUseType)
        Select Case lng领用ID
            Case -1
                MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Case -2
                MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
        End Select
        If lng领用ID <= 0 Then Exit Function
    End If

    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        If blnDelOpt Then
            blnDo = True
        Else
            If intPrintFormat = 0 Then   '以缺省票据格式显示
                intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
            End If
            SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", intPrintFormat
            '由于没有格式的传入,因此,需要强制缺省到指定格式
            blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
            '取出选择的格式
            intPrintFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
    End If
    
    If blnDo Then
        '取下一个票据号码
        If Not gblnStrictCtrl Then
            '有可能是第一次使用
            Do
                blnInput = False
                '非严格控制时直接从本地读取
                strInvoice = UCase(zlDatabase.GetPara("当前收费票据号", glngSys, 1121, ""))
                If strInvoice = "" Then
                    strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                    vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                    "", frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                Else
                    strInvoice = zlCommFun.IncStr(strInvoice)
                    strInvoice = UCase(InputBox("请确认使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
                    
                '用户取消输入,允许打印
                If strInvoice = "" Then
                    If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnValid = True
                Else
                    '检查输入有效性
                    If blnInput Then
                        If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                            MsgBox "输入的票据号码长度应该为 " & gbytFactLength & " 位！", vbInformation, gstrSysName
                        Else
                            blnValid = True
                        End If
                    Else
                        blnValid = True
                    End If
                End If
            Loop While Not blnValid
        Else
            Do
                '根据票据领用读取
                blnInput = False
                strInvoice = GetNextBill(lng领用ID)
                If strInvoice = "" Then
                    '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                    '30386:打印了发票的,必需重打再发出
                    If frmInputBox.InputBox(frmParent, "开始发票号", "无法根据票据领用情况获取将要使用的开始票据号，" & _
                                    vbCrLf & "请你输入将要使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                     blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                Else
                    '30386
                    If frmInputBox.InputBox(frmParent, "开始发票号", "请确认使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                                    blnDelOpt, frmParent.Left + 1500, frmParent.Top + 1500) = False Then
                                Exit Function
                    End If
                    blnInput = True
                End If
                
                '用户取消输入,不打印
                If strInvoice = "" Then Exit Function
                
                '检查输入有效性
                If blnInput Then
                    lng领用IDTemp = GetInvoiceGroupID(1, 1, lng领用ID, lngShareUseID, strInvoice, strUseType)
                    If lng领用IDTemp = -3 Then
                        MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                    Else
                        lng领用ID = lng领用IDTemp
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            Loop While Not blnValid
        End If
        
        '1-新单打印,2-重打,3-退费打印,4-补打票据(只有:2-按系统预定规则和3-用户自定规则时才转入),6-退费票据(红票)打印
        If blnDelOpt Then
            Call frmPrint.ReportPrint(6, lng结算序号, "", "", lng领用ID, lngShareUseID, strInvoice, _
                DateDel, , , False, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType)
        Else
            Call frmPrint.ReportPrint(6, lng结算序号, "", "", lng领用ID, lngShareUseID, strInvoice, _
                 zlDatabase.Currentdate, , , False, intPrintFormat, blnVirtualPrint, blnDelRecord, strUseType)
        End If
        PrintDelCharge = True
    End If
End Function

Public Function zlGetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "", _
    Optional ByRef lngRestNum As Long, _
    Optional ByRef lngNextUseID As Long, _
    Optional ByRef bytErr As Byte) As Long
    '功能：获取张数够用并且指定票据在其可用范围内的领用ID
    '入参：
    '    bytKind        =   票种
    '    intNum         =   要打印的票据张数
    '    lngLastUseID   =   上次使用的领用ID
    '    lngShareUseID  =   本地参数指定的共用ID
    '    strBill        =   当前票据号，用于检查领用批次的票据范围
    '    strUseType     =   使用类别
    '出参：
    '    lngRestNum     =   上次使用批次剩余的票据数
    '    lngNextUseID   =   下一个可用批次的领用ID
    '    bytErr         =
    '                    1 - 没有自用(用完或本批不够且下一批也不够，或未领用),未设置共用
    '                    2 - 没有自用(用完或不够，或未领用),设置的共用已用完或不够
    '                    3 - 指定票据号不在当前所有可用领用批次的有效票据号范围内
    '                    4 - 指定批次的票据已用完
    '                    5 - 指定批次的票据不够用,下一个可用批次是自用票据,领用ID记录在lngNextUseID中
    '                    6 - 指定批次的票据不够用,下一个可用批次是共用票据,领用ID记录在lngNextUseID中
    '返回：
    '    >0   =   成功，可用的领用ID
    '    =0   =   失败
    '修改:冉俊明,自动更换票据批次,减少发票的浪费
    '修改日期:2015-04-21
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    Dim lngShortCount As Long '当前批次缺的票据张数，从下一个批次取
    Dim blnNext As Boolean '是否下一个批次
    
    On Error GoTo errH
    lngShortCount = intNum
    lngRestNum = 0: lngNextUseID = 0: bytErr = 0
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSQL = "" & _
        "   Select 前缀文本,开始号码,终止号码,剩余数量" & _
        "   From 票据领用记录 " & _
        "   Where 票种=[1] And 剩余数量>0 And ID=[2]  " & _
        "           And (Nvl(使用类别,'LXH')=[3] Or  使用类别 Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then '目前的票据号可能和上次不同，所以需要检查范围
                lngShortCount = lngShortCount - Val(Nvl(!剩余数量))
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !前缀文本
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                        blnTmp = True
                    End If
                End If
                If strBill = "" Or strBill <> "" And Not blnTmp Then
                    blnNext = True
                    lngRestNum = Val(Nvl(!剩余数量))
                    zlGetInvoiceGroupID = lngLastUseID
                    bytErr = IIf(lngRestNum < intNum, 5, 0)
'                    Exit Function'这里无论当前批次是否足够都不退出，
                                    '是因为预定义分配票号检查时还不知道需要多少张票据，始终传进来的都是1张，
                                    '无法知道当前批次是否足够，所以始终取出下一个批次（若果有）
                Else
                    bytErr = 3
                End If
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前指定批次已用完
                blnNext = True
                lngRestNum = Val(Nvl(!剩余数量))
                zlGetInvoiceGroupID = lngLastUseID
                bytErr = 4
            End If
        End With
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSQL = "" & _
    "   Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量" & _
    "   From 票据领用记录" & _
    "   Where 票种 = [1] And 剩余数量 >0 And 领用人 = [2]  " & _
    "           And (Nvl(使用类别,'LXH')=[3] Or  使用类别 Is NULL ) " & _
    "           And 使用方式 = 1" & _
    IIf(lngLastUseID > 0, "           And ID <> [4]", "") & _
    "   Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 desc, 开始号码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, UserInfo.姓名, IIf(strUseType = "", "LXH", strUseType), lngLastUseID)
    With rsTmp
        If .RecordCount > 0 Then
            For i = 1 To .RecordCount
                lngShortCount = lngShortCount - Nvl(!剩余数量)
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !前缀文本
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                        blnTmp = True
                    End If
                End If
                If blnNext Or strBill = "" Or strBill <> "" And Not blnTmp Then '第一次使用时没有当前票据号
                    If blnNext Then
                        If lngShortCount > 0 Then
                            bytErr = 5
                        Else
                            bytErr = IIf(lngRestNum < intNum, 5, bytErr)
                            lngNextUseID = Nvl(!ID)
                            Exit Function
                        End If
                    Else
                        blnNext = True
                        lngRestNum = Val(Nvl(!剩余数量))
                        zlGetInvoiceGroupID = Nvl(!ID)
                        bytErr = IIf(lngRestNum < intNum, 5, 0)
'                        Exit Function'这里无论这个批次是否足够都不退出，
                                    '是因为预定义分配票号检查时还不知道需要多少张票据，始终传进来的都是1张，
                                    '无法知道当前批次是否足够，所以始终取出下一个批次（若果有）
                    End If
                Else
                    bytErr = 3
                End If
                .MoveNext
            Next
        Else
            bytErr = IIf(lngShortCount > 0, IIf(blnNext, 5, 1), bytErr)
        End If
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 And lngShareUseID <> lngLastUseID Then
        strSQL = "" & _
        "   Select 前缀文本,开始号码,终止号码,剩余数量" & _
        "   From 票据领用记录  " & _
        "   Where 票种=[1] And 剩余数量>0 And ID=[2] " & _
        "   And (Nvl(使用类别,'LXH')=[3] Or  使用类别 Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                lngShortCount = lngShortCount - Nvl(!剩余数量)
                If strBill <> "" Then
                    blnTmp = False
                    strPre = "" & !前缀文本
                    If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                        blnTmp = True
                    End If
                End If
                If blnNext Or strBill = "" Or strBill <> "" And Not blnTmp Then '第一次使用时没有当前票据号
                    If blnNext Then
                        If lngShortCount > 0 Then
                            bytErr = 6
                        Else
                            bytErr = IIf(lngRestNum < intNum, 6, bytErr)
                            lngNextUseID = lngShareUseID
                            Exit Function
                        End If
                    Else
                        If lngShortCount > 0 Then
                            bytErr = 2
                        Else
                            bytErr = 0
                            lngRestNum = Val(Nvl(!剩余数量))
                            zlGetInvoiceGroupID = lngShareUseID
                            Exit Function
                        End If
                    End If
                Else
                    bytErr = 3
                End If
            Else
                bytErr = 2
            End If
        End With
    Else
        bytErr = IIf(lngShortCount > 0, IIf(blnNext, 6, 1), bytErr)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlCheckInvoiceValied(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "", _
    Optional ByVal lngShareUseID As Long, Optional strUseType As String = "", _
    Optional ByRef lngRestNum As Long, _
    Optional ByRef lngNextUseID As Long, Optional ByRef strNextInvoiceNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:
    '    lng领用ID = 领用id
    '    intNum = 页数
    '    strInvoiceNO = 输入的发票号
    '    lngShareUseID = 本地参数指定的共用ID
    '    strUseType = 使用类别
    '出参:lng领用ID-领用ID
    '    lngRestNum     =   上次使用批次剩余的票据数
    '    lngNextUseID = 下一个可用批次的领用ID
    '    strNextInvoiceNO = 下一个可用批次的下一个票号
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '修改:冉俊明
    '修改日期:2015-04-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim bytErr As Byte
    
    strNextInvoiceNO = ""
    lng领用ID = zlGetInvoiceGroupID(1, intNum, lng领用ID, lngShareUseID, strInvoiceNO, strUseType, _
                                        lngRestNum, lngNextUseID, bytErr)
    If lngNextUseID <> 0 Then strNextInvoiceNO = GetNextBill(lngNextUseID)
    If lng领用ID > 0 And bytErr = 0 Then zlCheckInvoiceValied = True: Exit Function
    'bytErr =
    '         1 - 没有自用(用完或本批不够且下一批也不够，或未领用),未设置共用
    '         2 - 没有自用(用完或不够，或未领用),设置的共用已用完或不够
    '         3 - 指定票据号不在当前所有可用领用批次的有效票据号范围内
    '         4 - 指定批次的票据已用完
    '         5 - 指定批次的票据不够用,下一个可用批次是自用票据,领用ID记录在lngNextUseID中
    '         6 - 指定批次的票据不够用,下一个可用批次是共用票据,领用ID记录在lngNextUseID中
    Select Case bytErr
        Case 1
            If Trim(strUseType) = "" Then
                MsgBox "你没有自用和共用的收费票据，请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "你没有自用和共用的『" & strUseType & "』收费票据，请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
            End If
        Case 2
            If Trim(strUseType) = "" Then
                MsgBox "本地的共用票据已经用完，请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            Else
                MsgBox "本地的共用票据的『" & strUseType & "』收费票据已经用完，请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End If
        Case 3
            MsgBox "当前票据号码 " & strInvoiceNO & " 不在可用领用批次的有效票据号范围内，请重新输入！", vbInformation, gstrSysName
        Case 4 '指定批次的票据已用完
            If strNextInvoiceNO <> "" Then
                If MsgBox("当前批次票据已使用完，是否使用开始票号为『" & strNextInvoiceNO & "』的下一个票据批次完成打印？" & vbCrLf & vbCrLf & _
                    "注意：请核对下一个票据批次的开始票号是否为『" & strNextInvoiceNO & "』，若不是，请选择“否”。若选择“是”，请及时更换发票！", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    zlCheckInvoiceValied = True: Exit Function
                End If
            End If
        Case 5, 6
            If strNextInvoiceNO <> "" Then
                If MsgBox("当前打印共需要 " & intNum & " 张票据，但当前票据号码 " & strInvoiceNO & " 所在批次的有效票据只有 " & lngRestNum & " 张。" & vbCrLf & _
                    "是否使用开始票号为『" & strNextInvoiceNO & "』的下一个" & IIf(bytErr = 5, "自用票据批次", "共用票据批次") & "完成打印？" & vbCrLf & vbCrLf & _
                    "注意：请核对下一个票据批次的开始票号是否为『" & strNextInvoiceNO & "』，若不是，请选择“否”。若选择“是”，请及时更换发票！", _
                    vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    zlCheckInvoiceValied = True: Exit Function
                End If
            Else '没有下一批可用的领用票据了
                MsgBox "当前票据剩余数量 " & lngRestNum & " 张不足本次打印所需数量 " & intNum & " 张，请先领用一批票据或设置本地共用票据。", vbInformation, gstrSysName
            End If
        Case Else
            MsgBox "票据领用信息访问失败！将来，你可以进行重打单据！", vbInformation, gstrSysName
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckDisable(objBill As ExpenseBill) As String
'功能：检查单据中的药品的禁忌情况
'返回：药品互相禁忌提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean, strMsg As String
    
    Err = 0: On Error GoTo errH:
    For p = 1 To objBill.Pages.Count
        strInfo = "": strIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If InStr(",5,6,7,", objBill.Pages(p).Details(i).收费类别) > 0 Then
                strIDs = strIDs & "," & objBill.Pages(p).Details(i).收费细目ID
            End If
        Next
        strIDs = Mid(strIDs, 2)
        If Not (strIDs = "" Or UBound(Split(strIDs, ",")) < 1) Then
            strSQL = _
                " Select /*+ RULE */  A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
                " From 诊疗互斥项目 A,药品规格 B," & _
                "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
                " Where A.项目ID=B.药名ID And B.药品ID  = j.Column_Value" & _
                " Having Count(Distinct A.项目ID)>1  " & _
                "  Group by A.组编号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strIDs)
            
            If Not rsTmp.EOF Then
                strGroup = ""
                For i = 1 To rsTmp.RecordCount
                    strGroup = strGroup & "," & rsTmp!组编号
                    rsTmp.MoveNext
                Next
                strGroup = Mid(strGroup, 2)
                
                For i = 0 To UBound(Split(strGroup, ","))
                    strSQL = _
                        "Select /*+ RULE */   Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
                        " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D," & _
                        "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
                        " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
                        "           And C.组编号=[1]" & _
                        "           And A.药品ID  = j.Column_Value" & _
                        " Order by C.类型,C.组编号,D.编码"
                        
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Val(Split(strGroup, ",")(i)), strIDs)
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "类型=1"
                        If rsTmp.RecordCount > 1 Then
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相慎用)：" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = "类型=2"
                        If rsTmp.RecordCount > 1 Then
                            blnStop = True
                            k = k + 1
                            strInfo = strInfo & vbCrLf & "第 " & k & " 组(互相禁用)：" & vbCrLf
                            For j = 1 To rsTmp.RecordCount
                                strInfo = strInfo & "[" & rsTmp!编码 & "]" & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", "(" & rsTmp!规格 & ")") & "                 " & vbCrLf
                                rsTmp.MoveNext
                            Next
                        End If
                        rsTmp.Filter = 0
                    End If
                Next
                If strInfo <> "" Then
                    If objBill.Pages.Count = 1 Then
                        strMsg = strMsg & vbCrLf & "单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo
                    Else
                        strMsg = strMsg & vbCrLf & "单据" & p & "中下列药品互相禁用或慎用：" & vbCrLf & strInfo
                    End If
                End If
            End If
        End If
    Next
    If strMsg <> "" Then
        If blnStop Then
            strMsg = strMsg & vbCrLf & "请修改禁用药品后再继续！"
        Else
            strMsg = strMsg & vbCrLf & "要继续吗？"
        End If
    End If
    CheckDisable = Mid(strMsg, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistIOClass(bytBill As Byte) As Long
'功能：判断是否存在指定处方单据类型的入出类别
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", bytBill)
    If Not rsTmp.EOF Then
        ExistIOClass = IIf(IsNull(rsTmp!类别ID), 0, rsTmp!类别ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetChargeTotal() As Currency
'功能：获取当前操作员当天内的收费总额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    '收费中的非医保结算金额之和
    strSQL = "" & _
    "   Select Sum(冲预交) as 金额 From 病人预交记录 a" & _
    "   Where 记录性质=3 And 操作员姓名=[1]" & _
    "       And 收款时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)" & _
    "       And Not Exists(Select 'X' From 结算方式 b Where 性质 IN(3,4) And a.结算方式=B.名称)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", UserInfo.姓名)
    If Not rsTmp.EOF Then
        GetChargeTotal = IIf(IsNull(rsTmp!金额), 0, rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlInit缺省部门()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化当前操作员的缺省部门
    '编制：刘兴洪
    '日期：2010-08-16 16:28:21
    '说明：31936
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If gstr所属部门ID <> "" Then Exit Sub
    strSQL = "Select 部门ID From 部门人员 Where 人员ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取操作员的所属部门", UserInfo.ID)
    Do While Not rsTmp.EOF
        gstr所属部门ID = gstr所属部门ID & "," & Nvl(rsTmp!部门ID)
        rsTmp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetStockInfo(lng药品ID As Long, bln药房 As Boolean, bln药库 As Boolean) As String
'功能：获取药品在各个药房，药库的库存信息
'参数："bln药房/bln药库"至少要有一个设置为真
'返回：描述信息
    Dim strSQL As String, strSQL2 As String, i As Integer
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Call zlInit缺省部门
    
    If bln药房 And bln药库 Then
        strSQL = "'中药房','西药房','成药房','中药库','西药库','成药库'"
    ElseIf bln药房 Then
        strSQL = "'中药房','西药房','成药房'"
    ElseIf bln药库 Then
        strSQL = "'中药库','西药库','成药库'"
    End If
    
    '排除多个性质的情况,不区分门诊、住院
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN(" & strSQL & ")"
    '药房不分批药品不管效期
    strSQL2 = "Select 部门ID From 部门性质说明 Where 工作性质 IN('西药房','成药房','中药房')"
    '不分批或分批药品
    strSQL = _
        " Select B.编码,B.名称,A.库房ID," & _
        " Nvl(Sum(A.可用数量),0)" & IIf(gbln药房单位, "/Nvl(C." & gstr药房包装 & ",1)", "") & " as 库存" & _
        " From 药品库存 A,(" & strSQL & ") B,药品规格 C" & _
        " Where A.库房ID=B.ID And A.药品ID=C.药品ID" & _
        " And ((A.效期 is NULL Or 效期>Trunc(Sysdate))" & _
        "   Or (Nvl(C.药房分批,0)=0 And A.库房ID IN(" & strSQL2 & ")))" & _
        " And A.性质=1 And A.药品ID=[1]" & _
        " Group by B.编码,B.名称,A.库房ID,Nvl(C." & gstr药房包装 & ",1)" & _
        " Having Sum(Nvl(A.可用数量,0))<>0" & _
        " Order By B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药品ID)
    strSQL = ""
    Do While Not rsTmp.EOF
        If InStr(1, gstr所属部门ID & ",", "," & Nvl(rsTmp!库房id) & ",") > 0 Or gbyt库存显示方式 = 0 Then
            '显示库存数:所属库房,或其他库房显示为库存数时
            strSQL = strSQL & "," & rsTmp!名称 & ":" & rsTmp!库存
        Else
            '非操作员库房,则显示为有无
            strSQL = strSQL & "," & rsTmp!名称 & ":" & IIf(Val(Nvl(rsTmp!库存)) > 0, "有", "无") & "库存."
        End If
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ChargeExistInsure(ByVal strNos As String, _
    Optional lng病人ID As Long, Optional lng结帐ID As Long, _
    Optional bln急诊 As Boolean, Optional ByVal bln退费 As Boolean) As Integer
'功能：判断收费(或退费)记录中是否存在指定的医保结算方式
'参数：strNO=收费单据号
'返回：如果存在则返回单据当时的险类及病人ID,结帐ID,是否急诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errH
            
    lng病人ID = 0: lng结帐ID = 0: bln急诊 = False
    strWhere = " And A.记录状态 " & IIf(bln退费, "= 2", " IN(1,3)")
    strSQL = "" & _
        " Select /*+cardinality(j,10)*/ b.记录id, b.险类, b.病人id, a.是否急诊" & vbNewLine & _
        " From 门诊费用记录 A, 保险结算记录 B, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where a.结帐id = b.记录id And Mod(a.记录性质, 10) = 1 And a.No = j.Column_Value And b.性质 = 1" & strWhere & vbNewLine & _
        "       And Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Replace(strNos, "'", ""))
    If Not rsTmp.EOF Then
        lng病人ID = Nvl(rsTmp!病人ID, 0)
        lng结帐ID = Nvl(rsTmp!记录ID, 0)
        bln急诊 = Nvl(rsTmp!是否急诊, 0) = 1
        ChargeExistInsure = Nvl(rsTmp!险类, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Public Sub Load费别(cbo As ComboBox, ByVal lng科室ID As Long, ByVal bln初诊 As Boolean, ByRef rsTmp As ADODB.Recordset)
'功能：读取可用身份唯一性费别并填写
'参数：lng科室ID-开单科室ID,0表示取适用于所有科室的费别,bln初诊-是否允许仅限初诊的费别,rsTmp-费别记录集

    Dim strSQL As String, i As Integer
    
    cbo.Clear
    On Error GoTo errH
    If rsTmp Is Nothing Then
        strSQL = _
                    " Select a.编码,a.名称,a.简码,Nvl(a.缺省标志,0) as 缺省,Nvl(a.仅限初诊,0) as 初诊,Nvl(b.科室ID,0) as 科室ID" & _
                    " From 费别 a,费别适用科室 b Where a.名称=b.费别(+)" & _
                    " And Nvl(a.服务对象,3) IN(1,3) And a.属性=1 And " & _
                    " Trunc(Sysdate) Between Nvl(a.有效开始,To_Date('1900-01-01','YYYY-MM-DD')) And Nvl(a.有效结束,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by a.编码"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    End If
    '适用科室:1-全部；2-指定
    strSQL = IIf(bln初诊, "", " And 初诊=0")
    rsTmp.Filter = IIf(lng科室ID = 0, "科室ID=0" & strSQL, _
                        "(科室ID=0 " & strSQL & ") OR (科室ID=" & lng科室ID & strSQL & ")")

    For i = 1 To rsTmp.RecordCount
        cbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If gstr费别 = rsTmp!名称 Then
            cbo.ListIndex = cbo.NewIndex
            cbo.ItemData(cbo.NewIndex) = 1
        End If
        If rsTmp!缺省 = 1 Then
            If cbo.ListIndex = -1 Then cbo.ListIndex = cbo.NewIndex
            cbo.ItemData(cbo.NewIndex) = 1
        End If
        '仅限初诊不会是本地和系统缺省
        If rsTmp!初诊 = 1 Then cbo.ItemData(cbo.NewIndex) = 2
        rsTmp.MoveNext
    Next
    'Load费别 = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Load动态费别(科室ID As Long) As String
'功能：权限指定科室读取当前有效的动态费别
'返回：费别串="三八节,五一节"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    'If 科室ID = 0 Then Exit Function  '为0时适用于所有科室
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", 科室ID)
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

Public Function BillCanModi(strNo As String, bytFlag As Byte) As Boolean
'功能：判断一张单据是否可以修改
'参数：bytFlag=记录性质
'说明：如果单据中存在分批或时价药品,则不允许修改(因为库存的问题)

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.NO" & _
        " From 门诊费用记录 A,药品规格 B,收费项目目录 C" & _
        " Where A.收费细目ID=B.药品ID And A.收费细目ID=C.ID" & _
        " And A.记录状态 IN(0,1,3) And A.收费类别 IN('5','6','7')" & _
        " And (Nvl(B.药房分批,0)=1 Or Nvl(C.是否变价,0)=1)" & _
        " And A.NO=[1] And A.记录性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    BillCanModi = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadPatiCardObj(ByVal lng病人ID As Long, strNo As String) As Detail
'功能：读取指定病人的就诊卡划价记录
'返回：strNO=划价单据号
'      就诊卡项目对象(未计算价格)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objDetail As Detail
    
    On Error GoTo errH
    
    strSQL = "Select A.NO,B.ID,B.类别,D.名称 as 类别名称,B.编码,B.名称," & _
        " B.规格,B.计算单位,B.费用类型,B.执行科室,Nvl(B.屏蔽费别,0) as 屏蔽费别," & _
        " Nvl(B.加班加价,0) as 加班加价,Nvl(B.是否变价,0) as 是否变价" & _
        " From 门诊费用记录 A,收费项目目录 B,收费特定项目 C,收费项目类别 D" & _
        " Where A.收费细目ID+0=B.ID And A.记录性质=1 And A.记录状态=0" & _
        " And A.收费细目ID+0=C.收费细目ID And C.特定项目='就诊卡'" & _
        " And A.操作员姓名 is NULL And A.病人ID=[1] And B.类别=D.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID)
    If Not rsTmp.EOF Then
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .编码 = rsTmp!编码
            .规格 = IIf(IsNull(rsTmp!规格), "", rsTmp!规格)
            .计算单位 = IIf(IsNull(rsTmp!计算单位), "", rsTmp!计算单位)
            
            .变价 = rsTmp!是否变价 = 1
            .加班加价 = rsTmp!加班加价 = 1
            .屏蔽费别 = rsTmp!屏蔽费别 = 1
            
            .类别 = rsTmp!类别
            .类别名称 = rsTmp!类别名称
            .名称 = rsTmp!名称
            
            .执行科室 = IIf(IsNull(rsTmp!执行科室), 0, rsTmp!执行科室)
            .类型 = IIf(IsNull(rsTmp!费用类型), "", rsTmp!费用类型)
        End With
        Set ReadPatiCardObj = objDetail
        strNo = rsTmp!NO
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillCanDelete(ByVal strNo As String, ByVal int记录性质 As Byte, _
    Optional ByRef blnHaveExe As Boolean, Optional ByVal strTime As String, Optional ByRef blnFlagPrint As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：判断一张单据是否可以退费或销帐
    '参数：strNO=单据号,,int记录性质=记录性质
    '说明：可以退费或销帐的条件
    '    1.费用未完全执行(执行状态=0,2)
    '    2.剩余数量不<>0
    '    3.以上条件排开误差费用
    '返回：
    '   blnHaveExe=是否存在已(完全/部份)执行的内容
    '   -1=操作失败
    '    0=可以退费或销帐
    '    1=该单据不存在
    '    2=已经全部完全执行(执行状态=1)
    '    3=未完全执行部分剩余数量为0
    '    blnFlagPrint=检查对应的条码是否已打印(检验医嘱中的采集方式已执行)
    '编制:刘兴洪
    '日期:2014-07-15 09:56:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    
    On Error GoTo errH
    strNo = Replace(strNo, "'", "")
    strWhere = " And A.记录性质=[2] "
    If int记录性质 = 1 Then
        strWhere = " And mod(A.记录性质,10)=[2]"
    End If
    '1.费用未完全执行(执行状态=0,2)
    strSQL = "Select Distinct Nvl(A.执行状态,0) as 执行状态,B.样本条码" & _
        " From 门诊费用记录 A,病人医嘱发送 B" & vbNewLine & _
        " Where Nvl(A.附加标志,0)<>9 And A.NO=[1]   And A.记录状态 IN(0,1,3) " & strWhere & vbNewLine & _
        " And A.医嘱序号=B.医嘱ID(+) And A.NO=B.NO(+) And A.记录性质=B.记录性质(+) And Nvl(A.费用状态,0)<>1 " & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int记录性质, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int记录性质)
    End If
    
    '单据不存在
    If rsTmp.EOF Then BillCanDelete = 1: Exit Function
    blnFlagPrint = Not IsNull(rsTmp!样本条码)
    
    '单据已经全部完全执行
    rsTmp.Filter = "执行状态<>1"
    If rsTmp.EOF Then BillCanDelete = 2 ': Exit Function
    
    '是否存在已(完全/部份)执行的内容
    rsTmp.Filter = "执行状态<>0"
    blnHaveExe = Not rsTmp.EOF

    
    '未完全执行部分剩余数量不<>0
    '从原始单据中找未完全执行的行次(部分退药的退费后执行状态=1,但退费记录执行状态<>1)
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    '记录状态=1,3时：0:未执行;1:完全执行;2:部份执行；记录状态=2时：-x:第x次退费
    strSQL = "" & _
    "   Select Nvl(价格父号,序号) as 序号" & _
    "   From 门诊费用记录" & _
    "   Where Nvl(附加标志,0)<>9 And NO=[1]  And Nvl(执行状态,0)<>1 And 记录状态 IN(0,1,3)" & Replace(strWhere, "A.", "") & _
            IIf(strTime <> "", " And 登记时间=[3]", "")
            
    strSQL = _
    "   Select 序号,收费细目ID,Sum(数量) as 剩余数  " & _
    "   From ( Select 记录性质,记录状态,执行状态,Nvl(价格父号,序号) as 序号,收费细目ID," & _
    "                 Avg(Nvl(付数,1)*数次) as 数量 " & _
    "          From 门诊费用记录" & _
    "          Where Nvl(附加标志,0)<>9 And NO=[1] " & Replace(strWhere, "A.", "") & _
    "                And Nvl(执行状态,0)<>1 And Nvl(价格父号,序号) IN(" & strSQL & ")" & _
    "          Group by 记录性质,记录状态,执行状态,Nvl(价格父号,序号),收费细目ID,结帐ID)" & _
    "   Group by 序号,收费细目ID  " & _
    "   Having Sum(数量)<>0"
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int记录性质, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int记录性质)
    End If
    If rsTmp.EOF Then BillCanDelete = 3
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    BillCanDelete = -1
End Function

Public Function BillDeleteAll(strNo As String, bytFlag As Byte, blnHaveExcutePrice As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定单据本次是否可以完全退费(所有剩余数不为零为行都满足准退数=剩余数)
    '入参:strNO-单据号
    '       bytFlag-记录性质(1-收费;2-记帐)
    '返回:全退返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-25 17:13:02
    '说明:要配合界面上是否全部序号退费判断是否真正的完全退费
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
          
    On Error GoTo errH
    '刘兴洪45685,58077
    '读取药品收发记录中的准退数
    strSQL1 = _
    "   Select 费用ID,Sum(Nvl(付数,1)*实际数量) as 准退数量" & _
    "   From 药品收发记录" & _
    "   Where NO=[1] And MOD(记录状态,3)=1 And 审核人 is NULL" & _
    "             And 单据 IN([3],[4])  " & _
    "   Group by 费用ID" & _
    "   Union ALL "
    If blnHaveExcutePrice Then
            '60735:在医嘱执行计价中存在数据时,则按医嘱执行计价中取数
            '77686,李南春,2014/9/18,单据类别限制
            strSQL1 = strSQL1 & _
            " Select Max(ID) As 费用id, Decode(Sign(Sum(数量)), -1, 0, Sum(数量)) As 准退数" & vbNewLine & _
            " From ( Select Decode(a.记录状态, 2, 0, a.Id) As ID, a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * Nvl(a.数次, 1) As 数量," & vbNewLine & _
            "              Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 原始数量" & vbNewLine & _
            "       From 门诊费用记录 A, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And" & vbNewLine & _
            "             a.No = [1] And mod(a.记录性质,10) = [3] And a.记录状态 In (1, 2, 3)　and a.价格父号 is null" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量" & vbNewLine & _
            "       From 门诊费用记录 A, 医嘱执行计价 B, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.收费类别) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From 病人医嘱执行  Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From 病人医嘱发送 Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1))" & vbNewLine & _
            "          And a.No = [1] And mod(a.记录性质,10) = [3] And a.记录状态 In (1, 3)　and a.价格父号 is null" & vbNewLine & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From 药品收发记录 Where 费用id = Q1.Id And instr( ',8,9,10,21,24,25,26,',','||单据||',')>0) " & vbNewLine & _
            " Group by 医嘱ID,收费细目ID  Having Max(ID)<>0"
    Else
         strSQL1 = strSQL1 & _
         " Select Max(ID) as 费用ID,decode(sign(Sum(数量)),-1,0,Sum(数量)) as 准退数 " & _
         " From (   Select J.ID,J.医嘱序号 as 医嘱ID,J.收费细目ID,nvl(J.付数,1)*nvl(J.数次,1)  as 数量 " & _
         "               From  门诊费用记录 J,病人医嘱记录 M " & _
         "               Where  J.医嘱序号=M.ID  " & _
         "                       And Exists(Select 1 From   病人医嘱发送 where 医嘱ID=J.医嘱序号 and  Nvl( 执行状态, 0) <> 1 And No||''=[1] and 记录性质+0=[2]) " & _
         "                       And Exists(Select 1 From   病人医嘱计价 A Where   A.医嘱ID=J.医嘱序号 and A.收费细目ID=J.收费细目ID And A.费用性质=0  and  Nvl( A.收费方式, 0) =0 ) " & _
         "                       And J.No=[1] and mod(J.记录性质,10)=[2] And J.记录状态 in (1,2,3) and J.价格父号 is null   " & _
         "                       And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
         "                       And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0  " & _
         "               Union all  " & _
         "               Select j.id, A.医嘱ID,a.收费细目ID,-1*nvl(a.数量,1)*nvl(C.本次数次,1) as 数量 " & _
         "                              From 病人医嘱计价 A,病人医嘱发送 B,病人医嘱执行 C,门诊费用记录 J,病人医嘱记录 M " & _
         "               where  A.医嘱ID=b.医嘱id  and b.医嘱id=c.医嘱id and b.发送号=c.发送号 And a.医嘱id=M.ID " & _
         "                       And Nvl(C.执行结果, 1) =1  And A.费用性质=0 and  Nvl( A.收费方式, 0) =0  And Nvl(b.执行状态, 0) <> 1 And B.No||''=[1] and B.记录性质+0=[2]  " & _
         "                       And a.医嘱id=J.医嘱序号 and a.收费细目id=j.收费细目id  " & _
         "                       And J.No=[1] and mod(J.记录性质,10)=[2] And J.记录状态 in (1,3) and J.价格父号 is null   " & _
         "                       And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
         "                       And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0   "
        
        '58077需要排开医嘱计划中不能正常收取的费用:
             '   0-正常收取，1-检验试管费用；2-一次发送只收取一次；3-当天只收取一次；4-当天未执行收取一次；5-当天只收取一次，排斥其他项目；6-当天未执行收取一次，排斥其他项目；7-每天首次不收取
        strSQL1 = strSQL1 & " Union All" & _
         "               Select j.Id, a.医嘱id, a.收费细目id, 0 As 数量 " & _
         "               From 病人医嘱计价 A,门诊费用记录 J , 病人医嘱记录 M " & _
         "               Where  a.医嘱id = M.ID and a.医嘱id = j.医嘱序号 And a.收费细目id = j.收费细目id  And A.费用性质=0  and Nvl(a.收费方式, 0) <> 0  " & _
         "                            And j.No =[1] And mod(j.记录性质,10) = [2]  And nvl(J.执行状态,0)=2 " & _
         "                            And j.记录状态 In (1, 3) And  j.价格父号 Is Null And Instr('5,6,7', j.收费类别) = 0  " & _
         "                            And Not Exists(Select 1 From 材料特性 Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
         "                            And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0  " & _
         "               ) " & _
         " group by 医嘱ID,收费细目ID Having Max(ID)<>0"
      End If
    
    '整张费用单据中剩余数不为0的行(明细到每一行)
    '执行状态原始记录上判断(部分退药且部分退费的记录)
    strSQL = _
    " Select Sum(A.ID) as ID,Sum(A.执行状态) as 执行状态," & _
    "               A.序号,A.收费类别,Sum(数量) as 剩余数量" & _
    " From (  Select    Decode(A.记录状态,2,0,A.ID) as ID," & _
    "                           Decode(A.记录状态,2,0,Nvl(A.执行状态,0)) as 执行状态," & _
    "                           A.序号,A.收费类别,Nvl(A.付数,1)*A.数次 as 数量" & _
    "               From 门诊费用记录 A" & _
    "               Where A.价格父号 is NULL And Nvl(A.附加标志,0)<>9 And mod(A.记录性质,10)=[2] And A.NO=[1]" & _
    "               ) A" & _
    "   Group by A.序号,A.收费类别" & _
    "   Having Nvl(Sum(数量),0)<>0"
                
    '有剩余数量无准退数量的有两种情况：
        '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应有剩余数量
        '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        'Decode(Instr(',4,5,6,7,',A.收费类别),0,A.剩余数量,
    strSQL = _
    "   Select A.序号 " & _
    "   From ( Select A.序号,A.剩余数量," & _
    "                           Decode(A.执行状态,1,0,Nvl(B.准退数量,A.剩余数量) ) as 准退数量" & _
    "               From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
    "               Where A.ID=B.费用ID(+)" & _
    "               ) A" & _
    " Where Nvl(A.准退数量,0)<>A.剩余数量"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, IIf(bytFlag = 2, "9", "8"), IIf(bytFlag = 2, "25", "24"))
    BillDeleteAll = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillDeleteAllNew(strNo As String, bytFlag As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定单据本次是否可以完全退费(所有剩余数不为零为行都满足准退数=剩余数)
    '入参:strNO-单据号
    '     bytFlag-记录性质(1-收费;2-记帐)
    '返回:全退返回true,否则返回False
    '编制:冉俊明
    '日期:2016-10-09 17:13:02
    '说明:要配合界面上是否全部序号退费判断是否真正的完全退费
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
          
    On Error GoTo errH
    '刘兴洪45685,58077,99715
    '读取药品收发记录中的准退数
    strSQL1 = _
        " Select 费用ID,Sum(Nvl(付数,1)*实际数量) as 准退数量" & vbNewLine & _
        " From 药品收发记录" & vbNewLine & _
        " Where NO=[1] And Mod(记录状态,3)=1 And 审核人 Is NULL And 单据 IN([3],[4])" & vbNewLine & _
        " Group by 费用ID"
        
    '求诊疗相关的准退数
    strSQL1 = strSQL1 & vbNewLine & _
        " Union ALL " & vbNewLine & _
        " Select Max(ID) As 费用ID, Nvl(Sum(数量), 0) As 准退数" & vbNewLine & _
        " From(Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, Decode(c.执行状态, 0, 1, 0) * c.数量 As 数量" & vbNewLine & _
        "      From 门诊费用记录 A, 病人医嘱发送 B, 医嘱执行计价 C, 病人医嘱记录 M" & vbNewLine & _
        "      Where a.医嘱序号 = b.医嘱id And b.医嘱id = c.医嘱id And b.医嘱ID = m.ID" & vbNewLine & _
        "            And b.发送号 = c.发送号 And a.收费细目id = c.收费细目id + 0 And a.价格父号 Is Null" & vbNewLine & _
        "            And Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0" & vbNewLine & _
        "            And Not Exists(Select 1 From 材料特性 C Where a.收费细目id = c.材料id And c.跟踪在用 = 1)" & vbNewLine & _
        "            And Instr(',C,D,F,G,K,',','||m.诊疗类别||',')=0" & vbNewLine & _
        "            And a.No = [1] And a.记录性质 = [2] And a.记录状态 In (1, 3) And b.记录性质 = [2]" & vbNewLine & _
        "     )" & vbNewLine & _
        " Group By 医嘱ID, 收费细目ID" & vbNewLine & _
        " Having Max(ID) <> 0"
    
    '整张费用单据中剩余数不为0的行(明细到每一行)
    '执行状态原始记录上判断(部分退药且部分退费的记录)
    '*无医嘱执行计价的部分退费无法判断准退数量，不允许退费(执行状态调整为1)
    strSQL = _
        " Select Max(A.ID) as ID,Max(A.执行状态) as 执行状态," & _
        "        A.序号,A.收费类别,Sum(数量) as 剩余数量" & _
        " From (Select Decode(A.记录状态,2,0,A.ID) as ID," & _
        "              Decode(A.记录状态,2,0,Nvl(A.执行状态,0)) as 执行状态," & _
        "              A.序号,A.收费类别,Nvl(A.付数,1)*A.数次 as 数量" & _
        "       From 门诊费用记录 A" & _
        "       Where mod(A.记录性质,10)=[2] And A.NO=[1] And Nvl(A.附加标志,0)<>9" & _
        "       Union All" & _
        "       Select 0 As ID,1 As 执行状态,a.序号,a.收费类别,0 As 数量" & vbNewLine & _
        "       From 门诊费用记录 A" & vbNewLine & _
        "       Where A.记录性质 = [2] And A.NO=[1] And A.记录状态 In (1, 3) And Nvl(A.执行状态, 0) = 2" & vbNewLine & _
        "             And Not Exists(Select 1" & vbNewLine & _
        "                            From 病人医嘱发送 B, 医嘱执行计价 C" & vbNewLine & _
        "                            Where b.医嘱id = A.医嘱序号 And b.No = A.No" & vbNewLine & _
        "                                  And b.医嘱id = c.医嘱id And b.发送号 = c.发送号" & vbNewLine & _
        "                                  And c.收费细目id + 0 = A.收费细目id And b.记录性质 = [2])" & vbNewLine & _
        "             And Instr('5,6,7', A.收费类别) = 0" & vbNewLine & _
        "             And Not Exists(Select 1 From 材料特性 Where 材料id = A.收费细目id And Nvl(跟踪在用, 0) = 1)" & _
        "      ) A" & _
        " Group by A.序号,A.收费类别" & _
        " Having Nvl(Sum(数量),0)<>0"
                
    '有剩余数量无准退数量的有两种情况：
        '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应有剩余数量
        '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
    strSQL = _
        " Select A.序号 " & _
        " From(Select A.序号,A.剩余数量," & _
        "             Decode(A.执行状态,1,0,Nvl(B.准退数量,A.剩余数量) ) as 准退数量" & _
        "      From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
        "      Where A.ID=B.费用ID(+)" & _
        "     ) A" & _
        " Where Nvl(A.准退数量,0)<>A.剩余数量"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, _
        IIf(bytFlag = 2, "9", "8"), IIf(bytFlag = 2, "25", "24"))
    BillDeleteAllNew = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function BillExistDelete(ByVal strNo As String, ByVal int记录性质 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定单据是否包含(部分)退费或销帐的内容
    '入参:strNO-指定单据号
    '     int记录性质-记录性质
    '返回:存在退费的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-18 10:38:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select NO From 门诊费用记录 Where NO=[1] And 记录性质=[2] And 记录状态=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, int记录性质)
    BillExistDelete = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function BillExistMoney(ByVal strNos As String, ByVal int记录性质 As Integer, _
    Optional bln已收费 As Boolean, Optional lng打印ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定单据的项目是否已经全部退完(剩余数量=0)
    '入参:strNOs=可能为一张单据,也可能为多张单据(多单据收费产生的),格式为:"'AAA','BBB','CCC',..."
    '     int记录性质-记录性质
    '     lng打印ID-
    '出参:
    '返回: True=没有全部退完，表示部分退费
    '      False=已全部退完
    '编制:刘兴洪
    '日期:2014-06-18 10:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, strTmp As String
    Dim strWhere As String
    
    On Error GoTo errH
        
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    '实行多单据一次结算后,医保病人退费后会产生记录性质为11的记录,"NO,记录状态,执行状态,序号"有重复,AVG计算数量就有问题,所以要加上"记录性质"
    
    If lng打印ID > 0 Then
        '重临时表中取数
        
        strTmp = Replace(strNos, "'", "")
        If int记录性质 = 1 Then
            strWhere = "And Mod(A.记录性质,10)=[1]"
        Else
            strWhere = "And A.记录性质 =[1]"
        End If
        If bln已收费 Then strWhere = strWhere & " And A.记录状态<>0 "
        
        strSQL = _
        " Select NO,序号,Sum(数量) as 剩余数量" & _
        " From ( Select A.NO,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
        "               Avg(Nvl(A.付数, 1) * A.数次) As 数量" & _
        "       From 门诊费用记录 A,(select NO From 临时票据打印内容 where ID=[3] and 性质=[1]) J" & _
        "       Where Nvl(A.附加标志,0)<>9 And A.NO=J.NO " & strWhere & _
        "       Group by A.NO,A.记录性质,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号))" & _
        " Group by NO,序号 " & _
        " Having Sum(数量)<>0 "
        
          
          Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int记录性质, strTmp, lng打印ID)
    ElseIf UBound(Split(strNos, ",")) > 0 Then
        strTmp = Replace(strNos, "'", "")
        If int记录性质 = 1 Then
            strWhere = "And Mod(A.记录性质,10)=[1]"
        Else
            strWhere = "And A.记录性质 =[1]"
        End If
        If bln已收费 Then strWhere = strWhere & " And A.记录状态<>0 "
        
        strSQL = _
        " Select /*+ rule */  NO,序号,Sum(数量) as 剩余数量" & _
        " From ( Select A.NO,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号," & _
        "               Avg(Nvl(A.付数, 1) * A.数次) As 数量" & _
        "       From 门诊费用记录 A,(Select Column_Value From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) J " & _
        "       Where Nvl(A.附加标志,0)<>9 And A.NO=J.Column_Value " & strWhere & _
        "       Group by A.NO,A.记录性质,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号))" & _
        " Group by NO,序号 " & _
        " Having Sum(数量)<>0 "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int记录性质, strTmp)
    Else
        If int记录性质 = 1 Then
            strWhere = "And Mod(记录性质,10)=[2]"
        Else
            strWhere = "And 记录性质 =[2]"
        End If
        If bln已收费 Then strWhere = strWhere & " And 记录状态<>0 "
            
        strSQL = _
        " Select NO,序号,Sum(数量) as 剩余数量" & _
        " From ( Select NO,记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
        "               Avg(Nvl(付数, 1) * 数次) As 数量" & _
        "        From 门诊费用记录" & _
        "        Where Nvl(附加标志,0)<>9 And NO=[1] " & strWhere & _
        "        Group by NO,记录性质,记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by NO,序号 " & _
        " Having Sum(数量)<>0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", Replace(strNos, "'", ""), int记录性质)
    End If
    BillExistMoney = Not rsTemp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBillRows(strNo As String, bytFlag As Byte) As Integer
'功能：获取一张费用单据中未作废的费用行数
'参数：bytFlag=记录性质
'说明：用于退费/销帐时判断部份退费/销帐,退费时要排开误差费用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
    " Select 序号,Sum(数量) as 剩余数量" & _
    " From ( Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号," & _
    "               Avg(Nvl(付数, 1) * 数次) As 数量" & _
    "        From 门诊费用记录" & _
    "        Where Nvl(附加标志,0)<>9 And NO=[1] And 记录性质=[2]" & _
    "        Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
    " Group by 序号 " & _
    " Having Sum(数量)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional intPage As Integer, Optional intRow As Integer) As Boolean
'功能：费用单据药品处方限量检查,适用于记帐单,收费
'参数：intPage,intRow=指定对某张单据某行进行检查,否则为全部检查
'说明：
'   1.全部没超过限量，返回真；如有超过药品，则在函数内提示，并返回假。
'   2.记帐表是为每个病人单独检查
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim dblTime As Double, strItemIDs As String '已经检查过了的药品
    Dim dbl剂量 As Double, i As Integer, p As Integer
    Dim str药品限量提示 As String
    
    CheckLimit = True
    Err = 0: On Error GoTo errH:
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, objBill.Pages.Count, intPage)
        '收集病人
        strItemIDs = ""
        For i = 1 To objBill.Pages(p).Details.Count
            If intRow = 0 Or (intRow > 0 And i = intRow) Then
                With objBill.Pages(p).Details(i)
                    '收集药品ID
                    If InStr(strItemIDs & ",", "," & .收费细目ID & ",") = 0 And InStr(",5,6,7,", .收费类别) > 0 Then
                        strItemIDs = strItemIDs & "," & .收费细目ID
                    End If
                End With
            End If
        Next
        If strItemIDs <> "" Then
            strItemIDs = Mid(strItemIDs, 2)
            strSQL = "Select A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
                " From 药品规格 A,诊疗项目目录 B" & _
                " Where A.药名ID=B.ID And A.药品ID IN (" & strItemIDs & ")"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
            strItemIDs = ""
            For i = 1 To objBill.Pages(p).Details.Count
                If intRow = 0 Or (intRow > 0 And i = intRow) Then
                    Set tmpDetail = objBill.Pages(p).Details(i)
                    If InStr(",5,6,7,", tmpDetail.收费类别) > 0 And tmpDetail.Detail.处方限量 > 0 Then
                        If InStr(strItemIDs, "," & tmpDetail.收费细目ID) = 0 Then
                            dblTime = 0
                            For Each curDetail In objBill.Pages(p).Details
                                If InStr(",5,6,7,", curDetail.收费类别) > 0 And tmpDetail.收费细目ID = curDetail.收费细目ID Then
                                    dblTime = dblTime + curDetail.付数 * curDetail.数次
                                End If
                            Next
                            rsTmp.Filter = "药品ID=" & tmpDetail.收费细目ID
                            If Not rsTmp.EOF Then
                                If gbln药房单位 Then
                                    dbl剂量 = dblTime * tmpDetail.Detail.药房包装 * rsTmp!剂量系数
                                Else
                                    dbl剂量 = dblTime * rsTmp!剂量系数
                                End If
                                If dbl剂量 > tmpDetail.Detail.处方限量 Then
                                    str药品限量提示 = IIf(objBill.Pages.Count = 1, "", "单据" & p & "中") & "药品 """ & tmpDetail.Detail.名称 & """ 的总剂量 " & _
                                        FormatEx(dbl剂量, 5) & rsTmp!剂量单位 & "(" & FormatEx(dblTime, 5) & IIf(gbln药房单位, tmpDetail.Detail.药房单位, tmpDetail.Detail.计算单位) & ") 超过处方限量 " & _
                                        FormatEx(tmpDetail.Detail.处方限量, 5) & rsTmp!剂量单位 & " ！" & vbCrLf & vbCrLf & "是否允许?允许并且退出窗口之前不再进行处方限量检查请选择[取消]"
                                
                                        str药品限量提示 = MsgBox(str药品限量提示, vbYesNoCancel + vbDefaultButton3 + vbInformation, gstrSysName)
                                        If str药品限量提示 = vbNo Then
                                            CheckLimit = False: Exit Function
                                        End If
                                        If str药品限量提示 = vbCancel Then
                                           gbln处方限量 = True
                                        End If
                                End If
                            End If
                            strItemIDs = strItemIDs & "," & tmpDetail.收费细目ID
                        End If
                    End If
                End If
            Next
        End If
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get发药部门IDs(strNo As String, Optional strType As String) As String
'功能：获取收费单据的药房ID或卫材的发料部门ID
'返回："23,45,656..."
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errH
    If strType = "" Then   '药品
        strSQL = "Select Distinct 执行部门ID From 门诊费用记录" & _
        " Where 记录性质=1 And 记录状态 IN(0,1) And 收费类别 IN('5','6','7') And NO=[1]"
    Else                   '卫材
        strSQL = "Select Distinct a.执行部门ID From 门诊费用记录 a,材料特性 b" & _
        " Where a.收费细目id=b.材料id And b.跟踪在用=1 And " & _
        "(a.记录性质=1 And a.记录状态 In (0,1) Or a.记录性质=2 And a.记录状态=1) And a.收费类别='4' And a.NO=[1]"
        '记录性质=1 收费时,有直接收费或提划价单收费,所以a.记录状态 In (0,1)
        '记录性质=2 门诊记帐 仅记帐时允许,记录状态=1,记帐划价单的审核通过zl_门诊记帐记录_Verify处理自动发料
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!执行部门ID) Then
            Get发药部门IDs = Get发药部门IDs & "," & rsTmp!执行部门ID
        End If
        rsTmp.MoveNext
    Loop
    Get发药部门IDs = Mid(Get发药部门IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillRepeat(lng领用ID As Long, byt票种 As Byte, strFactNO As String) As Boolean
'功能：在使用新票号之前，检查是否重复
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 号码 From 票据使用明细" & _
        " Where 领用ID=[1] And 票种=[2] And 号码=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng领用ID, byt票种, strFactNO)
    CheckBillRepeat = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlCheckIsInvoiceListPrinted(ByVal strNo As String, Optional blnNOMoved As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否按明细打印的
    '入参:strNo-单据号
    '       blnNOMoved-是否在历史数据中
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-16 09:52:05
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, strPrintTable As String, strSQL1 As String
    
    On Error GoTo errHandle
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Function
    
    strPrintTable = "票据打印明细"
    '应根据最后一次打印的情况来定
    strSQL = "" & _
    "   Select NO  " & _
    "   From " & strPrintTable & " A, Table( f_Str2list([1])) J  " & _
    "   Where A.票种=1  And A.NO = J.Column_Value And Rownum=1"
    If blnNOMoved Then
        strSQL = Replace(strSQL, strPrintTable, "H" & strPrintTable)
        'strSql = strSql & " Union ALL " & strSQL1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp.EOF Then Exit Function
    zlCheckIsInvoiceListPrinted = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetMultiNOs(ByVal strNo As String, _
    Optional lng打印ID As Long, _
    Optional blnNOMoved As Boolean, _
    Optional bln按结算序号返回 As Boolean, _
    Optional bln历史表同步查 As Boolean = False) As String
    '功能：根据一张收费单据的NO，返回同一次打印的多张NO
    '参数：blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
    '          bln按结算序号返回-是否按结算序号返回
    '          bln历史表同步查-是否连接历史表一起查询
    '返回：格式如"'AAA','BBB','CCC',..."
    '      如果指定了"lng打印ID",则返回
    '说明：用于多单据收费
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long, strNos As String
    
    On Error GoTo errHandle
    lng打印ID = 0
    If bln按结算序号返回 Then
        strSQL = "Select Distinct A.NO,0 as ID" & vbNewLine & _
                " From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
                " Where A.结帐ID = B.结帐ID" & vbNewLine & _
                "       And B.结算序号 In (Select Max(A.结算序号)" & vbNewLine & _
                "                          From  病人预交记录 A, 门诊费用记录 B" & vbNewLine & _
                "                          Where A.结帐ID = B.结帐ID And Mod(b.记录性质, 10) =1 And B.记录状态<>2 And B.NO = [1] )"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
            strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        ElseIf bln历史表同步查 Then
            strSQL1 = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
            strSQL1 = Replace(strSQL1, "病人预交记录", "H病人预交记录")
            strSQL = strSQL & " Union ALL " & vbNewLine & strSQL1
        End If
    Else
        '应根据最后一次打印的情况来定
        strSQL = "Select ID, NO" & vbNewLine & _
                " From 票据打印内容" & vbNewLine & _
                " Where 数据性质 = 1" & vbNewLine & _
                "       And ID In (Select ID" & vbNewLine & _
                "                 From (Select b.Id" & vbNewLine & _
                "                       From 票据使用明细 A, 票据打印内容 B" & vbNewLine & _
                "                       Where a.打印id = b.Id And a.性质 = 1 And a.原因 In (1, 3) And b.数据性质 = 1 And b.No = [1]" & vbNewLine & _
                "                       Order By a.使用时间 Desc)" & vbNewLine & _
                "                 Where Rownum < 2)"
        If blnNOMoved Then
            strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
            strSQL = Replace(strSQL, "票据使用明细", "H票据使用明细")
        ElseIf bln历史表同步查 Then
            strSQL1 = Replace(strSQL, "票据打印内容", "H票据打印内容")
            strSQL1 = Replace(strSQL1, "票据使用明细", "H票据使用明细")
            strSQL = strSQL & " Union ALL " & vbNewLine & strSQL1
        End If
    End If
    strSQL = strSQL & vbNewLine & _
            " Order by NO"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    
    If Not rsTmp.EOF Then
        lng打印ID = Nvl(rsTmp!ID, 0) '可能没有
        For i = 1 To rsTmp.RecordCount
            strNos = strNos & ",'" & rsTmp!NO & "'"
            rsTmp.MoveNext
        Next
        GetMultiNOs = Mid(strNos, 2)
    Else
        GetMultiNOs = "'" & strNo & "'"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub GetErrorItem(ByRef lng误差细目id As Long, ByRef str收据费目 As String, _
    Optional ByVal strPricrGrade As String)
'功能：获取误差项的收费细目id,收据费目
'调用：收费时检查是否设了误差项,对导入划价单,计算票据张数
'说明：该项目不应撤档，也应为变价项目(未管)。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    Set rsTmp = zlGetSpecialItemFee("误差项", strPricrGrade)
    If Not rsTmp.EOF Then
        lng误差细目id = Val(Nvl(rsTmp!收费细目ID))
        str收据费目 = Nvl(rsTmp!收据费目)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function SeekPatiBill(ByVal lng病人ID As Long) As Long
'功能：根据病人搜寻病人的划价单据(系统指定天数内)
'返回：单据数量
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strDeptLimitWhere As String
    
    On Error GoTo errH
    '96357
    If gTy_Module_Para.str本机收费执行科室 <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Exists(Select 1" & vbNewLine & _
            "      From 门诊费用记录 M" & vbNewLine & _
            "      Where m.记录性质 = a.记录性质 And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str本机收费执行科室 & ",', ','||m.执行部门id||',') > 0)"
    ElseIf gTy_Module_Para.str已设置收费执行科室 <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Not Exists(Select 1" & vbNewLine & _
            "      From 门诊费用记录 M" & vbNewLine & _
            "      Where m.记录性质 = a.记录性质 And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str已设置收费执行科室 & ",', ','||m.执行部门id||',') > 0)"
    End If
    
    strSQL = "Select Count(a.NO) as 单据数 From 门诊费用记录 A" & vbNewLine & _
            " Where a.记录性质=1 And a.记录状态=0" & vbNewLine & _
            "       And a.划价人 is Not NULL And a.操作员姓名 IS NULL" & vbNewLine & _
            "       And a.病人ID=[1] And a.登记时间+0>=Sysdate-" & gintSeekDays & _
           strDeptLimitWhere
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID)
    If Not rsTmp.EOF Then SeekPatiBill = Nvl(rsTmp!单据数, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check药房上班安排() As Boolean
'功能：检查医院的药房是否使用了上班安排
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Count(B.部门ID) as NUM From 部门性质说明 A,部门安排 B" & _
        " Where A.部门ID=B.部门ID And A.工作性质 IN('西药房','成药房','中药房')"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlOutExse")
    If Not rsTmp.EOF Then
        Check药房上班安排 = Nvl(rsTmp!Num, 0) <> 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID( _
    ByVal str类别 As String, ByVal lng项目id As Long, ByVal int执行科室类型 As Integer, _
    ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, Optional ByVal int范围 As Integer = 1, _
    Optional ByVal lng西药房 As Long, Optional ByVal lng成药房 As Long, Optional ByVal lng中药房 As Long, _
    Optional ByVal lng执行科室ID As Long, Optional ByVal lng病人病区ID As Long) As Long
'功能：获取收费项目的执行科室
'参数：int范围=1.门诊,2-住院
'      lng执行科室ID=指定的缺省执行科室ID(用于药品和卫材)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte
    
    On Error GoTo errH
    
    If str类别 = "4" Then
        '有执行科室设置时
        strSQL = _
            " Select Distinct" & _
            "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
            " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
            "       And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2]  " & _
            "               Or Exists(select 1 From 病区科室对应 M where A.开单科室ID=M.病区ID And M.科室ID=[2] ))" & _
            "       And A.收费细目ID=[3]" & _
            " Order by B.服务对象,C.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", int范围, lng病人科室ID, lng项目id)
        If Not rsTmp.EOF Then
            Get收费执行科室ID = rsTmp!执行科室ID    '3:如果都没有，则返回第一个可用的执行科室(与医生站不同)
            
            '1:缺省为指定的(医嘱的)执行科室,不管是否服务于病人科室
            rsTmp.Filter = "执行科室ID=" & lng执行科室ID
            '2:其它可服务于病人科室的执行科室
            If rsTmp.EOF Then
                '2.1:尝试缺省为病人科室
                If lng执行科室ID <> lng病人科室ID Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人科室ID
                End If
                '2.2:尝试缺省为病人病区
                If rsTmp.EOF Then
                    If lng病人病区ID <> 0 And lng病人病区ID <> lng病人科室ID And lng病人病区ID <> lng执行科室ID Then
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人病区ID
                    End If
                End If
            End If
            '2.3:可服务于病人科室的一个执行科室
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If str类别 = "5" Then
            str药房 = "西药房": lng药房 = lng西药房
        ElseIf str类别 = "6" Then
            str药房 = "成药房": lng药房 = lng成药房
        ElseIf str类别 = "7" Then
            str药房 = "中药房": lng药房 = lng中药房
        End If
        
        '药品从系统指定的储备药房中找
        If Not gbln药房上班安排 Then
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                "       And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                "       And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
            strSQL = _
                " Select Distinct" & _
                "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                " Where A.执行科室ID+0=B.部门ID And B.工作性质=[1]" & _
                "       And B.服务对象 IN([2],3) And B.部门ID=C.ID" & _
                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                "       And D.部门ID=C.ID And D.星期=[5]" & _
                "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                "       And A.收费细目ID=[4]" & _
                " Order by B.服务对象,C.编码"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", str药房, int范围, lng病人科室ID, lng项目id, bytDay)
        If Not rsTmp.EOF Then
            Get收费执行科室ID = rsTmp!执行科室ID
            rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng执行科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0 And 执行科室ID=" & lng执行科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng药房
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0 And 执行科室ID=" & lng药房
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    Else
        Select Case int执行科室类型
            Case 0 '0-无明确科室
                Get收费执行科室ID = UserInfo.部门ID
            Case 1 '1-病人所在科室
                Get收费执行科室ID = lng病人科室ID
            Case 2 '2-病人所在病区
                If int范围 = 1 Then
                    Get收费执行科室ID = lng病人科室ID
                Else
                    Get收费执行科室ID = lng病人病区ID
                End If
            Case 3 '3-操作员所在科室
                Get收费执行科室ID = UserInfo.部门ID
            Case 4 '4-指定科室
                strSQL = "" & _
                "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
                "   From 收费执行科室  A,部门表 C" & _
                "   Where A.收费细目ID=[1]  And A.执行科室ID+0=C.ID  " & _
                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                "   Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng项目id, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                     '缺省取操作员所在科室
                    rsTmp.Filter = "开单科室ID=" & UserInfo.部门ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 6 '6-开单人所在科室
                Get收费执行科室ID = lng开单科室ID
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get费用摘要(ByVal strNo As String, ByVal bytFlag As Byte, ByVal int序号 As Integer) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定行数据的摘要
    '入参：strNO-单据号
    '      bytFlag-记录性质
    '      int序号-行号
    '出参：
    '返回：摘要
    '编制：刘兴洪
    '日期：2010-03-03 15:19:36
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 摘要 From 门诊费用记录" & _
        " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=[2] And Nvl(价格父号,序号)=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, int序号)
    If Not rsTmp.EOF Then Get费用摘要 = Nvl(rsTmp!摘要)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisAdviceMoney(ByVal strNo As String, ByVal bytFlag As Byte, _
    lng医嘱ID As Long, lng发送号 As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：判断一张单据是否医嘱的附加费用
    '入参：int记录性质=对应门诊费用记录.记录性质
    '出参：医嘱ID
    '      发送号
    '返回：是医嘱附费,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-03-03 15:22:06
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    lng医嘱ID = 0: lng发送号 = 0
    
    On Error GoTo errH
            
    strSQL = " Select 医嘱序号 From 门诊费用记录 Where Rownum=1 And 记录状态 IN(0,1,3) And NO=[1] And 记录性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    
    If Not rsTmp.EOF Then lng医嘱ID = Nvl(rsTmp!医嘱序号, 0)
    If lng医嘱ID <> 0 Then
    
        strSQL = "Select 发送号 From 病人医嘱附费  Where 医嘱ID=[3] And NO=[1] And 记录性质=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag, lng医嘱ID)
        
        If Not rsTmp.EOF Then lng发送号 = rsTmp!发送号
    End If
    
    If lng医嘱ID <> 0 And lng发送号 <> 0 Then
        BillisAdviceMoney = True
    Else
        lng医嘱ID = 0: lng发送号 = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistDrug(ByVal strNo As String, ByVal bytFlag As Byte) As Long
'功能：判断一张单据中是否存在未分配发药窗口的药品,有则回药房ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select 执行部门ID From 门诊费用记录" & _
    " Where 收费类别 IN('5','6','7') And NO=[1] And 记录状态 IN(0,1) And 记录性质=[2]" & _
    "       And 执行部门ID is Not NULL And 发药窗口 is NULL" & _
    " Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    If Not rsTmp.EOF Then
        BillExistDrug = rsTmp!执行部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillDrugDept(ByVal strNo As String, lng西药房 As Long, lng成药房 As Long, lng中药房 As Long) As Boolean
'功能：读取一张收费单据的药房ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "" & _
    " Select 收费类别,执行部门ID " & _
    " From 门诊费用记录" & _
    " Where 收费类别 IN('5','6','7') And NO=[1] And 记录状态 IN(0,1,3) And 记录性质=1" & _
    "       And 执行部门ID is Not NULL" & _
    " Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        If rsTmp!收费类别 = "5" Then
            lng西药房 = rsTmp!执行部门ID
        ElseIf rsTmp!收费类别 = "6" Then
            lng成药房 = rsTmp!执行部门ID
        ElseIf rsTmp!收费类别 = "7" Then
            lng中药房 = rsTmp!执行部门ID
        End If
    End If
    BillDrugDept = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Read个人帐户结算(ByVal lng结帐ID As Long) As Currency
'功能：读取指定结算记录中个人帐户支付的金额(冲预交方式不算)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 冲预交 From 病人预交记录 A,结算方式 B" & _
        " Where A.结算方式=B.名称 And B.性质=3" & _
        " And A.记录性质 Not IN(1,11) And A.结帐ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng结帐ID)
    If Not rsTmp.EOF Then
        Read个人帐户结算 = Nvl(rsTmp!冲预交, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Bill未收费(ByVal strNo As String, _
    ByVal bytFlag As Byte) As Boolean
    '功能：检查指定单据中是否存在未收费(未审核)的费用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select id From 门诊费用记录 Where NO=[1] And 记录性质=[2] And 记录状态=0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, bytFlag)
    Bill未收费 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillPay(ByVal strNo As String, _
    ByRef cur冲预交款 As Currency, ByRef cur应缴金额 As Currency) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费单据预交冲款合计,及缴款合计
    '入参:
    '出参:cur冲预交款-返回冲预交;cur应缴金额-返回应缴金额(现金的和非医保类（含三方卡等))
    '返回: 0(暂时未返回)
    '编制:刘兴洪
    '日期:2014-06-18 10:50:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "" & _
    " Select Sum(Decode(Mod(b.记录性质, 10), 1, b.冲预交, 0)) 冲预交," & vbNewLine & _
    "        Sum(Decode(b.记录性质,3,Decode(c.性质, 1, b.冲预交, 2, b.冲预交, 0),0)) 缴款金额" & vbNewLine & _
    " From 病人预交记录 B, 结算方式 C" & vbNewLine & _
    " Where b.结算方式 = c.名称" & vbNewLine & _
    "       And b.结帐id In (Select 结帐id From 门诊费用记录 Where Mod(记录性质, 10) = 1 And NO = [1])"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        cur冲预交款 = Val("" & rsTmp!冲预交)
        cur应缴金额 = Val("" & rsTmp!缴款金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceName(ByVal strNo As String) As String
'功能：获取收费单据原非医保结算方式名称
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "" & _
    " Select C.名称 " & _
    " From 门诊费用记录 A,病人预交记录 B,结算方式 C" & _
    " Where A.结帐ID=B.结帐ID And B.记录性质=3 And B.结算方式=C.名称" & _
    "       And Nvl(C.性质,1) IN(1,2) And A.记录性质=1 And A.记录状态 IN(1,3)" & _
    "       And A.NO=[1] And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetBalanceName = Nvl(rsTmp!名称)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDelBalanceID(ByVal strNo As String) As Long
'功能：获取退费记录的结帐ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 结帐ID From 门诊费用记录 Where NO=[1] And 记录性质=1 And 记录状态=2 And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then GetDelBalanceID = Val("" & rsTmp!结帐ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNo As String) As Boolean
'功能：判断指定的记帐单据中的状态是否一致,即是否同时存在审核和未审核的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
    " Select Count(Distinct 登记时间) as 时间数," & _
    "        Sum(Decode(记录状态,0,1,0)) as 未审核," & _
    "        Sum(Decode(记录状态,0,0,1)) as 已审核" & _
    " From 门诊费用记录" & _
    " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!未审核, 0) <> 0 And Nvl(rsTmp!已审核, 0) <> 0 Then
            BillIdentical = False
        ElseIf Nvl(rsTmp!时间数, 0) > 1 Then
            BillIdentical = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function AuditingWarn(ByVal strPrivs As String, ByVal strNo As String, ByVal str序号 As String) As Boolean
'功能：审核划价单时，对费用进行报警
'参数：str序号=指定单据中要审核的行号,为空表示所有行
    Dim rsWarn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, j As Long, str类别s As String
    Dim cur当日额 As Currency, cur金额 As Currency, cur余额 As Currency
    Dim strWarn As String, intWarn As Integer
    
    strSQL = "" & _
    " Select A.门诊标志, A.姓名, A.病人id , E.预交余额 - E.费用余额 As 余额, B.担保额, C.编码 As 付款码," & vbNewLine & _
    "        A.收费类别, D.名称 As 类别名称, Sum(A.实收金额) As 金额, Zl_Patiwarnscheme(A.病人id) As 适用病人" & vbNewLine & _
    " From 门诊费用记录 A, 病人信息 B, 医疗付款方式 C, 收费项目类别 D, 病人余额 E" & vbNewLine & _
    " Where A.记录性质 = 2 And A.记录状态 = 0 And A.NO = [1] And A.收费类别 = D.编码 And A.病人id = E.病人id(+) And" & vbNewLine & _
    "       E.性质(+) = 1 And A.病人id = B.病人id And B.医疗付款方式 = C.名称(+)" & vbNewLine & _
            IIf(str序号 <> "", " And Instr([2],','||Nvl(A.价格父号,A.序号)||',')>0", "") & _
    " Group By Nvl(A.价格父号, A.序号), A.门诊标志, A.姓名, A.病人id,  B.担保额, E.预交余额, E.费用余额, C.编码," & vbNewLine & _
    "         A.收费类别, D.名称"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo, "," & str序号 & ",")
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If InStr(str类别s, rsTmp!收费类别 & rsTmp!类别名称) = 0 Then
                str类别s = str类别s & "," & rsTmp!收费类别 & rsTmp!类别名称
            End If
            cur金额 = cur金额 + rsTmp!金额
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
        str类别s = Mid(str类别s, 2)
        
        If cur金额 > 0 Then
            Set rsWarn = GetUnitWarn(rsTmp!适用病人, "0")
                        
            cur当日额 = GetPatiDayMoney(rsTmp!病人ID)
            cur余额 = Nvl(rsTmp!余额, 0)
            If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(0, rsTmp!病人ID) + cur金额
            '分类报警
            For j = 0 To UBound(Split(str类别s, ","))
                intWarn = BillingWarn(strPrivs, rsTmp!姓名, rsTmp!适用病人, rsWarn, _
                    cur余额, cur当日额, cur金额, Nvl(rsTmp!担保额, 0), _
                    Left(Split(str类别s, ",")(j), 1), Mid(Split(str类别s, ",")(j), 2), strWarn)
                If intWarn = 2 Or intWarn = 3 Then Exit Function
            Next
        End If
    End If
    
    AuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'功能：检查卫生材料的灭菌效期是否过期
'说明：blnAsk=表示是否询问是否继续,否则为提醒
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, minDate As Date
    Dim strSQL As String, strName As String
    
    CheckValidity = True
    
    '仅一次性材料才判断
    '因为可能各批次灭菌效期不同,检查要用到的批次中最小的效期
    strSQL = _
        " Select C.名称,Nvl(B.批次,0) as 批次," & _
        " B.可用数量 as 库存,B.灭菌效期,Sysdate as 时间" & _
        " From 材料特性 A,药品库存 B,收费项目目录 C" & _
        " Where A.材料ID=B.药品ID And A.材料ID=C.ID And A.一次性材料=1" & _
        " And B.性质=1 And Nvl(B.可用数量,0)>0 And A.灭菌效期 is Not NULL" & _
        " And A.材料ID=[1] And B.库房ID=[2]" & _
        " Order by Nvl(B.批次,0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng材料ID, lng库房ID)
    If Not rsTmp.EOF Then
        strName = rsTmp!名称
        Curdate = rsTmp!时间
        minDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!灭菌效期 < minDate Then
                minDate = rsTmp!灭菌效期
            End If
            If Nvl(rsTmp!库存, 0) < dbl数量 Then
                dbl数量 = dbl数量 - Nvl(rsTmp!库存, 0)
            Else
                dbl数量 = 0
            End If
            If dbl数量 = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If Curdate > minDate Then
            If blnAsk Then
                If MsgBox("卫生材料""" & strName & """的灭菌效期""" & Format(minDate, "yyyy-MM-dd") & """已过期,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "提醒：" & vbCrLf & vbCrLf & "卫生材料""" & strName & """的灭菌效期""" & Format(minDate, "yyyy-MM-dd") & """已过期。", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistRegist(ByVal lng病人ID As Long, Optional ByRef blnPeisPriceBill As Boolean, _
    Optional ByVal blnSaveBillCheck As Boolean) As Boolean
    '功能：判断指定病人是存在有效的挂号记录
    '入参:
    '   blnSaveBillCheck - 是否保存时检查，体检检查
    '出参：
    '   blnPeisPriceBill-该病人是否存在体检划价单
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    '天数无限制时表示不检查,不然很慢
    If gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 And gTy_System_Para.Sy_Reg.bytNoDayseMergency = 0 Then ExistRegist = True: Exit Function
    
    If blnSaveBillCheck = False Then
        '体检病人不限制是否挂号
        '102660,判断是否体检病人，不再通过“病人医嘱记录"来判断是否体检病人了，主要根据划价单中是否包含体检费用来判断
        strSQL = "Select 1" & vbNewLine & _
                " From 门诊费用记录" & vbNewLine & _
                " Where 记录性质 = 1 And 记录状态 = 0 And Nvl(门诊标志, 0) = 4 And 病人id = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID)
        If Not rsTmp.EOF Then
            blnPeisPriceBill = True
            ExistRegist = True: Exit Function
        End If
    End If
    
    strSQL = "Select 1 From 病人挂号记录  " & _
            " Where RowNum<2 And 病人ID=[1] and 记录性质=1 and 记录状态=1 " & zlGetRegEventsCons(, , True)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID)
    ExistRegist = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckRegisted(ByVal lng病人ID As Long, Optional ByRef blnPeisPriceBill As Boolean, _
    Optional ByVal blnSaveBillCheck As Boolean) As Boolean
'功能:收费时,检查病人是否挂号
'入参:
'   blnSaveBillCheck - 是否保存时检查，体检检查
'出参：
'   blnPeisPriceBill-该病人是否存在体检划价单
    blnPeisPriceBill = False
    
    If gbytUnRegevent = 0 Then CheckRegisted = True: Exit Function
    If lng病人ID <> 0 Then
        If Not ExistRegist(lng病人ID, blnPeisPriceBill, blnSaveBillCheck) Then lng病人ID = 0
    End If
           
    If lng病人ID = 0 Then
        If gbytUnRegevent = 1 Then
            If MsgBox("病人没有挂号,你确定要进行收费吗?", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                CheckRegisted = False: Exit Function
            End If
        Else
            Call MsgBox("病人没有挂号,不允许进行收费.", vbInformation, gstrSysName)
            CheckRegisted = False: Exit Function
        End If
    End If
    CheckRegisted = True
End Function


Public Function CheckAddedItem(ByVal lng病人ID As Long, Optional str病人姓名 As String) As Boolean
'功能：判断指定病人是否需要自动加收指定收费项目
'      不进行自动加收的条件:
'       1.没有设定挂号单有效天数(为零)
'       2.没有设定自动加收项目,或该项目ID已不存在
'       3.挂号单有效天数内,没有挂号,没有自动加收
'       4.本地参数设置了不允许输入“其它”类别
    Dim rsTmp As ADODB.Recordset, strSQL As String, strID As String
    Dim strWhere As String
    On Error GoTo errH
    If (gTy_System_Para.Sy_Reg.bytNODaysGeneral = 0 And gTy_System_Para.Sy_Reg.bytNoDayseMergency) Or glngAddedItem = 0 Then Exit Function
    If gstr收费类别 <> "" And InStr(1, "," & gstr收费类别 & ",", ",'Z',") = 0 Then Exit Function
    
    If lng病人ID = 0 Then
        strID = "And 姓名 = [2]"
    Else
        strID = "And 病人id + 0 = [1]"
    End If
    
    strWhere = "   (nvl(急诊,0) =0 and 发生时间 > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & ")) "
    strWhere = strWhere & "  Or  (nvl(急诊,0) =1 and 发生时间 > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNoDayseMergency & ")) "
    strWhere = " And (" & strWhere & ") "
    
    strSQL = "" & _
    " Select 1" & vbNewLine & _
    " From Dual" & vbNewLine & _
    " Where     Not Exists (Select 1 From 病人挂号记录 Where 记录性质=1 and 记录状态=1 and  Rownum < 2 " & strID & strWhere & ")" & vbNewLine & _
    "       And Not Exists (Select 1" & vbNewLine & _
    "                       From 门诊费用记录" & vbNewLine & _
    "                       Where Rownum < 2 " & strID & " And 登记时间 > Trunc(Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & ") " & vbNewLine & _
    "                             And 收费细目id + 0 = [3] And 记录性质 = 1 And 记录状态 = 1) " & _
    "       And Exists (Select 1 From 收费项目目录 Where ID = [3])"
    
    '刘兴洪 问题:34717    日期:2010-12-20 16:09:26
    '由于门诊费用中无法区分是否已经加收了指定的收费项目,因此,还是普通挂号有效天数为准
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng病人ID, str病人姓名, glngAddedItem)
    CheckAddedItem = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ItemUnderSet(ByVal str类别 As String, ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal dbl库存 As Double) As Boolean
'功能：检查指定药品/材料在指定库房的库存是否低于储备下限
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '无记录或字段为零/空表示未设置
    If str类别 = "4" Then
        strSQL = "Select 库房ID,材料ID,上限,下限,盘点属性,库房货位 From 材料储备限额 Where 材料ID=[1] And 库房ID=[2] And Nvl(下限,0)<>0 And 下限>[3]"
    Else
        strSQL = "Select 库房ID,药品ID,上限,下限,盘点属性,库房货位 From 药品储备限额 Where 药品ID=[1] And 库房ID=[2] And Nvl(下限,0)<>0 And 下限>[3]"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng药品ID, lng库房ID, dbl库存)
    If Not rsTmp.EOF Then ItemUnderSet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistFact(strNo As String) As Boolean
'功能：判断指定单据中是否存在工本费
'参数：strNO=多张单据中的一张
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    '应取最后一次打印的最大号码
    strSQL = "Select Max(ID) From 票据打印内容 Where 数据性质=1 And NO=[1]"
    strSQL = "Select Count(A.ID) as NUM From 门诊费用记录 A,票据打印内容 B" & _
        " Where A.NO=B.NO And A.记录性质=1 And A.记录状态=1 And Nvl(A.附加标志,0)=8 And B.数据性质=1 And B.ID=(" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If Not rsTmp.EOF Then
        BillExistFact = Nvl(rsTmp!Num, 0) > 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckSingleBalance(ByVal strNo As String) As Boolean
'功能：判断指定单据中是否只有一种非医保结算方式(冲预交除外)
'       :strNO(格式为"E01,E02"):问题:34035
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNo = Replace(strNo, "'", "")
    CheckSingleBalance = True
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.结算方式) num" & vbNewLine & _
    " From 病人预交记录 A, 结算方式 B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.记录性质 = 3 And A.记录状态 In (1, 3) " & _
    "           And A.结算方式 = B.名称 And B.性质 In (1, 2)  And A.NO = J.Column_Value"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNo)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckTest(ByVal strNos As String, Optional ByVal DatBegin As Date, Optional ByVal DatEnd As Date) As Boolean
    '功能：检查当前选择要收费的单据中是否存在皮试结果为阳性或还没有皮试的记录
    '参数：strNOs=单据号(格式为"E01,E02")',多张单据时,须传入DatBegin和DatEnd参数
    '返回：为False时表示不允许进行收费
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strIF As String
    Dim i As Long, strInfo As String, strTmp As String, blnHaveMateria As Boolean
    Dim str诊疗项目ID As String, rsTest As ADODB.Recordset
    Dim strTable As String
    CheckTest = True
    
    '62020:有单据，就不应该存在日期范围的查找
'    If InStr(1, strNos, ",") > 0 Then
'        strIF = " And A.登记时间 Between [2] And [3] And Instr(','||[1]||',',','||A.NO||',')>0"
'    Else
'        strIF = " And A.NO = [1]"
'    End If
'
    strSQL = _
    " Select /*+ rule */ Distinct B.ID,B.医嘱内容,B.皮试结果,a.收费类别,B.诊疗项目ID" & _
    " From 门诊费用记录 A,病人医嘱记录 B,诊疗项目目录 C,Table(f_Str2list([1])) J" & _
    " Where A.记录性质=1 And A.记录状态=0  And A.划价人 IS Not NULL And A.操作员姓名 IS NULL" & _
    "       And A.NO = J.Column_Value " & _
    "       And A.医嘱序号=B.ID And B.诊疗项目ID=C.ID And C.类别='E' And C.操作类型='1'" & _
    "       And (B.皮试结果 is NULL Or B.皮试结果='(+)')"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNos, DatBegin, DatEnd)
    If Not rsTmp.EOF Then
        '问题:33110
        str诊疗项目ID = ""
        For i = 1 To rsTmp.RecordCount
            If InStr(1, str诊疗项目ID & ",", "," & Nvl(rsTmp!诊疗项目ID) & ",") = 0 Then
                    str诊疗项目ID = str诊疗项目ID & "," & Nvl(rsTmp!诊疗项目ID)
            End If
            rsTmp.MoveNext
        Next
        If Len(str诊疗项目ID) > 2000 Then
            strTable = " (Select distinct B.项目ID as 诊疗ID,B.用法ID From 诊疗项目目录 A,诊疗用法用量 B Where A.ID=B.用法ID and nvl(B.性质,0)=0 and  A.ID In (" & Mid(str诊疗项目ID, 2) & ") ) J"
        ElseIf str诊疗项目ID <> "" Then
            str诊疗项目ID = Mid(str诊疗项目ID, 2)
            strTable = " (Select distinct B.项目ID as 诊疗ID,B.用法ID From Table(Cast(f_num2list([4]) As Zltools.t_Numlist )) A ,诊疗用法用量 B Where a.Column_Value=B.用法ID and nvl(B.性质,0)=0) J"
        Else
            strTable = " (Select -1 as 诊疗ID,0 as 用法ID From dual) J"
        End If
        strSQL = "" & _
        "   Select /*+ rule */ distinct  M.编码,M.名称,M.规格,J.用法ID,J.诊疗ID" & _
        "   From 门诊费用记录 A,收费项目目录 M,药品规格 B,诊疗项目目录 C,Table(f_Str2list([1])) M ," & _
                    strTable & _
        "   Where A.记录性质=1 And A.记录状态=0 And  A.划价人 IS Not NULL And A.操作员姓名 IS NULL " & _
        "              And A.收费类别 In('5','6','7')  And A.NO = M.Column_Value" & _
        "              And A.收费细目ID =M.ID " & _
        "              And a.收费细目ID=b.药品ID and B.药名ID=C.ID And  B.药名ID=J.诊疗ID "
        Set rsTest = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNos, DatBegin, DatEnd, str诊疗项目ID)
        If rsTest.RecordCount > 0 Then blnHaveMateria = True
        
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            strTmp = rsTmp!医嘱内容 & "：" & IIf(IsNull(rsTmp!皮试结果), "无皮试结果", "结果为阳性(+)")
            If InStr(1, strInfo, strTmp) = 0 Then
                strInfo = strInfo & vbCrLf & strTmp & "　　"
                rsTest.Filter = "用法ID=" & Val(Nvl(rsTmp!诊疗项目ID))
                If Not rsTest.EOF Then
                    strInfo = strInfo & "药品:" & Nvl(rsTest!名称) & "　　"
                End If
            End If
            rsTmp.MoveNext
        Next
        If blnHaveMateria Then
            strInfo = "收费单据中，以下皮试无结果或为阳性：" & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许收费!"
            MsgBox strInfo, vbInformation, gstrSysName
        Else
            strInfo = "收费单据中，以下皮试无结果或为阳性：" & vbCrLf & strInfo & vbCrLf & vbCrLf & " 是否继续收费?"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckTest = False
                Exit Function
            End If
        End If
        CheckTest = Not blnHaveMateria
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetPriceBills(ByVal lngPatient As Long, ByVal lngRegDept As Long, _
                ByVal DatBegin As Date, ByVal DatEnd As Date, _
                Optional blnAddDiagnose As Boolean = False, _
                Optional bytDefaultSel As Byte = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人指定划价时间范围内的划价单
    '入参: lngPatient-病人ID
    '        lngRegDept-挂号科室,通过划价单输入,并且参数要求检查挂号科室时,才会传入挂号科室
    '        blnAddDiagnose-加入诊断(问题:33685)
    '        bytDefaultSel-缺省选择(0-当日;1-有效天数;2-所有)
    '返回:划价单记录集
    '编制:刘兴洪
    '日期:2011-03-14 10:40:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strSelected As String, strIF As String
    Dim strSubTable As String, strTbDiagnose As String
    Dim blnDeptLimit As Boolean
    Dim strDeptLimitWhere As String
    
    On Error GoTo errH
    '96357
    If gTy_Module_Para.str本机收费执行科室 <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Exists(Select 1" & vbNewLine & _
            "      From 门诊费用记录 M" & vbNewLine & _
            "      Where m.记录性质 = a.记录性质 And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str本机收费执行科室 & ",', ','||m.执行部门id||',') > 0)"
    ElseIf gTy_Module_Para.str已设置收费执行科室 <> "" Then
        strDeptLimitWhere = vbNewLine & _
            " And Not Exists(Select 1" & vbNewLine & _
            "      From 门诊费用记录 M" & vbNewLine & _
            "      Where m.记录性质 = a.记录性质 And m.No = a.No" & vbNewLine & _
            "           And Instr('," & gTy_Module_Para.str已设置收费执行科室 & ",', ','||m.执行部门id||',') > 0)"
    End If
    
    'decode(A.执行状态,-1,NULL,'√'):问题:38281
    '上次未选择的,执行状态为-1表示暂停执行
    If lngRegDept > 0 Then
        strSelected = "Decode(A.开单部门ID," & lngRegDept & ",decode(A.执行状态,-1,NULL,'√'),NULL)"
    Else
        If bytDefaultSel = 0 Then
            '36438:gintSeekDays
            strSelected = "Decode(Sign(Trunc(Sysdate)-Trunc(A.登记时间)),0,decode(A.执行状态,-1,NULL,'√'),-1,decode(A.执行状态,-1,NULL,'√'),NULL)"
        ElseIf bytDefaultSel = 1 Then   '缺省有效天数
            '36438:gintSeekDays
            strSelected = "Decode(Sign( Sysdate-" & gintSeekDays & "-A.登记时间),0,decode(A.执行状态,-1,NULL,'√'),-1,decode(A.执行状态,-1,NULL,'√'),NULL)"
        Else
            strSelected = " decode(A.执行状态,-1,NULL,'√') "
        End If
    End If
    strIF = " And A.病人ID=[1] And A.登记时间 Between [2] And [3] "
    strTbDiagnose = ""
    '将:Wmsys.Wm_Concat改为了f_List2Str(Cast(collect ()))的方式.原因是oracle10g目前只是测试版
    '问题:38528
    If blnAddDiagnose Then '加入医嘱
        strTbDiagnose = "" & _
        "         ,( Select distinct A.NO,  f_List2str(Cast(COLLECT(distinct Q.诊断描述 ) as t_Strlist))  as 诊断" & vbNewLine & _
        "           From A  ,病人诊断医嘱 J,病人诊断记录 Q  " & vbNewLine & _
        "           Where  A.医嘱序号=J.医嘱ID and J.诊断ID=Q.ID " & vbNewLine & _
        "           Group by  A.NO  ) C"
    End If
    strSubTable = "" & _
                " Select " & strSelected & " as 选择 ,A.开单部门ID, " & vbNewLine & _
                "       A.NO,A.开单人,A.姓名,A.性别,A.年龄,A.应收金额,A.实收金额," & vbNewLine & _
                "       A.划价人,A.登记时间  As 划价时间,nvl(B.相关ID,A.医嘱序号 ) as 医嘱序号, " & vbNewLine & _
                "       decode(C.ID,NULL,0,1) as 皮试" & vbNewLine & _
                " From 门诊费用记录 A,病人医嘱记录 B,诊疗项目目录 C" & vbNewLine & _
                " Where A.记录性质=1 And A.记录状态=0 And A.医嘱序号=B.ID(+)" & vbNewLine & _
                "       And  B.诊疗项目ID=C.ID(+) And C.类别(+)='E' And C.操作类型(+)='1' " & strIF & vbNewLine & _
                strDeptLimitWhere
   
   strSQL = _
        " Select * From ( with A as ( " & strSubTable & ")   " & vbNewLine & _
        " Select " & IIf(blnAddDiagnose, "       nvl(Max(C.诊断),' ') as 诊断,", "") & vbNewLine & _
        "       A.选择,A.NO as 单据号,B.名称 as 开单科室,A.开单人 as 医生,Ltrim(A.姓名) as 姓名,A.性别,A.年龄," & vbNewLine & _
        "       ltrim(To_Char(Sum(A.应收金额),'99999" & gstrDec & "')) as 应收金额," & vbNewLine & _
        "       ltrim(To_Char(Sum(A.实收金额),'99999" & gstrDec & "')) as 实收金额," & vbNewLine & _
        "       A.划价人,To_Char(A.划价时间,'YYYY-MM-DD HH24:MI:SS') as 划价时间, " & vbNewLine & _
        "       Decode(nvl(Max(A.皮试),0),1,'√','') as 皮试" & vbNewLine & _
        " From A,部门表 B" & vbCrLf & strTbDiagnose & vbNewLine & _
        " Where  A.开单部门ID=B.ID  " & IIf(blnAddDiagnose, " And A.NO=C.NO(+)", "") & vbNewLine & _
        " Group by A.选择,A.NO,B.名称,A.开单人,A.姓名,A.性别,A.年龄,A.划价人,A.划价时间,A.开单部门ID) " & vbNewLine & _
        " Order by 单据号 Desc"
        '" Order by  划价时间 Desc"'102748,划价时间降序调整为按单据号降序
    Set GetPriceBills = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngPatient, DatBegin, DatEnd)
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReadABCNum(ByVal strPrivs As String) As Boolean
'功能:获取中药输入快捷
'参数：strPrivs=用于根据权限控制是否读取小数快捷部份
'返回：可以输入的快捷字母
    Dim strSQL As String
        
    On Error GoTo errH
    
    If InStr(strPrivs, "药品输入小数") > 0 Then
        strSQL = "Select Upper(名称) as 名称,数值 From 中药输入快捷 Order by 名称"
    Else
        strSQL = "Select Upper(名称) as 名称,数值 From 中药输入快捷 Where Trunc(数值)=数值 Order by 名称"
    End If
    
    Set grsABCNum = New ADODB.Recordset 'Filter在New时清除
    Call zlDatabase.OpenRecordset(grsABCNum, strSQL, "mdlPublic")
    
    '获取可以输入的快捷字母
    gstrABC = ""
    Do While Not grsABCNum.EOF
        '快捷字母只有一位,且只能为字母
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(grsABCNum!名称, 1)) > 0 Then
            gstrABC = gstrABC & Left(grsABCNum!名称, 1)
        End If
        grsABCNum.MoveNext
    Loop
    ReadABCNum = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ConvertABCtoNUM(ByVal strInput As String) As String
'功能：根据中药输入快捷定义，将输入转换为数字
'规则：1.快捷字母前可以输入正负号
'      2.不能同时输入多个快捷字母
'      3.快捷字母不能与数字,小数点混合输
'      4.不满足以上规则则返回0
    Dim strBit As String, strNum As String, i As Long
    Dim blnABC As Boolean, blnNum As Boolean
    
    If strInput = "" Then ConvertABCtoNUM = "": Exit Function
    strInput = UCase(strInput)
    
    For i = 1 To Len(strInput)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(strInput, i, 1)) > 0 Then
            If blnNum Or blnABC Then strNum = "0": Exit For
            
            grsABCNum.Filter = "名称='" & Mid(strInput, i, 1) & "'"
            If Not grsABCNum.EOF Then
                strBit = FormatEx(grsABCNum!数值, 5)
            Else
                strNum = "0": Exit For
            End If
            
            blnABC = True
        ElseIf InStr("0123456789.", Mid(strInput, i, 1)) > 0 Then
            If blnABC Then strNum = "0": Exit For
            strBit = Mid(strInput, i, 1)
            blnNum = True
        Else
            strBit = Mid(strInput, i, 1)
        End If
        strNum = strNum & strBit
    Next
    ConvertABCtoNUM = strNum
End Function

Public Function CheckDeptIsMedTech(ByVal lngDeptID As Long) As Boolean
'功能：检查指定的部门是否是医技或体检性质
'参数:
    Dim rsTmp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    
    strSQL = "Select 1 From 部门性质说明 Where 部门id = [1] And 工作性质 In('检查','检验','手术','治疗','营养','体检')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
    CheckDeptIsMedTech = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMediCareItem(ByVal lng收费细目ID As Long, ByVal int险类 As Integer, _
    ByVal str收费项目名称 As String, ByVal bln定价 As Boolean, Optional ByVal strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断收费项目是否设置了保险支付项目
    '入参:bln定价-当前是否为定价
    '     dbl价格-定价的价格
    '出参:
    '返回:1.对码的返回true,否则返回False
    '     2.不检查医保对码,返回true
    '     3.定价的且价格=0的不检查
    '编制:刘兴洪
    '日期:2010-01-07 14:44:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset, strSQL As String, rs价格 As ADODB.Recordset, dbl价格 As Double
    Dim strWherePriceGrade As String
        
    CheckMediCareItem = True
    If gbyt医保对码检查 = 0 Then Exit Function
    On Error GoTo errH
    
    '刘兴洪 问题:27286 定价的价格为零的不进行检查对码 日期:2010-01-07 15:13:45
    If bln定价 Then
        If strPriceGrade <> "" Then
            strWherePriceGrade = _
                "      And (b.价格等级 = [2]" & vbNewLine & _
                "          Or (b.价格等级 Is Null" & vbNewLine & _
                "              And Not Exists(Select 1" & vbNewLine & _
                "                             From 收费价目" & vbNewLine & _
                "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
                "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
        Else
            strWherePriceGrade = " And b.价格等级 Is Null"
        End If
        strSQL = _
            "Select b.现价" & vbNewLine & _
            "From 收费价目 B" & vbNewLine & _
            "Where b.收费细目id = [1]" & vbNewLine & _
            "      And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        Set rs价格 = zlDatabase.OpenSQLRecord(strSQL, "获取当前价格", lng收费细目ID, strPriceGrade)
        If rs价格.EOF = False Then
            dbl价格 = Val(Nvl(rs价格!现价))
        Else
            dbl价格 = 0
        End If
        If dbl价格 = 0 Then Exit Function
    End If
     
    strSQL = "Select 收费细目ID From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng收费细目ID, int险类)
        
    If rsTmp.RecordCount = 0 Then
        If gbyt医保对码检查 = 1 Then
            If MsgBox("没有设置""" & str收费项目名称 & """对应的保险项目,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbyt医保对码检查 = 2 Then
            MsgBox "没有设置""" & str收费项目名称 & """对应的保险项目!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional bln未找到增加 As Boolean = False, Optional strOra过程 As String, Optional strWhere As String, _
    Optional bln站点 As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '     bln站点-是否进行站点限制
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str编码 As String, str名称 As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
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
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    str名称 = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  "
    End If
    gstrSQL = gstrSQL & strWhere & IIf(bln站点, zl_获取站点限制, "") & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If bln未找到增加 Then
            If zlCommFun.IsCharChinese(str名称) = False Then GoTo NOAdd::
            If MsgBox("注意:" & vbCrLf & _
                   "     未找到相关的" & strTable & ",是否增加“" & str名称 & "”？", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str编码, str名称, strTable & "增加", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str名称
                    End If
                End With
            Else
                If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                objCtl.Tag = str名称
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgbox "没有找到满足条件的" & strTable & ",请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!名称)
            Else
                .Text = IIf(blnOnlyName, Nvl(rsTemp!名称), Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称))
            End If
        End With
    Else
        If zlControl.IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!名称)
        objCtl.Tag = Nvl(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str编码 As String, str名称 As String, _
    Optional strTittle As String = "增加项目", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加项目信息(只针对有编码,名称的信息增加(只增加：编码和名称,简码)
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int编码 As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的" & strTable & "，你要把它加入" & strTable & "中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int编码 = rsTemp!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str名称)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str编码 = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 
Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '返回:返回加匹配串%dd%,并且是大写
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False, Optional blnNotTran As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNotTran-不处理事务
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNotTran = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlIsExistsSquareCard(ByVal strNos As String, Optional bln存在全退 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否为卡结算单据
    '入参:strNos-单据号(可以为多张,用逗号分离)
    '        bln存在全退-true:表示只检查是否存在全退的单存;False-只检查刷卡记录
    '出参:
    '返回:存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    Dim intType As Integer
    On Error GoTo errHandle
    intType = -1
    If bln存在全退 Then intType = 0
    
    '55064
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "Select /*+ rule */  B.ID  " & _
    " From 病人预交记录 B, Table(f_Str2list([1])) J " & _
    " Where B.NO = J.Column_Value And  " & _
    "       (  (Nvl(B.结算卡序号, 0) <> 0 And Exists(Select 1 From 消费卡类别目录 Where 编号=nvl(B.结算卡序号,0) And nvl(是否全退,0)<>[2] And Nvl(是否退现,0)=0 ) ) " & _
    "          Or (Nvl(B.卡类别id, 0) <> 0 And Exists(Select 1 From 医疗卡类别 Where ID=nvl(B.卡类别ID,0) And nvl(是否全退,0)<>[2]  And nvl(是否退现,0)=0) )  " & _
    "       ) And B.记录性质 = 3 And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查收费单是否存在刷卡记录", strNoIns, intType)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsOnly北京医保() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否仅为有北京医保
    '返回:是的,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-12 09:42:08
    '问题:27331
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrTmp As Variant, i As Long, strTemp As String
    arrTmp = Split(GetSetting("ZLSOFT", "公共全局", "本地支持的医保", ""), ",")
    strTemp = ""
    For i = 0 To UBound(arrTmp)
        If IsNumeric(arrTmp(i)) Then
            strTemp = strTemp & "," & Val(arrTmp(i))
            'If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For '检查外挂模式
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlIsOnly北京医保 = strTemp = "920"  '见问题:问题:26982
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
    On Error GoTo errHandle

    If mlng部门编码平均长度 = 0 Then
        strSQL = "Select Avg(length(编码)) As 长度 From 部门表"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取部门编码的平均长度")
        mlng部门编码平均长度 = Val(Nvl(rsTemp!长度))
    End If
    '由于编码长度可能过长,无法显示部门的名称,因此自动显示和不显示编码,当大于5时,不显示.小于5时,显示
   zlIsShowDeptCode = mlng部门编码平均长度 <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
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
    Dim strIDs As String, str简码 As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
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
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省,服务对象," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:人员选择选择器
    '入参:cboSel-指定的部门选择部件
    '     rsPerson-指定的人员信息(ID,编号,姓名,简码)
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有-所有名称(所有人,所有操作员等)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String, strLike As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编号) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!姓名)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编号)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!姓名))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboSel
End Function
Public Function zlCheckIsPrintInvoice(ByVal strNos As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查票据是否存在打印情况
    '入参：strNOs       =   指定要重打的单据号，带引号，可能是多个单据号，为"'AAA','BBB',..."的形式
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-05-27 22:04:21
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim cllPro As Collection, varData() As Variant
    
    On Error GoTo errHandle
    strNos = Replace(strNos, "'", "")
    
    If Len(strNos) <= 4000 Then
        strSQL = "" & _
        "   Select /*+ rule */ Max(A.ID) as ID " & _
        "   From 票据打印内容 A,Table( f_Str2list([1])) J " & _
        "   Where A.数据性质=1   And A.NO=J.Column_Value"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos)
    Else
        If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
        If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
        strSQL = "With 单据信息 as (" & strSQL & ")" & vbCrLf
        
        strSQL = "" & strSQL & _
        "   Select Max(A.ID) as ID " & _
        "   From 票据打印内容 A,单据信息 J " & _
        "   Where A.数据性质=1   And A.NO=J.NO"
        
        strSQL = "Select * From (" & strSQL & ")"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "mdlOutExse", varData)
    End If
    If rsTemp.RecordCount <> 0 Then
        zlCheckIsPrintInvoice = Val(Nvl(rsTemp!ID)) <> 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet诊疗收费对照(ByVal str医嘱序号 As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID,获取相应的诊疗收费对照项目
    '入参:str医嘱序号-多个时，用逗号分隔
    '返回:诊疗收费对照的数据集
    '编制:刘兴洪
    '日期:2010-11-03 14:12:35
    '问题:33634
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo errHandle
     strSQL = "" & _
    "   Select /*+ RULE */ A.Id as 医嘱序号,B.诊疗项目ID,B.收费项目ID as 收费细目ID,b.收费数量,b.固有对照,b.从属项目 " & _
    "   From  病人医嘱记录 A,诊疗收费关系 B,Table(f_num2list([1])) J" & _
    "   Where   a.ID=J.Column_Value   And a.诊疗项目id=b.诊疗项目ID And nvl(b.固有对照,0)=1"
    Set zlGet诊疗收费对照 = zlDatabase.OpenSQLRecord(strSQL, "获取医嘱所对应的收费关系", str医嘱序号)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub

Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub


Public Function zlGetRegEventsCons(Optional strFieldName As String = "急诊", _
    Optional strAliaName As String = "", Optional bln发生时间 As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号项目的限制条件
    '入参:strFieldName-包含别外还字段(如急诊)
    '       strAliaName:别名
    '出参:
    '返回:返回条件
    '编制:刘兴洪
    '日期:2010-12-20 16:33:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strTimeName As String
    strFieldName = IIf(strAliaName <> "", strAliaName & ".", "") & strFieldName
    strTimeName = IIf(strAliaName <> "", strAliaName & ".", "") & IIf(bln发生时间, "发生时间", "登记时间")
    
    With gTy_System_Para.Sy_Reg
        strWhere = ""
        If .bytNODaysGeneral <> 0 Or .bytNoDayseMergency <> 0 Then
            If .bytNODaysGeneral <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=0  And " & strTimeName & ">Trunc(Sysdate-" & .bytNODaysGeneral & "))"
            Else
                strWhere = strWhere & " Or  nvl(" & strFieldName & ",0)=0   "
            End If
            If .bytNoDayseMergency <> 0 Then
                strWhere = strWhere & " Or ( nvl(" & strFieldName & ",0)=1  And " & strTimeName & ">Trunc(Sysdate-" & .bytNoDayseMergency & "))"
            Else
                strWhere = strWhere & " Or nvl(" & strFieldName & ",0)=1  "
            End If
        End If
        If strWhere <> "" Then
            strWhere = " And  (" & Mid(strWhere, 4) & ")"
        End If
    End With
    zlGetRegEventsCons = strWhere
End Function


Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, _
    lng西药房 As Long, lng成药房 As Long, lng中药房 As Long, _
    Optional lng病人ID As Long = 0, Optional int险类 As Integer = 0, _
    Optional bln药房单位 As Boolean, Optional lng开单部门ID As Long = 0, Optional byt婴儿费 As Byte, _
    Optional int门诊标志 As Integer, Optional bln加班加价 As Boolean = False, _
    Optional ByVal lngUnitID As Long, Optional int范围 As Integer, _
    Optional str划价人 As String = "", Optional str开单人 As String = "", _
    Optional ByVal str药品价格等级 As String, _
    Optional ByVal str卫材价格等级 As String, Optional ByVal str普通价格等级 As String, _
    Optional ByVal lng主页ID As Long, Optional ByVal lng科室ID As Long, Optional ByVal lng病区ID As Long) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取费用单据到单据对象中
    '入参:rsSel-选中的成套项目
    '       lngUnitID    当前操作病区ID
    '      int范围=1.门诊,2-住院
    '出参:
    '返回:存放单据信息的单据对象
    '编制:刘兴洪
    '日期:2010-09-02 16:17:54
    '说明:因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    '       不包含已停用收费细目
    '问题:    '问题:34465
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str收费细目ID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng病人科室ID As Long, str摘要 As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, str药房IDs As String, str停用项目序号 As String, strPrivs As String
    Dim curModiMoney As Currency
    Dim strAdvance As String, strInfo As String
    Dim dblAllTime As Double, dblCurTime As Double, dbl加班加价率 As Double, lngLastPati As Long
    Dim colSerial As New Collection
    Dim bytType As Byte '0-门诊;1-住院;2-门诊或住院
    Dim strTable  As String
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dbl剩余数量 As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    Set colSerial = New Collection
    '价格等级
    If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And d.价格等级 = [14])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And d.价格等级 = [15])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And d.价格等级 = [16])" & vbNewLine & _
            "            Or (d.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From 收费价目" & vbNewLine & _
            "                                Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And 价格等级 = [14])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And 价格等级 = [15])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And 价格等级 = [16])))))"
    Else
        strWherePriceGrade = " And d.价格等级 Is Null"
    End If
    
    With rsSel
        str收费细目ID = "": j = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Len(str收费细目ID) > 1990 And j <= 10 Then
                strValue(j) = Mid(str收费细目ID, 2)
                strSubItem = strSubItem & " Union ALL " & _
                " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
                str收费细目ID = "": j = j + 1
            End If
            str收费细目ID = str收费细目ID & "," & Val(Nvl(!收费细目ID))
            .MoveNext
        Loop
    End With
    
    If str收费细目ID <> "" Then
        If j > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From 收费项目目录 Where id in (" & Mid(str收费细目ID, 2) & ")"
        Else
            strValue(j) = Mid(str收费细目ID, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 收费细目ID From Table(f_Num2List([" & j + 1 & "])) B "
        End If
    End If
    
    gstrSQL = "" & _
       "   Select /*+ rule */ A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
       "   From 收费从属项目 A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.主项id = D.收费细目id "
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOutExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as 病人ID,收费细目ID From (" & strSubItem & ")"

    If lng主页ID = 0 Then
        gstrSQL = "" & _
        " Select  X.药品ID,W.材料ID,W.跟踪在用," & _
        "       F.费别,F.姓名,F.性别,F.年龄,F.担保额," & _
        "       '' as 床号,F.门诊号 as 标识号,F.病人ID,0 as 主页ID,0 as 病人病区ID,0 as 病人科室ID," & _
        "       B.类别 as 收费类别,A.收费细目ID," & _
        "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(H.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
        "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象, B.费用类型  费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
        "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
        "       Decode(B.类别,'4',1,X." & gstr药房包装 & ") as 药房包装," & _
        "       Decode(B.类别,'4',B.计算单位,X." & gstr药房单位 & ") as 药房单位," & _
        "       Decode(b.类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,B.录入限量, " & _
        "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,X.中药形态,x.剂量系数,M1.计算单位 as 剂量单位" & _
        "   From  (" & strTable & ") A ,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,病人信息 F, " & _
        "          收费项目别名 H,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
        " Where  A.收费细目ID=D.收费细目ID And A.收费细目ID=B.ID " & _
        "       And b.类别=C.编码 And A.收费细目ID=X.药品ID(+) and X.药名ID=M1.ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
        "       And A.收费细目ID=H.收费细目ID(+) And H.码类(+)=1 And H.性质(+)=[12]" & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.病人ID=F.病人ID(+)" & _
        "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) " & vbNewLine & _
        "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                strWherePriceGrade
    Else
        gstrSQL = "" & _
        " Select  X.药品ID,W.材料ID,W.跟踪在用," & _
        "       Nvl(G.费别,F.费别) As 费别,Nvl(G.姓名,F.姓名) As 姓名,Nvl(G.性别,F.性别) As 性别,Nvl(G.年龄,F.年龄) As 年龄,F.担保额," & _
        "       G.出院病床 as 床号,F.门诊号 as 标识号,F.病人ID,G.主页ID,G.当前病区ID as 病人病区ID,G.出院科室ID as 病人科室ID," & _
        "       B.类别 as 收费类别,A.收费细目ID," & _
        "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(H.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
        "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象, B.费用类型  费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
        "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
        "       Decode(B.类别,'4',1,X." & gstr药房包装 & ") as 药房包装," & _
        "       Decode(B.类别,'4',B.计算单位,X." & gstr药房单位 & ") as 药房单位," & _
        "       Decode(b.类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,B.录入限量, " & _
        "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,X.中药形态,x.剂量系数,M1.计算单位 as 剂量单位" & _
        "   From  (" & strTable & ") A ,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,病人信息 F, " & _
        "       病案主页 G,收费项目别名 H,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
        " Where  A.收费细目ID=D.收费细目ID And A.收费细目ID=B.ID " & _
        "       And b.类别=C.编码 And A.收费细目ID=X.药品ID(+) and X.药名ID=M1.ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
        "       And A.收费细目ID=H.收费细目ID(+) And H.码类(+)=1 And H.性质(+)=[12]" & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.病人ID=F.病人ID(+) And F.病人ID=G.病人ID(+) And G.主页ID(+) = [17]" & _
        "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) " & vbNewLine & _
        "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                strWherePriceGrade
    End If
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOutExse", strValue(0), strValue(1), strValue(2), _
        strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), _
        IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1), lng病人ID, str药品价格等级, str卫材价格等级, str普通价格等级, lng主页ID)
     '没有记录就是空单子
    Set objBill = New ExpenseBill
 
    Set objBill.Pages(1).Details = New BillDetails
    With rsSel
            i = 1
            If .RecordCount <> 0 Then .MoveFirst
NextRecord: Do While Not .EOF
            '检查收费项目是否停用或服务于门诊病人
            '主项停用时,不导从项
            rsItems.Filter = "收费细目ID=" & Val(Nvl(!收费细目ID))
            If rsItems.EOF Then '未找到.不加入
                 .MoveNext
                GoTo NextRecord:
            End If

            '检查收费项目是否停用或服务于门诊病人
            '主项停用时,不导从项
            If InStr(",5,6,7,", rsItems!收费类别) = 0 Then
                If InStr(1, str停用项目序号 & ",", "," & !从属父号 & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!收费细目ID, 1) Then
                        str停用项目序号 = str停用项目序号 & "," & !序号
                        MsgBox "成套收费项目中的第" & !序号 & "行收费项目:" & rsItems!名称 & "" & vbCrLf & _
                            "已停用或不再服务于病人,将不会被导入." & IIf(IsNull(!从属父号), "如果有从属项目,也不会被导入.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            End If
        
            '处理单据主体=====================================================
            If i = 1 Then
                objBill.NO = ""
                objBill.Pages(1).NO = "" '要清空以便修改时表明是直接输入的费用
                objBill.Pages(1).开单部门ID = lng开单部门ID
                objBill.Pages(1).开单人 = str开单人
                objBill.Pages(1).医嘱序号 = 0
                
                objBill.病人ID = Val(Nvl(rsItems!病人ID))
                objBill.主页ID = Val(Nvl(rsItems!主页ID))
                objBill.病区ID = IIf(lng病区ID = 0, Val(Nvl(rsItems!病人病区ID)), lng病区ID)
                objBill.科室ID = IIf(lng科室ID = 0, Val(Nvl(rsItems!病人科室id)), lng科室ID)
                objBill.姓名 = Nvl(rsItems!姓名)
                objBill.性别 = Nvl(rsItems!性别)
                objBill.年龄 = Nvl(rsItems!年龄)
                objBill.标识号 = Val(Nvl(rsItems!标识号))
                objBill.床号 = Nvl(rsItems!床号)
                objBill.费别 = Nvl(rsItems!费别)
                objBill.门诊标志 = int门诊标志
                objBill.加班标志 = IIf(bln加班加价, 1, 0)
                objBill.婴儿费 = byt婴儿费
                objBill.划价人 = str划价人
                objBill.操作员编号 = UserInfo.编号
                objBill.操作员姓名 = UserInfo.姓名
                objBill.发生时间 = zlDatabase.Currentdate
                objBill.登记时间 = objBill.发生时间
                objBill.多病人单 = 0
            End If
            
            '处理收费细目=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                        
            '处理序号,从属父号
            intCurNo = intCurNo + 1
            objBillDetail.序号 = intCurNo '实际是行号
            colSerial.Add Array(Val(Nvl(!收费细目ID)), intCurNo), "_" & !序号
            objBillDetail.从属父号 = Nvl(!从属父号, 0) '因为可能排序乱了,先记录原来的,后面再处理
            
            '使用原定的动态费别
            objBillDetail.费别 = Nvl(rsItems!费别)
            objBillDetail.收费类别 = Nvl(rsItems!收费类别)
            objBillDetail.收费细目ID = Nvl(rsItems!收费细目ID)
            objBillDetail.计算单位 = Nvl(rsItems!计算单位)
            objBillDetail.付数 = IIf(Val(Nvl(!付数)) = 0, 1, Val(Nvl(!付数)))
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln药房单位 Then
                objBillDetail.数次 = Nvl(!数量, 0) / Nvl(rsItems!药房包装, 1)
            Else
                objBillDetail.数次 = Nvl(!数量, 0)
            End If
            objBillDetail.原始数量 = objBillDetail.付数 * objBillDetail.数次
            
            objBillDetail.发药窗口 = ""     '需要进一步确认
            
            objBillDetail.附加标志 = 0
            
            objBillDetail.摘要 = ""
            
            '卫材和药品部分
            '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
            If objBillDetail.收费类别 = "4" Then
                lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, objBill.科室ID)
                If lngDoUnit = 0 Then lngDoUnit = lng开单部门ID
            End If
            
            '病人科室ID
            lng病人科室ID = objBill.科室ID
            If lng病人科室ID = 0 Then lng病人科室ID = lng开单部门ID
            objBillDetail.Detail.执行科室 = IIf(IsNull(rsItems!执行科室), 0, rsItems!执行科室)
            
            lngDoUnit = Get收费执行科室ID(objBillDetail.收费类别, objBillDetail.收费细目ID, _
                objBillDetail.Detail.执行科室, lng病人科室ID, lng开单部门ID, int范围, _
                IIf(lng西药房 = 0, glng西药房, lng西药房), _
                IIf(lng成药房 = 0, glng成药房, lng成药房), _
                IIf(lng中药房 = 0, glng中药房, lng中药房), _
                lngDoUnit, lngUnitID)
            
            objBillDetail.执行部门ID = lngDoUnit
            
            objBillDetail.原始执行部门ID = objBillDetail.执行部门ID     '用于修改时快速判断库存
            
            objBillDetail.Detail.ID = !收费细目ID
            objBillDetail.Detail.编码 = Nvl(rsItems!编码)
            objBillDetail.Detail.变价 = (Val(Nvl(rsItems!是否变价)) = 1)
            objBillDetail.Detail.从项数次 = 0 '!!!目前忽略从属项目,当作独立项目
            objBillDetail.Detail.固有从属 = 0 '!!!目前忽略从属项目,当作独立项目
            objBillDetail.Detail.规格 = Nvl(rsItems!规格)
            objBillDetail.Detail.计算单位 = Nvl(rsItems!计算单位)
            
            objBillDetail.Detail.药房单位 = Nvl(rsItems!药房单位)
            objBillDetail.Detail.药房包装 = Nvl(rsItems!药房包装, 1)
            
            If InStr(",4,5,6,7,", rsItems!收费类别) > 0 Then
                dblStock = GetStock(Val(Nvl(!收费细目ID)), objBillDetail.执行部门ID)
            Else
                dblStock = 0
            End If

            If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln药房单位 Then dblStock = dblStock / objBillDetail.Detail.药房包装
            objBillDetail.Detail.库存 = dblStock
            
            
            objBillDetail.Detail.加班加价 = (Val(Nvl(rsItems!加班加价)) = 1)
            objBillDetail.Detail.类别 = Nvl(rsItems!类别)
            objBillDetail.Detail.类别名称 = Nvl(rsItems!类别名称)
            objBillDetail.Detail.名称 = Nvl(rsItems!名称)
            objBillDetail.Detail.商品名 = Nvl(rsItems!商品名)
            objBillDetail.Detail.屏蔽费别 = (Val(Nvl(rsItems!屏蔽费别)) = 1)
            objBillDetail.Detail.说明 = ""
            objBillDetail.Detail.类型 = IIf(IsNull(rsItems!费用类型), "", rsItems!费用类型)
            objBillDetail.Detail.诊疗名称 = Nvl(rsItems!诊疗名称)
            objBillDetail.Detail.中药形态 = Val(Nvl(rsItems!中药形态))
            
            If objBillDetail.从属父号 <> 0 Then
                'A.主项id, A.从项id, A.固有从属, A.从项数次 "
                rsOthers.Filter = "主项ID=" & colSerial("_" & !从属父号)(0) & " And 从项ID=" & objBillDetail.收费细目ID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.从项数次 = Val(Nvl(rsOthers!从项数次))
                    objBillDetail.Detail.固有从属 = Val(Nvl(rsOthers!固有从属))
                End If
            End If
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 Then
                objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                objBillDetail.Detail.处方限量 = Get处方限量(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.录入限量 = Val(Nvl(rsItems!录入限量))
            
            objBillDetail.Detail.药名ID = Val(Nvl(rsItems!药名ID))
            objBillDetail.Detail.变价 = Val(Nvl(rsItems!是否变价)) = 1
            objBillDetail.Detail.分批 = Val(Nvl(rsItems!分批)) = 1
            objBillDetail.Detail.跟踪在用 = Val(Nvl(rsItems!跟踪在用)) = 1
            objBillDetail.Detail.剂量单位 = Nvl(rsItems!剂量单位)
            objBillDetail.Detail.剂量系数 = Val(Nvl(rsItems!剂量系数))
            '问题:41136
            str摘要 = objBillDetail.摘要
'            If lng病人ID <> 0 And intInsure <> 0 Then '90304
                str摘要 = gclsInsure.GetItemInfo(intInsure, lng病人ID, objBillDetail.收费细目ID, str摘要, 1, , "|1")
                objBillDetail.摘要 = str摘要
'            Else
'                objBillDetail.摘要 = ""
'            End If
            
            '处理价格部份=====================================================
            If rsItems.RecordCount > 0 Then rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '按照现有的价格设置重新计算
                If InStr(",5,6,7,", rsItems!收费类别) > 0 Or (rsItems!收费类别 = "4" And Nvl(rsItems!跟踪在用, 0) = 1) Then
                    '----------------------------------------------------------------------------------------------
                    '时价药品计算价格(分批可不分批)
                    dblAllTime = Val(Nvl(!数量))     '这里是售价数量
                    If dblAllTime <> 0 Or Val(Nvl(rsItems!是否变价)) = 1 Then
                        Set rsPrice = zlDatabase.OpenSQLRecord("Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual", _
                                    "获取药品当前售价", CLng(!收费细目ID), objBillDetail.执行部门ID, dblAllTime)
                        If rsPrice.EOF Then
                            '获取价格失败
'                            If !收费类别 = "4" Then
'                                MsgBox "卫生材料""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
'                            Else
'                                MsgBox "药品""" & Nvl(!名称) & """获取价格失败！", vbInformation, gstrSysName
'                            End If
                            objBillIncome.标准单价 = 0
                        Else
                            strPrice = Nvl(rsPrice!Price) & "|||"
                            varPrice = Split(strPrice, "|")
                            objBillIncome.标准单价 = Val(varPrice(0))
                            dbl剩余数量 = Val(varPrice(2))
                            
                            If dbl剩余数量 <> 0 And Val(Nvl(rsItems!是否变价)) = 1 Then
                                '数量未分解完毕
'                                If rsItems!收费类别 = "4" Then
'                                    MsgBox "时价卫生材料""" & rsItems!名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
'                                Else
'                                    MsgBox "时价药品""" & rsItems!名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
'                                End If
                                objBillIncome.标准单价 = 0
                            End If
                        End If
                    Else
                        objBillIncome.标准单价 = 0
                    End If
                ElseIf Val(Nvl(rsItems!是否变价)) = 1 Then
                    If Abs(Val(Nvl(!单价))) > Abs(Val(Nvl(rsItems!现价))) Or Abs(Val(Nvl(!单价))) = 0 Then
                        objBillIncome.标准单价 = Val(Nvl(rsItems!缺省价格))
                    Else
                        objBillIncome.标准单价 = Val(Nvl(!单价))
                    End If
                Else
                objBillIncome.标准单价 = Val(Nvl(rsItems!现价))
                End If
                                    
                If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln药房单位 Then
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(rsItems!药房包装, 1), gstrFeePrecisionFmt)
                Else
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrFeePrecisionFmt)
                End If
                objBillIncome.现价 = Val(Nvl(rsItems!现价))  '现价原价对药品变价无用
                objBillIncome.原价 = Val(Nvl(rsItems!原价))
                objBillIncome.收入项目ID = Val(Nvl(rsItems!现收入ID))
                objBillIncome.收入项目 = Nvl(rsItems!收入项目)
                objBillIncome.收据费目 = Nvl(rsItems!现费目)
                
                '应收金额=单价*付次*数次
                objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                
                '附加手术费率用计算(所有收入项目)
                If 0 = 1 And Nvl(rsItems!收费类别) = "F" Then
                    objBillIncome.应收金额 = objBillIncome.应收金额 * IIf(Val(Nvl(rsItems!附术收费率)) = 0, 1, Val(Nvl(rsItems!附术收费率)) / 100)
                End If
                
                '加班费用率计算
                dbl加班加价率 = 0
                If bln加班加价 And Val(Nvl(rsItems!加班加价)) = 1 Then
                    dbl加班加价率 = Val(Nvl(rsItems!加班加价)) / 100
                    objBillIncome.应收金额 = objBillIncome.应收金额 + objBillIncome.应收金额 * dbl加班加价率
                End If
                objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                
                '计算实收金额
                If Val(Nvl(rsItems!屏蔽费别)) = 1 Then
                    objBillIncome.实收金额 = objBillIncome.应收金额
                Else
                    '使用原定的动态费别
                    objBillIncome.实收金额 = ActualMoney(objBillDetail.费别, Val(Nvl(rsItems!现收入ID)), objBillIncome.应收金额, _
                        objBillDetail.收费细目ID, objBillDetail.执行部门ID, objBillDetail.原始数量, dbl加班加价率)
                End If
                
                With objBillIncome
                    '获取项目保险信息,仅医保病人才算
                    If int险类 <> 0 Then
                        strAdvance = objBillDetail.摘要 & "||" & objBillDetail.原始数量
                        strInfo = gclsInsure.GetItemInsure(objBill.病人ID, objBillDetail.收费细目ID, .实收金额, True, int险类, strAdvance)
                        If strInfo <> "" Then
                            objBillDetail.保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                            objBillDetail.保险大类ID = Val(Split(strInfo, ";")(1))
                            .统筹金额 = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                            objBillDetail.保险编码 = CStr(Split(strInfo, ";")(3))
                            
                            If UBound(Split(strInfo, ";")) >= 4 Then
                                If CStr(Split(strInfo, ";")(4)) <> "" Then objBillDetail.摘要 = CStr(Split(strInfo, ";")(4))
                                If UBound(Split(strInfo, ";")) >= 5 Then
                                    If Split(strInfo, ";")(5) <> "" Then objBillDetail.Detail.类型 = Split(strInfo, ";")(5)
                                End If
                            End If
                        End If
                    End If
                    objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .统筹金额
                End With
                '判断下一条记录是否属于当前行
                int序号 = !序号
                i = i + 1
                rsItems.MoveNext
            Loop
           
            With objBillDetail
                objBill.Pages(1).Details.Add .费别, .Detail, .收费细目ID, .序号, .从属父号, .收费类别, .计算单位, .发药窗口, _
                    .付数, .数次, .附加标志, .执行部门ID, .InComes, , .保险项目否, .保险大类ID, .保险编码, .摘要, .原始数量, .原始执行部门ID
                
                '设置工本费
                If objBill.Pages(1).Details(objBill.Pages(1).Details.Count).附加标志 = 8 Then
                    objBill.Pages(1).Details(objBill.Pages(1).Details.Count).工本费 = True
                End If
            End With
            .MoveNext
        Loop
    End With
     '再重新处理从属父号
     With objBill.Pages(1)
        For i = 1 To .Details.Count
            If .Details(i).从属父号 <> 0 Then
                 .Details(i).从属父号 = colSerial("_" & .Details(i).从属父号)(1)
            End If
        Next
    End With
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
Errhand:
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
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
Errhand:
End Function



Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln负数检查, bln零检查, hWnd, str项目)
End Function
Public Function CheckNegative(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal lngItem As Long, ByVal lngExecuteDept As Long, _
    ByVal dblNum As Double, ByVal dbl药房包装 As Double, ByVal strPrivs As String, Optional strStartDate As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人本次次住院的收费项目的数量合计是否足够冲销
    '入参:lngNum-输入的负数数量，如果是药品，根据参数转换成售价单位再传入如果同一单据输入相同的项目和执行科室的有多行，此时不检查，保存之前再检查
    '       strPrivs-权限串
    '       strStartDate-查询的日期范围的开始时间到当前时间
    '出参:
    '返回:足够或有权限冲负数时,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-03-18 11:43:09
    '问题:36558
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl未结 As Double, dbl已结 As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '目前只支持门诊留观病人的检查
    If strStartDate = "" Then strStartDate = "2000-01-01"
    '暂不设置该权限
'    If InStr(1, strPrivs, ";负数记帐不检查发生项目;") > 0 Then
'        '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
'        CheckNegative = True: Exit Function
'    End If
    
    '记录性质 In(2,3)取掉结帐作废的情况:  :28029
    On Error GoTo errH
    CheckNegative = True
    
    strSQL = "" & _
            "   Select Nvl(Sum(Nvl(付数, 1) * 数次),0) As 数量," & vbNewLine & _
            "           Sum(decode(结帐ID,NULL,0,1)* Nvl( 付数,1)* 数次) as 结帐数量  " & _
            "   From 门诊费用记录" & vbNewLine & _
            "   Where  记录性质 =2 and 记帐费用 = 1 And 价格父号 Is Null  And 记录状态<>0  " & _
            "               And 病人id = [1] " & vbNewLine & _
            "               And 收费细目id+0 = [3] And 执行部门id+0 = [4] And 登记时间+0>=[5]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, lngItem, lngExecuteDept, CDate(strStartDate))
    
    If Not rsTmp.EOF Then
        If RoundEx(Abs(dblNum), 8) > RoundEx(Val(Nvl(rsTmp!数量)), 8) Then
                MsgBox "销帐数量大于该病人在当前执行科室的记帐数量" & FormatEx(rsTmp!数量 / IIf(gbln药房单位, dbl药房包装, 1), 5) & "。", vbInformation, gstrSysName
                CheckNegative = False: Exit Function
        End If
        
        '暂不管
'        Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
'        Case 0  '允许
'        Case 1   '提醒
'            dbl未结 = RoundEx((Val(Nvl(rstmp!数量)) - Val(Nvl(rstmp!结帐数量))) / IIf(gbln药房单位, dbl药房包装, 1), 8)
'            dbl已结 = RoundEx(Val(Nvl(rstmp!结帐数量)) / IIf(gbln药房单位, dbl药房包装, 1), 8)
'            If RoundEx(Abs(dblNum), 8) > RoundEx(dbl未结, 8) Then
'                If MsgBox("销帐数量(" & FormatEx(RoundEx(Abs(dblNum) / IIf(gbln药房单位, dbl药房包装, 1), 8), 5) & _
'                        ") 中包含了已经结帐部分(未结:" & FormatEx(dbl未结, 5) & "; 已结:" & FormatEx(dbl已结, 5) & ") 。" & vbCrLf & _
'                    " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    CheckNegative = False: Exit Function
'                End If
'            End If
'        Case 2   '禁止
'                dbl未结 = RoundEx((Val(Nvl(rstmp!数量)) - Val(Nvl(rstmp!结帐数量))) / IIf(gbln药房单位, dbl药房包装, 1), 8)
'                dbl已结 = RoundEx(Val(Nvl(rstmp!结帐数量)) / IIf(gbln药房单位, dbl药房包装, 1), 8)
'                If RoundEx(Abs(dblNum), 8) > RoundEx(dbl未结, 8) Then
'                    Call MsgBox("销帐数量(" & FormatEx(RoundEx(Abs(dblNum) / IIf(gbln药房单位, dbl药房包装, 1), 8), 5) & _
'                        ") 中包含了已经结帐部分(未结:" & FormatEx(dbl未结, 5) & "; 已结:" & FormatEx(dbl已结, 5) & ") ,不能继续。" & vbCrLf & _
'                    "", vbInformation + vbOKOnly, gstrSysName)
'                    CheckNegative = False: Exit Function
'                End If
'        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, _
    Optional str使用类别 As String = "", Optional ByRef intPrintFormatOld As Integer, _
    Optional ByRef blnPatiPrintBill As Boolean = False, Optional ByVal blnDelFeePrintBill As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '入参:blnPatiPrintBill-获取按病人补打票据格式
    '   blnDelFeePrintBill - 获取退费发票格式(91998)
    '出参:intPrintFormatOld-返回老票据打印格式(票号分配规则为根据实际打印分票票号方式所打印的格式)(56963)
    '返回:打印格式(序号)
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim intFormat As Integer, intFormat1 As Integer
    Dim intNewPrintFormat As Integer
    
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    If blnDelFeePrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("退费发票格式", glngSys, lngModule, ""))
    ElseIf blnPatiPrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("按病人补打发票格式", glngSys, lngModule, ""))
    Else
        strShareTypeFormat = Trim(zlDatabase.GetPara("收费发票格式", glngSys, lngModule, ""))
    End If
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intFormat1 = intFormat
        If Trim(varTemp(0)) = str使用类别 And intFormat <> 0 Then
            intNewPrintFormat = intFormat: GoTo GetOLdFormat:
        End If
    Next
    intNewPrintFormat = intFormat1
    '获取旧发票格式(56963)
GetOLdFormat:
    If gTy_Module_Para.byt票据分配规则 = 0 Or blnDelFeePrintBill Then
        '根据实际打印分配票号时,则还是原来方式打印不用处理
        intPrintFormatOld = intNewPrintFormat
        zl_GetInvoicePrintFormat = intNewPrintFormat
        Exit Function
    End If
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    strShareTypeFormat = Trim(zlDatabase.GetPara("收费发票格式(老)", glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        intFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intFormat1 = intFormat
        If Trim(varTemp(0)) = str使用类别 And intFormat <> 0 Then
            intPrintFormatOld = intFormat
            zl_GetInvoicePrintFormat = intNewPrintFormat
            Exit Function
        End If
    Next
    intPrintFormatOld = intFormat1
    zl_GetInvoicePrintFormat = intNewPrintFormat
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional str使用类别 As String = "", Optional ByVal blnDelFeePrintBill As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '出参:int打印方式-打印方式()
    '   blnDelFeePrintBill - 获取退费发票格式(91998)
    '返回:0-不打印;1-自动打印;2-提示打印
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    If blnDelFeePrintBill Then
        strShareTypeFormat = Trim(zlDatabase.GetPara("退费发票打印方式", glngSys, lngModule, ""))
    Else
        strShareTypeFormat = Trim(zlDatabase.GetPara("收费发票打印方式", glngSys, lngModule, ""))
    End If
    '格式:使用类别1,打印方式1|使用类别2,打印方式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        intPrintMode = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
        If Trim(varTemp(0)) = str使用类别 Then
            zl_GetInvoicePrintMode = intPrintMode: Exit Function
        End If
    Next
    zl_GetInvoicePrintMode = intPrintMode1
End Function

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
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Function GetChargeBalance(ByVal strNos As String, _
    Optional ByVal lng结算序号 As Long = 0, Optional lng结帐ID As Long, _
    Optional blnHistory As Boolean = False, _
    Optional strDelTime As String, Optional intSign As Integer = 1) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费的相关结算方式
    '入参:判断顺序:strNos-->lng结算序号-->lng结帐ID
    '       strDelTime-退费时间表
    '       intSign:统计符:1和-1;
    '返回:收费相关的结算方式(性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类)
    '       字段:A.结帐ID,A.NO,A.性质,A.结算性质,A.结算方式,A.结算金额,
    '               A.卡类别ID,A.名称,A.是否全退,A.是否退现,A.结算号码,A.卡号,A.交易流水号,
    '               A.交易说明,A.结算序号,A.校对标志
    '编制:刘兴洪
    '日期:2011-08-28 21:38:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable  As String, strFeeTab As String, strPreTab As String
    Dim dtDelDate As Date
    strFeeTab = IIf(blnHistory, "H门诊费用记录", "门诊费用记录")
    strPreTab = IIf(blnHistory, "H病人预交记录", "病人预交记录")
    
    On Error GoTo errHandle
    
   dtDelDate = CDate("1991-01-01")
    If strDelTime <> "" Then dtDelDate = CDate(strDelTime)
    If strNos <> "" Then
        strTable = "Select distinct M.NO,M.结帐ID From " & strFeeTab & " M," & strPreTab & " J,Table( f_Str2list([2])) Q Where M.结帐ID=J.结帐ID and M.NO=Q.Column_value And M.记录性质=1"
        strTable = strTable & "" & _
        IIf(strDelTime = "", " And M.记录状态 In(1,3)", " And M.记录状态=2") & _
        IIf(strDelTime <> "", " And M.登记时间=[3]", "")
    Else
        strTable = "Select distinct M.NO,M.结帐ID From " & strFeeTab & " M," & strPreTab & " J,Table( f_Num2list([1])) Q Where M.结帐ID=J.结帐ID and " & IIf(lng结算序号 = 0, "J.结帐ID", "J.结算序号") & "=Q.Column_value And M.记录性质=1"
    End If
    
    strSQL = "" & _
    "   Select   W.NO,A.结帐ID, " & _
    "       Case  When Mod(A.记录性质,10)=1 then 1  " & _
    "                 When nvl(A.卡类别ID,0)>0 then   3 " & _
    "                 When nvl(A.结算卡序号,0)>0 then 4 " & _
    "                 When  nvl(B.性质,0)=3 or nvl(B.性质,0)=4 then 2 " & _
    "                 When  C.结算方式 Is not null   then 5 else 0 End  as 性质," & _
    "       nvl(b.性质,1) as 结算性质,B.应付款, " & _
    "       Decode(Mod(A.记录性质,10),1,'预存款',A.结算方式) as 结算方式, " & _
    "       " & intSign & "*nvl(A.冲预交,0) as 冲预交," & _
    "       Nvl(I.是否退款验卡,0) as 是否退款验卡," & _
    "       nvl(nvl(A.卡类别ID,A.结算卡序号),0) as 卡类别ID, nvl(I.名称,L.名称) as 名称, " & _
    "       nvl(nvl(I.是否全退,L.是否全退) ,0) as 是否全退,nvl(nvl(I.是否退现,L.是否退现),0) as 是否退现," & _
    "       A.结算号码,A.摘要,nvl(A.卡号,A.单位帐号) as 卡号,nvl(A.交易流水号,A.结算号码) as 交易流水号,A.交易说明,A.结算序号,C.医院编码,nvl(A.校对标志,0) 校对标志" & _
    "   From " & strPreTab & " A,  (" & strTable & ") W, " & _
    "           医疗卡类别 I,消费卡类别目录 L ,结算方式 B,  " & _
    "           (Select 结算方式 ,医院编码 From 一卡通目录 Where 启用=1 ) C " & _
    "   where A.结帐ID=W.结帐ID   " & _
    "               And A.卡类别ID=I.ID(+) and A.结算卡序号=L.编号(+) " & _
    "               And A.结算方式=B.名称(+)   " & _
    "               And A.结算方式=C.结算方式(+)"
    strSQL = "" & _
    "   Select /*+ Rule*/  A.结帐ID,A.NO,A.性质,A.结算性质,A.应付款,A.结算方式,nvl(sum(A.冲预交),0) as 结算金额, " & _
    "           A.是否退款验卡,A.卡类别ID,A.名称,A.是否全退,A.是否退现,A.结算号码,Max(A.摘要) as 摘要,A.卡号,A.交易流水号," & _
    "           A.交易说明, A.结算序号,A.校对标志 " & _
    "   From (" & strSQL & ") A" & _
    "   Group by  A.结帐ID,A.NO,A.性质,A.应付款,A.结算方式,A.卡类别ID ,A.名称,A.是否全退,A.是否退现,A.结算号码, " & _
    "            A.是否退款验卡,A.卡号,A.交易流水号,A.交易说明,A.结算序号 ,A.结算性质,A.医院编码, A.校对标志" & _
    "   Having nvl(sum(A.冲预交),0)<>0"
    '异常单据的结算方式(不含预交款)
    Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", IIf(lng结算序号 = 0, lng结帐ID, lng结算序号), Replace(strNos, "'", ""), dtDelDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function isCheckExiseSingularity(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一次收费是否存在异常的作废单据
    '入参:strNo-单据号
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-01 12:02:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  1  " & _
    "   From 病人预交记录 B,门诊费用记录 A,病人预交记录 C  " & _
    "   Where  B.结帐ID=A.结帐ID and B.结算序号=C.结算序号  " & _
    "               And C.NO||''<>[1] And C.记录状态<>1  And A.NO=[1] And Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取指定单据一次收费是否存在已经作废的单据", strNo)
    isCheckExiseSingularity = rsTemp.RecordCount <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlExeCuteBillNoSplit(ByVal bln模拟计算 As Boolean, ByVal int操作类型 As Integer, _
    ByVal lng领用ID As Long, ByVal strNos As String, ByVal lng病人ID As Long, _
    ByVal str起始发票号 As String, ByVal datFeeDate As Date, Optional ByVal byt票种 As Byte = 1, _
    Optional ByRef str发票号 As String, Optional int发票张数 As Integer, _
    Optional ByVal lngNext领用ID As Long, Optional ByVal strNext起始发票号 As String, Optional lng打印ID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行分配票号的过程
    '入参::1-正常打印票据;2-补打票据;3-重打票据;4-退费收回票据并重新发出票据
    '       strNos-单据号(用逗号分开)
    '       str发票号-发票号(多个用逗号分离):(3-重打;4-部分退费时,传入)
    '       lng打印ID- lng打印ID<>0时，表示根据临时表“临时票据打印内容”所对应的NO来产生票据
    '出参:str发票号-本次打印的发票号(多个用逗号分离)
    '       int发票张数-发票张数
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-03-27 10:10:41
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCmd As ADODB.Command, i As Long
    Dim objPara(0 To 13) As ADODB.Parameter
    Dim varParaName As Variant, varParaValue As Variant, varTemp As Variant
    Dim strParaName As String, varValue As Variant
    Dim intDataType As ADODB.DataTypeEnum
    Dim Paradt As ParameterDirectionEnum
    Dim intMaxSize As Integer
    Dim strLog As String
   
    On Error GoTo errHandle
    Set objCmd = New ADODB.Command
    '  Zl_Invoice_Autoallot(
    '  操作类型_In   Number,
    '  模拟计算_In   Number,
    '  票种_In       票据使用明细.票种%Type,
    '  领用id_In     票据使用明细.领用id%Type,
    '  病人id_In     门诊费用记录.病人id%Type,
    '  Nos_In        Varchar2,
    '  起始发票号_In 门诊费用记录.实际票号%Type,
    '  使用人_In     票据使用明细.使用人%Type,
    '  使用时间_In   票据使用明细.使用时间%Type,
    '  Next领用id_In 票据使用明细.领用id%Type := 0,
    '  Next票据号_In 票据使用明细.号码%Type := Null,
    '  发票号_In     In Out Varchar2,
    '  发票张数_In   Out Number
   varParaName = Split("操作类型_In|N|IN,模拟计算_In|N|IN,票种_In|N|IN,领用id_In|N|IN," & _
                    "病人id_In|N|IN,Nos_In|C|IN,起始发票号_In|C|IN,使用人_In|C|IN," & _
                    "使用时间_In|D|IN,Next领用id_In|N|IN,Next票据号_In|C|IN,发票号_In|C|INOUT,发票张数_In|N|OUT,打印id_In|N|IN", ",")
                    
                    
   varParaValue = Split(int操作类型 & ";" & IIf(bln模拟计算, 1, 0) & ";1;" & lng领用ID & ";" & _
                    lng病人ID & ";" & strNos & ";" & str起始发票号 & ";" & UserInfo.姓名 & ";" & _
                    Format(datFeeDate, "YYYY-MM-DD HH:MM:SS") & ";" & lngNext领用ID & ";" & strNext起始发票号 & ";" & str发票号 & ";" & int发票张数 & ";" & lng打印ID, ";")
  
                               
    For i = 0 To UBound(varParaName)
        '参数名|参数类型|入出参
        varTemp = Split(varParaName(i) & "||||", "|")
        strParaName = varTemp(0)    '参数名
        Select Case Trim(varTemp(1))
        Case "C" '字符
             varValue = Replace(CStr(varParaValue(i)), "'", "")
            ' 如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
            intMaxSize = LenB(StrConv(varValue, vbFromUnicode))
            If intMaxSize <= 2000 Then
                intMaxSize = IIf(intMaxSize <= 200, 200, 2000)
                intDataType = adVarChar
            Else
                If intMaxSize < 4000 Then intMaxSize = 4000
                intDataType = adLongVarChar
            End If
            strLog = strLog & ",'" & varValue & "'"
        Case "D"
             intDataType = adDBTimeStamp
             varValue = CDate(varParaValue(i))
             strLog = strLog & ",to_date('" & varParaValue(i) & "','yyyy-mm-dd hh24:mi:ss') "
        Case Else
             intDataType = adVarNumeric
             varValue = CLng(varParaValue(i))
             strLog = strLog & "," & varValue
             intMaxSize = 30
        End Select
        Select Case Trim(varTemp(2))
        Case "IN" '字符
             Paradt = adParamInput
        Case "INOUT"
             Paradt = adParamInputOutput
        Case Else
             Paradt = adParamOutput
        End Select
        
        If varTemp(1) = "D" Then
            Set objPara(i) = objCmd.CreateParameter(strParaName, _
              intDataType, Paradt)
        Else
            Set objPara(i) = objCmd.CreateParameter(strParaName, _
              intDataType, Paradt, intMaxSize)
        End If
        If Paradt <> adParamOutput Then
          objPara(i).Value = varValue
        End If
        objCmd.Parameters.Append objPara(i)
    Next
    If strLog <> "" Then strLog = Mid(strLog, 2)
    strLog = "Zl_Invoice_Autoallot(" & strLog & ")"
    objCmd.CommandText = "Zl_Invoice_Autoallot"
    objCmd.CommandType = adCmdStoredProc
    Set objCmd.ActiveConnection = gcnOracle
    Call SQLTest(App.ProductName, "票号分配", strLog)
    objCmd.Execute
    Call SQLTest
    str发票号 = Nvl(objPara(UBound(varParaName) - 2).Value)
    int发票张数 = Val(Nvl(objPara(UBound(varParaName) - 1).Value))
    zlExeCuteBillNoSplit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInvoiceFromNOs(ByVal strInvioceNos As String, _
    Optional bln含历史表空间 As Boolean = False, _
    Optional ByRef str结算序号 As String = "", Optional cllInvoiceNoInfor As Collection) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据发票号,获取对应的单据号
    '入参:strInvioceNos-发票号,多个可以用逗号分离:A0001,A0002
    '出参:str结算序号-返回这张单据的多个结算序号(如果该发票涉及多次收费的)
    '       cllInvoiceNoInfor-array(No,序号)
    '返回:成功返回传入的发票所涉及的单据号
    '编制:刘兴洪
    '日期:2013-04-12 15:59:32
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNos As String
    Dim strSQL1 As String, strSQL As String
    
    On Error GoTo errHandle
    
    Set cllInvoiceNoInfor = New Collection
    str结算序号 = ""
    If gTy_Module_Para.byt票据分配规则 <> 0 Then
        strSQL = "" & _
        "   Select  /*+ RULE */  A.NO,Max(A.序号) as 序号,Max(C.结算序号) as 结算序号" & _
        "   From 票据打印明细 A,门诊费用记录 B,病人预交记录 C,Table( f_Str2list([1])) J" & _
        "   Where A.票号=J.Column_Value and 票种=1 and A.是否回收<>1" & _
        "           And A.No=B.NO And B.记录性质=1  And nvl(B.记录状态,0)<>2 And B.结帐ID=C.结帐ID" & _
        "   Group by A.NO"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strInvioceNos)
        strNos = ""
        With rsTemp
            Do While Not .EOF
                strNos = strNos & "," & Nvl(!NO)
                str结算序号 = str结算序号 & "," & Val(Nvl(!结算序号))
                cllInvoiceNoInfor.Add Array(Nvl(!NO), Nvl(!序号))
                .MoveNext
            Loop
            If str结算序号 <> "" Then str结算序号 = Mid(str结算序号, 2)
            If strNos <> "" Then
                zlInvoiceFromNOs = Mid(strNos, 2)
                Exit Function
            End If
        End With
    End If
    strSQL = "" & _
    "   Select NO  " & _
    "   From 票据打印内容 A, " & _
    "           (   Select Max(M.打印ID) as 打印ID " & _
    "               From  票据使用明细 M ,Table( f_Str2list([1])) J  " & _
    "               Where M.票种=1 And M.性质=1 And M.号码=J.Column_Value  " & _
    "               Group by M.号码" & _
    "               )  Q" & _
    "   Where A.数据性质=1  And ID=Q.打印ID "
    strSQL1 = Replace(Replace(strSQL, "票据打印内容", "H票据打印内容"), "票据使用明细", "H票据使用明细")
    
    strSQL = "" & _
    "   Select  /*+ RULE */   Distinct NO " & _
    "   From (" & strSQL & " Union ALL " & strSQL1 & ") " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strInvioceNos)
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(!NO)
            .MoveNext
        Loop
        If strNos <> "" Then
            zlInvoiceFromNOs = Mid(strNos, 2)
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetReclaimInvoice(ByVal strNoInfor As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次需要回收的票据
    '入参:strNoInfor-单据号信息(单据号1:序号1(1..n);单据号2:序号2(1..n)
    '       str序号-单据中的序号,多个用逗号分离
    '返回:获取成功,返回本次需要回收的票据
    '编制:刘兴洪
    '日期:2013-03-27 18:27:13
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoices As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim strNos As String, varData As Variant, varTemp As Variant
    Dim i As Long, blnFind As Boolean, j As Long
    Dim str关联序号 As String, rsInvoice As ADODB.Recordset
    Dim cllNos As Collection
    Dim strNo As String
    Dim varValue() As Variant
    
    On Error GoTo errHandle
    
    '根据实际打印分配票号时,不返回具体的票据
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Function

    
    Set rsInvoice = New ADODB.Recordset
    rsInvoice.Fields.Append "发票号", adVarChar, 50, adFldIsNullable
    rsInvoice.CursorLocation = adUseClient
    rsInvoice.LockType = adLockOptimistic
    rsInvoice.CursorType = adOpenStatic
    rsInvoice.Open

    If strNoInfor = "" Then Exit Function
    strNoInfor = Replace(strNoInfor, "'", "")
    varData = Split(strNoInfor, ";")
  
    Set cllNos = New Collection
    strNos = ""
    For i = 0 To UBound(varData)
        strNo = Split(varData(i) & ":", ":")(0)
        If Len(strNos & "," & strNo) > 4000 Then
            strNos = Mid(strNos, 2)
            cllNos.Add strNos
            strNos = ""
        End If
        strNos = strNos & "," & strNo
    Next
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
        cllNos.Add strNos
    End If
    
    If cllNos.Count <= 1 Then
        strSQL = "" & _
        "   Select  /*+ RULE */  A.NO, A.票号,A.序号,A.关联票号序号" & _
        "   From 票据打印明细 A,Table( f_Str2list([1])) J" & _
        "   Where A.NO=Column_Value and 票种=1 and 是否回收<>1  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strNos)
    Else
        If zlFromCollectBulidSQL(cllNos, strSQL, varValue) = False Then Exit Function
        strSQL = "With 单据信息 as (" & strSQL & ")" & vbCrLf
        strSQL = strSQL & " Select * From 单据信息 "
        
        strSQL = "" & _
        "   Select  A.NO, A.票号,A.序号,A.关联票号序号" & _
        "   From 票据打印明细 A,(" & strSQL & ") J" & _
        "   Where A.NO=J.NO and 票种=1 and 是否回收<>1  "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "获取对应单据的发票号", varValue)
    End If


    If rsTemp.RecordCount = 0 Then Exit Function
    With rsTemp
        Do While Not .EOF
            If InStr(strInvoices & ",", "," & Nvl(!票号)) = 0 Then
                blnFind = False
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & ":", ":")
                    If varTemp(0) = Nvl(!NO) Then
                         If varTemp(1) = "" Then
                            If Val(Nvl(!关联票号序号)) <> 0 And InStr(str关联序号 & ",", "," & Val(Nvl(!关联票号序号)) & ",") = 0 Then
                                str关联序号 = str关联序号 & "," & Val(Nvl(!关联票号序号))
                            End If
                            blnFind = True: Exit For
                          End If
                         varTemp = Split(varTemp(1), ",")
                         For j = 0 To UBound(varTemp)
                             If InStr("," & Nvl(!序号) & ",", "," & varTemp(j) & ",") > 0 Then
                                    If Val(Nvl(!关联票号序号)) <> 0 And InStr(str关联序号 & ",", "," & Val(Nvl(!关联票号序号)) & ",") = 0 Then
                                        str关联序号 = str关联序号 & "," & Val(Nvl(!关联票号序号))
                                    End If
                                    blnFind = True: Exit For
                             End If
                         Next
                    End If
                Next
                If blnFind Then
                    rsInvoice.AddNew
                    rsInvoice!发票号 = Nvl(!票号)
                    rsInvoice.Update
                    strInvoices = strInvoices & "," & Nvl(!票号)
                End If
            End If
            .MoveNext
        Loop
        If strInvoices <> "" Then strInvoices = Mid(strInvoices, 2)
    End With
    If str关联序号 <> "" Then
            '关联的票据也要显示出来
            str关联序号 = Mid(str关联序号, 2)
            gstrSQL = "" & _
             "   Select  /*+ RULE */ distinct  A.票号" & _
             "   From 票据打印明细 A,Table( f_Num2list([1])) J" & _
             "   Where A.关联票号序号=J.Column_Value And A.票种=1 and A.是否回收<>1 " & _
             "              And A.票号 Not In(Select Column_Value From Table( f_Str2list([2])) )"
             Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取对应单据的发票号", str关联序号, strInvoices)
             With rsTemp
                Do While Not .EOF
                    rsInvoice.AddNew
                    rsInvoice!发票号 = Nvl(!票号)
                    rsInvoice.Update
                    strInvoices = strInvoices & "," & Nvl(!票号)
                    .MoveNext
                Loop
             End With
     End If
     '排序
     rsInvoice.Sort = "发票号"
     With rsInvoice
        If .RecordCount <> 0 Then .MoveFirst
        strInvoices = ""
        Do While Not .EOF
            strInvoices = strInvoices & "," & Nvl(!发票号)
            .MoveNext
        Loop
        .Close
     End With
     If strInvoices <> "" Then strInvoices = Mid(strInvoices, 2)
     Set rsInvoice = Nothing
    zlGetReclaimInvoice = strInvoices
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function zlGetFromNoTOInvoice(ByVal strNos As String) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据所对应的发票号
    '入参:strNos-单据号,可以为多个,多个时用逗号分离
    '返回:成功返回指定单据所对应的发票号的记录集
    '编制:刘兴洪
    '日期:2013-05-06 16:17:50
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    'If strNos <> "" Then strNos = Replace(Mid(strNos, 2), "'", "")
    strSQL = "" & _
    "   With C_单据 as (Select Column_Value as NO From Table( f_Str2list([1])) )" & _
    "   Select  /*+ RULE */  A.NO, A.票号,A.序号,A.关联票号序号" & _
    "   From 票据打印明细 A,C_单据 J" & _
    "   Where A.NO=J.NO and 票种=1 and 是否回收<>1  " & _
    "   Union ALL " & _
    "   Select A.NO, A.票号,A.序号,A.关联票号序号" & _
    "   From 票据打印明细 A, " & _
    "               (Select 关联票号序号 From 票据打印明细 A,C_单据 M  " & _
    "                Where A.NO=M.NO and A.票种=1 and A.是否回收<>1 ) J" & _
    "   Where A.关联票号序号=J.关联票号序号 And A.票种=1 and A.是否回收<>1 "
   strSQL = "" & _
   "    Select /*+ RULE */ distinct  A.NO, A.票号,A.序号,A.关联票号序号  " & _
   "    From (" & strSQL & ") A"
    Set zlGetFromNoTOInvoice = zlDatabase.OpenSQLRecord(strSQL, "获取对应单据的发票号", strNos)
End Function
Public Function zlCheckDrugIsPutDrug(ByVal strNos As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对药品摆药进行检查
    '入参:strNos-单据号,多个用逗号分隔
    '返回:没有摆药或摆药允许选择为true时,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-16 12:44:49
    '问题:47400
    ' 调用:退费时调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    If gTy_Module_Para.byt药品摆药退费方式 = 0 Then zlCheckDrugIsPutDrug = True: Exit Function
    
    strSQL = "Select  /*+ rule */  1 From 未发药品记录 A,Table(f_str2List([1])) J Where A.NO=J.Column_Value And A.单据 in (8,24) And A.配药人 Is NOT NULL And Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查药品是否摆药!", strNos)
    If rsTemp.EOF Then zlCheckDrugIsPutDrug = True: Exit Function
    If gTy_Module_Para.byt药品摆药退费方式 = 1 Then
        '禁止退费
        MsgBox "在退费单据中已经存在摆药的单据,不允许对已经摆药的单据进行退费!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '提示
    If MsgBox("在退费单据中已经存在摆药的单据,是否对已经摆药的单据进行退费?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    zlCheckDrugIsPutDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckIsExcuteData(ByVal strNos As String, ByVal byt记录性质 As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医嘱执行计价是否有数据
    '入参:strNOs-收费单号,多个用逗号分离
    '返回:有数据返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-25 13:57:56
    '问题:60735
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */  1 " & _
    "   From 门诊费用记录 A, 医嘱执行计价 B, Table(f_Str2list([2])) J " & _
    "   Where a.医嘱序号 = b.医嘱id And mod(a.记录性质,10) = [1] And a.No = j.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查医嘱执行计价是否有数据", byt记录性质, strNos)
    zlCheckIsExcuteData = Not rsTemp.EOF
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlStringSort(ByVal strSort As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:字符串排序
    '入参:strSort-用逗号分离
    '出参:
    '返回:返回排序后的字付串
    '编制:刘兴洪
    '日期:2013-05-07 18:23:34
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    Dim strTemp As String, intCount As Integer
    Dim i As Long, j As Long
    varTemp = Split(strSort, ",")
    intCount = UBound(varTemp)
    For i = 0 To intCount
        For j = i + 1 To intCount
            If varTemp(i) > varTemp(j) Then
                strTemp = varTemp(i)
                varTemp(i) = varTemp(j)
                varTemp(j) = strTemp
            End If
        Next
    Next
    strTemp = ""
    For i = 0 To intCount
        strTemp = strTemp & "," & varTemp(i)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    zlStringSort = strTemp
End Function
 
Public Sub zlDebugWriteFile(ByVal strLogText As String)
    Dim objLogFile As FileSystemObject
    Dim objLogText As TextStream
    Dim strTmp As String
    If OS.IsDesinMode = False Then Exit Sub
    
    Set objLogFile = New FileSystemObject
    On Local Error Resume Next
    Set objLogText = objLogFile.OpenTextFile(gstrDBUser & "_" & Format(Date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
    On Local Error GoTo 0
    If Not objLogText Is Nothing Then
        strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
        objLogText.WriteLine strTmp
        objLogText.WriteLine strLogText
    End If
    objLogText.Close
    Set objLogText = Nothing
    Set objLogFile = Nothing
End Sub

Public Function zlBillErrIsCanDel(ByVal lng结帐ID As Long, ByRef bytDelErrType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:指定的异常收费单据是否能作废
    '出参:bytErrType-退费错误类型:1-异常收费单作废;2-正常的退费
    '编制:刘兴洪
    '日期:2012-03-01 01:04:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim blnDel As Boolean, strNo As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select NO, Max(记录状态) as 记录状态" & _
    "   From 门诊费用记录  " & _
    "   Where 结帐ID=[1] And nvl(费用状态,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取异常收费单据", lng结帐ID)
    blnDel = False
    If Not rsTemp.EOF Then
        blnDel = Val(Nvl(rsTemp!记录状态)) = 2
        strNo = Nvl(rsTemp!NO)
        bytDelErrType = 2
        If blnDel Then
            strSQL = "" & _
             "   Select  1 " & _
             "   From 门诊费用记录  " & _
             "   Where NO=[1] And 记录性质=1 And 记录状态 in (1,3) " & _
             "     And nvl(费用状态,0)=1 "
             
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取异常收费单据", strNo)
            If Not rsTemp.EOF Then bytDelErrType = 1
        End If
    End If
    zlBillErrIsCanDel = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str病人类型 As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人类型,设置不同病人类型的显示颜色
    '入参:objPatiControl-病人控件(文本框,标签)
    '    str病人类型-病人类型
    '    lngDefaultColor-缺省病人的显示颜色
    '返回:True-设置颜色成功，False-失败
    '编制:李南春
    '日期:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str病人类型 <> "" Then
        lngColor = zlDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function
Public Function zlFormatNum(ByVal dblMoney As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取格式化串(比如:.03 格式为0.03,123格式为123)
    '入参:dblMoney-格式化金额
    '返回:返回格式化串(比如:.03 格式为0.03,123格式为123)
    '编制:刘兴洪
    '日期:2014-07-30 15:29:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    Dim strMoney As String
'    If dblMoney = 0 Then Exit Function
    strTemp = Format(dblMoney, "###0.00######;-###0.00######;;")
    If strTemp = "" Then Exit Function
    strMoney = strTemp
    For i = Len(strTemp) To 1 Step -1
        If Val(Mid(strTemp, i, 1)) <> 0 Or Mid(strTemp, i, 1) = "." Then Exit For
        strMoney = Mid(strTemp, 1, i - 1)
    Next
    If Right(strMoney, 1) = "." Then strMoney = Mid(strMoney, 1, Len(strMoney) - 1)
    zlFormatNum = strMoney
End Function

Public Function GetMedicareStr(colBalance As Collection, Optional ByVal intPage As Integer, _
    Optional ByVal intBeforePage As Integer) As String
'功能：返回保险结算方式串,"结算方式|金额||...."
'参数：intPage=是否指定单据,否则为所有单据
'      intBeforePage=计算该单据及以前的单据
'说明：该函数以colBalance为准计算,对于医保划价收费也是
    Dim i As Integer, p As Integer, strTmp As String
    Dim varData As Variant, curMoney As Currency
    Dim rsTemp As New ADODB.Recordset, strBalance As String, varBalance As Variant
    
    Err = 0: On Error GoTo Errhand:
    rsTemp.Fields.Append "结算方式", adVarChar, 20, adFldIsNullable
    rsTemp.Fields.Append "金额", adCurrency, , adFldIsNullable
    rsTemp.CursorLocation = adUseClient
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorType = adOpenStatic
    rsTemp.Open
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, IIf(intBeforePage = 0, colBalance.Count, intBeforePage), intPage)
        For i = 0 To UBound(colBalance(p))
            '结算方式;原始(最大)金额;可否修改;有效金额
            varData = Split(colBalance(p)(i), ";")
            If varData(0) <> "" Then
                If InStr(strBalance & ";", ";" & varData(0) & ";") = 0 Then
                    strBalance = strBalance & ";" & varData(0)
                End If
                rsTemp.AddNew
                rsTemp!结算方式 = varData(0)
                rsTemp!金额 = Val(varData(3))
                rsTemp.Update
            End If
        Next
    Next
    If strBalance <> "" Then
        strBalance = Mid(strBalance, 2)
        varBalance = Split(strBalance, ";")
        For i = 0 To UBound(varBalance)
            curMoney = 0
            rsTemp.Filter = "结算方式='" & varBalance(i) & "'"
            Do While Not rsTemp.EOF
                curMoney = curMoney + Nvl(rsTemp!金额)
                rsTemp.MoveNext
            Loop
            strTmp = strTmp & "||" & varBalance(i) & "|" & curMoney
        Next
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    GetMedicareStr = Mid(strTmp, 3)
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SetBalanceVal(colBalance As Collection, ByVal intPage As Integer, _
    ByVal strBalance As String, Optional ByRef strNone As String)
    '功能：设置指定编号单据指定保险结算方式的有效值
    '参数：
    '       strBalance-根据结算方式字符串设置结算方式记录集，格式：结算方式1|金额1||结算方式2|金额2||...
    '说明：该函数以colBalance为准计算,对于医保划价收费也是
    '说明：用于正常医保收费修改保险结算金额；及划价单医保收费设置个人帐户等结算金额
    Dim arrValue As Variant, arrPage As Variant
    Dim strTmp As String, i As Long, j As Long
    Dim varBalance As Variant, varTemp As Variant
    Dim strItem As String, curVal As Currency
    Dim blnFind As Boolean, rs结算方式 As ADODB.Recordset
    
    If strBalance = "" Then Exit Sub
    
    Set rs结算方式 = Get结算方式("收费")
    arrPage = Array()
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(varBalance(i) & "|||", "|")
        strItem = varTemp(0): curVal = Val(varTemp(1))
        '必须已设置该结算方式,且为医保类的结算方式
        rs结算方式.Filter = "名称='" & strItem & "' And 性质<>1 And 性质<>2"
        If rs结算方式.EOF Then
            '记录医保有但本地没有的结算方式
            If InStr(strNone & ",", "," & strItem & ",") = 0 Then
                strNone = strNone & "," & strItem
            End If
        Else
            If colBalance.Count > 0 Then
                If UBound(colBalance(intPage)) >= 0 Then
                    blnFind = False
                    For j = 0 To UBound(colBalance(intPage))
                        '结算方式;原始(最大)金额;可否修改;有效金额
                        arrValue = Split(colBalance(intPage)(j), ";")
                        If arrValue(0) = strItem Then blnFind = True
                        If arrValue(0) = strItem And arrValue(3) <> curVal Then
                            strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & Format(curVal, "0.00")
                        Else
                            strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & arrValue(3)
                        End If
                        
                        ReDim Preserve arrPage(UBound(arrPage) + 1)
                        arrPage(UBound(arrPage)) = strTmp
                    Next
                    
                    If Not blnFind Then
                        ReDim Preserve arrPage(UBound(arrPage) + 1)
                        '结算方式;原始(最大)金额;可否修改;有效金额
                        arrPage(UBound(arrPage)) = strItem & ";" & curVal & ";" & Val(varTemp(2)) & ";" & curVal
                    End If
                Else
                     '无内容时强行增加:不支持预结算或医保划价收费时用
                    ReDim Preserve arrPage(UBound(arrPage) + 1)
                    arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";" & Val(varTemp(2)) & ";" & Format(curVal, "0.00")
                End If
            Else
                 '无内容时强行增加:不支持预结算或医保划价收费时用
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";" & Val(varTemp(2)) & ";" & Format(curVal, "0.00")
            End If
        End If
    Next

    colBalance.Remove intPage '集合元素不能直接修改
    If colBalance.Count >= intPage Then
        colBalance.Add arrPage, , intPage
    Else
        colBalance.Add arrPage
    End If
End Sub

Public Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal strNos As String) As Boolean
    '功能:将门诊信息写入卡中
    '入参：
    '    frmMain - 调用窗体
    '    lngModul - 模块号
    '    strPrivs - 权限串
    '    objSquareCard - 医疗卡对象
    '    strNOs - 单据号，格式：'A0001','A0002','A0003',...或A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng病人ID As Long, lng结算序号 As Long
    
    Err = 0: On Error GoTo errH:
    '问题:56615
    If InStr(strPrivs, ";门诊信息写卡;") = 0 Then Exit Function
    
    strSQL = "Select Distinct A.病人ID,B.结算序号" & _
        " From 门诊费用记录 A,病人预交记录 B,Table( f_Str2list([1])) J" & _
        " Where A.结帐ID=B.结帐ID And A.NO=J.Column_Value And  Nvl(A.附加标志,0)<>9 And A.记录性质 = 1 " & _
        "       And A.记录状态 in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取单据结算序号", Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Do While Not rsTemp.EOF
        lng病人ID = Val(Nvl(rsTemp!病人ID))
        lng结算序号 = Val(Nvl(rsTemp!结算序号))
        '调用健康卡写卡接口
        If lng病人ID <> 0 And lng结算序号 <> 0 Then
            Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng病人ID, lng结算序号)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function zlSaveTempPrintData(ByVal strNos As String, ByVal lng领用ID As Long, ByVal strFactNO As String, ByRef lng打印ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存临时的打印数据
    '入参:strNos-单据号
    '     lng领用ID-领用ID
    '     strFactNo-开始发票号
    '出参:lng打印ID-返回打印ID
    '返回:
    '编制:刘兴洪
    '日期:2016-05-03 16:44:43
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    
    On Error GoTo errHandle
        '处理临时数据
    Set cllPro = New Collection
    blnTrans = True
    If SaveTempPrintDataTocCllPro(strNos, strFactNO, lng领用ID, lng打印ID, cllPro) = False Then Exit Function
    zlExecuteProcedureArrAy cllPro, "保存临时票据打印内容"
    zlSaveTempPrintData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function SaveTempPrintDataTocCllPro(ByVal strNos As String, ByVal str开始发票号 As String, ByVal lng领用ID As Long, _
    ByRef lng打印ID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将单据号信息，分解成保存到临时票据打印内容中
    '入参:strNos-单据号字符串，多个用逗号分隔
    '     str发票号-开始发票号
    '     lng领用ID-领用ID
    '出参:cllPro-保存临时票据打印内容的过程.
    '     lng打印ID-返回打印ID
    '返回:成功返回true,否则返回false
    '编制:刘兴洪
    '日期:2016-04-27 17:48:42
    '说明：产生临时的票据打印内容，主要是因为按病人补打票据时，可能因单据号超过4000，而自定义报表有所限制，因此，改为用临时表来处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim cllData As Collection
    
    On Error GoTo errHandle
    
    lng打印ID = zlDatabase.GetNextId("票据打印内容")
    Set cllData = New Collection
    If zlGetSplitString4000(strNos, cllData) = False Then Exit Function
    
    For i = 1 To cllData.Count
        '    Zl_临时票据打印内容_Insert
        strSQL = "Zl_临时票据打印内容_Insert("
        '    打印id_In     票据打印内容.Id%Type,
        strSQL = strSQL & lng打印ID & ","
        '    No_In         Varchar2,
        strSQL = strSQL & "'" & cllData(i) & "',"
        '    领用id_In     临时票据打印内容.领用id%Type,
        strSQL = strSQL & "" & lng领用ID & ","
        '    开始发票号_In 临时票据打印内容.开始票号%Type
        strSQL = strSQL & "'" & str开始发票号 & "')"
        zlAddArray cllPro, strSQL
    Next
    SaveTempPrintDataTocCllPro = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSplitString4000(ByVal strSplitData As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将strSplitData数据，按4000个字符分解，以便保存在数据库(每个分隔的字符要小于10)
    '入参:strSplitData-要分解的数据,用逗号分离
    '出参:cllPro-返回给集
    '返回:分解成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-05-04 09:43:32
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    If Len(strSplitData) <= 4000 Then
        cllPro.Add strSplitData: zlGetSplitString4000 = True: Exit Function
    End If
    
    Do While True
        If Len(strSplitData) < 4000 Then
            cllPro.Add strSplitData: zlGetSplitString4000 = True: Exit Function
        End If
        
        i = InStr(3950, strSplitData, ",")
        If i = 0 Then Exit Do
        
        strTemp = Mid(strSplitData, 1, i - 1)
        strSplitData = Mid(strSplitData, i + 1)
        cllPro.Add strTemp
    Loop
    If strSplitData <> "" Then cllPro.Add strSplitData
    zlGetSplitString4000 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Sub zlDeleteTempPrintData(lng打印ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除临时打印数据
    '入参:lng打印ID
    '编制:刘兴洪
    '日期:2016-05-03 14:36:47
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error Resume Next
    strSQL = "Zl_临时票据打印内容_Delete( " & lng打印ID & ")"
    zlDatabase.ExecuteProcedure strSQL, "删除临时票据打印内容数据"
    Err = 0
End Sub

Public Function zlIsOnePatiPrint(ByVal strNo As String, ByRef strPrintNos As String, ByRef blnOnePatiPrint As Boolean, Optional ByVal blnNOMoved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定单据是否按病人一次打印的
    '入参:strNo-需要重打NO
    '     blnNOMoved-是否转入历史表空间
    '出参:返回本次一次打印人NO,多个用逗号分离
    '     blnOnePatiPrint-如果是按病人一次打印，返回true,否则返回False
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-05-03 17:12:20
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strNos As String
    strSQL = "" & _
    "   Select A2.NO,Max(nvl(A2.打印类型,0)) as 打印类型  " & _
    "   From  票据打印内容 A1,票据打印内容 A2  " & _
    "   Where A1.ID=A2.ID and A1.数据性质=A2.数据性质  And A1.NO=[1] And A1.数据性质=1" & _
    "   Group By A2.NO"
    If blnNOMoved Then strSQL = Replace(strSQL, "票据打印内容", "H票据打印内容")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据指定单据获取一起打印的所有单据", strNo)
    blnOnePatiPrint = False
    
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            If blnOnePatiPrint = False And Val(Nvl(rsTemp!打印类型)) = 1 Then blnOnePatiPrint = True
            .MoveNext
        Loop
    End With
    strPrintNos = strNos
    If strPrintNos <> "" Then strPrintNos = Mid(strPrintNos, 2)
    zlIsOnePatiPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlFromCollectBulidSQL(ByVal cllData As Collection, ByRef strBoundSQL As String, ByRef varData() As Variant, _
    Optional ByVal strAliaName As String = "NO", Optional ByVal blnNumber As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:构建集合，构建多个绑定值的SQL
    '入参:cllData-集合
    '     strAliaName-别名
    '     blnNumber-是否数字
    '出参:strBoundSQL-绑定的SQL
    '       varData-集合值
    '返回:组合成功，返回True,否则返回False
    '编制:刘兴洪
    '日期:2016-05-04 11:58:37
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varValue() As Variant
    Dim strSQL As String, strTable As String, i As Long
    

    On Error GoTo errHandle
    ReDim varValue(0 To cllData.Count - 1)
    For i = 1 To cllData.Count
        If blnNumber Then
            strTable = "Table(f_Num2list([" & i & "]))"
        Else
            strTable = "Table(f_Str2list([" & i & "]))"
        End If
        strSQL = strSQL & _
        " UNION ALL " & vbCrLf & _
        " Select Column_Value as " & strAliaName & " From " & strTable
        If blnNumber Then
            varValue(i - 1) = Val(cllData(i))
        Else
            varValue(i - 1) = CStr(cllData(i))
        End If
    Next
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    varData = varValue
    strBoundSQL = strSQL
    zlFromCollectBulidSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlChargeBillIsAllDel(ByVal strNos As String, Optional ByVal lng打印ID As Long = 0, Optional ByRef blnAllDel As Boolean, Optional strNotDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否所有单据都全都了
    '入参:strNos-(lng打印ID=0时使用)指定的单据号,多个用逗号分隔，比如:A0001,A0002,...
    '     lng打印ID -根据打印ID来检查
    '出参:blnAllDel-全部退完，返回true,否则返回False
    '     strNotDelNos-未退完的单据号
    '返回:如果获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-05-05 11:21:31
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim varValue() As Variant, cllBound As Collection
    
    On Error GoTo errHandle
    If lng打印ID > 0 Then
        strSQL = " " & _
        "   Select b.No,B.序号, Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) As 数量  " & _
        "   From 门诊费用记录 B, (Select Distinct NO From 临时票据打印内容 Where ID = [1]) A " & _
        "   Where Mod(b.记录性质, 10) = 1 And b.No = a.No And b.价格父号 Is Null " & _
        "   Having Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) <> 0 " & _
        "   Group By b.No,B.序号"
        
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(数量,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据票据打印内容来获取单据是否全部退完", lng打印ID)
    ElseIf Len(strNos) <= 4000 Then
        strSQL = " " & _
        "   Select b.No,B.序号, Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) As 数量  " & _
        "   From 门诊费用记录 B " & _
        "   Where Mod(b.记录性质, 10) = 1  And b.价格父号 Is Null " & _
        "         And b.No in (select Column_Value From Table(f_Str2list([1])))" & _
        "   Having Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) <> 0 " & _
        "   Group By b.No,B.序号"
        
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(数量,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据票据打印内容来获取单据是否全部退完", strNos)
    Else
        If zlGetSplitString4000(strNos, cllBound) = False Then Exit Function
        If zlFromCollectBulidSQL(cllBound, strSQL, varValue) = False Then Exit Function
    
        strSQL = " With 单据信息 as (" & strSQL & ") "
        strSQL = strSQL & vbCrLf & _
            "   Select b.No,B.序号, Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) As 数量  " & _
            "   From 门诊费用记录 B,单据信息 A" & _
            "   Where Mod(b.记录性质, 10) = 1  And b.价格父号 Is Null " & _
            "         And b.No =A.NO " & _
            "   Having Sum(Nvl(b.付数, 0) * Nvl(b.数次, 0)) <> 0 " & _
            "   Group By b.No,B.序号"
        strSQL = "Select Distinct  NO From (" & strSQL & ") Where nvl(数量,0)<>0  "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "根据票据打印内容来获取单据是否全部退完", varValue)
     End If
         
     With rsTemp
        blnAllDel = True
        strNotDelNos = ""
        Do While Not .EOF
            If blnAllDel Then blnAllDel = False
            strNotDelNos = strNotDelNos & "," & Nvl(rsTemp!NO)
            .MoveNext
        Loop
        If strNotDelNos <> "" Then strNotDelNos = Mid(strNotDelNos, 2)
     End With
     zlChargeBillIsAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Upgrade医嘱执行计价执行状态(ByVal strNos As String) As Boolean
    '功能：修正"医嘱执行计价.执行状态"
    '入参：
    '   strNos 单据号，格式:A001,A002,A003,...
    '返回：已修正则返回True，否则返回False
    '问题号:99715
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str医嘱IDs As String, var医嘱IDs As Variant
    Dim i As Long
    
    On Error GoTo errHandler
    str医嘱IDs = ""
    strSQL = " Select /*+cardinality(j,10)*/ Distinct a.医嘱序号 As 医嘱ID, a.No" & vbNewLine & _
        " From 门诊费用记录 A, 病人医嘱发送 B, 医嘱执行计价 C, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where b.医嘱id = a.医嘱序号 And b.No = a.No" & vbNewLine & _
        "       And b.医嘱id = c.医嘱id And b.发送号 = c.发送号 And c.收费细目id + 0 = a.收费细目id" & vbNewLine & _
        "       And a.No = j.Column_Value And a.记录性质 = 1 And a.记录状态 In (1, 3) And a.医嘱序号 Is Not Null" & vbNewLine & _
        "       And b.记录性质 = 1 And c.执行状态 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断医嘱执行计价执行状态是否已修正", strNos)
    If rsTemp.RecordCount = 0 Then
        Upgrade医嘱执行计价执行状态 = True
        Exit Function
    End If
    
    '搜集医嘱ID
    Do While Not rsTemp.EOF
        If InStr(";" & str医嘱IDs & ";", ";" & Nvl(rsTemp!医嘱id) & "," & Nvl(rsTemp!NO) & ";") = 0 Then
            str医嘱IDs = str医嘱IDs & ";" & Nvl(rsTemp!医嘱id) & "," & Nvl(rsTemp!NO)
        End If
        rsTemp.MoveNext
    Loop
    If str医嘱IDs = "" Then
        Upgrade医嘱执行计价执行状态 = True
        Exit Function
    End If
    
    '修正数据
    str医嘱IDs = Mid(str医嘱IDs, 2)
    var医嘱IDs = Split(str医嘱IDs, ";")
    For i = 0 To UBound(var医嘱IDs)
        'Zl_医嘱执行计价_修正(
        strSQL = "Zl_医嘱执行计价_修正("
        '  医嘱id_In   病人医嘱执行.医嘱id%Type,
        strSQL = strSQL & "" & Split(var医嘱IDs(i), ",")(0) & ","
        '  No_In       病人医嘱发送.No%Type,
        strSQL = strSQL & "'" & Split(var医嘱IDs(i), ",")(1) & "',"
        '  记录性质_In 病人医嘱发送.记录性质%Type
        strSQL = strSQL & "" & "1" & ")"
        zlDatabase.ExecuteProcedure strSQL, "修正数据"
    Next
    
    Upgrade医嘱执行计价执行状态 = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePlugIn(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Function zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String) As Boolean
'功能：外挂部件出错处理，同时判断是否为非接口方法不存在的错误
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    Err.Clear
End Function

Public Function CreatePublicDrug(ByVal lngSys As Long, _
    cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
    '功能：动态创建药品公共部件
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    If Not gobjPublicDrug Is Nothing Then CreatePublicDrug = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjPublicDrug = CreateObject("zlPublicDrug.clsPublicDrug")
    
    Err = 0: On Error GoTo errHandler
    If gobjPublicDrug Is Nothing Then
        MsgBox "药品公共部件（zlPublicDrug）创建失败，请与系统管员联系！", vbInformation, gstrSysName
        Exit Function
    End If
    'Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPublicDrug.zlInitCommon(lngSys, cnOracle, strDBUser) = False Then
        MsgBox "药品公共部件（zlPublicDrug）初始化失败，请与系统管员联系！", vbInformation, gstrSysName
        Set gobjPublicDrug = Nothing: Exit Function
    End If
    CreatePublicDrug = True
    Exit Function
errHandler:
    Set gobjPublicDrug = Nothing
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckChargeItemByPlugIn(objPlugIn As Object, _
    lngSys As Long, ByVal lngModule As Long, _
    ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据外挂部件对收费项目有效性进行检查
    '入参:lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '     intType:0-门诊;1-住院
    '     intMode:0-录入明细时的常规检查;1-保存单据前的汇总检查
    '     rsDetail-病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室,
    '                  执行科室ID、单据性质（1-收费单,2-记帐单)、是否划价(1-划价;0-正常的收费及记帐单)
    '     strExpend-待以后扩展，暂无用
    '出参:strExpend-待以后扩展，暂无用
    '返回:数据合法返回true,否则返回False
    '编制:冉俊明
    '日期:2017-04-19 10:09:26
    '问题号:105189
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    '1.没有外挂部件时，认为检查通过
    '2.外挂部件中无CheckChargeItem接口，也认为检查通过
    If objPlugIn Is Nothing Then CheckChargeItemByPlugIn = True: Exit Function
    
    On Error Resume Next
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intType, intMode, rsDetail, strExpend) = False Then
        '注意，接口不存在时也会进入
        If Err <> 0 Then
            If Err.Number = 438 Then '接口不存在，认为检查通过
                CheckChargeItemByPlugIn = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "CheckChargeItem")
        End If
        Exit Function
    End If
    CheckChargeItemByPlugIn = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PatiValiedCheckByPlugIn(ByVal lngModule As Long, ByVal lng病人ID As Long) As Boolean
    '调用外挂接口 PatiValiedCheck 检查病人信息
    '问题号:102234,138602
    '说明：
    '   1.没有外挂部件时，认为检查通过
    '   2.外挂部件中无PatiValiedCheck接口，也认为检查通过
    '   3.未建档病人不检查
    
    If gobjPlugIn Is Nothing Then PatiValiedCheckByPlugIn = True: Exit Function
    If lng病人ID = 0 Then PatiValiedCheckByPlugIn = True: Exit Function
    
    On Error Resume Next
    'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
    '功能：检查当前病人是否是指定的特殊病人
    '返回：true时允许继续操作，False时不允许操作
    '参数：
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐。
    '      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
    '      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
    '      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号"，出生日期 格式:2016-11-11 12:12:12
    '                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
    '      strReserve=保留参数,用于扩展使用
    On Error Resume Next
    If gobjPlugIn.PatiValiedCheck(glngSys, lngModule, 3, lng病人ID, 0, "") = False Then
        '注意，接口不存在时也会进入
        If Err <> 0 Then
            If Err.Number = 438 Then '接口不存在，认为检查通过
                PatiValiedCheckByPlugIn = True: Exit Function
            End If
            Call zlPlugInErrH(Err, "PatiValiedCheck")
        End If
        Exit Function
    End If
    PatiValiedCheckByPlugIn = True
End Function

Public Function GetPriceGradeFromNos(ByVal strNos As String, Optional ByVal lng结帐ID As Long) As String
    '功能：根据单据号及站点获取普通项目价格等级
    '入参：
    '   strNos 单据号，带引号，可能是多个单据号，为"'AAA','BBB',..."的形式
    Dim strPriceGrade As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varNos As Variant
    
    On Error GoTo errHandler
    If lng结帐ID = 0 Then
        varNos = Split(Replace(strNos, "'", ""), ",")
        If UBound(varNos) = -1 Then Exit Function
        
        strSQL = _
            "Select a.病人id, b.名称 As 付款方式" & vbNewLine & _
            "From 门诊费用记录 A, 医疗付款方式 B" & vbNewLine & _
            "Where a.付款方式 = b.编码 And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.No = [1] And Rownum < 2"
    Else
        strSQL = _
            "Select a.病人id, b.名称 As 付款方式" & vbNewLine & _
            "From 门诊费用记录 A, 医疗付款方式 B" & vbNewLine & _
            "Where a.付款方式 = b.编码 And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And a.结帐ID = [2] And Rownum < 2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人ID及医疗付款方式", CStr(varNos(0)), lng结帐ID)
    If rsTemp.EOF Then Exit Function
    
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(rsTemp!病人ID)), 0, Nvl(rsTemp!付款方式), , , strPriceGrade)

    GetPriceGradeFromNos = strPriceGrade
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    gstr药品价格等级 = "": gstr卫材价格等级 = "": gstr普通价格等级 = ""
    If gobjPublicExpense Is Nothing Then
        Set gobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If gobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    gintPriceGradeStartType = gobjPublicExpense.zlGetPriceGradeStartType()
    If gintPriceGradeStartType = 0 Then Exit Sub
    '读取站点价格等级
    Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", gstr药品价格等级, gstr卫材价格等级, gstr普通价格等级)
End Sub

Public Function UserIsClinic(ByVal lng人员ID As Long) As Boolean
    '判断当前操作员是否为临床部门人员
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = _
        "Select 1" & vbNewLine & _
        "From 部门人员 A,部门表 B, 部门性质说明 C" & vbNewLine & _
        "Where a.部门id = b.Id And b.id = c.部门id" & vbNewLine & _
        "      And c.工作性质 In ('临床', '检查', '检验', '手术', '治疗', '产科')" & vbNewLine & _
        "      And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
        "      And (b.站点='" & gstrNodeNo & "' Or b.站点 is Null) " & vbNewLine & _
        "      And a.人员id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lng人员ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    UserIsClinic = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlSelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
     ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示成套项目选择器(选择器入口)
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '出参:rsOutSel-成功时,返回选择的成套项目(有字段:细目ID,编码,名称,序号,从属父号,执行科室....)
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-08 16:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPublicExpense Is Nothing Then
        Call CreatePublicExpenseObject(lngModule)
    End If
    If gobjPublicExpense Is Nothing Then Exit Function
    zlSelectWholeItems = gobjPublicExpense.zlSelectWholeItems(frmMain, lngModule, strPrivs, rsOutSel)
End Function

Public Sub ZlShowBillFormat(ByVal lngModule As Long, lblFormat As Label, ByVal intFormat As Integer)
    '功能：显示票据格式名称
    '入参：
    '   lngModule - 模块号
    '   lblFormat - 显示票据格式的标签对象
    '   intFormat - 票据格式序号
    '返回：票据格式的名称
    Dim strFormatName As String
    
    On Error GoTo errHandler
    strFormatName = ZlGetBillFormat(lngModule, intFormat)
    If strFormatName = "" Then
        lblFormat.Visible = False
    Else
        lblFormat.Caption = "票据:" & strFormatName
        lblFormat.Visible = True
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ZlGetBillFormat(ByVal lngModule As Long, ByVal intFormat As Integer) As String
    '功能：获取票据格式名称
    '入参：
    '   lngModule - 模块号
    '   intFormat - 票据格式序号
    '返回：票据格式的名称
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo errHandler
    If lngModule = 1124 Then
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1124"
    Else
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1121_1"
    End If
    
    If intFormat = 0 Then '以缺省票据格式显示
        intFormat = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    strSQL = _
        "Select b.说明" & vbNewLine & _
        "From zlReports A, zlRPTFMTs B" & vbNewLine & _
        "Where a.Id = b.报表id And a.编号 = [1] And b.序号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取报表格式说明", strRptName, intFormat)
    If rsTmp.EOF Then Exit Function
    
    ZlGetBillFormat = Nvl(rsTmp!说明)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlChargeSaveValied_Plugin(ByVal lngModule As Long, ByVal int记录性质 As Integer, ByVal bln门诊 As Boolean, _
    ByVal bln划价单 As Boolean, ByVal strNos As String, ByVal rsSaveItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用外挂，检查保存数据的合法性
    '入参:lngModule-模块号
    '     int记录性质-1-收费单;2-记帐单
    '     bln划价单-是否当前是保存的划价单
    '     strNOs-门诊收费时，传入的划价单号（对本次收费的划价单号)
    '     rsSaveItems=当前保存的项目集，(字段 :病人ID，主页ID,单据序号, 序号,价格父号,收费细目ID，收入项目id，付数 ，数次，标准单价，应收金额 ，
    '                                            实收金额，发生时间，项目编码，项目名称，费用类别,开单部门ID,开单人,执行部门ID)
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.没有外挂部件时，认为检查通过
    '2.外挂部件中无CheckChargeItem接口，也认为检查通过
    If gobjPlugIn Is Nothing Then zlChargeSaveValied_Plugin = True: Exit Function
    
    On Error Resume Next
    If gobjPlugIn.ChargeSaveValied(glngSys, lngModule, int记录性质, bln门诊, bln划价单, strNos, rsSaveItems) = False Then
        '注意，接口不存在时也会进入
        If Err <> 0 Then
            If Err.Number = 438 Then '接口不存在，认为检查通过
                zlChargeSaveValied_Plugin = True
                Exit Function
            End If
            Call zlPlugInErrH(Err, "ChargeSaveValied")
            Err = 0: On Error GoTo 0
        End If
        Exit Function
    End If
    zlChargeSaveValied_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng病人ID, ByVal lng主页ID As Long, _
                                    ByVal bln门诊 As Boolean, ByVal int记录性质 As Integer, ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用外挂，检查保存数据的合法性
    '入参:     lngSys , lngModual = 当前调用接口的主程序系统号及模块号
    '   lng病人ID（记帐表时，传入0)
    '   lng主页ID（记帐表时，传入0)
    '   bln门诊 -是否门诊费用
    '   int记录性质-1-收费;2-记帐
    '   strNOs-单据号,多个用逗号分隔
    '编制:刘兴洪
    '日期:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1.没有外挂部件时，认为检查通过
    '2.外挂部件中无CheckChargeItem接口，也认为检查通过
    If gobjPlugIn Is Nothing Then Exit Sub
    
    On Error Resume Next
    Call gobjPlugIn.ChargeSaveAfter(glngSys, lngModule, lng病人ID, lng主页ID, bln门诊, int记录性质, strNos)
    If Err = 0 Then Exit Sub
    
    '注意，接口不存在时也会进入
    If Err.Number = 438 Then Exit Sub  '接口不存在，认为检查通过
    Call zlPlugInErrH(Err, "ChargeSaveAfter")
    Err = 0: On Error GoTo 0
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef str划价Nos As String, ByRef rsItems As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前需要保存的数据明细(此过程，主要是应用于外挂接口,如果没有外挂号，则直接返回True,记录集返回Nothing)
    '入参:objBills-单据对象
    '出参:str划价Nos-返回当前收费所涉及的划价单
    '     rsItems-返回当前需要保存的数据集(字段 :病人ID，主页ID,单据序号, 序号,价格父号,收费细目ID，收入项目id，付数 ，数次，标准单价，应收金额 ，
    '                                            实收金额，发生时间，项目编码，项目名称，费用类别,开单部门ID,开单人,执行部门ID)
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-14 11:41:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBillDetail As BillDetail  '单据的收费细目对象
    Dim objBillIncome As BillInCome
    Dim int价格父号 As Integer
    Dim p As Long, int序号 As Integer
    
    On Error GoTo errHandle
    
    Set rsItems = Nothing
    str划价Nos = ""
    
    If gobjPlugIn Is Nothing Then zlGetSaveDataItems_Plugin = True: Exit Function
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "序号", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "价格父号", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "收费项目ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "收入项目ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "付数", adDouble, , adFldIsNullable
    rsItems.Fields.Append "数次", adDouble, , adFldIsNullable
    rsItems.Fields.Append "标准单价", adDouble, , adFldIsNullable
    rsItems.Fields.Append "应收金额", adDouble, , adFldIsNullable
    rsItems.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsItems.Fields.Append "发生时间", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "项目编码", adVarChar, 30, adFldIsNullable
    rsItems.Fields.Append "项目名称", adVarChar, 200, adFldIsNullable
    rsItems.Fields.Append "费用类别", adVarChar, 2, adFldIsNullable
    rsItems.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "开单人", adVarChar, 20, adFldIsNullable
    rsItems.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
    
     '对每张单据独立执行保存
    For p = 1 To objBills.Pages.Count
       
        If objBills.Pages(p).NO = "" Then
            int序号 = 0
            For Each objBillDetail In objBills.Pages(p).Details
                If objBillDetail.数次 <> 0 Then
                    int价格父号 = 0
                    For Each objBillIncome In objBillDetail.InComes
                      int序号 = int序号 + 1 '当前记录序号
                       rsItems.AddNew
                       rsItems!病人ID = objBills.病人ID
                       rsItems!主页ID = objBills.主页ID
                       rsItems!单据序号 = p
                       rsItems!序号 = int序号
                       rsItems!价格父号 = IIf(int价格父号 = 0, Null, int序号)
                       rsItems!收费项目ID = objBillDetail.收费细目ID
                       rsItems!收入项目ID = objBillIncome.收入项目ID
                       rsItems!付数 = objBillDetail.付数
                       rsItems!数次 = objBillDetail.数次
                       rsItems!标准单价 = objBillIncome.标准单价
                       rsItems!应收金额 = objBillIncome.应收金额
                       rsItems!实收金额 = objBillIncome.实收金额
                       rsItems!发生时间 = Format(objBills.发生时间, "yyyy-mm-dd HH:MM:SS")
                       rsItems!项目编码 = objBillDetail.Detail.编码
                       rsItems!项目名称 = objBillDetail.Detail.名称
                       rsItems!费用类别 = objBillDetail.收费类别
                       rsItems!执行部门ID = objBillDetail.执行部门ID
                       rsItems!开单部门ID = objBills.Pages(p).开单部门ID
                       rsItems!开单人 = objBills.Pages(p).开单人
                       rsItems.Update
                      If int价格父号 = 0 Then int价格父号 = int序号
                    Next     '每一行收费项目
                End If
            Next
        Else
            str划价Nos = str划价Nos & "," & objBills.Pages(p).NO
        End If
    Next  '下一张单
    If str划价Nos <> "" Then str划价Nos = Mid(str划价Nos, 2)
    
    zlGetSaveDataItems_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


