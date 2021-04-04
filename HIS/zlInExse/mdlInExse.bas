Attribute VB_Name = "mdlInExse"
Option Explicit '要求变量声明
Public Enum gBalanceBill
    g_Ed_门诊结帐 = 0
    g_Ed_住院结帐 = 1
    g_Ed_重新结帐 = 2
    g_Ed_取消结帐 = 3
    g_Ed_结帐作废 = 4
    g_Ed_重新作废 = 5
    g_Ed_单据查看 = 6
End Enum

Public gcolPrivs As Collection              '记录内部模块的权限
'============费用系统参数=====================
'医保控制f
Public gclsInsure As New clsInsure
Public gstr医保费用类型 As String '医保病人允许的费用类型
Public gstr公费费用类型 As String '公费病人允许的费用类型
Public gbyt医保对码检查 As Byte '0-不进行检查、1-检查并提醒未对码项目、2-检查并禁止未对码项目
Public gbln简码切换 As Boolean '35242
'刷卡控制
Public gdbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbyt预存款退费验卡 As Byte '预存款退费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
Public gbln消费卡退费验卡 As Boolean '消费卡退费时是否刷卡验证

'LED控制
Public gblnLED As Boolean        '结帐时是否启用LED设备报价
Public gblnLedWelcome As Boolean '是否在结帐输完病人后提示欢迎信息
Public gobjKernel As Object
'票据控制
Public gblnStrictCtrl As Boolean '是否严格票据管理
Public gbytFactLength As Byte '票据号码长度

Public gobjBillPrint As Object '第三方票据打印部件
Public gblnBillPrint As Boolean '第三方票据打印部件是否可用

Public gobjTax As Object '税控打印接口对象
Public gblnTax As Boolean '本机是否使用税控打印
Public gstrTax As String
Public gblnNurseStation As Boolean
Public gblnPrintByPatient As Boolean '合约单位结帐按病人分别打印票据
Public gbytInvoiceKind As Byte      '0-住院医疗费收据,1-门诊医疗费收据
Public gbytFeePrintSet As Byte      '0-不打印;1-打印提示;2-打印但不提示

'费用计算控制
Public gBytMoney As Byte '收费分币处理方法
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gbln从项汇总折扣 As Boolean '从属项目汇总计算折扣
Public gbln住院单位 As Boolean      '药品按住院单位,否则按售价单位
Public gcurMaxMoney As Currency '单笔费用最大提醒金额


'药房相关控制
Public gblnStock As Boolean '指定药房时是否限定输入药品的库存
Public gbln分离发药 As Boolean '是否门诊收费与发药分离
Public gint卫材发料控制 As Integer    '记帐完成后是否自动发料:0-不自动发料，1-自动发料，2-本科室开单时自动发料

'药房、窗口控制
Public glng西药房 As Long '指定的西药房,0为动态分配
Public glng中药房 As Long '指定的中药房,0为动态分配
Public glng成药房 As Long '指定的成药房,0为动态分配
Public glng发料部门 As Long '指定的卫材发料部门,0为动态分配
Public gbln药房上班安排 As Boolean '是否启用了上班安排
Public gbytSendMateria As Byte '0-记帐后不发药,1-自动发药,2-提示发药
Public gbytMediOutMode As Byte '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
Public gbln不显示无库存卫材 As Boolean

'分离发药时要检查库存的药房
Public gstr西药房 As String
Public gstr成药房 As String
Public gstr中药房 As String

Public gbln其它药房 As Boolean '是否显示其它药房库存
Public gbln其它药库 As Boolean '是否显示其它药库库存


'单据输入控制
Public gstrLike As String  '项目匹配方法,%或空
Public gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
Public gblnMyStyle As Boolean '使用个性化风格

Public gint诊断输入 As Integer
Public gbln收费类别 As Boolean '是否首先输入类别
Public gblnFeeKindCode As Boolean '不输类别时,首位当作收费类别简码

Public gbln开单人 As Boolean '记帐是否必须输入开单人
Public gbln它科人 As Boolean '记帐是否可以输入它科的开单人
Public gstrMatchMode As String  '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求

'中药输入快捷
Public grsABCNum As ADODB.Recordset
Public gstrABC As String '输入允许的快捷字母

'操作控制
Public gbytBilling As Byte '0-记帐,1-划价,2-审核
Public gstrModiNO As String '修改后产生的新单据号
Public gbytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
Public gbytWarn As Byte '记帐报警返回值
Public gbln报警包含划价费用 As Boolean '记帐报警包含划价费用

Public gblnPrice As Boolean '是否允许保存存为划价单
Public gbln门诊留观 As Boolean '留观病人记帐
Public gbln住院留观 As Boolean
Public gbln每次住院新住院号 As Boolean
Public gobjPati As Object

'医嘱相关
Public gbln药疗划价单 As Boolean
Public gbln其他划价单 As Boolean
Public gbln执行后审核 As Boolean

'输入控制
Public gstr收费类别 As String '可输入的收费类别
Public gblnPay As Boolean '中药是否输入付数
Public gblnTime As Boolean '变价是否输入付数
Public gbln护士 As Boolean '开单人是否显示护士
Public gblnFromDr As Boolean '开单人确定科室
Public gbyt开单人显示 As Byte
'刘兴洪 问题:????    日期:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000

'打印控制
Public gbln记帐打印 As Boolean
Public gbln划价打印 As Boolean
Public gbln审核打印 As Boolean
 
'医技执行
Public gbln本人执行 As Boolean '执行者本人登记
Public gblnExe医嘱 As Boolean
Public gstrExe来源 As String
Public gstrExe类别 As String
Public gbytExe门诊单据类型 As Byte
Public gbytExe住院单据类型 As Byte
Public gbytExe体检单据类型 As Byte
Public gbytExe打印方式 As Byte
Public gbln执行后发料 As Boolean            '跟踪在用的卫材在执行后自动发料,取消执行后自动退料

Public gobjCustBill As Object               '自定义记帐单对象
Public gbln处方限量 As Boolean             '处方限量检查,如果输入时选择了允许,则保存时,不再检查
Public grs收费类别 As ADODB.Recordset
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--定义系统参数
Private Type TY_System_para_Balance
    bln刷卡输入密码 As Boolean  '是否刷卡输入密码
    bln在院不准结帐 As Boolean '1-在院不准结帐,0-在院允许结帐
    bytAuditing As Byte  '记帐未审核单据的结帐处理:0-不检查,1-检查并提示,2-检查并禁止
    byt检查未执行 As Byte    '出院和结帐出院时检查是否有未执行项目及未发药品:0-不检查,1-检查并提示,2-检查并禁止
    byt检查未发药 As Byte   '在出院结帐及病人入出管理中出院时是否检查病人的未发药品项目:0-不检查,1-检查并提示,2-检查并禁止
    byt门诊检查未执行 As Byte    '门诊结帐时检查是否有未执行项目及未发药品:0-不检查,1-检查并提示,2-检查并禁止
    byt门诊检查未发药 As Byte   '门诊结帐时是否检查病人的未发药品项目:0-不检查,1-检查并提示,2-检查并禁止
    bln医生允许才能出院 As Boolean '医生下达出院医嘱才允许病人出院
End Type

Private Type Ty_System_Para
     byt药品名称显示 As Byte   '药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
     byt输入药品显示 As Byte  '输入药品显示（通过输入简码方式进入选择器时药品名称的显示）：0-按输入匹配显示，1-固定显示通用名和商品名
     int数据补录时限 As Integer '数据补录时限
     byt病人审核方式 As Byte '49501:病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
     bln未入科禁止记账 As Boolean '51612
     byt条码卫材识别控制 As Byte   '是否仅条码识别::1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
     strCardPass As String '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
     TY_Balance As TY_System_para_Balance
End Type
Public gTy_System_Para As Ty_System_Para


'==============结帐参数===============
'结帐参数
Public gblnAutoOut As Boolean '在院病人结帐后是否自动出院
Public gbln医生允许才能出院 As Boolean '医生下达出院医嘱才允许病人出院
Public gintOutDay As Integer '结帐可选择出院病人天数
Public gbytAuditing As Byte  '记帐未审核单据的结帐处理:0-不检查,1-检查并提示,2-检查并禁止
Public gblnZero As Boolean '结帐时是否处理零费用
Public gstrCardPass As String '刷卡时要求输入密码,'0000000000'依位顺序表示各个场合,分别为:1.门诊挂号,2.门诊划价,3.门诊收费,4.门诊记帐,5.入院登记,6.住院记帐,7.病人结帐,8.病人预交款,9.检验技师站,10.影像医技站.'
Public gbyt检查未执行 As Byte    '出院和结帐出院时检查是否有未执行项目及未发药品:0-不检查,1-检查并提示,2-检查并禁止
Public gbyt检查未发药 As Byte   '在出院结帐及病人入出管理中出院时是否检查病人的未发药品项目:0-不检查,1-检查并提示,2-检查并禁止
Public gint费用时间 As Integer '0-按登记时间,1-按发生时间
Public gbln在院不准结帐 As Boolean '1-在院不准结帐,0-在院允许结帐
Public gbln仅用指定预交款 As Boolean  '仅使用指定住院次数的预交款
Public gbln多次住院弹出结帐设置 As Boolean '有多次住院费用的病人自动弹出结帐设置
Public gbyt结帐检查代收款项 As Byte '出院结帐时检查病人的代收款项,0-禁止,1-提醒
Public gbln中途结帐退预交 As Boolean '中途结帐缺省退预交款
Public gstr结算方式显示顺序 As String   '32322
Public gbyt结帐时输血费检查 As Byte   '34260
'=======系统控制相关变量============
Public Enum 医院业务
    support门诊预算 = 0
    
    support门诊退费 = 1
    support预交退个人帐户 = 2
    
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
    support出院无实际交易 = 29       '出院接口中是否要与接口商进行交易
    support允许部分冲销明细 = 32    '允许针对住院记帐处方的每笔明细进行部分冲销
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support住院结算作废 = 34        'HIS始终认为住院支持结算作废，如果不支持需医保接口内部处理，返回假即可；增加该参数是为了配合GetCapability交易来检查各种结算方式是否支持全退
    support负数记帐 = 35            '是否允许负数记帐，操作员首先要拥有负数记帐的权限。此参数缺省为真，不支持的接口需单独处理
    support结帐_指定住院次数 = 36   '是否支持指定住院次数进行医保结算
    support结帐_指定日期范围 = 37   '是否支持指定结帐日期范围进行医保结算
    support结帐_设置婴儿费条件 = 38 '是否允许设置婴儿费条件
    
    support门诊结帐 = 41            '是否支持门诊医保病人的记帐费用使用门诊结帐来完成
    support结帐_指定科室 = 42           '是否允许在结帐设置界面中指定科室
    support结帐_指定费用项目 = 43       '是否允许在结帐设置界面中指定费用项目
    support结帐_结帐设置后调用接口 = 44 '如果为真则在结帐设置后才调用住院虚拟结算，之前不调用
    support结帐_指定费用类型 = 45        '是否允许在结帐设置界面中指定费用类型
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support多单据一次结算 = 47      '多单据预结算时，医保接口仅在最后一次调用时返回结算结果，HIS中再分摊到每张单据上
        
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support门诊结帐_结帐设置后调用接口 = 49          '是否在结帐设置后才调用门诊虚拟结算接口,此参数为true，表示允许门诊结帐时进行结帐条件设置，且设置后调用虚拟结算接口
    support实时监控 = 60             '是否启用费用实时监控
    support退费后打印回单 = 65   '医保病人是否退费后打印回单:问题
    support结帐作废后打印回单 = 66      '结帐作废后打印回单
    support批量中途结帐 = 84   '医保是否支持批量中途结帐:81661
    support允许一次结多次住院费用 = 88  '允许住院病人对多次住院费用进行一次结算,问题号:114915
End Enum

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
Private mlng部门编码平均长度 As Long
Public gobjSquare As SquareCard  '卡结算部件  42301
Public gobjPlugIn As Object
Public gobjPublicDrug As Object '药品公共部件,105875
Public gobjPublicExpense As Object  '费用公共部件
Public gobjPublicExpenseBillOperation As Object
Public gintPriceGradeStartType As Integer
Public gstr药品价格等级 As String
Public gstr卫材价格等级 As String
Public gstr普通价格等级 As String
Public gobjCharge As Object '门诊费用部件 zl9OutExse.clsOutExse

Public glngInstanceCount As Long '进程总数

Public Function MakeDetailRecord(ByRef objBill As ExpenseBill, ByVal str开单人 As String, ByVal str开单科室 As String, _
    ByVal intMode As Integer, ByVal intPrice As Integer, Optional ByVal lngRow As Long) As ADODB.Recordset
'功能：根据单据对象内容创建一个明细记录集信息(以售价单位)
'字段：病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室,执行科室ID,
'          单据性质（1-收费单,2-记帐单),是否划价(1-划价;0-正常的收费及记帐单)
'参数：intPage=指定的单据,lngRow=指定的行，不指定时包含所有单据的所有行
'          intMode:单据性质（1-收费单,2-记帐单)
'          intPrice:是否划价(1-划价;0-正常的收费及记帐单)

    Dim i As Integer, j As Integer
    Dim intB As Integer, intE As Integer, blnNew As Boolean
    Dim dbl单价 As Double, cur实收 As Currency
    Dim rsTmp As New ADODB.Recordset
    
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
    
    
    If lngRow = 0 Then
        intB = 1
        intE = objBill.Details.Count
    Else
        intB = lngRow
        intE = lngRow
    End If
    
    For i = intB To intE
        dbl单价 = 0: cur实收 = 0
        With objBill.Details(i)
            If lngRow = 0 Then
                If .病人ID = 0 Then
                    rsTmp.Filter = "收费细目ID=" & .收费细目ID
                Else    '记帐表
                    rsTmp.Filter = "病人ID=" & .病人ID & " And 收费细目ID=" & .收费细目ID
                End If
                blnNew = rsTmp.RecordCount = 0
            Else
                blnNew = True
            End If
                            
            If blnNew Then
                rsTmp.AddNew
                
                If .病人ID = 0 Then
                    rsTmp!病人ID = objBill.病人ID
                    rsTmp!主页ID = objBill.主页ID
                Else    '记帐表
                    rsTmp!病人ID = .病人ID
                    rsTmp!主页ID = .主页ID
                End If
                
                rsTmp!收费类别 = .收费类别
                rsTmp!收费细目ID = .收费细目ID
                rsTmp!执行科室ID = .执行部门ID
                rsTmp!单据性质 = intMode
                rsTmp!是否划价 = intPrice
                
                For j = 1 To .InComes.Count
                    dbl单价 = dbl单价 + .InComes(j).标准单价
                    cur实收 = cur实收 + .InComes(j).实收金额
                Next
                If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                    '从药房单位转换为售价单位
                    rsTmp!数量 = IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.住院包装
                    rsTmp!单价 = Format(dbl单价 / .Detail.住院包装, gstrFeePrecisionFmt)
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
                If InStr(",5,6,7,", .收费类别) > 0 And gbln住院单位 Then
                    '从药房单位转换为售价单位
                    rsTmp!数量 = rsTmp!数量 + IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.住院包装
                    rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价 / .Detail.住院包装, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                Else
                    rsTmp!数量 = rsTmp!数量 + IIf(.付数 = 0, 1, .付数) * .数次
                    rsTmp!单价 = Format((rsTmp!单价 + Format(dbl单价, gstrFeePrecisionFmt)) / 2, gstrFeePrecisionFmt)
                End If
                rsTmp!实收金额 = rsTmp!实收金额 + Format(cur实收, gstrDec)
            End If
            
            rsTmp.Update
        End With
    Next
    
    rsTmp.Filter = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeDetailRecord = rsTmp
End Function

Public Function GetVBalance(ByVal byt业务类型 As Byte, strPrivs As String, int险类 As Integer, lng病人ID As Long, Optional strTime As String, _
     Optional DateBegin As Date, Optional DateEnd As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional bytBaby As Byte, Optional blnOnly门诊 As Boolean, Optional bytKind As Byte, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, Optional strChargeType As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定病人未结帐细目明细(按收费细目)
    '入参：lng病人ID-病人ID,
    '      strTime： 医保病人只能设置住院次数和费用期间 [strTime=住院次数串,"0,1,2,3",0表示门诊]
    '      DateBegin,DateEnd： 结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      bytBaby：0-所有费用,1-病人费用,2以及上-第bytBaby-1个婴儿费用]
    '      blnOnly门诊：仅门诊记帐费用
    '      bytKind：0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    '      strItem:收据费目串,'西药费','中药费',...
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：所有费用(含未设置),"'类型1','类型2',..."
    '      byt业务类型-0-门诊业务;1-住院业务
    '出参：
    '返回：成功=记录集,失败=Nothing
    '编制：刘兴洪
    '日期：2010-03-06 10:39:50
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String, bytType As Byte '0-门诊,1-住院,2-门诊和住院
    Dim strWherePage As String '住院次数条件
    Dim strWhereMzPage As String
    On Error GoTo errH
    
    strPrivs = ";" & strPrivs & ";"
    'Modified by ZYB 2002-10-30
    blnRelation = gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类)
     
    strCond = " And A.病人ID=[1]"
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
    strWhereMzPage = IIf(strTime = "", "", " And Instr([2],',0,')>0")   '36004
    
    If DateBegin <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(gint费用时间 = 0, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
 
    '刘兴洪:2010-03-06 11:23:52: Or A.结帐ID is Not NULL 这个包含了已结帐的明细,在我的分析来看,是错的,但又没听说医保又问题,因此暂不更改!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null Or A.结帐ID is Not NULL)"
    
    strCond = strCond & IIf(bytBaby = 0, "", IIf(bytBaby = 1, " And Nvl(A.婴儿费,0)=0", " And A.婴儿费=[6]"))
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.收费类别||''',')>0")   '34260
    
    If bytKind = 1 Then
        strCond = strCond & " And A.门诊标志=4"
    Else
        If InStr(strPrivs, ";住院费用结帐;") = 0 Or blnOnly门诊 Then strCond = strCond & " And A.门诊标志<>2"
        If InStr(strPrivs, ";门诊费用结帐;") = 0 Then strCond = strCond & " And A.门诊标志<>1"
        If bytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
    End If
    If byt业务类型 = 0 Or bytKind = 1 Then
        bytType = 0
    Else
        bytType = 1 '42027
    End If
''    '获取费用获取范围类型
''    If bytKind = 1 Then '仅体检费用
''        bytType = 0
''    ElseIf (InStr(strPrivs, ";住院费用结帐;") = 0 Or blnOnly门诊) Then  '门诊部分的处理
''            If InStr(strPrivs, ";门诊费用结帐;") = 0 Then
''                '无权限,又处理门诊结帐数据的:
''                ' a: 3-其他(就诊卡等额外的收费);4-体检
''                bytType = IIf(bytKind = 0, 1, 0) '如果是就诊卡,就读住院费用记录,否则读门诊费用记录
''            Else
''                '有门诊结算权限
''                'a: 1-门诊,3-其他(就诊卡等额外的收费);4-体检
''                bytType = IIf(bytKind = 0, 2, 0)
''            End If
''    ElseIf InStr(strPrivs, ";门诊费用结帐;") = 0 Then    '住院结算,但不能结帐门诊的
''        '2-住院;3-其他(就诊卡等额外的收费);4-体检
''        bytType = IIf(bytKind = 0, 1, 2)
''    Else  '门诊和住院
''        '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
''        bytType = 2
''    End If
    '结算要求：记录性质,记录状态,NO、序号、收费类别、收费细目ID,收费名称、计算单位、开单部门、规格、产地、数量、价格、金额、医生,
    '          发生时间,登记时间,婴儿费,医保项目编码、保险大类ID、保险项目否、是否上传,是否急诊
    '注意：由于结算只能针对有保险项目编码的,所以在与保险支付项目连接时不用(+)
    '   数次为零指按费别重算后产生的打折冲减记录,这类单据明细不上传
    
    '临时更改：结帐作废后该SQL不对,单价暂时改为"金额/数量"
    If blnOnly门诊 Then
        '门诊结帐
        '不传单据号,只有门诊收费时取划价单才用到单据号
        '一正一负的冲销费用不传,虽然不存在单笔部分结帐的情况,但仍然要用sum(实收金额)-sum(结帐金额),因为结帐作废产生的记录没有实收金额
        If bytType = 2 Then
            strTable = "" & _
            "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
            "              Avg(Nvl(付数, 0) * 数次) As 数量, Avg(标准单价) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
            "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & vbNewLine & _
            "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0 " & strCond & strWhereMzPage & vbNewLine & _
            "       Group By NO, Mod(记录性质, 10), 记录状态, Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊," & vbNewLine & _
            "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
            "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 " & _
            "       UNION ALL " & vbCrLf & _
            "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
            "              Avg(Nvl(付数, 0) * 数次) As 数量, Avg(标准单价) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
            "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & vbNewLine & _
            "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0 " & strCond & strWherePage & vbNewLine & _
            "       Group By NO, Mod(记录性质, 10), 记录状态, Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊," & vbNewLine & _
            "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
            "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 "
            
        Else
            If bytType = 0 Then
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A")
            Else
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A")
            End If
            
            strTable = "" & _
            "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
            "              Avg(Nvl(付数, 0) * 数次) As 数量, Avg(标准单价) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
            "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
            "       From " & strTable & vbNewLine & _
            "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0" & IIf(bytType = 1, " And A.主页ID Is Not Null ", "") & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & vbNewLine & _
            "       Group By NO, Mod(记录性质, 10), 记录状态, Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊," & vbNewLine & _
            "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
            "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 "
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
                "              Avg(Nvl(付数, 0) * 数次) As 数量, Avg(标准单价) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
                "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
                "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & vbNewLine & _
                "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0 And (Mod(A.记录性质,10) = 5 And A.主页ID Is Null)" & strCond & vbNewLine & _
                "       Group By NO, Mod(记录性质, 10), 记录状态, Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 摘要, 是否急诊," & vbNewLine & _
                "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
                "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 "
            End If
        End If
        
        strSQL = "" & _
        " Select Sysdate As 结算时间, A.病人id, A.收费类别, A.收据费目, A.计算单位, A.收费细目id," & vbNewLine & _
        "       B.大类id 保险支付大类id, B.是否医保 是否医保, B.项目编码 保险编码, Sum(A.数量) As 数量, Avg(A.单价) As 单价," & vbNewLine & _
        "       Sum(A.实收金额) As 实收金额, Sum(A.统筹金额) As 统筹金额, Max(A.摘要) 摘要, Max(A.是否急诊) 是否急诊," & vbNewLine & _
        "       Max(A.开单部门id) 开单部门id, Max(A.执行部门id) 执行部门id, Max(A.开单人) 开单人,Max(A.费用类型) 费用类型" & vbNewLine & _
        " From ( " & strTable & ") A, 保险支付项目 B, 收费项目目录 C " & vbNewLine & _
        " Where A.收费细目id = C.ID And A.收费细目id = B.收费细目id" & IIf(blnRelation, "(+)", "") & " And B.险类" & IIf(blnRelation, "(+)", "") & " = [5] " & vbNewLine & _
                    IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
        " Group By A.收费细目id, A.病人id, A.收费类别, A.收据费目, A.计算单位, B.大类id, B.是否医保, B.项目编码" & vbNewLine & _
        " Having Sum(A.实收金额) <> 0"
    Else
        '住院结帐
        If bytType = 2 Then
            strTable = ""
            If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '包含门诊:33189
                    strTable = "" & _
                    "       Select Mod(A.记录性质, 10) As 记录性质, A.记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
                    "              -1*NULL as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
                    "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
                    "              Avg(A.标准单价) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
                    "              A.登记时间, Nvl(A.是否上传, 0) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, A.摘要,A.费用类型" & vbNewLine & _
                    "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & ", 收入项目 B" & vbNewLine & _
                    "       Where A.记录状态 <> 0 And A.记帐费用 = 1 And A.收入项目id = B.ID And A.数次 <> 0 " & strCond & strWhereMzPage & vbNewLine & _
                    "       Group By Mod(A.记录性质, 10), A.记录状态, A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id," & vbNewLine & _
                    "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
                    "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否上传, 0), Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),　A.摘要,A.费用类型" & vbNewLine & _
                    "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　" & _
                    "       UNION ALL " & vbCrLf
            End If
            strTable = strTable & "" & _
            "       Select Mod(A.记录性质, 10) As 记录性质, A.记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
            "              A.主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
            "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
            "              Avg(A.标准单价) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
            "              A.登记时间, Nvl(A.是否上传, 0) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,　Nvl(A.保险项目否, 0) As 保险项目否, A.摘要,A.费用类型" & vbNewLine & _
            "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " , 收入项目 B" & vbNewLine & _
            "       Where A.记录状态 <> 0 And A.记帐费用 = 1 And A.收入项目id = B.ID And A.数次 <> 0 " & strCond & strWherePage & vbNewLine & _
            "       Group By Mod(A.记录性质, 10), A.记录状态, A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id, A.主页id," & vbNewLine & _
            "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
            "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否上传, 0), Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),A.摘要,A.费用类型" & vbNewLine & _
            "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0"
        Else
            If bytType = 0 Then
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A")
            Else
                strTable = IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A")
            End If
            strTable = "" & _
            "       Select Mod(A.记录性质, 10) As 记录性质, A.记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
            "              " & IIf(bytType = 0, "-1*NULL", "A.主页id") & " as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
            "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
            "              Avg(A.标准单价) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
            "              A.登记时间, Nvl(A.是否上传, 0) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, A.摘要,A.费用类型" & vbNewLine & _
            "       From " & strTable & " , 收入项目 B" & vbNewLine & _
            "       Where A.记录状态 <> 0 And A.记帐费用 = 1" & IIf(bytType = 1, " And A.主页ID Is Not Null ", "") & " And A.收入项目id = B.ID And A.数次 <> 0 " & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & vbNewLine & _
            "       Group By Mod(A.记录性质, 10), A.记录状态, A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id" & IIf(bytType = 0, "", ", A.主页id") & "," & vbNewLine & _
            "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
            "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否上传, 0), Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),　A.摘要,A.费用类型" & vbNewLine & _
            "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　"
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                "       Select Mod(A.记录性质, 10) As 记录性质, A.记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
                "              " & IIf(bytType = 0, "-1*NULL", "A.主页id") & " as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
                "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
                "              Avg(A.标准单价) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
                "              A.登记时间, Nvl(A.是否上传, 0) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, A.摘要,A.费用类型" & vbNewLine & _
                "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " , 收入项目 B" & vbNewLine & _
                "       Where A.记录状态 <> 0 And A.记帐费用 = 1 And (Mod(A.记录性质,10) = 5 And A.主页ID Is Null) And A.收入项目id = B.ID And A.数次 <> 0 " & strCond & vbNewLine & _
                "       Group By Mod(A.记录性质, 10), A.记录状态, A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id" & IIf(bytType = 0, "", ", A.主页id") & "," & vbNewLine & _
                "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
                "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否上传, 0), Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),　A.摘要,A.费用类型" & vbNewLine & _
                "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　"
            End If
        End If
        
        strSQL = "Select A.记录性质, A.记录状态, A.NO, A.序号, A.门诊标志, A.病人id, A.主页id, A.婴儿费, C.项目编码 As 医保项目编码," & vbNewLine & _
                "       A.保险编码, A.保险大类id, A.收费类别, A.收费细目id, Nvl(E.名称, B.名称) As 收费名称, A.计算单位," & vbNewLine & _
                "       X.名称 As 开单部门, B.规格, B.产地, A.数量, A.标准单价 As 价格, A.金额," & vbNewLine & _
                "       A.医生, A.发生时间, A.登记时间, A.是否上传, A.是否急诊, A.保险项目否, A.摘要,A.费用类型" & vbNewLine & _
                "From ( " & strTable & ") A, 收费项目目录 B, 保险支付项目 C, 收费项目别名 E,部门表 X" & vbNewLine & _
                "Where A.收费细目id = B.ID And B.ID = C.收费细目id" & IIf(blnRelation, "(+)", "") & " And C.险类" & IIf(blnRelation, "(+)", "") & " = [5] And A.开单部门id = X.ID " & vbNewLine & _
                        IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.费用类型,Nvl(B.费用类型,'无'))||''',')>0") & _
                "      And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1)
    End If
    Set GetVBalance = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, "," & strTime & ",", DateBegin, DateEnd, int险类, bytBaby - 1, "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    '问题:strDeptIDs:42478
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBalance(ByVal byt业务类型 As Byte, strPrivs As String, lng病人ID As Long, Optional strTime As String, Optional strDeptIDs As String, _
    Optional strClass As String, Optional DateBegin As Date, Optional DateEnd As Date, Optional bytBaby As Byte, Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
    Optional blnOnly门诊 As Boolean, Optional bytKind As Byte, _
    Optional bln消费卡启用 As Boolean = False, Optional strChargeType As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定病人未结帐金额细目(按每收入项目行)
    '入参：lng病人ID-病人ID,
    '      strTime：住院次数串,"0,1,2,3",0表示门诊
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：""-所有费用(含未设置),"'类型1','类型2',..."
    '      strItem：收据费目串,'西药费','中药费',...
    '      bytBaby：0-所有费用,1-病人费用,2以及上-第bytBaby-1个婴儿费用
    '      DateBegin,DateEnd：结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnZero：是否读取零费用
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      blnOnly门诊：仅门诊记帐费用
    '      bytKind：  0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    '      bln消费卡启用-启用了消费卡(需要返回一些字段:A.)
    '     strChargeType:""表示所有费用,否则为指定收费类别的费用;如:5,6,7等  '34260
    '     byt业务类型-0-门诊业务;1-住院业务
    '出参：
    '返回：成功=记录集,失败=Nothing
    '编制：刘兴洪
    '日期：2010-03-06 13:21:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, strCond As String, strCond2 As String
    Dim strTable As String, bytType As Byte '0-门诊,1-住院,2-门诊和住院
    Dim strWherePage As String '住院次数条件
    Dim strWhereMzPage As String
        
    strCond = " And A.病人ID=[1]"
    
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
    strWhereMzPage = IIf(strTime = "", "", " And Instr([2],',0,')>0")   '门诊
    
    If Not DateBegin = CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(gint费用时间 = 0, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([5],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(bytBaby = 0, "", IIf(bytBaby = 1, " And Nvl(A.婴儿费,0)=0", " And A.婴儿费=[6]"))
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([9],','''||A.收费类别||''',')>0")   '34260
    
    If bytKind = 1 Then
        strCond = strCond & " And A.门诊标志=4"
    Else
        If InStr(strPrivs, ";住院费用结帐;") = 0 Or blnOnly门诊 Then strCond = strCond & " And A.门诊标志<>2"
        If InStr(strPrivs, ";门诊费用结帐;") = 0 Then strCond = strCond & " And A.门诊标志<>1"
        If bytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
    End If
        
    
    strCond2 = strCond   '已经结过帐的,不管是否上传都要取,所以先把这个条件记录下来,第二个子查询用
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null) "
    Else
        strCond = strCond & " And A.结帐ID Is Null "
    End If
    
    ' bytKind:0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    If byt业务类型 = 0 Or bytKind = 1 Then
        bytType = 0
    Else
         bytType = 1
    End If

''    '获取费用获取范围类型
''    If bytKind = 1 Then '仅体检费用
''        bytType = 0
''    ElseIf (InStr(strPrivs, "住院费用结帐") = 0 Or blnOnly门诊) Then  '门诊部分的处理
''            If InStr(strPrivs, "门诊费用结帐") = 0 Then
''                '无权限,又处理门诊结帐数据的:
''                ' a: 3-其他(就诊卡等额外的收费);4-体检
''                bytType = IIf(bytKind = 0, 1, 0) '如果是就诊卡,就读住院费用记录,否则读门诊费用记录
''            Else
''                '有门诊结算权限
''                'a: 1-门诊,3-其他(就诊卡等额外的收费);4-体检
''                bytType = IIf(bytKind = 0, 2, 0)
''            End If
''    ElseIf InStr(strPrivs, "门诊费用结帐") = 0 Then    '住院结算,但不能结帐门诊的
''        '2-住院;3-其他(就诊卡等额外的收费);4-体检
''        bytType = IIf(bytKind = 0, 1, 2)
''    Else  '门诊和住院
''        '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
''        bytType = 2
''    End If
    
    
    '住院,科室,时间,[单据号],项目,费目,婴儿费,[ID],[序号],[记录性质],[记录状态],[执行状态],[A.主页ID],[A.开单部门ID],[登记时间],未结金额,结帐金额,[类型]
    If blnZero Then
        If bytType = 2 Then
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,1 as 标志,'门诊' as 住院," & _
            "                -1*NULL as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                     IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
            " From 门诊费用记录 A " & _
            " Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWhereMzPage & _
            " Union all " & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次') as 住院," & _
            "                A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                     IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
            " From 住院费用记录 A " & _
            " Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWherePage & _
            ""
        Else
            If bytType = 0 Then
                    strTable = " From 门诊费用记录 A Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWhereMzPage
            Else
                    strTable = " From 住院费用记录 A Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWherePage
            End If
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态," & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & _
            "               " & IIf(bytType = 0, " -1*NULL", "A.主页ID") & " as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                     IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
            " " & strTable & IIf(bytType = 1, " And A.主页ID Is Not Null ", "")
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态," & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & _
                "               " & IIf(bytType = 0, " -1*NULL", "A.主页ID") & " as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
                "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                         IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
                " From 住院费用记录 A Where A.记录状态<>0 And (Mod(A.记录性质,10) = 5 And A.主页ID Is Null) And A.记帐费用=1" & strCond
            End If
        End If
    Else
    
        '该子查询用于过滤掉第一次结帐时一正一负的费用
        '后续结帐时,即使一正一负,也要拿出来结了
        If bytType = 2 Then
            strTable = ""
            If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '包含门诊
                strTable = "" & _
                " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,1 as 标志,'门诊'  as 住院," & _
                "               -1*NULL as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
                "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                             IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
                " From  门诊费用记录 A," & _
                "      ( Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
                "        From  门诊费用记录 A" & _
                "        Where A.记录状态<>0  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0" & strCond & _
                "        Group by A.NO,A.序号,A.记录性质 Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
                " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond & _
                " Union ALL "
            End If
            strTable = strTable & "" & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次') as 住院," & _
            "         A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "         Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                      IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
            " From  住院费用记录 A," & _
            "      ( Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
            "        From  住院费用记录 A" & _
            "        Where A.记录状态<>0  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0" & strCond & strWherePage & _
            "        Group by A.NO,A.序号,A.记录性质 Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
            " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond & strWherePage
            
        Else
            strTable = "" & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态," & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & _
            "                " & IIf(bytType = 0, "-1*NULL", "A.主页ID") & " as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                         IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
            " From  " & IIf(bytType = 0, "门诊费用记录 A", " 住院费用记录 A") & "," & _
            "      ( Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
            "        From " & IIf(bytType = 0, "门诊费用记录 A", " 住院费用记录 A") & _
            "        Where A.记录状态<>0  And A.记帐费用=1" & IIf(bytType = 1, " And A.主页ID Is Not Null ", "") & " And Nvl(A.实收金额,0)<>0" & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage) & _
            "        Group by A.NO,A.序号,A.记录性质 Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
            " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond & IIf(bytType = 0, strWhereMzPage, strWherePage)
            
            If bytType = 0 Then
                strTable = strTable & " Union " & _
                " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态," & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & _
                "                " & IIf(bytType = 0, "-1*NULL", "A.主页ID") & " as 主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
                "                Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目,Nvl(A.实收金额,0) as 未结金额,费用类型, A.收费类别" & _
                             IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数次 * A.付数 As 数量, A.标准单价, A.统筹金额, A.保险大类id", "") & _
                " From   住院费用记录 A" & "," & _
                "      ( Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
                "        From 住院费用记录 A" & _
                "        Where A.记录状态<>0 And (Mod(A.记录性质,10) = 5 And A.主页ID Is Null)  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0" & strCond & _
                "        Group by A.NO,A.序号,A.记录性质 Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
                " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond
            End If
        End If
    End If
    
    If bytType = 2 Then
        strSQL = ""
        If InStr(1, "," & strTime & ",", ",0,") > 0 Or strTime = "" Then  '包含门诊
            strSQL = "" & _
            "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,A.记录状态,A.执行状态,1 as 标志," & _
            "                '门诊'  as 住院,-1*NULL as 主页ID," & _
            "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
            "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额,A.费用类型 , max(A.收费类别) as 收费类别" & _
                            IIf(bln消费卡启用, ",max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人,avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价, avg(A.统筹金额) as 统筹金额,max( A.保险大类id) as 保险大类id ", "") & _
            "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A") & " " & _
            "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 And (Nvl(A.实收金额, 0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)" & strCond2 & strWhereMzPage & _
            "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0  " & _
            "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and Sum(Nvl(A.结帐金额,0)) =0 And Mod(Count(*),2)=0) " & _
            "                    Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
            "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.记录状态,A.执行状态," & _
            "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型" & _
            "         Union all  "
        End If
        strSQL = strSQL & _
        "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,A.记录状态,A.执行状态,2 as 标志," & _
        "               Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次') as 住院,A.主页ID," & _
        "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
        "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额,A.费用类型, max(A.收费类别) as 收费类别 " & _
                        IIf(bln消费卡启用, ",max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人,avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价, avg(A.统筹金额) as 统筹金额,max( A.保险大类id) as 保险大类id ", "") & _
        "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " " & _
        "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 And (Nvl(A.实收金额, 0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)" & strCond2 & strWherePage & _
        "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
        "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and Sum(Nvl(A.结帐金额,0)) =0 And Mod(Count(*),2)=0) " & _
        "                    Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
        "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.记录状态,A.执行状态,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次'),A.主页ID," & _
        "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型" & _
        ""
    Else
        If bytType = 0 Then
            strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录", 2, "", True, ""), "门诊费用记录")
        Else
            strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", 2, "", True, ""), "住院费用记录")
        End If
        
        strSQL = "" & _
        "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,A.记录状态,A.执行状态," & _
        "               " & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & IIf(bytType = 0, "-1*NULL", "A.主页ID") & " as 主页ID," & _
        "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
        "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额,A.费用类型, max(A.收费类别) as 收费类别 " & _
                        IIf(bln消费卡启用, ",max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人,avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价, avg(A.统筹金额) as 统筹金额,max( A.保险大类id) as 保险大类id ", "") & _
        "        FROM " & strSQL & " A" & _
        "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 " & IIf(bytType = 1, " And A.主页ID Is Not Null ", "") & " And (Nvl(A.实收金额, 0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)" & strCond2 & IIf(bytType = 0, strWhereMzPage, strWherePage) & _
        "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
        "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and  Sum(Nvl(A.结帐金额,0))=0  And Mod(Count(*),2)=0) " & _
        "                     Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
        "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.记录状态,A.执行状态," & IIf(bytType = 0, "", "Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次'),A.主页ID,") & _
        "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型" & _
        ""
        If bytType = 0 Then
            strSQL = strSQL & " Union " & _
            "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,A.记录状态,A.执行状态," & _
            "               " & IIf(bytType = 0, "1 as 标志,'门诊'", "2 as 标志,Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次')") & " as 住院," & IIf(bytType = 0, "-1*NULL", "A.主页ID") & " as 主页ID," & _
            "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
            "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额,A.费用类型, max(A.收费类别) as 收费类别 " & _
                            IIf(bln消费卡启用, ",max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人,avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价, avg(A.统筹金额) as 统筹金额,max( A.保险大类id) as 保险大类id ", "") & _
            "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", 2, "", True, ""), "住院费用记录") & " A" & _
            "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 And (Mod(A.记录性质,10) = 5 And A.主页ID Is Null) And (Nvl(A.实收金额, 0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)" & strCond2 & _
            "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
            "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and  Sum(Nvl(A.结帐金额,0))=0  And Mod(Count(*),2)=0) " & _
            "                     Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
            "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.记录状态,A.执行状态," & IIf(bytType = 0, "", "Decode(A.主页ID,NULL,'门诊','第'||A.主页ID||'次'),A.主页ID,") & _
            "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型" & _
            ""
        End If
    End If
    
    '
    '问题:48305,61527: 多增加 And Mod(Count(*),2)=0
    '   Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and  Sum(Nvl(A.结帐金额,0))=0 ) " & _

    strTable = strTable & " Union ALL " & strSQL
    
    strSQL = _
        "Select A.标志,A.住院,Nvl(B.名称,'未知') as 科室,A.时间,A.NO as 单据号 ,Nvl(E.名称,C.名称) as 项目,A.收据费目 as 费目, A.婴儿费,A.ID,A.序号,A.记录性质,A.记录状态,A.执行状态,A.主页ID,A.开单部门ID,A.登记时间," & _
        "       Nvl(A.未结金额,0) 未结金额,Nvl(A.未结金额,0) 结帐金额,Nvl(A.费用类型,C.费用类型) as 类型, A.收费类别" & _
                IIf(bln消费卡启用, ",A.费别, A.执行部门id, A.开单人,A.数量, A.标准单价 as 价格, A.统筹金额, A.保险大类id,A.收费细目ID,C.计算单位", "") & _
        " From (  " & strTable & ") A,部门表 B,收费项目目录 C,收入项目 D,收费项目别名 E " & _
        " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID And A.收入项目ID=D.ID " & IIf(strClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        " Order by A.时间 Desc,A.住院,A.NO Desc,A.记录性质,A.序号"
    
    'Mod(Count(*),2)=1是为了区别打折后实收金额为零的费用在结帐后是否作废或再次结帐
    On Error GoTo errH
    Set GetBalance = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, "," & strTime & ",", DateBegin, DateEnd, _
                    "," & strDeptIDs & ",", bytBaby - 1, "," & strItem & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人余额(lng病人ID As Long, bytType As Byte, Optional int预交类别 As Integer = 2) As Currency
'功能：获取指定病人的预交款余额
'参数：bytType:0-费用余额,1-预交余额
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    strSQL = "Select Nvl(sum(费用余额),0) as 费用余额,Nvl(sum(预交余额),0) as 预交余额" & _
        " From 病人余额 Where 性质=1 And 病人ID=[1]  " & IIf(int预交类别 = 0, "", " And 类型=[2] ")
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, int预交类别)
    If Not rsTmp.EOF Then Get病人余额 = IIf(bytType = 0, rsTmp!费用余额, rsTmp!预交余额)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function Chk病人审核(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：判断病人是否已审核
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Nvl(审核标志,0) as 审核标志" & _
        " From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    
    '49501
    If gTy_System_Para.byt病人审核方式 = 0 Then
        Chk病人审核 = (rsTmp!审核标志 >= 1)
    Else
        Chk病人审核 = (rsTmp!审核标志 > 1)
    End If

    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function check医生下达出院医嘱(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：判断病人是否处于预出院状态,且存在有效的出院(转院、死亡)医嘱才允许出院(有效的医嘱是指开始执行时间与预出院时间相同，且处于已发送状态[医嘱状态=8])。
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select a.Id" & vbNewLine & _
            "From 病人医嘱记录 a, 病人变动记录 b, 病案主页 c, 诊疗项目目录 d" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And a.医嘱状态 = 8 And a.病人id = b.病人id And a.主页id = b.主页id And" & vbNewLine & _
            "           a.开始执行时间 = b.开始时间+0 And b.开始原因 = 10 And b.病人id = c.病人id And b.主页id = c.主页id And" & vbNewLine & _
            "           c.状态 = 3 And d.类别='Z' And d.操作类型 In ('5', '6', '11') And a.诊疗项目id = d.Id"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    check医生下达出院医嘱 = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillingWarn(strPrivs As String, str姓名 As String, lng病区ID As Long, str适用病人 As String, _
    rsWarn As ADODB.Recordset, cur余额 As Currency, cur当日额 As Currency, _
    cur单据金额 As Currency, cur担保 As Currency, str类别 As String, _
    ByVal str类别名 As String, ByRef str已报类别 As String, _
    Optional bln多病人 As Boolean, Optional ByVal blnPrice As Boolean, _
    Optional curItemMoney As Currency = 0, _
    Optional blnNotCheck类别 As Boolean = False) As Integer
'功能:对病人记帐进行报警提示
'参数:
'     str姓名=病人姓名,用于提示
'     lng病区ID=病人病区ID,用于选择适用的病区报警设置，0表示没有确定病区，仅所有病区的报警设置适用
'     str适用病人=根据病人身份返回的记帐报警适用方案
'     rsWarn=当前病区记帐报警设置记录
'     cur余额=病人余额,用于累计报警
'     cur当日额=病人当日发生的费用额,用于每日报警
'     cur单据金额=病人单据中输入的费用
'     cur担保=病人担保费用额,用于累计报警
'     str类别=当前要检查的类别,用于分类报警
'     str类别名=类别名称,用于提示
'     blnPrice=欠费时是否允许强制保存为划价单,用于记帐或划价有
'     curItemMoney-当笔金额(如果传入<>0 ,则需要判断当笔情况,如果超出金额,则允许用户继续,否则根据报警方式进行):刘兴洪:24491
'     blnNotCheck类别:不对类别进行检查(主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
'返回:0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
'     5.报警提示后用户选择继续,但只允许保存存为划价单
'     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    Dim i As Integer, byt标志 As Byte
    Dim bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    
    '报警参数检查
    If rsWarn.State = 0 Then Exit Function '20030709
    rsWarn.Filter = "适用病人='" & str适用病人 & "' And 病区ID=" & lng病区ID
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志1 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志2 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志3 <> "-" And blnNotCheck类别 Then Exit Function
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
                        '如果已报方式存在两种:低于报警值时,预交耗尽时,则只需最后一种,因为最后一种总是最后发生
                        'Exit For
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
                            '如果已报类别存在两种:低于报警值时,预交耗尽时,则只需最后一种,因为最后一种总是最后发生
                            'Exit For
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
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？" & _
                            IIf(gbytBilling = 0 And blnPrice, vbCrLf & vbCrLf & "提示:你可以选择继续记帐,并将当前单据保存为划价单,等病人缴费后再审核。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = IIf(gbytBilling = 0, 5, 1)  '1  :问题:28515
                        End If
                    Else
                        If gbytBilling = 0 And blnPrice Then
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & _
                                Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。", vbInformation, gstrSysName
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        End If
                        BillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur余额 + cur担保 - cur单据金额 < 0 Then
                    
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > 0 Then
                             '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                            If MsgBox("注意" & vbCrLf & _
                                       "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                         
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If blnPrice Then
                                If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。" & _
                                    vbCrLf & vbCrLf & IIf(gbytBilling = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    BillingWarn = 2
                                Else
                                    BillingWarn = IIf(gbytBilling = 0, 5, 1)
                                End If
                            Else
                                MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                BillingWarn = 3
                            End If
                        Else
                            If gbytBilling = 0 And blnPrice Then
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。" & _
                                    vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。", vbInformation, gstrSysName
                            Else
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                            End If
                            BillingWarn = 4
                        End If
                    ElseIf cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                        '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                         If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                             '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                             '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                            If MsgBox("注意" & vbCrLf & _
                                       "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ", 是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                            End If
                            Exit Function
                         End If
                        
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？" & _
                                IIf(gbytBilling = 0 And blnPrice, vbCrLf & vbCrLf & "提示:你可以选择继续记帐,并将当前单据保存为划价单,等病人缴费后再审核。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = 1
                            End If
                        Else
                            If gbytBilling = 0 And blnPrice Then
                                MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & _
                                    Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                    vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。", vbInformation, gstrSysName
                            Else
                                MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                            End If
                            BillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur余额 + cur担保 - cur单据金额 < 0 Then
                            '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                             If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > 0 Then
                                 '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                                 '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                                If MsgBox("注意" & vbCrLf & _
                                           "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                                    BillingWarn = 2
                                Else
                                    BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                                End If
                                Exit Function
                             End If
                             
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                                If blnPrice Then
                                    If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。" & _
                                        vbCrLf & vbCrLf & IIf(gbytBilling = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        BillingWarn = 2
                                    Else
                                        BillingWarn = IIf(gbytBilling = 0, 5, 1)
                                    End If
                                Else
                                    MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                    BillingWarn = 3
                                End If
                            Else
                                If gbytBilling = 0 And blnPrice Then
                                    MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。" & _
                                        vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。", vbInformation, gstrSysName
                                Else
                                    MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                                End If
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
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur余额 + cur担保 - (cur单据金额 - curItemMoney) > Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽,是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
            
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        If blnPrice Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。" & _
                                vbCrLf & vbCrLf & IIf(gbytBilling = 0, "要将当前单据保存为划价单吗？", "要强制保存划价单吗？"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                BillingWarn = 2
                            Else
                                BillingWarn = IIf(gbytBilling = 0, 5, 1)
                            End If
                        Else
                            MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                            BillingWarn = 3
                        End If
                    Else
                        If gbytBilling = 0 And blnPrice Then
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。" & _
                                vbCrLf & vbCrLf & "提示:你可以选择将当前单据保存为划价单,等病人缴费后再审核。", vbInformation, gstrSysName
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        End If
                        BillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur当日额 + cur单据金额 - curItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                     
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        If MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1
                        End If
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        BillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    '24491 刘兴洪:表示预交款扣除本笔金额已经耗尽  ,则按原来的规则来处理,否则只提示
                     If curItemMoney <> 0 And cur当日额 + cur单据金额 - curItemMoney < Val(Nvl(rsWarn!报警值)) Then
                         '只提示: gbytBilling As Byte '0-记帐,1-划价,2-审核
                         '主要是检查如下情况:可能录入当笔项目时，要更改数次、附加项目等，从而导致实收金额的减少，就有不可能满足报警条件
                        If MsgBox("注意" & vbCrLf & _
                                   "    病人“" & str姓名 & "” 当日费用:" & Format(cur当日额 + cur单据金额, gstrDec) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",是否继续录入该项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            BillingWarn = 2
                        Else
                            BillingWarn = 1 ' IIf(gbytBilling = 0, 0, 1)
                        End If
                        Exit Function
                     End If
                     
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
    If BillingWarn = 1 Or BillingWarn = 4 Or BillingWarn = 5 Then
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


Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long) As Double
'功能：获取单据中指定药品在同一药房多行的数量合
'参数： lng药房ID-0表示分离发药时,不限定药房检查
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).收费细目ID = lng药品ID Then
            If IIf(lng药房ID <> 0, objBill.Details(i).执行部门ID = lng药房ID, 1 = 1) Then
                dblCount = dblCount + objBill.Details(i).付数 * objBill.Details(i).数次
            End If
        End If
    Next
    GetDrugTotal = dblCount
End Function

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long) As Double
'功能：获取单据中指定药品在同一药房多行的原始数量和
'参数： lng药房ID-0表示分离发药时,不限定药房检查
    Dim i As Integer, dblCount As Double
    
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).收费细目ID = lng药品ID Then
            If IIf(lng药房ID <> 0, objBill.Details(i).原始执行部门ID = lng药房ID, 1 = 1) Then
                dblCount = dblCount + objBill.Details(i).原始数量
            End If
        End If
    Next
    GetOriginalTotal = dblCount
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

Public Sub NurseDeposit(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    Optional bln余额退款 As Boolean = True, Optional ByVal bytPrepayType As Byte = 2)
    '调用预交款管理
    '入参：
    '   bytPrepayType-预交类型(0-门诊和住院;1-门诊;2-住院)
    On Error GoTo errH
    If gobjPati Is Nothing Then
        Set gobjPati = CreateObject("zl9Patient.clsPatient")

    End If
    If gobjPati Is Nothing Then Exit Sub
    
    Call gobjPati.NurseDeposit(glngSys, gcnOracle, frmMain, gstrDBUser, lng病人ID, lng主页ID, bln余额退款, bytPrepayType)
    Set gobjPati = Nothing
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetMultiStock(ByVal lng药品ID As Long, ByVal str药房IDs As String) As Double
'功能：获取指定药房指定药品库存(以零售单位)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    '获取库存(不分批或分批),药房不分批(批次=0,这里为药房)不管效期
    strSQL = _
        " Select Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
        " Where (Nvl(A.批次,0)=0 Or A.效期 is NULL Or A.效期>Trunc(Sysdate))" & _
        " And A.性质=1 And A.药品ID=[1] And Instr([2],','||A.库房ID||',')>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng药品ID, "," & str药房IDs & ",")
    If Not rsTmp.EOF Then GetMultiStock = rsTmp!库存
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckDisable(objBill As ExpenseBill) As String
'功能：检查单据中的药品的禁忌情况
'返回：药品互相禁忌提示信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim i As Long, j As Long, k As Long
    Dim strGroup As String, strIDs As String
    Dim blnStop As Boolean
    
    Err = 0: On Error GoTo errH:
    For i = 1 To objBill.Details.Count
        If InStr(",5,6,7,", objBill.Details(i).收费类别) > 0 Then
            strIDs = strIDs & "," & objBill.Details(i).收费细目ID
        End If
    Next
    strIDs = Mid(strIDs, 2)
    If strIDs = "" Or UBound(Split(strIDs, ",")) < 1 Then Exit Function
    
    strSQL = _
        " Select /*+ RULE */  A.组编号,Count(Distinct A.项目ID) as 禁忌数" & _
        " From 诊疗互斥项目 A,药品规格 B," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        " Where A.项目ID=B.药名ID And B.药品ID  = j.Column_Value" & _
        " Having Count(Distinct A.项目ID)>1  " & _
        "  Group by A.组编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strIDs)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            strGroup = strGroup & "," & rsTmp!组编号
            rsTmp.MoveNext
        Next
        strGroup = Mid(strGroup, 2)
        
        For i = 0 To UBound(Split(strGroup, ","))
            strSQL = _
            "Select /*+ RULE */ Distinct C.类型,C.组编号,D.编码,D.名称,D.规格" & _
            " From 药品规格 A,诊疗项目目录 B,诊疗互斥项目 C,收费项目目录 D," & _
            "          (Select Column_Value From Table(Cast(f_num2list([2]) As Zltools.t_Numlist ))) J " & _
            " Where A.药名ID=B.ID And B.ID=C.项目ID And A.药品ID=D.ID" & _
            "           And C.组编号=[1]" & _
            "           And A.药品ID=  j.Column_Value " & _
            " Order by C.类型,C.组编号,D.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", Val(Split(strGroup, ",")(i)), strIDs)
            
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
            If blnStop Then
                CheckDisable = "发现单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo & vbCrLf & "请修改禁用药品后再继续！"
            Else
                CheckDisable = "发现单据中下列药品互相禁用或慎用：" & vbCrLf & strInfo & vbCrLf & "要继续吗？"
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetDeposit(lng病人ID As Long, _
    Optional blnDateMoved As Boolean, Optional strTime As String, _
    Optional ByVal bln门诊转住院 As Boolean = False, _
    Optional ByVal strPepositDate As String = "", _
    Optional int预交类别 As Integer = 0, _
    Optional rs结算方式 As ADODB.Recordset) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人剩余预交款明细
    '入参:strTime-住院次数,如:1,2,3
    '        bln门诊转住院-是否门诊费用转住院(只能充指定的预交)
    '        strPepositDate-指定的预交日期
    '       int预交类别-0-门诊和住院;1-门诊;2- 住院
    '出参:
    '返回: 预交明细数据
    '编制:刘兴洪
    '日期:2011-03-31 14:58:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strSub1 As String
    Dim strWherePage As String, strTable As String
    Dim strWhere As String, strDate As String
    Dim strPara As String, intPara As Integer
    Dim str排序 As String, strData() As String
    Dim int方式 As Integer, rsDeposit As ADODB.Recordset
    Dim i As Integer
    Dim str类型排序 As String, int类型排序 As Integer
    Dim str条件排序 As String, strDecode As String
    Dim strHead As String
    On Error GoTo errH

    strSQL = ""
    If int预交类别 = 1 Then strTime = ""    '69500
    
    strWherePage = IIf(strTime = "", "", " And instr(','||[2]||',',','||Nvl(A.主页ID,0)||',')>0")
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("病人预交记录"), "病人预交记录 A")
    strWhere = "": strDate = "2000-01-01 00:00:00"
    
    strPara = zlDatabase.GetPara("冲预交缺省顺序", glngSys, 1137, "0")
    intPara = Val(Split(strPara & "|", "|")(0))
    
    If strPepositDate <> "" Then
        If IsDate(strPepositDate) Then
            strDate = strPepositDate
            strWhere = " And A.收款时间=[3]"
        End If
    End If
    
    If int预交类别 <> 0 Then
        strWhere = strWhere & " And A.预交类别 =[4]"
    End If
    
    If bln门诊转住院 Then
        strWhere = strWhere & " And A.摘要='门诊转住院预交'"
    End If

    If intPara = 0 Then
        '默认排序
        '性质=5:代扣费
        strSQL = "" & _
        " Select a.No, a.票据号, a.Id, a.金额, a.记录状态, a.预交id, a.日期, a.结算方式, " & vbNewLine & _
        "       a.卡类别id, a.结算卡序号, decode(nvl(A.结算卡序号,0),0,0,1)  as 是否消费卡,a.卡号, a.交易流水号, a.交易说明, " & vbNewLine & _
        "       c.是否转帐及代扣 As 转帐及代扣,  Nvl(c.名称, q.名称)  As 卡类别名称, Nvl(C.是否退现, Q.是否退现)  As 是否退现, " & vbNewLine & _
        "       Nvl(C.是否全退, Q.是否全退) As 是否全退, c.是否缺省退现," & vbNewLine & _
        "       b.性质 As 结算性质,  Sign(Nvl(a.金额, 0)) As 标志" & vbNewLine & _
        " From (Select a.No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态," & vbNewLine & _
        "              Min(Decode(a.结帐id, Null, a.Id, 0) * Decode(a.记录状态, 1, 1, 0)) As ID," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.Id), 0)) As 预交id," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.实际票号), Null)) As 票据号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, To_Char(a.收款时间, 'yyyy-mm-dd hh24:mi:ss')), Null)) As 日期," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.结算方式), Null)) As 结算方式," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.卡类别id), Null)) As 卡类别id," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.结算卡序号), Null)) As 结算卡序号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.卡号), Null)) As 卡号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.交易流水号), Null)) As 交易流水号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.交易说明), Null)) As 交易说明" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where a.记录性质 In (1, 11) And a.病人id = [1] " & strWhere & strWherePage & _
        "       Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0" & vbNewLine & _
        "       Group By a.No) A, 结算方式 B, 医疗卡类别 C,消费卡类别目录 Q" & vbNewLine & _
        " Where a.结算方式 = b.名称(+) And a.卡类别id = c.Id(+) And a.结算卡序号 = q.编号(+) And b.性质 <> 5" & vbNewLine & _
        " Order By 标志, 日期, NO"
    Else
        '按结算类别排序
        strData = Split(Split(strPara & "|", "|")(1), ",")
        int类型排序 = 1
        For i = 0 To UBound(strData)
            int方式 = Val(Split(strData(i) & ":", ":")(1))
            If Split(strData(i) & ":", ":")(0) = "现金类结算" Then
                If int类型排序 = 1 Then
                    str类型排序 = "Decode(Nvl(b.性质,0),1,1"
                Else
                    str类型排序 = str类型排序 & ",1," & int类型排序
                End If
                
                Select Case int方式
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 1, A.日期, Null) As 排序现金"
                    str条件排序 = str条件排序 & ",排序现金"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 1, A.金额, Null) As 排序现金"
                    str条件排序 = str条件排序 & ",排序现金"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 1, A.金额, Null) As 排序现金"
                    str条件排序 = str条件排序 & ",排序现金 Desc"
                End Select
                
                int类型排序 = int类型排序 + 1
            End If
            
            If Split(strData(i) & ":", ":")(0) = "其他类结算" Then
                If int类型排序 = 1 Then
                    str类型排序 = "Decode(Nvl(b.性质,0),2,1,3,1,4,1,6,1,7,1"
                Else
                    str类型排序 = str类型排序 & ",2," & int类型排序 & ",3," & int类型排序 & ",4," & int类型排序 & ",6," & int类型排序 & ",7," & int类型排序
                End If
                
                Select Case int方式
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 2, A.日期, 3, A.日期, 4, A.日期, 6, A.日期, 7, A.日期, Null) As 排序其他"
                    str条件排序 = str条件排序 & ",排序其他"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 2, A.金额, 3, A.金额, 4, A.金额, 6, A.金额, 7, A.金额, Null) As 排序其他"
                    str条件排序 = str条件排序 & ",排序其他"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 2, A.金额, 3, A.金额, 4, A.金额, 6, A.金额, 7, A.金额, Null) As 排序其他"
                    str条件排序 = str条件排序 & ",排序其他 Desc"
                End Select
                
                int类型排序 = int类型排序 + 1
            End If
            
            If Split(strData(i) & ":", ":")(0) = "三方卡类结算" Then
                If int类型排序 = 1 Then
                    str类型排序 = "Decode(Nvl(b.性质,0),8,1"
                Else
                    str类型排序 = str类型排序 & ",8," & int类型排序
                End If
                
                Select Case int方式
                Case 0
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 8, A.日期, Null) As 排序三方"
                    str条件排序 = str条件排序 & ",排序三方"
                Case 1
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 8, A.金额, Null) As 排序三方"
                    str条件排序 = str条件排序 & ",排序三方"
                Case 2
                    strDecode = strDecode & ",Decode(NVL(B.性质, 0), 8, A.金额, Null) As 排序三方"
                    str条件排序 = str条件排序 & ",排序三方 Desc"
                End Select
                
                int类型排序 = int类型排序 + 1
            End If
        Next i
        str类型排序 = str类型排序 & ",Null) As 类型排序"
        
        str排序 = "Order By 标志,类型排序" & str条件排序 & ",No"
        
        strSQL = "" & _
        "Select " & str类型排序 & strDecode & ", a.No, a.票据号, a.Id, a.金额, a.记录状态, a.预交id, a.日期, a.结算方式, " & _
        "       a.卡类别id, a.结算卡序号, decode(nvl(A.结算卡序号,0),0,0,1)  as 是否消费卡,a.卡号, a.交易流水号, a.交易说明, " & vbNewLine & _
        "       c.是否转帐及代扣 As 转帐及代扣,  Nvl(c.名称, q.名称)  As 卡类别名称, Nvl(C.是否退现, Q.是否退现)  As 是否退现, " & vbNewLine & _
        "        Nvl(C.是否全退, Q.是否全退) As 是否全退, c.是否缺省退现,b.性质 As 结算性质,  Sign(Nvl(a.金额, 0)) As 标志" & vbNewLine & _
        " From (Select a.No, Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) As 金额, Min(记录状态) As 记录状态," & vbNewLine & _
        "              Min(Decode(a.结帐id, Null, a.Id, 0) * Decode(a.记录状态, 1, 1, 0)) As ID," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.Id), 0)) As 预交id," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.实际票号), Null)) As 票据号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, To_Char(a.收款时间, 'yyyy-mm-dd hh24:mi:ss')), Null)) As 日期," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.结算方式), Null)) As 结算方式," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.卡类别id), Null)) As 卡类别id," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.结算卡序号), Null)) As 结算卡序号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.卡号), Null)) As 卡号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.交易流水号), Null)) As 交易流水号," & vbNewLine & _
        "              Max(Decode(a.记录性质, 1, Decode(a.记录状态, 2, Null, a.交易说明), Null)) As 交易说明" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where a.记录性质 In (1, 11) And a.病人id = [1] " & strWhere & strWherePage & _
        "       Having Sum(Nvl(金额, 0) - Nvl(冲预交, 0)) <> 0" & vbNewLine & _
        "       Group By a.No) A, 结算方式 B, 医疗卡类别 C,消费卡类别目录 Q" & vbNewLine & _
        " Where a.结算方式 = b.名称(+) And a.卡类别id = c.Id(+) And a.结算卡序号 = q.编号(+) And b.性质 <> 5" & vbNewLine & str排序
    End If

    Set rsDeposit = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, strTime, strDate, int预交类别)
    
    Set GetDeposit = rsDeposit
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMaxDate(lng病人ID As Long, lng主页ID As Long, Optional int原因 As Integer) As Date
'功能：获取转科病人最大的上次变动时间
'参数：int原因=返回上次变动的原因
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    int原因 = 0
    
    strSQL = " Select 开始时间,开始原因 From 病人变动记录" & _
             " Where 开始时间 is Not NULL And 终止时间 is NULL And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIf(IsNull(rsTmp!开始时间), GetMaxDate, rsTmp!开始时间)
        int原因 = Nvl(rsTmp!开始原因, 0)
    End If
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
    '收费分币处理方式
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    gBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))
    
    gbln分离发药 = zlDatabase.GetPara(16, glngSys) = "1"
    gblnStock = zlDatabase.GetPara(18, glngSys) = "1"
    
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytFactLength = Val(Split(strValue, "|")(2))
    
    gbyt检查未执行 = Val(zlDatabase.GetPara(22, glngSys))
    gbyt检查未发药 = Val(zlDatabase.GetPara(154, glngSys))  '33048
    gbytBillOpt = Val(zlDatabase.GetPara(23, glngSys))
    
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnStrictCtrl = Mid(strValue, 3, 1) = "1"
    
    gbln在院不准结帐 = zlDatabase.GetPara(31, glngSys) = "1"
    
    '一卡通消费验证
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    gdbl预存款消费验卡 = Val(Split(strValue, "|")(0))
    gbyt预存款退费验卡 = Val(Split(strValue, "|")(1))
    gbln消费卡退费验卡 = zlDatabase.GetPara(282, glngSys) = "1"
    
    gbln执行后发料 = True ' zlDatabase.GetPara(33, glngSys) = "1"
    
    strValue = zlDatabase.GetPara(41, glngSys)
    gstr医保费用类型 = "'" & Replace(strValue, "|", "','") & "'"
    strValue = zlDatabase.GetPara(42, glngSys)
    gstr公费费用类型 = "'" & Replace(strValue, "|", "','") & "'"
    
    gbln医生允许才能出院 = zlDatabase.GetPara(43, glngSys) = "1"
    
    '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
            
    gstrCardPass = zlDatabase.GetPara(46, glngSys, , "0000000000")
    gbln本人执行 = zlDatabase.GetPara(51, glngSys) = "1"
    gbln开单人 = zlDatabase.GetPara(52, glngSys) = "1"
    gbln它科人 = zlDatabase.GetPara(53, glngSys) = "1"
    gbytAuditing = Val(zlDatabase.GetPara(58, glngSys))
    
    gbyt医保对码检查 = Val(zlDatabase.GetPara(59, glngSys))
    gcurMaxMoney = Val(zlDatabase.GetPara(60, glngSys))
    gint卫材发料控制 = Val(zlDatabase.GetPara(63, glngSys)) '
    gint诊断输入 = Val(zlDatabase.GetPara(65, glngSys, , 1)) Mod 10
    
    gbln收费类别 = zlDatabase.GetPara(72, glngSys) = "1"
    gbln药疗划价单 = zlDatabase.GetPara(79, glngSys) = "1"
    gbln其他划价单 = zlDatabase.GetPara(80, glngSys) = "1"
    ' 81参数:该参数至少在10.03以前就存在，未找到BUG号。审核划价单的目的是确认费用，执行之后，如果不确认费用，就还需要人工单独去审核划价单。从业务特性来说，本参数没有必要存在，应该都处理为执行后自动审核划价单。程序相关控制按勾上此参数进行处理
    gbln执行后审核 = True ' zlDatabase.GetPara(81, glngSys) = "1"
    
    gbln从项汇总折扣 = zlDatabase.GetPara(93, glngSys) = "1"
    gbln报警包含划价费用 = zlDatabase.GetPara(98, glngSys) = "1"
                    
    '刘兴洪 问题:????    日期:2010-12-06 23:38:53
    '费用单价保留位数
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '当不输类别时,输入费用项目时,首位当作类别简码
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1" And Not gbln收费类别
    
    gbln每次住院新住院号 = zlDatabase.GetPara(145, glngSys)
    gbytMediOutMode = Val(zlDatabase.GetPara(150, glngSys))
    gbln不显示无库存卫材 = zlDatabase.GetPara(316, glngSys) = "1"
    
    '个人全局参数
    '-------------------------------------------------------------------------------------------------
    '问题:27990
    With gTy_System_Para
        .byt输入药品显示 = Val(zlDatabase.GetPara("输入药品显示")) '0-按输入匹配显示，1-固定显示通用名和商品名
        .byt药品名称显示 = Val(zlDatabase.GetPara("药品名称显示"))  '：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        .int数据补录时限 = Val(zlDatabase.GetPara(158, glngSys, , "24"))    '问题:33744
        .byt病人审核方式 = Val(zlDatabase.GetPara(185, glngSys, , "0"))
        .bln未入科禁止记账 = Val(zlDatabase.GetPara(215, glngSys, , "0")) = 1    '51612
        .byt条码卫材识别控制 = Val(zlDatabase.GetPara(320, glngSys, , "0"))      '1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
        With .TY_Balance
            .bln刷卡输入密码 = Mid(gstrCardPass, 7, 1) = "1"
            .bytAuditing = Val(zlDatabase.GetPara(58, glngSys))
            .byt检查未执行 = Val(zlDatabase.GetPara(22, glngSys))
            .byt检查未发药 = Val(zlDatabase.GetPara(154, glngSys))
            .byt门诊检查未发药 = Val(zlDatabase.GetPara(265, glngSys))
            .byt门诊检查未执行 = Val(zlDatabase.GetPara(266, glngSys))
            .bln医生允许才能出院 = zlDatabase.GetPara(43, glngSys) = "1"
            .bln在院不准结帐 = zlDatabase.GetPara(31, glngSys) = "1"
        End With
    End With
    InitSysPar = True
End Function

Public Sub zlInit药房()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化药房的相关参数
    '编制:刘兴洪
    '日期:2010-01-25 21:29:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glng中药房 = Val(zlDatabase.GetPara("缺省中药房", glngSys, 1150))
    glng西药房 = Val(zlDatabase.GetPara("缺省西药房", glngSys, 1150))
    glng成药房 = Val(zlDatabase.GetPara("缺省成药房", glngSys, 1150))
    glng发料部门 = Val(zlDatabase.GetPara("缺省发料部门", glngSys, 1150))
    
    gbln其它药房 = zlDatabase.GetPara("显示其它药房库存", glngSys, 1150) = "1"
    gbln其它药库 = zlDatabase.GetPara("显示其它药库库存", glngSys, 1150) = "1"
    
    '分离发药时的检查
    gstr西药房 = zlDatabase.GetPara("西药房选择", glngSys, 1150)
    gstr成药房 = zlDatabase.GetPara("成药房选择", glngSys, 1150)
    gstr中药房 = zlDatabase.GetPara("中药房选择", glngSys, 1150)
End Sub

Public Sub InitLocPar(lngModul As Long)
'功能：初始化模块参数
'参数：无
    Dim strValue As String
    On Error Resume Next
   
   'a.本机注册表存储的模块参数
    '----------------------------------------------------------------------------------------
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
        gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
    End If
    
    'b.数据库存储的公共全局参数
    '----------------------------------------------------------------------------------------
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    strValue = zlDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
        
        
        
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    If lngModul = 1137 Then '结帐
        gbytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, lngModul, "0"))
        'glngShareUseID = Val(zlDatabase.GetPara("共用结帐票据批次", glngSys, lngModul, "0"))
        gbytFeePrintSet = Val(zlDatabase.GetPara("结帐明细打印", glngSys, lngModul, "0"))
        gbyt结帐时输血费检查 = Val(zlDatabase.GetPara("结帐时输血费检查", glngSys, lngModul, "0"))
    ElseIf lngModul = 1142 Then
        strValue = zlDatabase.GetPara("医技病人来源", glngSys, lngModul, "111")
        '处理旧数据
        If Len(strValue) = 1 Then
            If strValue = "0" Then
                strValue = "111"
            ElseIf strValue = "1" Then
                strValue = "101"
            Else
                strValue = "010"
            End If
        End If
        gstrExe来源 = strValue
        
        gstrExe类别 = zlDatabase.GetPara("医技执行类别", glngSys, lngModul)
        gbytExe门诊单据类型 = Val(zlDatabase.GetPara("医技门诊单据类型", glngSys, lngModul, "2"))
        gbytExe住院单据类型 = Val(zlDatabase.GetPara("医技住院单据类型", glngSys, lngModul, "2"))
        gbytExe体检单据类型 = Val(zlDatabase.GetPara("医技体检单据类型", glngSys, lngModul, "2"))
        gbytExe打印方式 = Val(zlDatabase.GetPara("执行登记单打印方式", glngSys, lngModul, "2"))
    End If
        
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Or lngModul = 1137 Then
        gblnLedWelcome = zlDatabase.GetPara("LED显示欢迎信息", glngSys, lngModul, "1") = "1"
    End If
        
    
    If lngModul = 1133 Or lngModul = 1134 Or lngModul = 1135 Then
        gstr收费类别 = zlDatabase.GetPara("收费类别", glngSys, 1150)
        gbln门诊留观 = zlDatabase.GetPara("门诊留观病人记帐", glngSys, 1150) = "1"
        gbln住院留观 = zlDatabase.GetPara("住院留观病人记帐", glngSys, 1150) = "1"
        
        gintOutDay = Val(zlDatabase.GetPara("出院病人天数", glngSys, 1150))
        
        
        Call zlInit药房
        
    
        gbyt开单人显示 = IIf(zlDatabase.GetPara("开单人显示方式", glngSys, 1150) = "2", 2, 1)
        gblnFromDr = zlDatabase.GetPara("科室医生", glngSys, 1150) = "0"
        
        gblnPrice = zlDatabase.GetPara("允许保存为划价单", glngSys, 1150) = "1"
        gbln住院单位 = zlDatabase.GetPara("记帐药品单位", glngSys, 1150) = "1"
        gbytSendMateria = Val(zlDatabase.GetPara("记帐后发药", glngSys, 1150))
    
        gblnPay = zlDatabase.GetPara("中药付数", glngSys, 1150) = "1"
        gblnTime = zlDatabase.GetPara("变价数次", glngSys, 1150) = "1"
        gbln护士 = zlDatabase.GetPara("显示护士", glngSys, 1150) = "1"
        
        '打印控制
        gbln记帐打印 = zlDatabase.GetPara("记帐打印", glngSys, lngModul) = "1"  '记帐打印不是1150的参数
        gbln划价打印 = zlDatabase.GetPara("划价打印", glngSys, 1150) = "1"
        gbln审核打印 = zlDatabase.GetPara("审核打印", glngSys, 1150) = "1"
        
        
        gbln药房上班安排 = Check药房上班安排
        
    ElseIf lngModul = 1137 Then
        gintOutDay = Val(zlDatabase.GetPara("出院病人天数", glngSys, lngModul))
        
        'gint结帐打印 = Val(zlDatabase.GetPara("普通病人结帐打印", glngSys, lngModul))
        gblnPrintByPatient = zlDatabase.GetPara("合约单位按病人打印", glngSys, lngModul) = "1"
        gbln中途结帐退预交 = zlDatabase.GetPara("中途结帐退预交", glngSys, lngModul) = "1"
        gblnAutoOut = zlDatabase.GetPara("在院病人结帐后自动出院", glngSys, lngModul) = "1"
        gblnZero = zlDatabase.GetPara("处理零费用", glngSys, lngModul) = "1"
        gbln仅用指定预交款 = zlDatabase.GetPara("仅用指定预交款", glngSys, lngModul) = "1"
        gbln多次住院弹出结帐设置 = zlDatabase.GetPara("多次住院弹出结帐设置", glngSys, lngModul) = "1"
        gint费用时间 = IIf(zlDatabase.GetPara("结帐费用时间", glngSys, lngModul) = "1", 1, 0)
        gbyt结帐检查代收款项 = zlDatabase.GetPara("结帐检查代收款项", glngSys, lngModul, , "0")
        '32322
        gstr结算方式显示顺序 = Trim(zlDatabase.GetPara("结算方式显示顺序", glngSys, lngModul, "非医保结算-有金额;非医保结算-无金额;医保结算-有金额且允许修改;医保结算-无金额且允许修改;医保结算-有金额且不允许修改;医保结算-无金额且不允许修改"))
    
    ElseIf lngModul = 1142 Then
        '医技执行参数
        gblnExe医嘱 = zlDatabase.GetPara("医技医嘱发送", glngSys, lngModul) = "1"
    End If
End Sub


Public Function ImportBill(strNO As String, blnBat As Boolean, frmParent As Object, _
    Optional blnModi As Boolean, Optional bln住院单位 As Boolean, Optional ByVal bln不读煎法 As Boolean, _
    Optional ByVal lngUnitID As Long, Optional ByVal bln不读执行性质 As Boolean = True, _
    Optional ByVal str药品价格等级 As String, _
    Optional ByVal str卫材价格等级 As String, Optional ByVal str普通价格等级 As String) As ExpenseBill
'功能：读取费用单据到单据对象中(目前忽略从属项目,当作独立项目)
'参数：
'      strNO=单据号
'      blnBat=是否多病人单(对记帐单有效)
'      blnInHos=是否只导入单据中在院病人记录(主要用于记帐表导入)  '此参数已取消,因为现修改为允许导入门诊病人的单据
'      blnModi=是否在修改单据时调用该过程(否则为导入)
'      bln不读煎法  简单记帐等不读煎法
'      lngUnitID    当前操作病区ID
'      bln不读执行性质   只有记帐单时,才需要取执行性质
'返回：存放单据信息的单据对象
'说明：因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
'      不管是导入还是修改单据,都不应包含已停用收费细目

    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset
    Dim rsPrice As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim i As Long, intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, str药房IDs As String, str停用项目序号 As String, strPrivs As String
    Dim curModiMoney As Currency
    
    Dim dblAllTime As Double, dblCurTime As Double, dbl加班加价率 As Double, dblPriceSingle As Double, lngLastPati As Long
    Dim colSerial As New Collection, dblPrice As Double
    Dim bytType As Byte '0-门诊;1-住院;2-门诊或住院
    Dim strTable  As String
    Dim str摘要 As String, strWherePriceGrade As String
    
    On Error GoTo errH
    '价格等级
    If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
        strWherePriceGrade = _
            "      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And d.价格等级 = [5])" & vbNewLine & _
            "            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And d.价格等级 = [6])" & vbNewLine & _
            "            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And d.价格等级 = [7])" & vbNewLine & _
            "            Or (d.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists (Select 1" & vbNewLine & _
            "                                From 收费价目" & vbNewLine & _
            "                                Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                      And ((Instr(';5;6;7;', ';' || b.类别 || ';') > 0 And 价格等级 = [5])" & vbNewLine & _
            "                                            Or (Instr(';4;', ';' || b.类别 || ';') > 0 And 价格等级 = [6])" & vbNewLine & _
            "                                            Or (Instr(';4;5;6;7;', ';' || b.类别 || ';') = 0 And 价格等级 = [7])))))"
    Else
        strWherePriceGrade = " And d.价格等级 Is Null"
    End If
    
    If blnBat Or blnModi Then  '多病人单或才需要修改的单单据,肯定是住院的
        strTable = "" & _
        "   Select A.记录性质,A.序号,A.从属父号,A.NO,A.记录状态,A.多病人单,A.婴儿费,A.开单部门ID,A.门诊标志,A.加班标志," & _
        "          A.附加标志,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        "          A.标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        "          A.收费类别,A.费用类型,A.病人ID,A.结论" & _
        "   From 住院费用记录 A " & _
        "   Where A.记录性质 in (1,2)  And A.记录状态 IN(0,1,3) And A.价格父号 Is Null " & IIf(Not blnModi, " And Nvl(A.数次,0)>=0", "") & _
        "         And A.NO=[1] And Nvl(A.多病人单,0)=[3] " & _
        ""
    Else
        strTable = "" & _
        "   Select A.记录性质,A.序号,A.从属父号,A.NO,A.记录状态,0 as 多病人单,A.婴儿费,A.开单部门ID,A.门诊标志,A.加班标志," & _
        "          A.附加标志,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        "          A.标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        "          A.收费类别,A.费用类型,A.病人ID,A.结论" & _
        "   From 门诊费用记录 A " & _
        "   Where A.记录性质 in (1,2)  And A.记录状态 IN(0,1,3) And A.价格父号 Is Null " & IIf(Not blnModi, " And Nvl(A.数次,0)>=0", "") & _
        "         And A.NO=[1]  " & _
        "   Union ALL " & _
        "   Select A.记录性质,A.序号,A.从属父号,A.NO,A.记录状态,A.多病人单,A.婴儿费,A.开单部门ID,A.门诊标志,A.加班标志," & _
        "          A.附加标志,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        "          A.标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        "          A.收费类别,A.费用类型,A.病人ID,A.结论" & _
        "   From 住院费用记录 A " & _
        "   Where A.记录性质 in (1,2)  And A.记录状态 IN(0,1,3) And A.价格父号 Is Null " & IIf(Not blnModi, " And Nvl(A.数次,0)>=0", "") & _
        "         And A.NO=[1] And Nvl(A.多病人单,0)=[3] "
        
    End If
    
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    '价格父号 is NULL:只取每个收费细目ID的第一行(药品只有一行),因为要重算价格
    strSQL = _
        " Select F.险类,X.药品ID,W.材料ID,W.跟踪在用," & _
        "       A.序号 As 序号,A.从属父号,A.NO,A.记录性质,A.记录状态,A.多病人单,A.婴儿费,G.费别,F.姓名,F.性别,F.年龄,F.担保额," & _
        "       G.出院病床 as 床号,F.住院号 as 标识号,F.病人ID,G.主页ID,G.当前病区ID as 病人病区ID,G.出院科室ID as 病人科室ID,A.开单部门ID,A.门诊标志,A.加班标志," & _
        "       G.病人性质,A.附加标志,A.收费类别,A.收费细目ID,A.发药窗口,Nvl(付数,1) as 付数,Nvl(A.数次,0) as 数次," & _
        "       A.标准单价,A.执行部门ID,A.划价人,A.开单人,A.操作员编号,A.操作员姓名,A.发生时间,A.登记时间,A.摘要," & _
        "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(H.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
        "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象,Nvl(A.费用类型,B.费用类型) 费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
        "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
        "       Decode(A.收费类别,'4',1,X.住院包装) as 住院包装,Decode(A.收费类别,'4',B.计算单位,X.住院单位) as 住院单位," & _
        "       Decode(A.收费类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,B.录入限量,A.结论,M1.名称 as 诊疗名称,X.中药形态,x.剂量系数,M1.计算单位 as 剂量单位" & _
        " From (" & strTable & ") A,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,病人信息 F, " & _
        "          病案主页 G,收费项目别名 H,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
        " Where  A.收费细目ID=D.收费细目ID And A.收费细目ID=B.ID " & _
        "       And A.收费类别=C.编码 And A.收费细目ID=X.药品ID(+) and X.药名ID=M1.ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
        "       And A.收费细目ID=H.收费细目ID(+) And H.码类(+)=1 And H.性质(+)=[4]" & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
        "       And A.病人ID=F.病人ID(+) And F.病人ID=G.病人ID(+) And F.主页ID=G.主页ID(+)" & _
        "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) " & _
                strWherePriceGrade
        
    If blnBat And Not blnModi Then
        strPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
        If InStr(1, strPrivs, ";所有病区;") = 0 Then strSQL = strSQL & " And G.当前病区ID = [2]"
    End If
        
    If Not gbln分离发药 Then
        strSQL = "Select * From (" & strSQL & ")" & IIf(blnBat, " Order by LPAD(床号,10,' '),病人ID,序号", " Order by 序号")
    Else
        '分离发药时排开时价和分批药品或卫材
        strSQL = "Select * From (" & strSQL & ") Where Not(Instr(',5,6,7,',收费类别)>0 And (分批=1 Or 是否变价=1))" & _
            IIf(blnBat, " Order by LPAD(床号,10,' '),病人ID,序号", " Order by 序号")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, lngUnitID, IIf(blnBat, 1, 0), _
        IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1), str药品价格等级, str卫材价格等级, str普通价格等级)
    
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    If rsTmp.RecordCount <> 0 Then
        '如果即有门诊的,也有住院的单据,则只能选择一种
        If Not blnModi Then
            rsTmp.Filter = "记录性质=1"
            i = rsTmp.RecordCount
            If i > 0 Then
                rsTmp.Filter = "记录性质=2"
                If rsTmp.RecordCount > 0 Then
                    If zlCommFun.ShowMsgbox("单据导入", "找到两张单据号为[" & strNO & "]的单据,请问您要导入", _
                            "!住院单据(&Z),门诊单据(&M)", frmParent, vbInformation) = "门诊单据" Then
                        rsTmp.Filter = "记录性质=1"
                    End If
                Else
                    rsTmp.Filter = ""   '门诊单据
                End If
            Else
                rsTmp.Filter = ""       '住院单据
            End If
        Else    '修改单据只允许住院的
            rsTmp.Filter = "记录性质=2"
        End If
        
        rsTmp.MoveFirst
        
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
                        If Not CheckFeeItemAvailable(!收费细目ID, 2) Then
                            str停用项目序号 = str停用项目序号 & "," & !序号
                            MsgBox "单据[" & strNO & "]中第" & !序号 & "行收费项目:" & !名称 & "" & vbCrLf & _
                                "已停用或不再服务于病人,将不会被导入." & IIf(IsNull(!从属父号), "如果有从属项目,也不会被导入.", ""), vbInformation, gstrSysName
                            .MoveNext
                            GoTo NextRecord
                        End If
                    End If
                End If
                If blnBat And Not blnModi Then
                    If InStr(1, strPrivs, ";所有病区;") > 0 And lngUnitID <> 0 And lngLastPati <> Val(!病人ID) Then
                        lngLastPati = !病人ID
                        If InStr(1, "," & lngUnitID & ",", "," & !病人病区ID & ",") = 0 Then
                            If MsgBox("病人""" & !姓名 & """当前不属于当前病区，是否导入该病人费用?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                GoTo NextRecord
                            End If
                        End If
                    End If
                End If
                
                
                '处理单据主体=====================================================
                If i = 1 Then
                    objBill.NO = !NO
                    objBill.病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
                    objBill.主页ID = IIf(IsNull(!主页ID), 0, !主页ID)
                    objBill.病区ID = IIf(IsNull(!病人病区ID), 0, !病人病区ID)
                    objBill.科室ID = IIf(IsNull(!病人科室id), 0, !病人科室id)
                    objBill.姓名 = IIf(IsNull(!姓名), "", !姓名)
                    objBill.性别 = IIf(IsNull(!性别), "", !性别)
                    objBill.年龄 = IIf(IsNull(!年龄), "", !年龄)
                    objBill.标识号 = IIf(IsNull(!标识号), 0, !标识号)
                    objBill.床号 = "" & !床号
                    objBill.费别 = IIf(IsNull(!费别), "", !费别)
                    objBill.门诊标志 = IIf(IsNull(!门诊标志), 0, !门诊标志)
                    objBill.加班标志 = IIf(IsNull(!加班标志), 0, !加班标志)
                    objBill.婴儿费 = IIf(IsNull(!婴儿费), 0, !婴儿费)
                    objBill.开单部门ID = IIf(IsNull(!开单部门ID), 0, !开单部门ID)
                    objBill.划价人 = IIf(IsNull(!划价人), "", !划价人)
                    objBill.开单人 = IIf(IsNull(!开单人), "", !开单人)
                    objBill.操作员编号 = IIf(IsNull(!操作员编号), "", !操作员编号)
                    objBill.操作员姓名 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
                    objBill.发生时间 = !发生时间
                    objBill.登记时间 = !登记时间
                    objBill.多病人单 = (IIf(IsNull(!多病人单), 0, !多病人单) = 1)
                End If
                
                '处理收费细目=====================================================
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                                
                '处理序号和从属父号
                intCurNo = intCurNo + 1
                objBillDetail.序号 = intCurNo '实际是行号
                colSerial.Add intCurNo, "_" & !序号 '记录原序号现在的行号
                objBillDetail.从属父号 = Nvl(!从属父号, 0) '因为可能排序乱了,先记录原来的,后面再处理
                
                objBillDetail.病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
                objBillDetail.主页ID = IIf(IsNull(!主页ID), 0, !主页ID)
                objBillDetail.婴儿费 = IIf(IsNull(!婴儿费), 0, !婴儿费) '记帐表时,每个病人不同
                objBillDetail.病区ID = IIf(IsNull(!病人病区ID), 0, !病人病区ID)
                objBillDetail.科室ID = IIf(IsNull(!病人科室id), 0, !病人科室id)
                objBillDetail.姓名 = IIf(IsNull(!姓名), "", !姓名)
                objBillDetail.性别 = IIf(IsNull(!性别), "", !性别)
                objBillDetail.年龄 = IIf(IsNull(!年龄), "", !年龄)
                objBillDetail.住院号 = IIf(IsNull(!标识号), 0, !标识号)
                objBillDetail.床号 = "" & !床号
                objBillDetail.费别 = IIf(IsNull(!费别), "", !费别)
                objBillDetail.担保额 = IIf(IsNull(!担保额), 0, !担保额)
                
                '目前仅用于记帐表
                objBillDetail.医疗付款 = Get病人医疗付款方式(IIf(IsNull(!病人ID), 0, !病人ID), IIf(IsNull(!主页ID), 0, !主页ID))
                
                objBillDetail.收费类别 = IIf(IsNull(!收费类别), "", !收费类别)
                objBillDetail.收费细目ID = IIf(IsNull(!收费细目ID), 0, !收费细目ID)
                objBillDetail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.付数 = Nvl(!付数, 1)
                If InStr(",5,6,7,", !收费类别) > 0 And bln住院单位 Then
                    objBillDetail.数次 = Nvl(!数次, 0) / Nvl(!住院包装, 1)
                Else
                    objBillDetail.数次 = Nvl(!数次, 0)
                End If
                objBillDetail.原始数量 = objBillDetail.付数 * objBillDetail.数次
                                
                If blnBat Then
                    objBillDetail.发药窗口 = IIf(IsNull(!险类), "", !险类)
                Else
                    objBillDetail.发药窗口 = IIf(IsNull(!发药窗口), "", !发药窗口)
                End If
                
                objBillDetail.附加标志 = IIf(IsNull(!附加标志), 0, !附加标志)
                objBillDetail.摘要 = IIf(IsNull(!摘要), "", !摘要)
            
                If InStr(",5,6,7,", !收费类别) > 0 And gbln分离发药 Then
                    objBillDetail.执行部门ID = 0
                Else
                    objBillDetail.执行部门ID = IIf(IsNull(!执行部门ID), 0, !执行部门ID)
                End If
                objBillDetail.原始执行部门ID = objBillDetail.执行部门ID
                
                '这里可能记录修改该单据的原单据号,而本身却要用于存放记帐表病人的费用情况
                If blnBat Then
                    blnLoad = objBill.Details.Count = 0
                    If Not blnLoad Then
                        blnLoad = objBillDetail.病人ID <> objBill.Details(objBill.Details.Count).病人ID
                    End If
                    If blnLoad Then
                        '费用信息
                        Set rsMoney = Nothing
                        If blnModi Then
                            '修改前的当前单据的病人费用金额
                            If gbytBilling = 0 Then
                                'int来源-1-门诊,2-住院
                                curModiMoney = GetBillMoney(2, strNO, objBillDetail.病人ID)
                            End If
                            
                            Set rsMoney = GetMoneyInfo(objBillDetail.病人ID, CDbl(curModiMoney), , 2)
                        Else
                            Set rsMoney = GetMoneyInfo(objBillDetail.病人ID, , , 2)
                        End If
                        If Not rsMoney Is Nothing Then
                            objBillDetail.就诊卡号 = rsMoney!预交余额 & "," & rsMoney!费用余额 & "," & rsMoney!预交余额 - rsMoney!费用余额
                        Else
                            objBillDetail.就诊卡号 = "0,0,0"
                        End If
                        '当日费用额
                        objBillDetail.就诊卡号 = objBillDetail.就诊卡号 & "," & GetPatiDayMoney(objBillDetail.病人ID)
                    Else
                        objBillDetail.就诊卡号 = objBill.Details(objBill.Details.Count).就诊卡号
                    End If
                End If
                
                objBillDetail.Detail.ID = !收费细目ID
                objBillDetail.Detail.编码 = !编码
                objBillDetail.Detail.变价 = (IIf(IsNull(!是否变价), 0, !是否变价) = 1)
                objBillDetail.Detail.从项数次 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.固有从属 = 0 '!!!目前忽略从属项目,当作独立项目
                objBillDetail.Detail.规格 = IIf(IsNull(!规格), "", !规格)
                objBillDetail.Detail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                
                objBillDetail.Detail.住院单位 = Nvl(!住院单位)
                objBillDetail.Detail.住院包装 = Nvl(!住院包装, 1)
                                
                
                If Not gbln分离发药 And InStr(",4,5,6,7,", !收费类别) > 0 Then
                    dblStock = GetStock(!收费细目ID, !执行部门ID)
                Else
                    dblStock = 0
                End If
                If InStr(",5,6,7,", !收费类别) > 0 And gbln分离发药 Then
                    str药房IDs = Decode(!收费类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                    If str药房IDs <> "" Then dblStock = GetMultiStock(!收费细目ID, str药房IDs)
                End If
                If InStr(",5,6,7,", !收费类别) > 0 And bln住院单位 Then dblStock = dblStock / Nvl(!住院包装, 1)
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
                objBillDetail.Detail.服务对象 = IIf(IsNull(!服务对象), 0, !服务对象)
                objBillDetail.Detail.类型 = IIf(IsNull(!费用类型), "", !费用类型)
                objBillDetail.Detail.诊疗名称 = Nvl(!诊疗名称)
                
                If InStr(",5,6,7,", !收费类别) > 0 Then
                    objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                    objBillDetail.Detail.处方限量 = Get处方限量(objBillDetail.Detail.ID)
                End If
                objBillDetail.Detail.录入限量 = Val("" & !录入限量)
                                    
                objBillDetail.Detail.药名ID = IIf(IsNull(!药名ID), 0, !药名ID)
                objBillDetail.Detail.变价 = IIf(IsNull(!是否变价), 0, !是否变价) = 1
                objBillDetail.Detail.分批 = IIf(IsNull(!分批), 0, !分批) = 1
                objBillDetail.Detail.跟踪在用 = Nvl(!跟踪在用, 0) = 1
                objBillDetail.Detail.要求审批 = 0
                objBillDetail.Detail.中药形态 = Val(Nvl(!结论))
                objBillDetail.Detail.剂量单位 = Nvl(!剂量单位)
                objBillDetail.Detail.剂量系数 = Val(Nvl(!剂量系数))
                
                '问题:41136
                str摘要 = objBillDetail.摘要
                '90304
                If Not blnModi Then
                    str摘要 = gclsInsure.GetItemInfo(Val(Nvl(!险类)), objBill.病人ID, objBillDetail.收费细目ID, str摘要, 2, , "|1")
                    objBillDetail.摘要 = str摘要
                End If
                
                '处理价格部份=====================================================
                Set objBillDetail.InComes = New BillInComes
                Do
                    '按照现有的价格设置重新计算'***
                    If !是否变价 = 1 Then
                        If InStr(",5,6,7,", !收费类别) > 0 Or (!收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                            '----------------------------------------------------------------------------------------------
                            '时价药品计算价格(分批可不分批)
                            dblAllTime = !付数 * !数次
                            If dblAllTime <> 0 Then
                                dblPrice = Get时价药品应收金额(objBillDetail.执行部门ID, CLng(!收费细目ID), dblAllTime, gstrDec, dblPriceSingle)
                                If dblAllTime <> 0 Then
                                    '数量未分解完毕
                                    If !收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & !名称 & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                Else
                                    '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                                    objBillIncome.标准单价 = IIf(dblPriceSingle = 0, Format(dblPrice / (!付数 * !数次), gstrFeePrecisionFmt), dblPriceSingle) '这里是售价价格
                                End If
                            Else
                                objBillIncome.标准单价 = 0
                            End If
                            '----------------------------------------------------------------------------------------------
                        Else
                            If Abs(!标准单价) > Abs(IIf(IsNull(!现价), 0, !现价)) Then
                                objBillIncome.标准单价 = IIf(IsNull(!缺省价格), 0, !缺省价格)
                            Else
                                objBillIncome.标准单价 = !标准单价
                            End If
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If

                    If InStr(",5,6,7,", !收费类别) > 0 And bln住院单位 Then
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(!住院包装, 1), gstrFeePrecisionFmt)
                    Else
                        objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrFeePrecisionFmt)
                    End If
                    objBillIncome.现价 = IIf(IsNull(!现价), 0, !现价) '现价原价对药品变价无用
                    objBillIncome.原价 = IIf(IsNull(!原价), 0, !原价)
                    objBillIncome.收入项目ID = IIf(IsNull(!现收入ID), 0, !现收入ID)
                    objBillIncome.收入项目 = IIf(IsNull(!收入项目), "", !收入项目)
                    objBillIncome.收据费目 = IIf(IsNull(!现费目), "", !现费目)
                    
                    '应收金额=单价*付次*数次
                    If !是否变价 = 1 And (InStr(",5,6,7,", !收费类别) > 0 Or !收费类别 = "4" And Nvl(!跟踪在用, 0) = 1) Then
                        objBillIncome.应收金额 = dblPrice '保证应收金额与零售金额没有误差
                    Else
                        objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                    End If
                    
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
                        objBillIncome.实收金额 = ActualMoney(objBillDetail.费别, !现收入ID, objBillIncome.应收金额, _
                            objBillDetail.收费细目ID, objBillDetail.执行部门ID, objBillDetail.原始数量, dbl加班加价率)
                    End If
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额
                    End With
                    
                    '判断下一条记录是否属于当前行
                    blnDo = False
                    int序号 = !序号
                    .MoveNext
                    If Not .EOF Then blnDo = (int序号 = !序号)
                    i = i + 1
                Loop While blnDo And Not .EOF
               
                With objBillDetail
                    objBill.Details.Add .Detail, .收费细目ID, .序号, .从属父号, .病人ID, .主页ID, .病区ID, .科室ID, .姓名, .性别, .年龄, .住院号, .床号, _
                        .费别, .病人性质, .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, .InComes, .就诊卡号, , .担保额, .医疗付款, , , , .摘要, .原始数量, .原始执行部门ID, .婴儿费
                    '分离发药时,Key设置为1,表示编辑时执行科室列不可进入
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 Then
                        objBill.Details(objBill.Details.Count).Key = 1
                    End If
                End With
            Loop
        End With
        
        '再重新处理从属父号
        For i = 1 To objBill.Details.Count
            If objBill.Details(i).从属父号 <> 0 Then
                objBill.Details(i).从属父号 = colSerial("_" & objBill.Details(i).从属父号)
            End If
        Next
    End If
    
    If Not bln不读煎法 Then
        If blnModi And Not blnBat Then  '仅记帐及记帐划价单的修改时(没有排开简单记帐,只有等它去读一回儿)
                    '读取中药煎法
            strSQL = "Select 外观 From 药品收发记录 Where NO=[1] And 单据=9" '9-记帐单处方发药；
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!外观) Then
                    objBill.煎法 = rsTmp!外观
                End If
            End If
        End If
    End If
    
    If Not bln不读执行性质 Then
        '刘兴洪 问题:27383 日期:2010-02-01 16:58:14
        strSQL = "Select max(扣率) as 扣率 From 药品收发记录 Where NO=[1] And 单据 =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(Not blnBat, 9, 10))
        objBill.执行性质 = 0
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!扣率) Then
                objBill.执行性质 = Mid(Nvl(rsTmp!扣率) & "00", 2, 1)
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
Public Function zlGetBalancePati(ByVal lng结帐ID As Long, ByRef lng病人ID As Long, lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定结帐的病人ID和主页ID
    '入参:
    '出参:lng病人ID,lng主页ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-03 18:48:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errH
    
    strSQL = "" & _
        "   Select 病人ID,Max(主页ID) as 主页ID From ( " & _
        "   Select distinct 病人ID,主页ID From 住院费用记录 Where 结帐ID=[1]  Union " & _
        "   Select distinct 病人ID,0 as 主页ID From 门诊费用记录 Where 结帐ID=[1]    ) Group by 病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng结帐ID)
    lng病人ID = 0: lng主页ID = 0
    If rsTemp.RecordCount > 0 Then
        lng病人ID = Nvl(rsTemp!病人ID, 0): lng主页ID = Nvl(rsTemp!主页ID, 0)
    End If
    zlGetBalancePati = rsTemp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBalance(strNO As String, frmParent As Object, lng结帐ID As Long, _
       Optional intInsure As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:当前收款记录重新打印一张票据
    '入参:blnMediCare-是否为保险结算票据
    '出参:
    '返回:打印成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-02 10:48:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoice As String
    Dim i As Long, lng领用ID As Long, lngPatientID As Long, lngPatientCount As Long
    Dim blnDo As Boolean, rsTmp As ADODB.Recordset, strRptName As String, intFormat As Integer
    Dim strUseType As String, intPrintMode As Integer
    Dim lngShareUseID As Long, lng病人ID As Long, lng主页ID As Long
    Dim objInvoice As clsInvoice, objFact As clsFactProperty
    Dim bytInvoiceKind As Byte
    Dim strKind As String, rsKind As ADODB.Recordset, bytKind As Byte
    
    '合约单位
    Set rsTmp = GetBanlancePatients(lng结帐ID)
    If rsTmp Is Nothing Then Exit Function
    lngPatientCount = rsTmp.RecordCount
    Set objInvoice = New clsInvoice
    Set objFact = New clsFactProperty
    strKind = "Select Nvl(结帐类型,0) As 类型 From 病人结帐记录 Where ID = [1] And Rownum < 2"
    Set rsKind = zlDatabase.OpenSQLRecord(strKind, "结帐类型", lng结帐ID)
    If Not rsKind.EOF Then bytKind = Val(rsKind!类型)
    If bytKind = 1 Then
        bytInvoiceKind = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, 1137, "0"))
    Else
        bytInvoiceKind = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, 1137, "0"))
    End If
    
    Call objInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)

    If lngPatientCount > 1 Then
        '合约单位结帐
        Call objInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, objFact, , , bytKind)
        objFact.使用类别 = zlDatabase.GetPara("合约单位结帐打印", glngSys, 1137)
    Else
        Call zlGetBalancePati(lng结帐ID, lng病人ID, lng主页ID)
        Call objInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), lng病人ID, lng主页ID, intInsure, objFact, , , bytKind)
    End If
    
    If Not gobjTax Is Nothing And gblnTax Then
        blnDo = True
    Else
        'bytInvoiceKind:结帐票据类型,0-住院票据;1-门诊票据
        strRptName = IIf(bytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
        If objFact.打印格式 = 0 Then   '以缺省票据格式显示
            objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
        End If
        SetReportPrintSet gcnOracle, glngSys, strRptName, "Format", objFact.打印格式
        '由于没有格式的传入,因此,需要强制缺省到指定格式
        blnDo = ReportPrintSet(gcnOracle, glngSys, strRptName, frmParent)
        '取出选择的格式
        objFact.打印格式 = Val(GetReportPrintSet(gcnOracle, glngSys, strRptName, gstrDBUser, 1, , "Format"))
    End If
    
    If blnDo Then
        If gblnPrintByPatient Then
            '合约单位
            lngPatientCount = rsTmp.RecordCount
        Else
            lngPatientCount = 1
        End If
        
        Call GetNextInvoice(frmParent, objInvoice, objFact, lngPatientCount, lng领用ID, strInvoice)
        If objFact.严格控制 And strInvoice = "" Then Exit Function
        objFact.LastUseID = lng领用ID
        
        If gblnPrintByPatient And lngPatientCount > 1 Then
            For i = 1 To rsTmp.RecordCount
                lngPatientID = rsTmp!病人ID
                '合药单位,按普通住院病人打印票据
                Call frmPrint.ReportPrint(2, strNO, lng结帐ID, objFact, strInvoice, , , , lngPatientID, objFact.打印格式)
                If i < rsTmp.RecordCount Then
                    strInvoice = ""
                    Call GetNextInvoice(frmParent, objInvoice, objFact, lngPatientCount + 1 - i, lng领用ID, strInvoice, i = 1)
                    If objFact.严格控制 And strInvoice = "" Then Exit Function
                End If
                rsTmp.MoveNext
            Next
        Else
            Call frmPrint.ReportPrint(2, strNO, lng结帐ID, objFact, strInvoice, , , , , objFact.打印格式)
        End If
        RePrintBalance = True
    End If
End Function

Public Sub GetNextInvoice(ByRef frmParent As Object, ByVal objInvoice As clsInvoice, ByRef objFact As clsFactProperty, _
    ByVal lngLeastNum As Long, ByRef lng领用ID As Long, ByRef strInvoice As String, _
    Optional ByRef blnFirst As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打结帐票据时,获取下一票据号
    '入参:blnFirst-按病人打印时，是否首次打印（仅首次打印提示确定票据号）
    '出参:lng领用ID-返回领用ID
    '        strInvoice-返回发票号
    '编制:刘兴洪
    '日期:2011-05-03 17:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean, blnInput As Boolean
    '如果严格控制票据使用
    If objFact.严格控制 Then
        If objInvoice.GetInvoiceGroupID(UserInfo.姓名, objFact.票种, lngLeastNum, objFact.LastUseID, objFact.共享批次ID, "", objFact.使用类别, lng领用ID) = False Then Exit Sub
        Select Case lng领用ID
            Case -1
                If objFact.使用类别 <> "" Then
                    MsgBox "你没有自用和共用『" & objFact.使用类别 & "』的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
            Case -2
                If objFact.使用类别 <> "" Then
                    MsgBox "本地的共用『" & objFact.使用类别 & "』的结帐票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
        End Select
        If lng领用ID <= 0 Then Exit Sub
    End If
        
    '取下一个票据号码
    If Not objFact.严格控制 Then
        '有可能是第一次使用
        Do
            blnInput = False
            '非严格控制时直接从本地读取
            strInvoice = UCase(zlDatabase.GetPara("当前结帐票据号", glngSys, 1137, ""))
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlStr.Increase(strInvoice)
                If blnFirst Then
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                                    strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
                    blnInput = True
                End If
            End If
                
            '用户取消输入,允许打印
            If strInvoice = "" Then
                If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                blnValid = True
            Else
                '检查输入有效性
                If blnInput Then
                    If zlCommFun.ActualLen(strInvoice) <> gbytFactLength Then
                        MsgBox "输入的票据号码长度应该为 " & objFact.票号长度 & " 位！", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        Exit Sub
    End If
    Do
        '根据票据领用读取
        blnInput = False
        Call objInvoice.zlGetNextBill(1137, lng领用ID, strInvoice)
        If strInvoice = "" Then
            '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
            strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                "", frmParent.Left + 1500, frmParent.Top + 1500))
            blnInput = True
        ElseIf blnFirst Then
            strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, _
                            strInvoice, frmParent.Left + 1500, frmParent.Top + 1500))
            blnInput = True
        End If
        '用户取消输入,不打印
        If strInvoice = "" Then Exit Sub
        
        '检查输入有效性
        If blnInput Then
            If objInvoice.GetInvoiceGroupID(UserInfo.姓名, objFact.票种, lngLeastNum, objFact.LastUseID, objFact.共享批次ID, strInvoice, objFact.使用类别, lng领用ID) = False Then Exit Sub
            If lng领用ID = -3 Then
                MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
End Sub

Public Function GetBanlancePatients(lng结帐ID As Long) As ADODB.Recordset
'功能：判断一张记帐单据是否批量记帐单
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "" & _
        "   Select 病人ID From 住院费用记录 Where 结帐ID=[1] Group by 病人ID Union " & _
        "   Select 病人ID From 门诊费用记录 Where 结帐ID=[1] Group by 病人ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng结帐ID)
    
    Set GetBanlancePatients = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisBatch(strNO As String) As Boolean
'功能：判断一张记帐单据是否批量记帐单
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(多病人单,0) as 多病人单 From 住院费用记录 Where 记录性质=2 And 记录状态 IN(0,1,3) And NO=[1] And RowNum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then BillisBatch = (rsTmp!多病人单 = 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisSimple(strNO As String, Optional bytType As Byte = 2) As Boolean
'功能：判断一张记帐单据是否为简单模式
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Distinct 发药窗口,收费类别,Nvl(数次,1) as 数次" & _
        " From 住院费用记录 Where 记录状态 IN(0,1,3)" & _
        " And 记录性质=[2] And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!数次 = 1 And rsTmp!收费类别 = "Z" And Nvl(rsTmp!发药窗口) = "Z" Then BillisSimple = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceID(strNO As String) As Long
'功能：获取一张结帐单据的ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ID From 病人结帐记录 Where 记录状态=1 And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then GetBalanceID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalanceDeposit(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
    '功能：获取一张结帐单据的冲预交记录
    Dim strSQL As String
    On Error GoTo errH
    strSQL = " " & _
    "   Select a.Id, a.单据号, a.票据号, To_Char(Max(b.收款时间), 'YYYY-MM-DD') As 日期, a.结算方式, " & _
    "          LTrim(To_Char(Max(a.冲预交), '9999999990.00')) As 金额, nvl(b.卡类别id, b.结算卡序号) as 卡类别Id ,min(decode(nvl(b.结算卡序号,0),0,0,1)) as 是否消费卡, Min(Nvl(c.名称, q.名称)) As 卡类别名称, " & _
    "          Min(Nvl(q.是否退现, c.是否退现)) As 是否退现, Min(Nvl(q.是否全退, c.是否全退)) As 是否全退, Min(c.是否缺省退现) As 是否缺省退现,min(C.是否转帐及代扣) as 是否转帐及代扣, Min(b.卡号) As 卡号, " & _
    "          Min(b.交易流水号) As 交易流水号, Min(b.交易说明) As 交易说明 " & _
    "   From (Select ID, NO As 单据号, 实际票号 As 票据号, To_Char(收款时间, 'YYYY-MM-DD') As 日期, 结算方式, Nvl(冲预交, 0) As 冲预交 " & _
    "          From 病人预交记录 " & _
    "          Where Mod(记录性质, 10) = 1 And 结帐id = [1] And Nvl(冲预交, 0) <> 0) A, 病人预交记录 B, 医疗卡类别 C, 消费卡类别目录 Q " & _
    "   Where a.单据号 = b.No And b.记录性质 = 1 And b.卡类别id = c.Id(+) And b.结算卡序号 = q.编号(+) " & _
    "   Group By a.Id, a.单据号, a.票据号, a.结算方式,nvl(b.卡类别id, b.结算卡序号) " & _
    "   Order By 日期, a.结算方式"
    If blnNOMoved Then strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
    Set GetBalanceDeposit = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngBalanceID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBalancePay(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
'功能：获取一张结帐单据的结算记录
    Dim strSQL As String
    On Error GoTo errH
    strSQL = _
            "Select A.结算方式,Ltrim(To_Char(A.冲预交,'9999999990.00')) as 金额," & _
            " A.结算号码,Nvl(B.性质,0) as 性质 From " & IIf(blnNOMoved, "H", "") & "病人预交记录 A,结算方式 B" & _
            " Where mod(A.记录性质,10)=2 And A.结帐ID=[1]" & _
            " And A.结算方式=B.名称(+) Order by A.结算方式"
        
    Set GetBalancePay = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngBalanceID)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!类别ID, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistErrRecord(lngID As Long) As Boolean
'功能：结帐作废时判断结帐时是否产生过误差,如果没有,则要新取单据号,用于生成误差费用记录
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select NO From 住院费用记录 Where Nvl(附加标志,0)=9 And 结帐ID=[1] Union " & _
             "Select NO From 门诊费用记录 Where Nvl(附加标志,0)=9 And 结帐ID=[1] "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngID)
    ExistErrRecord = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureInfo(lng病人ID As Long) As String
'功能：获取住院病人保险帐户信息
'返回："险类名;医保号"
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '连接病案主页,确保本次住院是保险病人,但不一定在院
    strSQL = "Select A.名称,B.医保号" & _
        " From 保险类别 A,保险帐户 B,病人信息 C,病案主页 D" & _
        " Where A.序号=B.险类 And B.病人ID=C.病人ID" & _
        " And B.险类=D.险类 And C.病人ID=D.病人ID" & _
        " And D.主页ID=C.主页ID And C.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If Not rsTmp.EOF Then GetInsureInfo = rsTmp!名称 & ";" & rsTmp!医保号
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetMinMaxDate(ByVal lngID As Long, dMin As Date, dMax As Date, Optional ByVal blnNOMoved As Boolean) As Boolean
'功能：根据结帐ID获取最大最小登记/发生时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If gint费用时间 = 0 Then
        
        strSQL = "Select Max(登记时间) as 最大,Min(登记时间) as 最小 From " & IIf(blnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1] Union all " & _
                 "Select Max(登记时间) as 最大,Min(登记时间) as 最小 From " & IIf(blnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
    Else
        strSQL = "Select Max(发生时间) as 最大,Min(发生时间) as 最小 From " & IIf(blnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1] Union all " & _
                 "Select Max(发生时间) as 最大,Min(发生时间) as 最小 From " & IIf(blnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
    End If
    
    strSQL = "Select Max(最大) as 最大,Min(最小) as 最小 From ( " & strSQL & ")"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngID)
    If Not rsTmp.EOF Then
        If IsNull(rsTmp!最大) Or IsNull(rsTmp!最小) Then Exit Function
        dMax = rsTmp!最大
        dMin = rsTmp!最小
        GetMinMaxDate = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDept(lng病人ID As Long) As Long
'功能：返回病人所属科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.出院科室ID From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDept = IIf(IsNull(rsTmp!出院科室ID), 0, rsTmp!出院科室ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get待发药清单(strNO As String, strTime As String, bln记帐表 As Boolean) As ADODB.Recordset
'功能：根据费用单据号,登记时间,获取待发药品清单
'说明：普通发药时为病人科室，急诊、医技则为开单科室。
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.库房ID,A.对方部门ID" & _
        " From 药品收发记录 A,住院费用记录 B" & _
        " Where A.NO=[1] And A.单据=[2] And Mod(A.记录状态,3)=1 And A.审核人 is NULL" & _
        " And A.NO=B.NO And A.费用ID=B.ID And B.记录状态<>0 And B.登记时间+0=[3]" & _
        " Order by A.药品ID"
    If strTime <> "" Then
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(bln记帐表, 10, 9), CDate(strTime))
    Else
        Set Get待发药清单 = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, IIf(bln记帐表, 10, 9))
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetStockInfo(lng药品ID As Long, bln药房 As Boolean, bln药库 As Boolean, Optional ByVal bln住院单位 As Boolean) As String
'功能：获取药品在各个药房，药库的库存信息
'参数："bln药房/bln药库"至少要有一个设置为真
'返回：描述信息
    Dim strSQL As String, strSQL2 As String, i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
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
        " Nvl(Sum(A.可用数量),0)" & IIf(bln住院单位, "/Nvl(C.住院包装,1)", "") & " as 库存" & _
        " From 药品库存 A,(" & strSQL & ") B,药品规格 C" & _
        " Where A.库房ID=B.ID And A.药品ID=C.药品ID" & _
        " And ((A.效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Or (Nvl(C.药房分批,0)=0 And A.库房ID IN(" & strSQL2 & ")))" & _
        " And A.性质=1 And A.药品ID=[1]" & _
        " Group by B.编码,B.名称,A.库房ID,Nvl(C.住院包装,1)" & _
        " Having Sum(Nvl(A.可用数量,0))<>0" & _
        " Order By B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng药品ID)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!名称 & ":" & rsTmp!库存
        rsTmp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    GetStockInfo = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetItemLog(ByVal int来源 As Integer, strNO As String, bytFlag As Byte, 序号 As Integer, Optional blnNOMoved As Boolean) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：读取单据行的结论
    '入参：int来源-1-门诊;2-住院
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-06 17:07:09
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If int来源 = 1 Then
        strSQL = "Select 结论 From " & IIf(blnNOMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录") & _
                " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=[2] And 序号=[3]"
    Else
        strSQL = "Select 结论 From " & IIf(blnNOMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录") & _
                " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=[2] And 序号=[3]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, 序号)
    
    If Not rsTmp.EOF Then GetItemLog = IIf(IsNull(rsTmp!结论), "", rsTmp!结论)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckNegative(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal lngItem As Long, ByVal lngExecuteDept As Long, _
    ByVal dblNum As Double, ByVal dbl住院包装 As Double, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人本次次住院的收费项目的数量合计是否足够冲销
    '入参:lngNum-输入的负数数量，如果是药品，根据参数转换成售价单位再传入如果同一单据输入相同的项目和执行科室的有多行，此时不检查，保存之前再检查
    '     strPrivs-权限串
    '出参:
    '返回:足够或有权限冲负数时,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-29 12:01:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl未结 As Double, dbl已结 As Double
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '问题:26951
    If InStr(1, strPrivs, ";负数记帐不检查发生项目;") > 0 Then
        '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
        CheckNegative = True: Exit Function
    End If
    
    '记录性质 In(2,3)取掉结帐作废的情况:  :28029
    On Error GoTo errH
    CheckNegative = True
    
   ' strSQL = "" & _
            "   Select Nvl(Sum(Nvl(付数, 1) * 数次),0) As 数量," & vbNewLine & _
            "           Sum(decode(结帐ID,NULL,0,1)* Nvl( 付数,1)* 数次) as 结帐数量  " & _
            "   From 住院费用记录" & vbNewLine & _
            "   Where  记录性质 In(2,3) and 记帐费用 = 1 And 价格父号 Is Null" & _
                    IIf(gbytBilling = 0, " And 记录状态<>0", "") & " And 病人id = [1] And 主页id = [2]" & vbNewLine & _
            "      And 收费细目id+0 = [3] And 执行部门id+0 = [4]"
    '问题:39836
    strSQL = " " & _
    "   Select Nvl(Sum(Decode(A.记录性质, 2, 1, 3, 1, 0) * Nvl(A.付数, 1) * A.数次), 0) As 数量, " & _
    "          Sum(Decode(nvl(Mod(M.记录状态, 3),1), 0, 1, 1, 1, -1) * Decode(A.结帐id, Null, 0, 1) * Nvl(A.付数, 1) * A.数次) As 结帐数量 " & _
    "   From 住院费用记录 A, 病人结帐记录 M " & _
    "   Where  A.结帐id = M.ID(+) And A.记帐费用 = 1 And A.价格父号 Is Null  " & IIf(gbytBilling = 0, " And A.记录状态<>0", "") & _
    "         And A.病人id = [1] And A.主页id = [2] And " & _
    "         A.收费细目id + 0 = [3] And A.执行部门id + 0 = [4] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, lngItem, lngExecuteDept)
    
    If Not rsTmp.EOF Then
        '问题:32106
        If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(Val(Nvl(rsTmp!数量)), 8)) Then
                MsgBox "销帐数量大于该病人本次住院在当前执行科室的记帐数量" & FormatEx(rsTmp!数量 / IIf(gbln住院单位, dbl住院包装, 1), 5) & "。", vbInformation, gstrSysName
                CheckNegative = False: Exit Function
        End If
        Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
        Case 0  '允许
        Case 1   '提醒
            dbl未结 = Val(FormatEx((Val(Nvl(rsTmp!数量)) - Val(Nvl(rsTmp!结帐数量))) / IIf(gbln住院单位, dbl住院包装, 1), 8))
            dbl已结 = Val(FormatEx(Val(Nvl(rsTmp!结帐数量)) / IIf(gbln住院单位, dbl住院包装, 1), 8))
            If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(dbl未结, 8)) Then
                If MsgBox("销帐数量(" & FormatEx(FormatEx(Abs(dblNum) / IIf(gbln住院单位, dbl住院包装, 1), 8), 5) & _
                        ") 中包含了已经结帐部分(未结:" & FormatEx(dbl未结, 5) & "; 已结:" & FormatEx(dbl已结, 5) & ") 。" & vbCrLf & _
                    " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    CheckNegative = False: Exit Function
                End If
            End If
        Case 2   '禁止
                dbl未结 = Val(FormatEx((Val(Nvl(rsTmp!数量)) - Val(Nvl(rsTmp!结帐数量))) / IIf(gbln住院单位, dbl住院包装, 1), 8))
                dbl已结 = Val(FormatEx(Val(Nvl(rsTmp!结帐数量)) / IIf(gbln住院单位, dbl住院包装, 1), 8))
                If Val(FormatEx(Abs(dblNum), 8)) > Val(FormatEx(dbl未结, 8)) Then
                    Call MsgBox("销帐数量(" & FormatEx(FormatEx(Abs(dblNum) / IIf(gbln住院单位, dbl住院包装, 1), 8), 5) & _
                        ") 中包含了已经结帐部分(未结:" & FormatEx(dbl未结, 5) & "; 已结:" & FormatEx(dbl已结, 5) & ") ,不能继续。" & vbCrLf & _
                    "", vbInformation + vbOKOnly, gstrSysName)
                    CheckNegative = False: Exit Function
                End If
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    CheckNegative = False
    Call SaveErrLog
End Function

Public Function GetPatientFeeItemTotal(ByVal lngPatient As Long, ByVal lngPage As Long, ByVal strNO As String) As ADODB.Recordset
'功能：获取指定单据的收费项目的记帐数据集合
'参数：
    Dim strSQL As String

    On Error GoTo errH
    '记录性质 In(2,3)-排除结帐作废的那种
    strSQL = "Select A.收费细目id, A.执行部门id, Sum(Nvl(A.付数, 1) * A.数次" & IIf(gbln住院单位, "/Nvl(X.住院包装,1)", "") & ") As 数量" & vbNewLine & _
            "From 住院费用记录 A,药品规格 X" & vbNewLine & _
            "Where A.记录性质 In(2,3) And A.记帐费用 = 1 And A.价格父号 Is Null And A.病人id = [1] And A.主页id = [2] And A.收费细目ID=X.药品ID(+)" & _
            IIf(gbytBilling = 0, " And A.记录状态<>0", "") & " And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From 住院费用记录 B" & vbNewLine & _
            "       Where NO = [3] And 记录性质 In(2,3) And A.收费细目id = B.收费细目id + 0 And A.执行部门id = B.执行部门id + 0)" & vbNewLine & _
            "Group By 收费细目id, 执行部门id"
    Set GetPatientFeeItemTotal = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngPatient, lngPage, strNO)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    
    Call SaveErrLog
End Function

Public Function GetNOFeeItem(ByVal strNO As String, ByVal bytFlag As Byte, Optional ByVal strRows As String) As ADODB.Recordset
'功能：获取指定单据的费用行的收费项目和执行科室
'参数：
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select 序号,收费类别,收费细目id, 执行部门id" & vbNewLine & _
            "From 住院费用记录 A" & vbNewLine & _
            "Where NO = [1] And 记录性质 = [2] And 价格父号 Is Null" & IIf(strRows = "", "", " And Instr(','||[3]||',',','||序号||',')>0")
    Set GetNOFeeItem = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, strRows)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillCanBeOperate(ByVal strNO As String, ByVal strPriv As String, _
    ByVal strNote As String, Optional ByVal strTime As String, _
    Optional str病人IDs As String, Optional ByVal bytType As Byte = 2, _
    Optional ByVal byt费用来源 As Byte) As Boolean
'功能：根据单据的病人信息判断是否有权限操作该单据
'参数：strNote=描述操作类型,用于提示。销帐时有特殊处理。
'      str病人IDs=允许时，返回允许操作的病人ID串,空为所有病人
'      byt费用来源 0-住院,1-门诊
'说明：主要是病人出院(或预出院)后,如果没有权限,则不允许操作
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnOut As Boolean
    Dim strInfo As String
    
    str病人IDs = ""
    
    If InStr(strPriv, ";出院未结强制记帐;") > 0 _
        And InStr(strPriv, ";出院结清强制记帐;") > 0 Then
        BillCanBeOperate = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '如果无对应主页,则当作已出院病人(如门诊病人医技记帐)
    If strNote Like "*销帐" Then
        '销帐操作时,只对可以销帐部份内容进行判断
        strSQL = _
            " Select 序号 From 住院费用记录" & _
            " Where 记录性质=[2] And NO=[1] And Nvl(执行状态,0)<>1 And 价格父号 is NULL" & _
            " Group by 序号 Having Nvl(Sum(Nvl(付数,1)*数次),0)<>0"
    ElseIf strNote Like "*审核" Then
        '审核操作时,只对未审核部份内容进行判断
        strSQL = _
            " Select 序号 From 住院费用记录" & _
            " Where 记录性质=2 And 价格父号 is NULL And 记录状态=0 And NO=[1]"
    End If
    strSQL = "Select Distinct 姓名,病人ID,主页ID From 住院费用记录" & _
        " Where 记录性质=[2] And NO=[1] And 记录状态 IN(0,1,3)" & _
        IIf(strTime <> "", " And 登记时间=[3]", "") & _
        IIf(strSQL <> "", " And Nvl(价格父号,序号) IN(" & strSQL & ")", "")

    strSQL = "Select B.病人ID,B.姓名," & _
    " Decode(A.病人ID,NULL,Sysdate,A.出院日期) as 出院日期," & _
    " Nvl(A.状态,0) as 状态,Nvl(C.费用余额,0) as 余额" & _
    " From 病案主页 A,(" & strSQL & ") B,病人余额 C" & _
    " Where B.病人ID=A.病人ID(+) And C.性质(+)=1 And C.类型(+)=2  And B.主页ID=A.主页ID(+) And B.病人ID=C.病人ID(+) And C.性质(+)=1 And C.类型(+)=2 "
    
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    End If
    
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!出院日期) Or rsTmp!状态 = 3 Then
            If rsTmp!余额 = 0 And InStr(strPriv, ";出院结清强制记帐;") = 0 Then
                strInfo = strInfo & vbCrLf & "病人""" & rsTmp!姓名 & """已出院(或预出院)且费用已经结清。"
            ElseIf rsTmp!余额 <> 0 And InStr(strPriv, ";出院未结强制记帐;") = 0 Then
                strInfo = strInfo & vbCrLf & "病人""" & rsTmp!姓名 & """已出院(或预出院)且费用尚未结清。"
            Else
                str病人IDs = str病人IDs & "," & rsTmp!病人ID
            End If
        Else
            str病人IDs = str病人IDs & "," & rsTmp!病人ID
        End If
        rsTmp.MoveNext
    Loop
    str病人IDs = Mid(str病人IDs, 2)
        
    '只有记帐表销帐可以部份继续
    If str病人IDs = "" Or (strInfo <> "" And strNote <> "销帐") Then
        MsgBox Mid(strInfo, 3) & vbCrLf & "你没有权限对单据""" & strNO & """进行" & strNote & "。", vbInformation, gstrSysName
        Exit Function
    Else
        If UBound(Split(str病人IDs, ",")) + 1 = rsTmp.RecordCount Then str病人IDs = ""
        If strInfo <> "" Then
            MsgBox Mid(strInfo, 3) & vbCrLf & "你只能对单据中其他病人的费用进行" & strNote & "。", vbInformation, gstrSysName
        End If
    End If
    
    BillCanBeOperate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function BillCanModi(strNO As String, bytFlag As Byte) As Boolean
'功能：判断一张单据是否可以修改
'参数：bytFlag=记录性质
'说明：如果单据中存在分批或时价药品,则不允许修改(因为库存的问题)
'***
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.ID" & _
        " From 住院费用记录 A,药品规格 B,收费项目目录 C" & _
        " Where A.收费细目ID=B.药品ID And A.收费细目ID=C.ID" & _
        " And A.记录状态 IN(0,1,3) And (Nvl(B.药房分批,0)=1 Or Nvl(C.是否变价,0)=1)" & _
        " And A.NO=[1] And A.记录性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillCanModi = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Read药品信息(lng药品ID As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select A.* From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng药品ID)
    If Not rsTmp.EOF Then Set Read药品信息 = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function BillCanDelete(ByVal strNO As String, ByVal bytFlag As Byte, _
    Optional ByVal blnBat As Boolean, Optional ByVal strTime As String, _
    Optional ByVal strPrivs As String, Optional ByRef blnFlagPrint As Boolean, _
    Optional ByVal byt费用来源 As Byte) As Integer
'功能：判断一张单据是否可以退费或销帐
'参数：strNO=单据号,bytFlay=记录性质,blnBat=是否多病人单,strTime=单据的登记时间
'      strPrivs=如果传入，则用于判断否有药品销帐或诊疗销帐权限(医技记帐,简单记帐可以不传)
'      byt费用来源 0-住院,1-门诊
'说明：可以退费或销帐的条件
'    1.费用未完全执行(执行状态=0,2)
'    2.剩余数量不<>0
'返回：
'   -1=操作失败
'    0=可以退费或销帐
'    1=单据不存在或没有该类别收费项目的销帐权限
'    2=已经全部完全执行(执行状态=1)
'    3=未完全执行部分剩余数量为0
'    blnFlagPrint=检查对应的条码是否已打印(检验医嘱中的采集方式已执行)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strFeeKind As String
    
    On Error GoTo errH
    '之前已检查,至少有一种销帐权限
    If strPrivs <> "" Then
        '55380
        Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
        blnYP = zlStr.IsHavePrivs(";" & strPrivs & ";", "药品销帐")
        blnZL = zlStr.IsHavePrivs(";" & strPrivs & ";", "诊疗销帐")
        blnWC = zlStr.IsHavePrivs(";" & strPrivs & ";", "卫材销帐")
        If blnYP And blnWC And blnZL Then
            '所有,不限制
        ElseIf blnYP And blnWC And Not blnZL Then
            strFeeKind = " And 收费类别   In('4','5','6','7')"
        ElseIf blnYP And Not blnWC And blnZL Then
            strFeeKind = " And 收费类别   <>'4'"
        ElseIf blnYP And Not blnWC And Not blnZL Then
            strFeeKind = " And 收费类别 In('5','6','7')"
        ElseIf Not blnYP And blnWC And blnZL Then
            strFeeKind = " And 收费类别 Not In('5','6','7')"
        ElseIf Not blnYP And Not blnWC And blnZL Then
            strFeeKind = " And 收费类别 Not In('4','5','6','7')"
        ElseIf Not blnYP And blnWC And Not blnZL Then
            strFeeKind = " And 收费类别 ='4'"
        End If
    End If
    
    '1.费用未完全执行(执行状态=0,2)
    strSQL = "Select Distinct Nvl(A.执行状态,0) as 执行状态,B.样本条码" & _
        " From 住院费用记录 A,病人医嘱发送 B" & vbNewLine & _
        " Where A.NO=[1] And A.记录性质=[2] And A.记录状态 IN(0,1,3)" & IIf(byt费用来源 = 0, " And Nvl(多病人单,0)=[3]", "") & vbNewLine & _
        " And A.医嘱序号=B.医嘱ID(+) And A.NO=B.NO(+) And A.记录性质=B.记录性质(+)" & _
        IIf(strTime <> "", " And 登记时间=[4]", "") & strFeeKind
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, IIf(blnBat, 1, 0), CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, IIf(blnBat, 1, 0))
    End If
    
    If rsTmp.EOF Then BillCanDelete = 1: Exit Function '单据不存在或没有该类别收费项目的销帐权限
    blnFlagPrint = Not IsNull(rsTmp!样本条码)
    
    '单据已经全部完全执行
    rsTmp.Filter = "执行状态<>1"
    If rsTmp.EOF Then BillCanDelete = 2 ': Exit Function
    
    
    '未完全执行部分剩余数量不<>0
    '从原始单据中找未完全执行的行次(部分退药的退费后执行状态=1,但退费记录执行状态<>1)
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
        " Select Nvl(价格父号,序号) as 序号" & _
        " From 住院费用记录" & _
        " Where Nvl(执行状态,0)<>1 And NO=[1] And 记录性质=[2] And 记录状态 IN(0,1,3)" & _
                IIf(strTime <> "", " And 登记时间=[3]", "")
    strSQL = _
        " Select 序号,收费细目ID,Sum(数量) as 剩余数 " & _
        " From ( Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号,收费细目ID," & _
        "               Avg(Nvl(付数,1)*数次) as 数量 " & _
        "        From 住院费用记录" & _
        "        Where NO=[1] And 记录性质=[2] And Nvl(执行状态,0)<>1 And Nvl(价格父号,序号) IN(" & strSQL & ")" & _
        "        Group by 记录状态,执行状态,Nvl(价格父号,序号),收费细目ID " & _
        "       )" & _
        " Group by 序号,收费细目ID  " & _
        " Having Sum(数量)<>0"
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then BillCanDelete = 3
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    BillCanDelete = -1
End Function

Public Function BillExistDelete(strNO As String, bytFlag As Byte) As Boolean
'功能：判断指定单据是否包含(部分)退费或销帐的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select NO From 住院费用记录 Where NO=[1] And 记录性质=[2] And 记录状态=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillExistDelete = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistMoney(strNO As String, bytFlag As Byte) As Boolean
'功能：判断指定单据的项目是否已经全部退完(剩余数量=0)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
        " Select 序号,Sum(数量) as 剩余数量" & _
        " From (" & _
            " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号,Avg(Nvl(付数, 1) * 数次) As 数量" & _
            " From 住院费用记录" & _
            " Where NO=[1] And 记录性质=[2]" & _
            " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    BillExistMoney = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckMediCareItem(ByVal lng收费细目ID As Long, ByVal int险类 As Integer, ByVal str收费项目名称 As String, ByVal bln定价 As Boolean, _
    Optional blnErrShowInsureName As Boolean = False, Optional ByVal strPriceGrade As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断收费项目是否设置了保险支付项目
    '入参:lng收费细目ID-收费细目ID
    '     int险类-险类
    '     str收费项目名称-收费项目名称
    '     blnErrShowInsureName-出现提示时,是否显示险类名称
    '返回:存在对码返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-07-09 11:45:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String, rs价格 As ADODB.Recordset, dbl价格 As Double
    Dim strInsureName As String, strWherePriceGrade As String
    
    CheckMediCareItem = True
    If gbyt医保对码检查 = 0 Then Exit Function
    
    If gclsInsure.GetCapability(support允许不设置医保项目, , int险类) Then Exit Function
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
        " Select  B.现价 " & _
        " From 收费价目 B " & _
        " Where   ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And B.收费细目ID=[1]" & vbNewLine & _
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
        strInsureName = ""
        If blnErrShowInsureName Then
            strSQL = "Select 名称  From 保险类别 where 序号=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", int险类)
            If Not rsTmp.EOF Then
                strInsureName = "『" & Nvl(rsTmp!名称) & "』"
            End If
        End If
        If gbyt医保对码检查 = 1 Then
            If MsgBox(strInsureName & "没有设置""" & str收费项目名称 & """对应的保险项目,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbyt医保对码检查 = 2 Then
            MsgBox strInsureName & "没有设置""" & str收费项目名称 & """对应的保险项目!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveNOAuditing(ByVal lng病人ID As Long, Optional ByVal strHosTimes As String) As Boolean
'功能：判断病人未结费用中是否存在未审核记帐费用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '77686,李南春,2014/9/18,单据类别限制
    If strHosTimes = "" Then
        strSQL = _
            "Select 1 From 住院费用记录 A" & _
                " Where 记帐费用=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And 病人ID=[1] And Not Exists" & _
                " (Select 1 From 药品收发记录 C Where A.ID = C.费用ID And Mod(C.记录状态, 3) = 1 And Nvl(C.摘要,'大一')='拒发' And instr( ',8,9,10,21,24,25,26,',','||C.单据||',')>0) And Not Exists" & _
                " (Select 1 From 病人医嘱发送 B Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱序号=B.医嘱ID And B.执行状态 = 2) And Rownum=1"
    Else
        strSQL = _
        "Select /*+ rule*/ 1 From 住院费用记录 A,Table(f_num2list([2])) B" & _
            " Where 记帐费用=1 And 记录状态=0 And Nvl(实收金额,0)<>0 And 病人ID=[1] And Not Exists" & _
            " (Select 1 From 药品收发记录 C Where A.ID = C.费用ID And Mod(C.记录状态, 3) = 1 And Nvl(C.摘要,'大一')='拒发' And instr( ',8,9,10,21,24,25,26,',','||C.单据||',')>0) And Not Exists" & _
            " (Select 1 From 病人医嘱发送 B Where A.NO=B.NO And A.记录性质=B.记录性质 And A.医嘱序号=B.医嘱ID And B.执行状态 = 2) And Rownum=1 And A.主页ID=B.COLUMN_VALUE"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, strHosTimes)
    HaveNOAuditing = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BalanceExistInsure(strNO As String, Optional ByRef bytFlag As Byte, Optional ByRef lng病人ID As Long) As Integer
'功能：判断结帐记录中是否存在指定的医保结算方式
'参数：strNO=收费单据号,bytFlag-医保结算性质:1-门诊，2-住院
'返回：如果存在,则返回病人险类
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    lng病人ID = 0
    On Error GoTo errH
    
    strSQL = "Select B.险类,B.性质,nvl(A.病人ID,B.病人ID) as 病人ID  From 病人结帐记录 A,保险结算记录 B" & _
       " Where A.记录状态 IN(1,3) And A.NO=[1]" & _
       "    And A.ID=B.记录ID And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    
    If Not rsTmp.EOF Then
        BalanceExistInsure = Val(IIf(IsNull(rsTmp!险类), 0, rsTmp!险类))
        lng病人ID = Val(Nvl(rsTmp!病人ID))
        bytFlag = Val("" & rsTmp!性质)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetBillInsures(strInsure As String, ByVal strNO As String, _
    Optional ByVal strTime As String, Optional ByVal blnAuditing As Boolean, _
    Optional ByVal blnGetNoneInsure As Boolean, Optional ByVal bytFlag As Byte = 2, _
    Optional ByVal byt费用来源 As Byte) As Boolean
'功能：获取记帐表中的险类串"10,20,30,...",也适用于记帐单
'参数：strNO=记帐单据号
'      blnAuditing=是否用于记帐审核,只检查未审核的部份内容
'      blnGetNoneInsure=是否将非保险费用返回为0险类
'      byt费用来源 0-住院,1-门诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strInsure = ""
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.险类,0) as 险类" & _
        " From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=[2] And A.记录状态" & IIf(blnAuditing, "=0", " IN(0,1,3)") & _
            IIf(blnGetNoneInsure, "", " And B.险类 is Not NULL") & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    Do While Not rsTmp.EOF
        strInsure = strInsure & "," & rsTmp!险类
        rsTmp.MoveNext
    Loop
    strInsure = Mid(strInsure, 2)
    GetBillInsures = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDelPriv(ByVal strNO As String, ByVal strPrivs As String, Optional ByVal strTime As String, _
        Optional ByVal bytFlag As Byte = 2, Optional ByVal bytMode As Byte = 1, _
        Optional ByVal byt费用来源 As Byte) As Boolean
'功能：检查是否权限冲销住院记帐单
'入参：
'      byt费用来源 0-住院,1-门诊
'参数: bytMode,部分权限不足时是否仅提示,1-允许继续,返回真,0-不允继续,返回假
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '只判断未销帐费用行
    strSQL = "Select Nvl(Sum(Decode(收费类别,'5',1,'6',1,'7',1,0)),0) as 药品数," & _
        " Nvl(Sum(Decode(收费类别,'4',1,0)),0) as 卫材数," & _
        " Nvl(Sum(Decode(收费类别,'4',0,'5',0,'6',0,'7',0,1)),0) as 诊疗数" & _
        " From 住院费用记录" & _
        " Where 记录性质=[2] And 记录状态 IN(0,1) And NO=[1]" & _
        IIf(strTime <> "", " And 登记时间=[3]", "")
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    
    If rsTmp.EOF Then CheckDelPriv = True: Exit Function
    '没有住院销帐权限时,菜单和按钮已设置为不可见
    '55380
    '55380
    Dim blnYP As Boolean, blnZL As Boolean, blnWC As Boolean
    Dim strNotPrivs As String, strNote As String
    
    blnYP = zlStr.IsHavePrivs(";" & strPrivs & ";", "药品销帐")
    blnZL = zlStr.IsHavePrivs(";" & strPrivs & ";", "诊疗销帐")
    blnWC = zlStr.IsHavePrivs(";" & strPrivs & ";", "卫材销帐")
    
    If blnYP = False And blnZL = False And blnWC = False Then
        MsgBox "你没有药品销帐或卫材销帐或诊疗销帐的权限,不能对单据[" & strNO & "]进行销帐！", vbInformation, gstrSysName
        Exit Function
    End If
    strNotPrivs = ""
    If Not blnYP Then strNotPrivs = strNotPrivs & "和药品销帐"
    If Not blnWC Then strNotPrivs = strNotPrivs & "和卫材销帐"
    If Not blnZL Then strNotPrivs = strNotPrivs & "和诊疗销帐"
    strNotPrivs = Mid(strNotPrivs, 2)
    strNote = ""
    
    If blnYP Then strNote = strNote & "或药品销帐"
    If blnWC Then strNote = strNote & "或卫材销帐"
    If blnZL Then strNote = strNote & "或诊疗销帐"
    strNote = Mid(strNote, 2)
    
    If rsTmp!药品数 > 0 And Not blnYP Then
        MsgBox "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!卫材数 > 0 And Not blnWC Then
        MsgBox "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    If rsTmp!诊疗数 > 0 And Not blnZL Then
        MsgBox "你没有" & strNotPrivs & "权限,只能对单据[" & strNO & "]中的" & strNote & "进行销帐！", vbInformation, gstrSysName
        If bytMode = 0 Then Exit Function
    End If
    CheckDelPriv = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get服务对象(lng收费细目ID As Long) As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 服务对象 From 收费项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng收费细目ID)
    If Not rsTmp.EOF Then Get服务对象 = IIf(IsNull(rsTmp!服务对象), 0, rsTmp!服务对象)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check留观病人(ByVal strNO As String, ByVal strPrivs As String, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2, _
    Optional ByVal byt费用来源 As Byte) As String
'功能：根据是否允许对留观病人人进行记帐,对记帐单/表进行检查
'入参：
'      byt费用来源 0-住院,1-门诊
'说明：主要用于记帐单/表修改,销帐。对于记帐表,只要存在一个留观病人无权限,则整单禁止
'返回：没有权限的留观病人,如"留观病人","门诊留观病人","住院留观病人"
    Dim rsTmp As ADODB.Recordset
    Dim bln门诊留观 As Boolean
    Dim bln住院留观 As Boolean
    Dim strSQL As String
    
    bln门诊留观 = gbln门诊留观 And InStr(strPrivs, ";门诊留观记帐;") > 0
    bln住院留观 = gbln住院留观 And InStr(strPrivs, ";住院留观记帐;") > 0
        
    If bln门诊留观 And bln住院留观 Then Exit Function
    
    If Not bln门诊留观 And Not bln住院留观 Then
        strSQL = "1,2"
    ElseIf Not bln门诊留观 Then
        strSQL = "1"
    ElseIf Not bln住院留观 Then
        strSQL = "2"
    End If
    
    On Error GoTo errH
    
    strSQL = "Select Distinct Nvl(B.病人性质,0) as 病人性质" & _
        " From 住院费用记录 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        " And A.NO=[1] And A.记录性质=[2]" & _
        " And Nvl(B.病人性质,0) IN(" & strSQL & ") And A.记录状态 IN(0,1,3)" & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    End If
    If Not rsTmp.EOF Then
        If rsTmp.RecordCount = 2 Then
            Check留观病人 = "留观病人"
        ElseIf rsTmp!病人性质 = 1 Then
            Check留观病人 = "门诊留观病人"
        ElseIf rsTmp!病人性质 = 2 Then
            Check留观病人 = "住院留观病人"
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillRows(strNO As String, bytFlag As Byte) As Integer
'功能：获取一张费用单据中未作废的费用行数
'参数：bytFlag=记录性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    strSQL = _
        " Select 序号,Sum(数量) as 剩余数量" & _
        " From (" & _
            " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号,Avg(Nvl(付数, 1) * 数次) As 数量" & _
            " From 住院费用记录" & _
            " Where NO=[1] And 记录性质=[2]" & _
            " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    If Not rsTmp.EOF Then GetBillRows = rsTmp.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetMaxBedLen(Optional lng部门ID As Long, Optional bln科室 As Boolean) As Integer
'功能：获取指定部门的床位号的最大长度
'参数：lng部门ID=病区ID或科室ID,为0表示所有病区或科室
'      bln占用=是否只管被占用的床
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not bln科室 Or lng部门ID = 0 Then
        strSQL = "Select Max(Lengthb(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 病区ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    Else
        strSQL = "Select Max(Lengthb(床号)) as 长度 From 床位状况记录 Where 状态='占用' And 科室ID" & IIf(lng部门ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng部门ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!长度), 0, rsTmp!长度)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckLimit(ByVal objBill As ExpenseBill, Optional intRow As Integer, Optional ByVal bln住院单位 As Boolean) As Boolean
'功能：费用单据药品处方限量检查,适用于记帐单/表
'说明：
'   1.全部没超过限量，返回真；如有超过药品，则在函数内提示，并返回假。
'   2.记帐表是为每个病人单独检查
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim tmpDetail As BillDetail, curDetail As BillDetail
    Dim i As Integer, j As Integer, dblTime As Double
    Dim dbl剂量 As Double, strItemIDs As String '已经检查过了的药品
    Dim strPatiIDs As String, arrPati As Variant '已经检查过了的病人
    Dim lng病人ID As Long, str姓名 As String
    Dim str药品限量提示 As String
    
    CheckLimit = True
    If objBill.Details.Count = 0 Then Exit Function
    Err = 0: On Error GoTo errH:
    '收集病人
    For i = 1 To objBill.Details.Count
        If intRow = 0 Or (intRow > 0 And i = intRow) Then
            With objBill.Details(i)
                '收集药品ID
                If InStr(strItemIDs & ",", "," & .收费细目ID & ",") = 0 And InStr(",5,6,7,", .收费类别) > 0 Then
                    strItemIDs = strItemIDs & "," & .收费细目ID
                End If
                '收集病人信息
                If InStr(strPatiIDs & ";", ";" & .病人ID & "," & .姓名 & ";") = 0 Then
                    strPatiIDs = strPatiIDs & ";" & .病人ID & "," & .姓名
                End If
            End With
        End If
    Next
    If strItemIDs = "" Then Exit Function
    strItemIDs = Mid(strItemIDs, 2)
    arrPati = Split(Mid(strPatiIDs, 2), ";")
        
    strSQL = "Select A.药品ID,A.剂量系数,B.计算单位 as 剂量单位" & _
        " From 药品规格 A,诊疗项目目录 B" & _
        " Where A.药名ID=B.ID And A.药品ID IN (" & strItemIDs & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    
    For i = 0 To UBound(arrPati)
        lng病人ID = Val(Split(arrPati(i), ",")(0))
        str姓名 = CStr(Split(arrPati(i), ",")(1))
        strItemIDs = ""
        For j = 1 To objBill.Details.Count
            If intRow = 0 Or (intRow > 0 And j = intRow) Then
                Set tmpDetail = objBill.Details(j)
                If InStr(",5,6,7,", tmpDetail.收费类别) > 0 And tmpDetail.Detail.处方限量 > 0 And tmpDetail.病人ID = lng病人ID Then
                    If InStr(strItemIDs, "," & tmpDetail.收费细目ID) = 0 Then
                        dblTime = 0
                        For Each curDetail In objBill.Details
                            If InStr(",5,6,7,", curDetail.收费类别) > 0 And tmpDetail.收费细目ID = curDetail.收费细目ID And curDetail.病人ID = lng病人ID Then
                                dblTime = dblTime + curDetail.付数 * curDetail.数次
                            End If
                        Next
                        rsTmp.Filter = "药品ID=" & tmpDetail.收费细目ID
                        If Not rsTmp.EOF Then
                            If bln住院单位 Then
                                dbl剂量 = dblTime * tmpDetail.Detail.住院包装 * rsTmp!剂量系数
                            Else
                                dbl剂量 = dblTime * rsTmp!剂量系数
                            End If
                            If dbl剂量 > tmpDetail.Detail.处方限量 Then
                                str药品限量提示 = IIf(str姓名 = "", "", """" & str姓名 & """ 的") & "药品 """ & tmpDetail.Detail.名称 & """ 的总剂量 " & _
                                    FormatEx(dbl剂量, 5) & rsTmp!剂量单位 & "(" & FormatEx(dblTime, 5) & IIf(bln住院单位, tmpDetail.Detail.住院单位, tmpDetail.Detail.计算单位) & ") 超过处方限量 " & _
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
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDiagnosticInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                                  ByVal str诊断类型 As String, ByVal str记录来源 As String) As ADODB.Recordset
'功能：获取指定病人的诊断记录'
'参数:
'诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);

    On Local Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " Select 诊断类型,记录来源,诊断描述,疾病ID,诊断ID,出院情况,是否疑诊 From 病人诊断记录 " & _
             " Where 病人ID=[1] And Nvl(主页ID,0)=[2]" & _
             " And 诊断次序=1 And instr([3],','||诊断类型||',')>0 And 记录来源 in (" & str记录来源 & ")" & _
             " Order by 记录日期 Desc"
            '诊断次序-出院时,病案主页整理中可能填写主要诊断,次要诊断等多条记录
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID, "," & str诊断类型 & ",")
    
    If Not rsTmp.EOF Then Set GetDiagnosticInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepCharacter(ByVal lngDepID As Long) As String
'功能：获取部门工作性质
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 工作性质 From 部门性质说明 Where 部门ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lngDepID)
    
    Do While Not rsTmp.EOF
        If InStr(1, GetDepCharacter & ",", "," & rsTmp!工作性质 & ",") = 0 Then
            GetDepCharacter = GetDepCharacter & "," & rsTmp!工作性质
        End If
        rsTmp.MoveNext
    Loop
    
    If GetDepCharacter <> "" Then GetDepCharacter = Mid(GetDepCharacter, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckInhibitiveByNurse(ByRef objBill As ExpenseBill, ByRef rs开单人 As ADODB.Recordset) As Boolean
'功能：判断指定单据中是否有护士禁止输入的内容
    Dim bln护士 As Boolean, i As Integer
    
    CheckInhibitiveByNurse = False
    If objBill.开单人 <> "" Then
        Call GetOperatorInfo(rs开单人, objBill.开单人, bln护士)
        If Not bln护士 Then Exit Function
        
        For i = 1 To objBill.Details.Count
            If InStr(",E,M,4,", objBill.Details(i).收费类别) = 0 Then
                CheckInhibitiveByNurse = True: Exit Function
            End If
        Next
    End If
End Function

Public Function CheckErrorItem() As Boolean
'功能：检查用于处理金额小数误差的项目是否设置正确
'说明：该项目不应撤档，也应为变价项目(未管)。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.ID,B.类别,B.编码,B.名称" & _
        " From 收费特定项目 A,收费项目目录 B" & _
        " Where A.特定项目='误差项' And A.收费细目ID=B.ID" & _
        " And (B.撤档时间 is NULL Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    CheckErrorItem = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Init开单人开单科室(ByRef cbo开单人 As ComboBox, ByRef cbo开单科室 As ComboBox, _
                            ByRef rs开单人 As ADODB.Recordset, ByRef rs开单科室 As ADODB.Recordset, _
                            ByVal strPrivs As String, ByVal bytUseType As Byte, ByVal lngDeptID As Long _
                            ) As Boolean
'功能:初始化开单人,开单科室列表,不触发Click事件
'参数:lngDeptID-当前操作的病区ID,所有病区时为0

    '1.开单人决定开单科室,不缺省开单人(除非仅有一个)
    If gblnFromDr Then
        Call FillDoctor(cbo开单人, rs开单人)
        If cbo开单人.ListCount = 1 Then Call zlControl.CboSetIndex(cbo开单人.hWnd, 0)
        
        Call FillDept(cbo开单科室, rs开单科室, rs开单人, strPrivs, bytUseType, lngDeptID)
        If cbo开单科室.ListCount = 0 Then
            MsgBox "没有初始化住院临床科室,请先到部门管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
        If cbo开单科室.ListIndex = -1 And cbo开单科室.ListCount = 1 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, 0)
    
    '2.开单科室决定开单人,显示缺省开单科室
    Else
        Call FillDept(cbo开单科室, rs开单科室, rs开单人, strPrivs, bytUseType, lngDeptID)
        If cbo开单科室.ListCount = 0 Then
            MsgBox "没有初始化住院临床科室,请先到部门管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '缺省显示当前病区,如果当前是所有病区,则显示第一个
        If lngDeptID <> 0 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lngDeptID))
        If cbo开单科室.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, 0)
        
        Call FillDoctor(cbo开单人, rs开单人, cbo开单科室.ItemData(cbo开单科室.ListIndex))
        If cbo开单人.ListCount = 1 Then Call zlControl.CboSetIndex(cbo开单人.hWnd, 0)
    End If
    
    If cbo开单科室.ListCount > 0 Then Call SetWidth(cbo开单科室.hWnd, GetWidth(cbo开单科室.hWnd) * 1.2)
    Init开单人开单科室 = True
End Function


Public Sub Set开单人开单科室(ByRef cbo开单人 As ComboBox, ByRef cbo开单科室 As ComboBox, _
           ByRef rs开单人 As ADODB.Recordset, ByRef rs开单科室 As ADODB.Recordset, _
           ByVal str开单人 As String, ByVal lng开单科室ID As Long _
           )
'功能：根据系统参数设置开单人和开单科室，但不触发其Click事件
'       主要目的在于禁止隐式调用Click时对开单人，开单科室的相互影响，及相关数据影响(例如：会改变单据对象中的对应值)
    Dim lng人员ID As Long, str部门名称 As String
    
    'a.开单人定开单科室
    If gblnFromDr Then
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True))
        
        If cbo开单人.ListIndex = -1 And str开单人 <> "" Then
            lng人员ID = GetPersonnelID(str开单人, rs开单人)
            cbo开单人.AddItem str开单人
            cbo开单人.ItemData(cbo开单人.NewIndex) = lng人员ID
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo开单人.NewIndex)
        End If
        
        If cbo开单人.ListIndex <> -1 Then
            cbo开单科室.Clear
            Call FillDept(cbo开单科室, rs开单科室, rs开单人, "", 0, 0, cbo开单人.ItemData(cbo开单人.ListIndex))
        End If
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        If cbo开单科室.ListIndex = -1 And lng开单科室ID <> 0 Then
            str部门名称 = GET部门名称(lng开单科室ID, rs开单科室)
            If str部门名称 <> "" Then
                cbo开单科室.AddItem str部门名称
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = lng开单科室ID
                Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo开单科室.NewIndex)
            End If
        End If
        
    'b.开单科室定开单人
    Else
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        If cbo开单科室.ListIndex = -1 And lng开单科室ID <> 0 Then
            str部门名称 = GET部门名称(lng开单科室ID, rs开单科室)
            If str部门名称 <> "" Then
                cbo开单科室.AddItem str部门名称
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = lng开单科室ID
                Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo开单科室.NewIndex)
            End If
        End If
        
        If cbo开单科室.ListIndex <> -1 Then
            cbo开单人.Clear
            Call FillDoctor(cbo开单人, rs开单人, lng开单科室ID)
        End If
        Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo.FindIndex(cbo开单人, str开单人, True))
        If cbo开单人.ListIndex = -1 And str开单人 <> "" Then
            lng人员ID = GetPersonnelID(str开单人, rs开单人)
            cbo开单人.AddItem str开单人
            cbo开单人.ItemData(cbo开单人.NewIndex) = lng人员ID
            Call zlControl.CboSetIndex(cbo开单人.hWnd, cbo开单人.NewIndex)
        End If
    End If
End Sub

Public Function SetDefaultDept(ByRef cbo开单科室 As ComboBox, ByRef rs开单科室 As ADODB.Recordset, _
                               ByRef rs开单人 As ADODB.Recordset, ByVal lng开单人ID As Long _
                                ) As Boolean
'功能:根据开单人设置缺省的开单科室,但不触发Click事件
'说明:缺省科室为"只服务于住院"时，可以定位缺省
'     或者开单人的所有科室都为同一优先排序级别时(如都是即服务于门诊或住院的)，可以定位缺省
'     否则,按编码排序,取第一个

    Dim i As Long, lng开单科室ID As Long, blnDo As Boolean, lng优先级 As Long
    
    If cbo开单科室.ListCount = 1 Then
        Call zlControl.CboSetIndex(cbo开单科室.hWnd, 0)
    Else
        rs开单人.Filter = "缺省=1 And ID=" & lng开单人ID
        If rs开单人.RecordCount > 0 Then lng开单科室ID = rs开单人!部门ID
        
        If rs开单科室.RecordCount > 1 And lng开单科室ID > 0 Then
            rs开单科室.MoveFirst
            For i = 1 To rs开单科室.RecordCount
                If lng开单科室ID = rs开单科室!ID And rs开单科室!优先级 = 1 Then blnDo = True: Exit For
                rs开单科室.MoveNext
            Next
            
            If Not blnDo Then
                blnDo = True
                rs开单科室.MoveFirst
                For i = 1 To rs开单科室.RecordCount
                    If lng优先级 <> rs开单科室!优先级 And lng优先级 <> 0 Then blnDo = False: Exit For
                    lng优先级 = rs开单科室!优先级
                    rs开单科室.MoveNext
                Next
            End If
            
            If blnDo Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, cbo.FindIndex(cbo开单科室, lng开单科室ID))
        End If
        
        If cbo开单科室.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo开单科室.hWnd, 0)
    End If
End Function

Public Sub FillDoctor(ByRef cbo开单人 As ComboBox, ByRef rs开单人 As Recordset, Optional ByVal lng科室ID As Long)
'功能：根据指定的开单科室ID读取并填写医生列表,但不缺省医生
    Dim strOldID As String
    
    cbo开单人.Clear
    Call GetDoctor(lng科室ID, gbln护士 And (gstr收费类别 = "" _
        Or gstr收费类别 Like "*'E'*" Or gstr收费类别 Like "*'M'*" Or gstr收费类别 Like "*'4'*"), rs开单人)
    
    Do While Not rs开单人.EOF
        '70857:刘尔旋,2014-03-07,开单人简码一致时存在加载重复的问题
        If InStr("," & strOldID & ",", "," & rs开单人!ID & ",") = 0 Then
            If gbyt开单人显示 = 1 Then
                cbo开单人.AddItem rs开单人!简码 & "-" & rs开单人!姓名
            Else
                cbo开单人.AddItem rs开单人!编号 & "-" & rs开单人!姓名
            End If
            cbo开单人.ItemData(cbo开单人.NewIndex) = rs开单人!ID
            strOldID = strOldID & rs开单人!ID & ","
        End If
        rs开单人.MoveNext
    Loop
End Sub

Public Sub FillDept(ByRef cbo开单科室 As ComboBox, ByRef rs开单科室 As Recordset, ByRef rs开单人 As Recordset, _
                   ByVal strPrivs As String, ByVal bytUseType As Byte, ByVal lngDeptID As Long, Optional ByVal lng人员ID As Long)
'功能：读取并加载科室列表,但不缺省科室
'参数： lngDeptID-当前操作的病区
'       lng人员ID=只读取指定人员所在科室(包含非缺省的)
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  '指定人员所属的多个部门
        
    cbo开单科室.Clear
    If rs开单科室 Is Nothing Then Call GetDoctorDept(rs开单科室, strPrivs, bytUseType, lngDeptID)
   
    If lng人员ID <> 0 Then
        If Not rs开单人 Is Nothing Then
            rs开单人.Filter = "ID=" & lng人员ID
            For i = 1 To rs开单人.RecordCount
                strDepts = strDepts & " OR ID=" & rs开单人!部门ID      'filter不支持in
                rs开单人.MoveNext
            Next
        End If
        If strDepts <> "" Then
            rs开单科室.Filter = Mid(strDepts, 4)
        Else
            rs开单科室.Filter = "ID=0" '人员没有设置部门,不显示开单科室
        End If
    Else
        rs开单科室.Filter = ""
    End If
    
    If rs开单科室.RecordCount > 0 Then
        For i = 1 To rs开单科室.RecordCount
            If lngOldDepID <> rs开单科室!ID Then   '一个部门可能同时属于产科和临床,不加载相同的
                cbo开单科室.AddItem IIf(zlIsShowDeptCode, rs开单科室!编码 & "-", "") & rs开单科室!名称
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = rs开单科室!ID
                lngOldDepID = rs开单科室!ID
            End If
            rs开单科室.MoveNext
        Next
    End If
End Sub


Public Function GetOperatorInfo(ByVal rs开单人 As Recordset, ByVal str姓名 As String, _
                                Optional ByRef bln护士 As Boolean, Optional int职务 As Integer) As Boolean
'功能：获取指定姓名开单人(医生或护士)的性质或职务
'返回：int职务:0-未设置；bln护士:是否只是护士
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    bln护士 = False: int职务 = 0
    If Not rs开单人 Is Nothing Then
        rs开单人.Filter = "姓名='" & str姓名 & "' " & IIf(gbln护士, "", " And 人员性质<>'护士'")
        If rs开单人.RecordCount > 0 Then
            int职务 = rs开单人!职务
            strSQL = rs开单人!人员性质
            If strSQL = "护士" Then bln护士 = True
            If strSQL = "医生" Then bln护士 = False
        End If
    Else
        strSQL = _
            " Select Nvl(A.聘任技术职务,0) as 职务,B.人员性质 From 人员表 A,人员性质说明 B" & _
            " Where A.ID=B.人员ID And B.人员性质 IN('医生','护士') And A.姓名=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", str姓名)
        If Not rsTmp.EOF Then
            int职务 = rsTmp!职务
            Do While Not rsTmp.EOF
                If rsTmp!人员性质 = "护士" Then bln护士 = True
                If rsTmp!人员性质 = "医生" Then bln护士 = False: Exit Do
                rsTmp.MoveNext
            Loop
        End If
    End If
    GetOperatorInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetDoctor(ByVal lng科室ID As Long, ByVal bln护士 As Boolean, ByRef rsTmp As ADODB.Recordset)
'功能：获取指定科室的医生
'参数：lng科室ID=指定科室ID,bln护士=是否也读取护士
    'Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '允许部门工作性质为非临床的医生或护士,因为可能该人所属的一个部门是非末级部门.
    If rsTmp Is Nothing Then
        strSQL = _
            "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
            " C.人员性质,Nvl(A.聘任技术职务,0) as 职务,B.缺省" & _
            " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
            " Where A.ID = B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID And C.人员性质 IN('医生','护士') " & _
            " And D.服务对象 IN(2,3) And D.工作性质 IN('临床','手术') And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by " & IIf(gbyt开单人显示 = 1, "简码", "编号") & ",缺省 Desc"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    End If
   
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

Public Sub GetDoctorDept(ByRef rs开单科室 As ADODB.Recordset, ByRef strPrivs As String, _
                        ByRef bytUseType As Byte, ByRef lngDeptID As Long)
'功能：获取所有开单科室
'参数：strPrivs-用于判断是否具有"门诊留观记帐"，"所有科室"的权限
'      bytUseType-'记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐
'      lngDeptID-当前病区ID或科室ID
    Dim strSQL As String
    
    On Error GoTo errH
    '可选开单科室(如果是医技科室,则包含门诊和住院的)
    If (InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "门诊留观记帐") > 0 And gbln门诊留观) Or bytUseType = 2 Then
        strSQL = "1,2,3"
    Else
        strSQL = "2,3"
    End If
    If bytUseType = 0 Or bytUseType = 1 Then
        strSQL = _
            "Select A.ID, A.编码, A.名称, A.简码, 0 As 缺省, B.工作性质, D.优先级" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B," & vbNewLine & _
            "     (Select 部门id, Max(Decode(服务对象, 2, 1, 2)) As 优先级 From 部门性质说明 Where 服务对象 <> 0 Group By 部门id) D" & vbNewLine & _
            "Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And A.ID = B.部门id" & vbNewLine & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And B.部门id = D.部门id And (B.服务对象 IN(" & strSQL & ") AND B.工作性质 IN('临床','手术') Or b.工作性质='产科')" & vbNewLine & _
            "Order By 优先级,编码"
    ElseIf bytUseType = 2 Then
        '医技科室记帐
        strSQL = _
            "Select A.ID, A.编码, A.名称, A.简码, 0 As 缺省, B.工作性质, D.优先级" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B," & vbNewLine & _
            "     (Select 部门id, Max(Decode(服务对象, 2, 1, 2)) As 优先级 From 部门性质说明 Where 服务对象 <> 0 Group By 部门id) D" & vbNewLine & _
            "Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And A.ID = B.部门id" & vbNewLine & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " And B.部门id = D.部门id And (B.服务对象 IN(" & strSQL & ") AND B.工作性质 IN('检查','检验','手术','治疗','营养') Or b.工作性质='产科')" & vbNewLine & _
            IIf(InStr(strPrivs, ";所有科室;") > 0, "", " And A.ID=[1] ") & vbNewLine & _
            "Order By 优先级,编码"
    End If
    Set rs开单科室 = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngDeptID)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Public Function GetLastAdviceTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Date
'功能：获取指定病人最后一条有效的医嘱的时间
'说明：用于病人出院时判断出院时间必须大于该时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetLastAdviceTime = CDate("1900-01-01")
    
    On Error GoTo errH
    
    '以长嘱最后执行时间为准判断,暂时排开持续性长嘱
    '临嘱有离院带药的情况,如以"出院"医嘱为准,出院时间本来就必须大于该变动时间。
    strSQL = "Select Max(Nvl(执行终止时间,Nvl(上次执行时间,开始执行时间))) as 时间" & _
        " From 病人医嘱记录" & _
        " Where Nvl(医嘱期效,0)=0 And 医嘱状态 Not IN(1,2,4)" & _
        " And Not (执行时间方案 is NULL And (Nvl(频率次数, 0) = 0 Or Nvl(频率间隔, 0) = 0 Or 间隔单位 is NULL))" & _
        " And 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetLastAdviceTime = rsTmp!时间
        End If
    End If
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
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "mdlInExse")
    If Not rsTmp.EOF Then
        Check药房上班安排 = Nvl(rsTmp!Num, 0) <> 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get收费执行科室ID(ByVal str类别 As String, ByVal lng项目id As Long, _
    ByVal int执行科室类型 As Integer, ByVal lng病人科室ID As Long, ByVal lng开单科室ID As Long, Optional ByVal int范围 As Integer = 2, _
    Optional ByVal lng发料部门ID As Long, Optional lng病人病区ID As Long, _
    Optional lng成套缺省执行科室 As Long = 0) As Long
    
    '功能：获取非药收费项目的执行科室
    '参数：int范围=1.门诊,2-住院
    '       lng发料部门ID=指定的缺省执行科室ID或病人病区ID(目前仅用于卫材)
    '       lng成套缺省执行科室-成套项目默认的执行科室:27327
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim str药房 As String, lng药房 As Long
    Dim bytDay As Byte, strIDs As String
    
    On Error GoTo errH
    If str类别 = "4" Then
        strSQL = _
        " Select Distinct" & _
        "   B.服务对象,C.编码,Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
        "       And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
        "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
        "       And (A.病人来源 is NULL Or A.病人来源=[1])" & _
        "       And ( A.开单科室ID is NULL Or A.开单科室ID=[2]   " & _
        "             Or Exists(select 1 From 病区科室对应 M where A.开单科室ID=M.病区ID And M.科室ID=[2]))" & _
        "       And A.收费细目ID=[3]" & _
        " Order by B.服务对象,C.编码"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", int范围, lng病人科室ID, lng项目id)
        If Not rsTmp.EOF Then
            Get收费执行科室ID = rsTmp!执行科室ID    '3:如果都没有，则返回第一个可用的执行科室(与医生站不同)
            '1:缺省为指定的(医嘱的)执行科室,不管是否服务于病人科室
            rsTmp.Filter = "执行科室ID=" & lng发料部门ID
            '2:其它可服务于病人科室的执行科室
            If rsTmp.EOF Then
                '2.0 如果成套中存在缺省的执行科室,则缺省为成套指定的缺省科室
                If lng成套缺省执行科室 <> 0 Then
                    rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                    If Not rsTmp.EOF Then
                            Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                    End If
                End If
                
                '2.1:尝试缺省为病人科室
                If lng发料部门ID <> lng病人科室ID Then
                    rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人科室ID
                End If
                '2.2:尝试缺省为病人病区
                If rsTmp.EOF Then
                    If lng病人病区ID <> 0 And lng病人病区ID <> lng病人科室ID And lng病人病区ID <> lng发料部门ID Then
                        rsTmp.Filter = "开单科室ID=" & lng病人科室ID & " And 执行科室ID=" & lng病人病区ID
                    End If
                End If
            End If
            '2.3:可服务于病人科室的一个执行科室
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
        End If
    ElseIf InStr(",5,6,7,", str类别) > 0 Then
        If gbln分离发药 Then Exit Function
        
        If str类别 = "5" Then
            str药房 = "西药房"
            lng药房 = glng西药房
        ElseIf str类别 = "6" Then
            str药房 = "成药房"
            lng药房 = glng成药房
        ElseIf str类别 = "7" Then
            str药房 = "中药房"
            lng药房 = glng中药房
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", str药房, int范围, lng病人科室ID, lng项目id, bytDay)
        If Not rsTmp.EOF Then
            strIDs = ""
            If lng成套缺省执行科室 <> 0 Then
                rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                If Not rsTmp.EOF Then
                        Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                End If
            End If
            rsTmp.Filter = "开单科室ID=" & lng病人科室ID
            If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!执行科室ID '收集用于动态分配
                If i = 1 Or rsTmp!执行科室ID = lng药房 Then
                    Get收费执行科室ID = rsTmp!执行科室ID
                    If rsTmp!执行科室ID = lng药房 Then
                        strIDs = "": Exit For
                    End If
                End If
                rsTmp.MoveNext
            Next
            strIDs = Mid(strIDs, 2)
            If UBound(Split(strIDs, ",")) <= 0 Then strIDs = ""
         End If
    Else
        Select Case int执行科室类型
            Case 0 '0-无明确科室
                '1 成套项目选择且存在缺省的执行科室的 成套项目的执行部门ID
                If lng成套缺省执行科室 <> 0 Then
                    Get收费执行科室ID = lng成套缺省执行科室: Exit Function
                End If
                '101736,手工记帐缺省执行科室
                '2 收费项目.缺省科室(手工记帐缺省执行科室)
                If int范围 = 2 Then
                    strSQL = "Select a.执行科室id" & vbNewLine & _
                            " From 收费执行科室 A, 部门表 C" & vbNewLine & _
                            " Where a.执行科室id + 0 = c.Id And a.收费细目id = [1]" & vbNewLine & _
                            "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
                            "       And (c.站点 = '" & gstrNodeNo & "' Or c.站点 Is Null)" & vbNewLine & _
                            "       And a.病人来源 = [2] And a.开单科室id Is Null"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng项目id, 2)
                    If Not rsTmp.EOF Then
                        If Val(Nvl(rsTmp!执行科室ID)) <> 0 Then
                            Get收费执行科室ID = Val(Nvl(rsTmp!执行科室ID)): Exit Function
                        End If
                    End If
                    '3 病人科室
                    If lng病人科室ID <> 0 Then Get收费执行科室ID = lng病人科室ID: Exit Function
                    '4 开单科室
                    If lng开单科室ID <> 0 Then Get收费执行科室ID = lng开单科室ID: Exit Function
                End If
                '5 操作员所属部门ID
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
                "   From 收费执行科室 A,部门表 C" & _
                "   Where A.执行科室ID+0=C.ID And  A.收费细目ID=[1]" & _
                "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
                "       And (A.病人来源 is NULL Or A.病人来源=[2])" & _
                "       And (A.开单科室ID is NULL Or A.开单科室ID=[3])" & _
                " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng项目id, int范围, lng病人科室ID)
                If Not rsTmp.EOF Then
                    If lng成套缺省执行科室 <> 0 Then
                         rsTmp.Filter = "执行科室ID=" & lng成套缺省执行科室
                         If Not rsTmp.EOF Then
                                 Get收费执行科室ID = rsTmp!执行科室ID: Exit Function
                         End If
                     End If
                    '缺省取操作员所在科室
                    rsTmp.Filter = "开单科室ID=" & UserInfo.部门ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=" & lng病人科室ID
                    If rsTmp.EOF Then rsTmp.Filter = "开单科室ID=0"
                    If Not rsTmp.EOF Then Get收费执行科室ID = rsTmp!执行科室ID
                End If
            Case 5 '院外执行(预留,程序暂未用)
            Case 6 '开单人科室
               Get收费执行科室ID = lng开单科室ID
        End Select
        If Get收费执行科室ID = 0 Then Get收费执行科室ID = UserInfo.部门ID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal int门诊标志 As Integer) As String
'功能：检查病人在医技科室是否还有未执行完成(未执行或正在执行)的项目
'返回：医技科室名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(2,[1],[2],-1,0,[3]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitExe", lng病人ID, lng主页ID, int门诊标志)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = Nvl(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lng检查离院带药 As Long = 0) As String
'功能：检查病人在药房是否还有未发药的药品或卫材
'返回：药房和发料部门名称
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],-1,[3]) as 内容 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng病人ID, lng主页ID, lng检查离院带药)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = Nvl(rsTmp!内容)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillisAdviceMoney(ByVal strNO As String, ByVal bytFlag As Byte, _
    lng医嘱ID As Long, lng发送号 As Long) As Boolean
'功能：判断一张单据是否医嘱的附加费用
'参数：int记录性质=对应住院费用记录.记录性质
'返回：医嘱ID,发送号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    lng医嘱ID = 0: lng发送号 = 0
    
    On Error GoTo errH
            
    strSQL = "Select 医嘱序号 From 住院费用记录" & _
        " Where Rownum=1 And 记录状态 IN(0,1,3) And NO=[1] And 记录性质=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag)
    If Not rsTmp.EOF Then lng医嘱ID = Nvl(rsTmp!医嘱序号, 0)
    If lng医嘱ID <> 0 Then
        strSQL = "Select 发送号 From 病人医嘱附费" & _
            " Where 医嘱ID=[3] And NO=[1] And 记录性质=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytFlag, lng医嘱ID)
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

Public Function PatiCanBilling(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String) As Boolean
    '功能：检查指定病人是否具有相关权限
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    On Error GoTo errH
    strSQL = "Select Nvl(b.姓名, a.姓名) As 姓名, b.出院日期, b.状态, Nvl(Sum(c.金额), 0) As 金额,b.病人性质 " & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人未结费用 C" & vbNewLine & _
            "Where a.病人id = [1] And a.病人id = b.病人id And b.主页id = [2] And b.病人id = c.病人id(+) And b.主页id = c.主页id(+)" & vbNewLine & _
            "Group By Nvl(b.姓名, a.姓名), b.出院日期, b.状态,b.病人性质"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        If Val(Nvl(rsTmp!病人性质)) = 1 And Not (gbln门诊留观 And InStr(strPrivs, ";门诊留观记帐;") > 0) Then
            strMsg = """" & rsTmp!姓名 & """为门诊留观病人，你没有权限对其进行记帐操作！"
        End If
        If Val(Nvl(rsTmp!病人性质)) = 2 And Not (gbln住院留观 And InStr(strPrivs, ";住院留观记帐;") > 0) Then
            strMsg = """" & rsTmp!姓名 & """为住院留观病人，你没有权限对其进行记帐操作！"
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
        
        If IsNull(rsTmp!出院日期) And Nvl(rsTmp!状态, 0) <> 3 Then PatiCanBilling = True: Exit Function
        If InStr(strPrivs, ";出院未结强制记帐;") = 0 Then
            If Nvl(rsTmp!金额, 0) <> 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用未结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        If InStr(strPrivs, ";出院结清强制记帐;") = 0 Then
            If Nvl(rsTmp!金额, 0) = 0 Then
                strMsg = """" & rsTmp!姓名 & """的费用已结清，当前已经出院(或预出院)。你不具有对该病人记帐的权限。"
            End If
        End If
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    PatiCanBilling = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxFact(ByVal strNO As String) As String
'功能：获取指定结帐单据发出的最大票据号
    Dim rsTmp As ADODB.Recordset
    Dim strSQL  As String
    
    On Error GoTo errH
    
    '应取最后一次打印的最大号码
    strSQL = "Select Max(ID) From 票据打印内容 Where 数据性质=3 And NO=[1]"
    strSQL = "Select Max(号码) as 号码 From 票据使用明细 Where 票种=3 And 性质=1 And 打印ID=(" & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If Not rsTmp.EOF Then GetMaxFact = Nvl(rsTmp!号码)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PatiHaveStorage(ByVal lng病人ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    strSQL = "Select A.结算方式,Sum(A.金额) as 金额" & _
        " From 病人预交记录 A,结算方式 B" & _
        " Where A.记录性质=1 And A.结算方式=B.名称 And B.性质=5 And A.病人ID=[1]" & _
        " Group by A.结算方式 Having Sum(A.金额)<>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            strMsg = strMsg & vbCrLf & rsTmp!结算方式 & "：" & Format(rsTmp!金额, "0.00")
            rsTmp.MoveNext
        Loop
    End If
    If strMsg <> "" Then
        If gbyt结帐检查代收款项 = 1 Then
            If MsgBox("还有以下代收费用没有退还病人：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "要继续结帐吗?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                PatiHaveStorage = True
            Else
                PatiHaveStorage = False
            End If
        Else
            MsgBox "还有以下代收费用没有退还病人：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "请先将费用退还给病人再结帐。", vbInformation, gstrSysName
            PatiHaveStorage = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillIdentical(ByVal strNO As String) As Boolean
'功能：判断指定的记帐单据中的状态是否一致,即是否同时存在审核和未审核的内容
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    BillIdentical = True
    
    On Error GoTo errH

    strSQL = _
        " Select Count(Distinct 登记时间) as 时间数," & _
        " Sum(Decode(记录状态,0,1,0)) as 未审核," & _
        " Sum(Decode(记录状态,0,0,1)) as 已审核" & _
        " From 住院费用记录" & _
        " Where 记录状态 IN(0,1,3) And NO=[1] And 记录性质=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
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

Public Function AuditingWarn(ByVal strPrivs As String, ByRef rsWarn As ADODB.Recordset, _
                            ByVal strNO As String, ByVal str序号 As String) As Boolean
'功能：审核划价单时，对费用进行报警
'参数：str序号=指定单据中要审核的行号,为空表示所有行
    Dim rsTmp As ADODB.Recordset
    Dim rsFee As ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, lng病人ID As Long
    Dim str病人IDs As String, str类别s As String
    Dim cur当日额 As Currency, cur金额 As Currency, cur余额 As Currency
    Dim strWarn As String, intWarn As Integer, bln医保 As Boolean
            
    strSQL = "Select A.门诊标志, A.姓名, A.病人id, A.主页id, A.病人病区id 病区id," & vbNewLine & _
            "       Decode(B.担保额, Null, B.担保额, Zl_Patientsurety(A.病人id, A.主页id)) 担保额," & vbNewLine & _
            "       Zl_Patiwarnscheme(A.病人id, A.主页id) As 适用病人, C.是否医保 As 付款码, A.收费类别, D.名称 As 类别名称," & vbNewLine & _
            "       Sum(A.实收金额) As 金额" & vbNewLine & _
            "From 住院费用记录 A, 病人信息 B, 医疗付款方式 C, 收费项目类别 D" & vbNewLine & _
            "Where A.记录性质 = 2 And A.记录状态 = 0 And A.NO = [1] And A.收费类别 = D.编码 And A.病人id = B.病人id And" & vbNewLine & _
            "      B.医疗付款方式 = C.名称(+)" & vbNewLine & _
            IIf(str序号 <> "", " And Instr([2],','||Nvl(A.价格父号,A.序号)||',')>0", "") & _
            "Group By Nvl(A.价格父号, A.序号), A.门诊标志, A.姓名, A.病人id, A.主页id, A.病人病区id, B.担保额, C.是否医保, A.收费类别," & vbNewLine & _
            "         D.名称" & vbNewLine & _
            "Order By A.病人id"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, "," & str序号 & ",")
    For i = 1 To rsTmp.RecordCount
        If InStr(str病人IDs & ",", "," & rsTmp!病人ID & ",") = 0 Then
            str病人IDs = str病人IDs & "," & rsTmp!病人ID
        End If
        rsTmp.MoveNext
    Next
            
    If str病人IDs <> "" Then
        str病人IDs = Mid(str病人IDs, 2)
        For i = 0 To UBound(Split(str病人IDs, ","))
            lng病人ID = Val(Split(str病人IDs, ",")(i))
            rsTmp.Filter = "病人ID=" & lng病人ID
            
            '取报警类别和金额
            str类别s = "": cur金额 = 0
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
                bln医保 = ("" & rsTmp!付款码 = "1")
                            
                cur当日额 = GetPatiDayMoney(lng病人ID)
                Set rsFee = GetMoneyInfo(lng病人ID, 0, bln医保, 2)
                If Not rsFee Is Nothing Then cur余额 = rsFee!预交余额 - rsFee!费用余额
                
                If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(1, lng病人ID) + cur金额
                
                '分类报警
                For j = 0 To UBound(Split(str类别s, ","))
                    intWarn = BillingWarn(strPrivs, rsTmp!姓名, Val("" & rsTmp!病区ID), rsTmp!适用病人, rsWarn, _
                        cur余额, cur当日额, cur金额, Nvl(rsTmp!担保额, 0), _
                        Left(Split(str类别s, ",")(j), 1), Mid(Split(str类别s, ",")(j), 2), strWarn)
                    If intWarn = 2 Or intWarn = 3 Then Exit Function
                Next
            End If
        Next
    End If
    AuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckExistsGathering(strNO As String) As Boolean
'功能:判断指定结帐单是否存在应收款的缴款记录
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select A.NO From 病人结帐记录 A, 病人缴款对照 B Where A.NO = [1] And A.ID = B.结帐id And Rownum = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    CheckExistsGathering = rsTmp.RecordCount > 0
    
Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatientDue(lng病人ID As Long) As Currency
'功能:获取指定病人的应收款余额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Patientdue([1]) Due From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If rsTmp.RecordCount > 0 Then GetPatientDue = rsTmp!Due
    
Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub LoadPatientBaby(ByRef cboBaby As ComboBox, ByVal lngPatient As Long, lngPatientPage As Long)
    Dim rsTmp As ADODB.Recordset, i As Long
    
    cboBaby.Clear
    cboBaby.AddItem "0-病人本人"
    cboBaby.ItemData(cboBaby.NewIndex) = 0
    Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
    
    If lngPatient <> 0 Then
        Set rsTmp = GetPatientBaby(lngPatient, lngPatientPage)
        With rsTmp
            For i = 1 To .RecordCount
                If Not IsNull(!婴儿姓名) Then
                    cboBaby.AddItem !序号 & "-" & !婴儿姓名
                Else
                    cboBaby.AddItem !序号 & "-第" & !序号 & "个婴儿"
                End If
                cboBaby.ItemData(cboBaby.NewIndex) = !序号
                .MoveNext
            Next
        End With
    End If
End Sub

Public Function GetPatientBaby(ByVal lngPatient As Long, lngPatientPage As Long) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 序号, 婴儿姓名 From 病人新生儿记录 Where 病人id = [1] And 主页ID = [2]"
    On Error GoTo errH
    Set GetPatientBaby = zlDatabase.OpenSQLRecord(strSQL, "读取新生儿记录", lngPatient, lngPatientPage)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValidity(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal dbl数量 As Double, Optional ByVal blnAsk As Boolean = True) As Boolean
'功能：检查卫生材料的灭菌效期是否过期
'说明：blnAsk=表示是否询问是否继续,否则为提醒
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, MinDate As Date
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng材料ID, lng库房ID)
    If Not rsTmp.EOF Then
        strName = rsTmp!名称
        Curdate = rsTmp!时间
        MinDate = CDate("3000-01-01")
            
        Do While Not rsTmp.EOF
            If rsTmp!灭菌效期 < MinDate Then
                MinDate = rsTmp!灭菌效期
            End If
            If Nvl(rsTmp!库存, 0) < dbl数量 Then
                dbl数量 = dbl数量 - Nvl(rsTmp!库存, 0)
            Else
                dbl数量 = 0
            End If
            If dbl数量 = 0 Then Exit Do
            rsTmp.MoveNext
        Loop

        If Curdate > MinDate Then
            If blnAsk Then
                If MsgBox("卫生材料""" & strName & """的灭菌效期""" & Format(MinDate, "yyyy-MM-dd") & """已过期,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckValidity = False
                End If
            Else
                MsgBox "提醒：" & vbCrLf & vbCrLf & "卫生材料""" & strName & """的灭菌效期""" & Format(MinDate, "yyyy-MM-dd") & """已过期。", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckRecalcRecord(ByVal strNO As String, Optional ByVal byt费用来源 As Byte) As Boolean
'功能：判断指定病人的指定单据是否存在按费别重算的冲减记录(数次为0的记录)
'入参：
'   byt费用来源 0-住院,1-门诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Count(A.ID) Num" & vbNewLine & _
            "From 住院费用记录 A," & vbNewLine & _
            "     (Select 病人id, 主页id, 病人病区id, 病人科室id, 收费细目id, 收入项目id, 开单部门id, 执行部门id, 发生时间" & vbNewLine & _
            "       From 住院费用记录" & vbNewLine & _
            "       Where NO = [1] And 记帐费用 = 1" & vbNewLine & _
            "       Group By 病人id, 主页id, 病人病区id, 病人科室id, 收费细目id, 收入项目id, 开单部门id, 执行部门id, 发生时间) B" & vbNewLine & _
            "Where A.记录性质 = 2 And A.数次 = 0 And A.病人id = B.病人id And A.主页id = B.主页id And" & vbNewLine & _
            "      A.病人病区id + 0 = B.病人病区id And A.病人科室id + 0 = B.病人科室id And A.收费细目id + 0 = B.收费细目id And" & vbNewLine & _
            "      A.收入项目id + 0 = B.收入项目id And A.开单部门id + 0 = B.开单部门id And A.执行部门id + 0 = B.执行部门id And" & vbNewLine & _
            "      A.发生时间 = B.发生时间"
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    If rsTmp.RecordCount > 0 Then CheckRecalcRecord = rsTmp!Num > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckNONegative(ByVal strNO As String, Optional ByVal bytType As Byte = 2, _
    Optional ByVal byt费用来源 As Byte) As Boolean
'功能：判断指定单据是否包含负数明细
'入参：
'   byt费用来源 0-住院,1-门诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 1 From 住院费用记录 Where NO = [1] And 记录性质 = [2] And 记录状态 = 1 And 数次 < 0"
    If byt费用来源 = 1 Then strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO, bytType)
    If rsTmp.RecordCount > 0 Then CheckNONegative = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function inBlackList(ByVal lng病人ID As Long) As String
'功能：判断病人是否在黑名单中,并反回加入原因
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    '加入编号:47663
    strSQL = "Select 编号, 加入原因 From 特殊病人 Where 撤消时间 is Null And 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID)
    If Not rsTmp.EOF Then inBlackList = rsTmp!编号 & "-" & rsTmp!加入原因
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ReadABCNum(ByVal strPrivs As String) As Boolean
'功能:获取中药输入快捷
'参数：strPrivs=用于根据权限控制是否读取小数快捷部份
'返回：可以输入的快捷字母
    Dim strSQL As String
        
    On Error GoTo errH
    
    If InStr(strPrivs, ";药品输入小数;") > 0 Then
        strSQL = "Select Upper(名称) as 名称,数值 From 中药输入快捷 Order by 名称"
    Else
        strSQL = "Select Upper(名称) as 名称,数值 From 中药输入快捷 Where Trunc(数值)=数值 Order by 名称"
    End If
    
    Set grsABCNum = New ADODB.Recordset 'Filter在New时清除
    Call zlDatabase.OpenRecordset(grsABCNum, strSQL, "mdlInExse")
    
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


Public Function GetExamineItem(ByVal strItems As String, ByVal lngMediCareID As Long) As ADODB.Recordset
'功能:返回指定险类的收费项目要求审批的记录集
'参数:strItems-收费细目ID串,例如:"2369,2367,2368"
'     lngMediCareID-险类,例如:901
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select A.收费细目id" & vbNewLine & _
            "From 保险支付项目 A ,Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) B" & vbNewLine & _
            "Where A.险类 = [1] And A.要求审批 = 1 And A.收费细目id = B.Column_Value"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngMediCareID, strItems)
    
    Set GetExamineItem = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRowByFeeItemID(ByRef ObjBillDetails As BillDetails, ByRef lngItemID As Long, Optional lngPatientID As Long) As Long
'功能:根据收费项目ID返回其在单据中的行号,如果有重复的,只返回第一个
    Dim i As Long
    
    For i = 1 To ObjBillDetails.Count
        If lngItemID = ObjBillDetails(i).收费细目ID And (ObjBillDetails(i).病人ID = lngPatientID Or lngPatientID = 0) Then
            GetRowByFeeItemID = i: Exit Function
        End If
    Next
End Function

Public Function CheckExamine(ByRef ObjBillDetails As BillDetails, _
    ByRef rsMedAudit As ADODB.Recordset, _
    ByRef lngMediCareID As Long, _
    Optional ByVal str姓名 As String = "") As Boolean
    '功能:根据给定的收费项目对象集和病人审批项目记录集检查相应的收费项目是否需要审批
    '入参:str姓名-为空时,提示为当前病人,不为空时,按传入提示病人
    Dim i As Long, j As Long, strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    For i = 1 To ObjBillDetails.Count
        strTmp = strTmp & "," & ObjBillDetails(i).收费细目ID
    Next
    Set rsTmp = GetExamineItem(Mid(strTmp, 2), lngMediCareID)
    
    strTmp = ""
    For i = 1 To rsTmp.RecordCount
        rsMedAudit.Filter = "项目ID=" & rsTmp!收费细目ID
        If rsMedAudit.RecordCount = 0 Then
            strTmp = strTmp & "," & GetRowByFeeItemID(ObjBillDetails, rsTmp!收费细目ID)
        ElseIf Not IsNull(rsMedAudit!可用数量) Then
            j = GetRowByFeeItemID(ObjBillDetails, rsTmp!收费细目ID)
            If ObjBillDetails(j).付数 * ObjBillDetails(j).数次 * IIf(gbln住院单位, ObjBillDetails(j).Detail.住院包装, 1) > rsMedAudit!可用数量 Then
                MsgBox IIf(str姓名 <> "", "病人:" & str姓名 & "在", "") & "第" & j & "行收费项目的数次超过了批准的使用限量" & FormatEx(rsMedAudit!可用数量 / IIf(gbln住院单位, ObjBillDetails(j).Detail.住院包装, 1), 5) & "。", vbInformation, gstrSysName
                CheckExamine = False: Exit Function
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If strTmp <> "" Then
        If str姓名 <> "" Then
            MsgBox "病人:" & str姓名 & "在第" & Mid(strTmp, 2) & "行收费项目要求审批,但未被批准使用!", vbInformation, gstrSysName
        Else
            MsgBox "第" & Mid(strTmp, 2) & "行收费项目要求审批,当前病人未被批准使用!", vbInformation, gstrSysName
        End If
        CheckExamine = False: Exit Function
    End If
    CheckExamine = True
End Function

Public Function CheckFeeItemLimitDept(ByVal lngFeeItem As Long, ByVal lngPatientUnit As Long, ByVal lngPatientDept As Long) As Boolean
'功能:检查收费项目,如果是主项,是否适用于当前病人科室或病区
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select 科室id From 收费适用科室 Where 项目id = [1] And (Select Count(从项id) From 收费从属项目 Where 主项id = [1]) > 0"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItem)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!科室ID = lngPatientUnit Or rsTmp!科室ID = lngPatientDept Then
                CheckFeeItemLimitDept = True
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        CheckFeeItemLimitDept = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillBeforIN(ByVal strNO As String) As Boolean
'功能：检查结帐单是否发生在本次住院之前
    Dim rsTmp As ADODB.Recordset, strSQL As String
     
    strSQL = "Select 1" & vbNewLine & _
        "From 病人结帐记录 A, 病人信息 B" & vbNewLine & _
        "Where A.NO = [1] And A.记录状态 = 1 And A.病人id = B.病人id And B.入院时间 > A.收费时间 And B.出院时间 Is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", strNO)
    CheckBillBeforIN = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckBillExistReplenishData(intTYPE As Integer, Optional lngBalance As Long, Optional strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在二次结算
    '返回:True-存在二次结算数据 False-不存在二次结算数据
    '入参:intType:0-收费数据，使用lngBalance为结算序号
    '     intType:1-收费数据，使用strNos为单据号
    '编制:刘尔旋
    '日期:2014-10-15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If intTYPE = 0 Then
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From 费用补充记录 A, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
        " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
        strSQL = strSQL & " Union " & _
        " Select 1 From 费用补充记录 Where 结算序号 = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", lngBalance)
    Else
        strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From 费用补充记录 A," & vbNewLine & _
        "      (Select Distinct 结帐id" & vbNewLine & _
        "       From 门诊费用记录" & vbNewLine & _
        "       Where Mod(记录性质, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
        " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查二次结算", strNos)
    End If
    If rsTmp.EOF Then
        CheckBillExistReplenishData = False
    Else
        CheckBillExistReplenishData = True
    End If
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
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
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


Public Function zlCreateFeeListStruc(ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建本地的费用记录集结构
    '入参:
    '出参:rsFeelists-返回本地记录集结构,同时打开了记录集的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 16:18:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set rsFeelists = New ADODB.Recordset
    rsFeelists.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsFeelists.Fields.Append "实际票号", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "结算时间", adDBTimeStamp, , adFldIsNullable
    rsFeelists.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsFeelists.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    rsFeelists.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsFeelists.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "数量", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "单价", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "保险支付大类ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "是否医保", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "保险编码", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "摘要", adVarChar, 4000, adFldIsNullable
    rsFeelists.Fields.Append "是否急诊", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "本次结算", adDouble, , adFldIsNullable
    rsFeelists.CursorLocation = adUseClient
    rsFeelists.LockType = adLockOptimistic
    rsFeelists.CursorType = adOpenStatic
    rsFeelists.Open
    zlCreateFeeListStruc = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

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
    Err = 0: On Error GoTo ErrHand:
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
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function



Public Function ImportWholeSet(frmParent As Object, ByVal intInsure As Integer, rsSel As ADODB.Recordset, Optional lng病人ID As Long = 0, _
     Optional bln住院单位 As Boolean, Optional lng开单部门ID As Long = 0, Optional byt婴儿费 As Byte, _
     Optional int门诊标志 As Integer, Optional bln加班加价 As Boolean = False, _
     Optional ByVal lngUnitID As Long, Optional int范围 As Integer, _
     Optional str划价人 As String = "", Optional str开单人 As String = "", _
     Optional lngPatiNums As Long = 1, Optional blnNurseStation As Boolean = False, _
    Optional ByVal str药品价格等级 As String, _
    Optional ByVal str卫材价格等级 As String, Optional ByVal str普通价格等级 As String, _
    Optional ByVal lng主页ID As Long, Optional ByVal lng科室ID As Long, Optional ByVal lng病区ID As Long) As ExpenseBill
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取费用单据到单据对象中
    '入参:rsSel-选中的成套项目
    '       lngUnitID    当前操作病区ID
    '      int范围=1.门诊,2-住院
    '      lngPatiNums-病人数(批量记帐有效)
    '出参:
    '返回:存放单据信息的单据对象
    '编制:刘兴洪
    '日期:2010-09-02 16:17:54
    '说明:因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    '       不包含已停用收费细目
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue(0 To 10) As String, strSubItem As String, str收费细目ID As String, j As Long
    Dim rsItems As ADODB.Recordset, rsOthers As ADODB.Recordset
    Dim lng病人科室ID As Long, str摘要 As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim rsTmp As ADODB.Recordset, rsPrice As ADODB.Recordset
    Dim rsMoney As ADODB.Recordset
    Dim lngDoUnit As Long
    Dim i As Long, intCurNo As Integer
    Dim int序号 As Integer, blnDo As Boolean, dblStock As Double
    Dim blnLoad As Boolean, strSQL As String, str药房IDs As String, str停用项目序号 As String, strPrivs As String
    Dim curModiMoney As Currency
    
    Dim dblAllTime As Double, dblCurTime As Double, dbl加班加价率 As Double, dblPriceSingle As Double, lngLastPati As Long
    Dim colSerial As New Collection, dblPrice As Double
    Dim bytType As Byte '0-门诊;1-住院;2-门诊或住院
    Dim strTable  As String, strWherePriceGrade As String
    
    On Error GoTo errH
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
       "   Select A.主项id, A.从项id, A.固有从属, A.从项数次 " & _
       "   From 收费从属项目 A, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.主项id = D.收费细目id "
    Set rsOthers = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInExse", strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    strSubItem = Mid(strSubItem, 11)
    strTable = " Select [13] as 病人ID,收费细目ID From (" & strSubItem & ")"
    
    gstrSQL = "" & _
    " Select  X.药品ID,W.材料ID,W.跟踪在用," & _
    "       G.费别,F.姓名,F.性别,F.年龄,F.担保额," & _
    "       G.出院病床 as 床号,F.住院号 as 标识号,F.病人ID,G.主页ID,G.当前病区ID as 病人病区ID,G.出院科室ID as 病人科室ID," & _
    "       G.病人性质,B.类别 as 收费类别,A.收费细目ID," & _
    "       B.计算单位,B.类别,C.名称 as 类别名称,B.编码,Nvl(H.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格,Nvl(B.是否变价,0) as 是否变价,B.加班加价," & _
    "       B.屏蔽费别,B.说明,B.执行科室,B.服务对象, B.费用类型  费用类型,D.现价,D.原价,D.缺省价格,D.收入项目ID as 现收入ID,E.名称 as 收入项目," & _
    "       E.收据费目 as 现费目,D.加班加价率,D.附术收费率,Nvl(W.诊疗ID,X.药名ID) as 药名ID," & _
    "       Decode(B.类别,'4',1,X.住院包装) as 住院包装,Decode(B.类别,'4',B.计算单位,X.住院单位) as 住院单位," & _
    "       Decode(b.类别,'4',Nvl(W.在用分批,0),Nvl(X.药房分批,0)) as 分批,B.录入限量, " & _
    "       M1.编码 as 诊疗编码,M1.名称 as 诊疗名称,X.中药形态,x.剂量系数,M1.计算单位 as 剂量单位" & _
    "   From  (" & strTable & ") A ,收费项目目录 B,收费项目类别 C,收费价目 D,收入项目 E,病人信息 F, " & _
    "          病案主页 G,收费项目别名 H,收费项目别名 E1,材料特性 W,药品规格 X,诊疗项目目录 M1" & _
    " Where  A.收费细目ID=D.收费细目ID And A.收费细目ID=B.ID " & _
    "       And b.类别=C.编码 And A.收费细目ID=X.药品ID(+) and X.药名ID=M1.ID(+) And A.收费细目ID=W.材料ID(+) And D.收入项目ID=E.ID" & _
    "       And A.收费细目ID=H.收费细目ID(+) And H.码类(+)=1 And H.性质(+)=[12]" & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3" & _
    "       And A.病人ID=F.病人ID(+) And F.病人ID=G.病人ID(+) And " & IIf(lng主页ID <> 0, " G.主页ID(+) = [17]", " F.主页ID=G.主页ID(+) ") & _
    "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) " & vbNewLine & _
    "       And Sysdate Between D.执行日期 And Nvl(D.终止日期,To_Date('3000-01-01','YYYY-MM-DD')) " & _
            strWherePriceGrade
    
    If Not gbln分离发药 Then
        gstrSQL = "Select * From (" & gstrSQL & ")"
    Else
        '分离发药时排开时价和分批药品或卫材
        gstrSQL = "Select * From (" & gstrSQL & ") Where Not( Instr(',5,6,7,',收费类别)>0 And (分批=1 Or 是否变价=1))"
    End If
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, "mdlExse", strValue(0), strValue(1), strValue(2), strValue(3), _
        strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10), _
        IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1), lng病人ID, str药品价格等级, str卫材价格等级, str普通价格等级, lng主页ID)
    '没有记录就是空单子
    Set objBill = New ExpenseBill
    Set objBill.Details = New BillDetails
    
    With rsSel
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
NextRecord: Do While Not .EOF
            '检查收费项目是否停用或服务于门诊病人
            '主项停用时,不导从项
            rsItems.Filter = "收费细目ID=" & Val(Nvl(!收费细目ID))
            If rsItems.EOF Then '未找到.不加入
                 .MoveNext
                GoTo NextRecord:
            End If
            If InStr(",5,6,7,", rsItems!收费类别) = 0 Then
                If InStr(1, str停用项目序号 & ",", "," & !从属父号 & ",") > 0 Then
                    .MoveNext
                    GoTo NextRecord
                Else
                    If Not CheckFeeItemAvailable(!收费细目ID, 2) Then
                        str停用项目序号 = str停用项目序号 & "," & !序号
                        MsgBox "成套收费项目中的第" & !序号 & "行收费项目:" & rsItems!名称 & "" & vbCrLf & _
                            "已停用或不再服务于病人,将不会被导入." & IIf(IsNull(!从属父号), "如果有从属项目,也不会被导入.", ""), vbInformation, gstrSysName
                        .MoveNext
                        GoTo NextRecord
                    End If
                End If
            Else
                If blnNurseStation Then
                    MsgBox "成套收费项目中的第" & !序号 & "行收费项目:" & rsItems!名称 & "" & vbCrLf & _
                        "为药品项目,护士站批量记帐时将不会被导入.", vbInformation, gstrSysName
                    .MoveNext
                    GoTo NextRecord
                End If
            End If
            
            If i = 1 Then
                objBill.NO = ""
                objBill.病人ID = Val(Nvl(rsItems!病人ID))
                objBill.主页ID = Val(Nvl(rsItems!主页ID))
                objBill.病区ID = IIf(lng病区ID = 0, Val(Nvl(rsItems!病人病区ID)), lng病区ID)
                objBill.科室ID = IIf(lng科室ID = 0, Val(Nvl(rsItems!病人科室id)), lng科室ID)
                objBill.姓名 = Nvl(rsItems!姓名)
                objBill.性别 = Nvl(rsItems!性别)
                objBill.年龄 = Nvl(rsItems!年龄)
                objBill.标识号 = Val(Nvl(rsItems!标识号))
                objBill.床号 = "" & rsItems!床号
                objBill.费别 = Nvl(rsItems!费别)
                objBill.门诊标志 = int门诊标志
                objBill.加班标志 = IIf(bln加班加价, 1, 0)
                objBill.婴儿费 = byt婴儿费
                objBill.开单部门ID = lng开单部门ID
                objBill.划价人 = str划价人
                objBill.开单人 = str开单人
                objBill.操作员编号 = UserInfo.编号
                objBill.操作员姓名 = UserInfo.姓名
                objBill.发生时间 = zlDatabase.Currentdate   ' !发生时间
                objBill.登记时间 = zlDatabase.Currentdate
                objBill.多病人单 = 0
                
            End If
            '处理收费细目=====================================================
            Set objBillDetail = New BillDetail
            Set objBillDetail.Detail = New Detail
                
            '处理序号和从属父号
            intCurNo = intCurNo + 1
            objBillDetail.序号 = intCurNo
            colSerial.Add Array(Val(Nvl(!收费细目ID)), intCurNo), "_" & !序号  '记录原序号现在的行号
            objBillDetail.从属父号 = Nvl(!从属父号, 0) '因为可能排序乱了,先记录原来的,后面再处理
            
            objBillDetail.病人ID = Val(Nvl(rsItems!病人ID))
            objBillDetail.主页ID = Val(Nvl(rsItems!主页ID))
            objBillDetail.婴儿费 = 0
            objBillDetail.病区ID = Val(Nvl(rsItems!病人病区ID))
            objBillDetail.科室ID = Val(Nvl(rsItems!病人科室id))
            objBillDetail.姓名 = Nvl(rsItems!姓名)
            objBillDetail.性别 = Nvl(rsItems!性别)
            objBillDetail.年龄 = Nvl(rsItems!年龄)
            objBillDetail.住院号 = Val(Nvl(rsItems!标识号))
            objBillDetail.床号 = "" & rsItems!床号
            objBillDetail.费别 = Nvl(rsItems!费别)
            objBillDetail.担保额 = Val(Nvl(rsItems!担保额))
            
            '目前仅用于记帐表
            objBillDetail.医疗付款 = Get病人医疗付款方式(objBillDetail.病人ID, objBillDetail.主页ID)
            
            objBillDetail.收费类别 = Nvl(rsItems!收费类别)
            objBillDetail.收费细目ID = Val(Nvl(!收费细目ID))
            objBillDetail.计算单位 = Nvl(rsItems!计算单位)
            objBillDetail.付数 = IIf(Val(Nvl(!付数)) = 0, 1, Val(Nvl(!付数)))
            If InStr(",5,6,7,", rsItems!收费类别) > 0 And bln住院单位 Then
                objBillDetail.数次 = Nvl(!数量, 0) / Nvl(rsItems!住院包装, 1)
            Else
                objBillDetail.数次 = Nvl(!数量, 0)
            End If
            
            objBillDetail.原始数量 = objBillDetail.付数 * objBillDetail.数次
            objBillDetail.发药窗口 = ""
            
            objBillDetail.附加标志 = 0 ' IIf(IsNull(!附加标志), 0, !附加标志)
            '卫材和药品部分
            '卫材执行科室缺省为病人病区,如果本地指定了,则为指定科室
            If objBillDetail.收费类别 = "4" Then
                lngDoUnit = IIf(glng发料部门 > 0, glng发料部门, objBillDetail.病区ID)
                If lngDoUnit = 0 Then lngDoUnit = lng开单部门ID
            End If
            
            '病人科室ID
            lng病人科室ID = objBillDetail.科室ID
            If lng病人科室ID = 0 Then lng病人科室ID = lng开单部门ID
            objBillDetail.Detail.执行科室 = IIf(IsNull(rsItems!执行科室), 0, rsItems!执行科室)
            objBillDetail.执行部门ID = Val(Nvl(!执行科室ID))
            lngDoUnit = Get收费执行科室ID(objBillDetail.收费类别, objBillDetail.收费细目ID, _
                 objBillDetail.Detail.执行科室, lng病人科室ID, lng开单部门ID, int范围, lngDoUnit, objBillDetail.病区ID, objBillDetail.执行部门ID)
            
            objBillDetail.执行部门ID = lngDoUnit

            If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln分离发药 Then
                objBillDetail.执行部门ID = 0
            End If
            objBillDetail.原始执行部门ID = objBillDetail.执行部门ID
            
            objBillDetail.Detail.ID = !收费细目ID
            objBillDetail.Detail.编码 = Nvl(rsItems!编码)
            objBillDetail.Detail.变价 = (Val(Nvl(rsItems!是否变价)) = 1)
            objBillDetail.Detail.从项数次 = 0
            objBillDetail.Detail.固有从属 = 0
            
            If Not gbln分离发药 And InStr(",4,5,6,7,", rsItems!收费类别) > 0 Then
                dblStock = GetStock(Val(Nvl(!收费细目ID)), objBillDetail.执行部门ID)
            Else
                dblStock = 0
            End If
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 And gbln分离发药 Then
                str药房IDs = Decode(rsItems!收费类别, "5", gstr西药房, "6", gstr成药房, "7", gstr中药房)
                If str药房IDs <> "" Then dblStock = GetMultiStock(!收费细目ID, str药房IDs)
            End If
            If InStr(",5,6,7,", rsItems!收费类别) > 0 And bln住院单位 Then dblStock = dblStock / Nvl(rsItems!住院包装, 1)
            objBillDetail.Detail.库存 = dblStock
            
            
            If objBillDetail.从属父号 <> 0 Then
                'A.主项id, A.从项id, A.固有从属, A.从项数次 "
                rsOthers.Filter = "主项ID=" & colSerial("_" & !从属父号)(0) & " And 从项ID=" & objBillDetail.收费细目ID
                If Not rsOthers.EOF Then
                    objBillDetail.Detail.从项数次 = Val(Nvl(rsOthers!从项数次))
                    objBillDetail.Detail.固有从属 = Val(Nvl(rsOthers!固有从属))
                End If
            End If
            
            objBillDetail.Detail.规格 = Nvl(rsItems!规格)
            objBillDetail.Detail.计算单位 = Nvl(rsItems!计算单位)
            
            objBillDetail.Detail.住院单位 = Nvl(rsItems!住院单位)
            objBillDetail.Detail.住院包装 = Val(Nvl(rsItems!住院包装))
            
            objBillDetail.Detail.加班加价 = 0 ' (IIf(IsNull(!加班加价), 0, !加班加价) = 1)
            objBillDetail.Detail.类别 = Nvl(rsItems!类别)
            objBillDetail.Detail.类别名称 = Nvl(rsItems!类别名称)
            objBillDetail.Detail.名称 = Nvl(rsItems!名称)
            objBillDetail.Detail.商品名 = Nvl(rsItems!商品名)
            objBillDetail.Detail.屏蔽费别 = (Val(Nvl(rsItems!屏蔽费别)) = 1)
            objBillDetail.Detail.说明 = ""
            objBillDetail.Detail.服务对象 = IIf(IsNull(rsItems!服务对象), 0, rsItems!服务对象)
            objBillDetail.Detail.类型 = IIf(IsNull(rsItems!费用类型), "", rsItems!费用类型)
            objBillDetail.Detail.诊疗名称 = Nvl(rsItems!诊疗名称)
            
            If InStr(",5,6,7,", rsItems!收费类别) > 0 Then
                objBillDetail.Detail.处方职务 = Get处方职务(objBillDetail.Detail.ID)
                objBillDetail.Detail.处方限量 = Get处方限量(objBillDetail.Detail.ID)
            End If
            objBillDetail.Detail.录入限量 = Val(Nvl(rsItems!录入限量))
            objBillDetail.Detail.药名ID = Val(Nvl(rsItems!药名ID))
            objBillDetail.Detail.变价 = Val(Nvl(rsItems!是否变价)) = 1
            objBillDetail.Detail.分批 = Val(Nvl(rsItems!分批)) = 1
            objBillDetail.Detail.跟踪在用 = Val(Nvl(rsItems!跟踪在用)) = 1
            objBillDetail.Detail.要求审批 = 0
            objBillDetail.Detail.中药形态 = Val(Nvl(rsItems!中药形态))
            objBillDetail.Detail.剂量单位 = Nvl(rsItems!剂量单位)
            objBillDetail.Detail.剂量系数 = Val(Nvl(rsItems!剂量系数))
         
            '问题:41136
            str摘要 = objBillDetail.摘要
'            If lng病人ID <> 0 Then '90304
                str摘要 = gclsInsure.GetItemInfo(intInsure, lng病人ID, objBillDetail.收费细目ID, str摘要, 2, , "|1")
                objBillDetail.摘要 = str摘要
'            Else
'                objBillDetail.摘要 = ""
'            End If
             '处理价格部份=====================================================
             rsItems.MoveFirst
            Set objBillDetail.InComes = New BillInComes
            Do While Not rsItems.EOF
                '按照现有的价格设置重新计算'***
                If Val(Nvl(rsItems!是否变价)) = 1 Then
                    If InStr(",5,6,7,", rsItems!收费类别) > 0 Or (rsItems!收费类别 = "4" And Nvl(rsItems!跟踪在用, 0) = 1) Then
                        '----------------------------------------------------------------------------------------------
                        '时价药品计算价格(分批可不分批)
                        dblAllTime = Val(Nvl(!数量)) * IIf(Val(Nvl(!付数)) = 0, 1, Val(Nvl(!付数))) * lngPatiNums
                        If dblAllTime <> 0 Then
                            dblPrice = Get时价药品应收金额(objBillDetail.执行部门ID, CLng(Nvl(!收费细目ID)), dblAllTime, gstrDec, dblPriceSingle)
                            If dblAllTime <> 0 Then
                                If Val(Nvl(!单价)) = 0 Then
                                    '数量未分解完毕
                                    If rsItems!收费类别 = "4" Then
                                        MsgBox "时价卫生材料""" & Nvl(rsItems!名称) & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    Else
                                        MsgBox "时价药品""" & Nvl(rsItems!名称) & """库存不足,无法计算价格！", vbInformation, gstrSysName
                                    End If
                                    objBillIncome.标准单价 = 0
                                Else
                                    objBillIncome.标准单价 = Val(Nvl(!单价))
                                End If
                            Else
                                '注意：货币型最多只能保留4位小数,且不四舍五入,所以需要手工舍入;而用其它型在计算精度上又有问题
                                objBillIncome.标准单价 = IIf(dblPriceSingle = 0, Format(dblPrice / (Val(Nvl(!数量))), gstrFeePrecisionFmt), dblPriceSingle)  '这里是售价价格
                            End If
                        Else
                            objBillIncome.标准单价 = 0
                        End If
                        '----------------------------------------------------------------------------------------------
                    Else
                        
                        If Abs(Val(Nvl(!单价))) > Val(Nvl(rsItems!现价)) Or Abs(Val(Nvl(!单价))) = 0 Then
                            objBillIncome.标准单价 = Val(Nvl(rsItems!缺省价格))
                        Else
                            objBillIncome.标准单价 = Val(Nvl(!单价))
                        End If
                    End If
                Else
                    objBillIncome.标准单价 = Val(Nvl(rsItems!现价))
                End If

                If InStr(",5,6,7,", rsItems!收费类别) > 0 And bln住院单位 Then
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价 * Nvl(rsItems!住院包装, 1), gstrFeePrecisionFmt)
                Else
                    objBillIncome.标准单价 = Format(objBillIncome.标准单价, gstrFeePrecisionFmt)
                End If
                
                objBillIncome.现价 = Val(Nvl(rsItems!现价))  '现价原价对药品变价无用
                objBillIncome.原价 = Val(Nvl(rsItems!原价))
                objBillIncome.收入项目ID = Val(Nvl(rsItems!现收入ID))
                objBillIncome.收入项目 = Nvl(rsItems!收入项目)
                objBillIncome.收据费目 = Nvl(rsItems!现费目)
                
                '应收金额=单价*付次*数次
                If Val(Nvl(rsItems!是否变价)) = 1 And (InStr(",5,6,7,", rsItems!收费类别) > 0 Or rsItems!收费类别 = "4" And Nvl(rsItems!跟踪在用, 0) = 1) Then
                    objBillIncome.应收金额 = dblPrice '保证应收金额与零售金额没有误差
                Else
                    objBillIncome.应收金额 = objBillIncome.标准单价 * objBillDetail.付数 * objBillDetail.数次
                End If
                
                '加班费用率计算
                dbl加班加价率 = 0
                If bln加班加价 And Val(Nvl(rsItems!加班加价)) = 1 Then
                    dbl加班加价率 = Val(Nvl(rsItems!加班加价)) / 100
                    objBillIncome.应收金额 = objBillIncome.应收金额 + objBillIncome.应收金额 * dbl加班加价率
                End If
                objBillIncome.应收金额 = Format(objBillIncome.应收金额, gstrDec)
                
                '计算实收金额
                If lng病人ID = 0 Then   '批量记帐(多个病人),所以此处不计算实收金额
                    objBillIncome.实收金额 = objBillIncome.应收金额
                Else
                    If Val(Nvl(rsItems!屏蔽费别)) = 1 Then
                        objBillIncome.实收金额 = objBillIncome.应收金额
                    Else
                        objBillIncome.实收金额 = ActualMoney(objBillDetail.费别, Val(Nvl(rsItems!现收入ID)), objBillIncome.应收金额, _
                            objBillDetail.收费细目ID, objBillDetail.执行部门ID, objBillDetail.原始数量, dbl加班加价率)
                    End If
                End If
                With objBillIncome
                    objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额
                End With
                '判断下一条记录是否属于当前行
                int序号 = !序号
                i = i + 1
                rsItems.MoveNext
            Loop
            
            With objBillDetail
                objBill.Details.Add .Detail, .收费细目ID, .序号, .从属父号, .病人ID, .主页ID, .病区ID, .科室ID, .姓名, .性别, .年龄, .住院号, .床号, _
                    .费别, .病人性质, .收费类别, .计算单位, .发药窗口, .付数, .数次, .附加标志, .执行部门ID, .InComes, .就诊卡号, , .担保额, .医疗付款, , , , .摘要, .原始数量, .原始执行部门ID, .婴儿费
                '分离发药时,Key设置为1,表示编辑时执行科室列不可进入
                If InStr(",5,6,7,", .收费类别) > 0 And gbln分离发药 Then
                    objBill.Details(objBill.Details.Count).Key = 1
                End If
            End With
            .MoveNext
        Loop
    End With
     '再重新处理从属父号
    For i = 1 To objBill.Details.Count
        If objBill.Details(i).从属父号 <> 0 Then
            objBill.Details(i).从属父号 = colSerial("_" & objBill.Details(i).从属父号)(1)
        End If
    Next
    Set ImportWholeSet = objBill
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub zlExecPrintSingleBill(ByVal frmMain As Object, ByVal lng病人ID As Long, _
    ByVal strPrivs As String, Optional str截止日期 As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印指定病人的催款单
    '编制:刘兴洪
    '日期:2010-10-29 16:50:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl催款金额 As Double, bytType As Byte
    If frmPatiPressMoney.zlPatiPressMoney(frmMain, glngModul, strPrivs, 0, "", lng病人ID, IIf(bytType = 2, 1, 2)) = False Then Exit Sub
End Sub
Public Sub zlPrintBedCard(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印床头卡
    '编制:刘兴洪
    '日期:2010-10-29 17:10:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", frmMain) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", frmMain, "病人ID=" & lng病人ID, "主页ID=" & lng主页ID, 2)
    End If
End Sub

Public Sub zlPrintDayDetail(ByVal frmMain As Object, ByVal int场合 As Integer, ByVal lng病人ID As Long, ByVal lng病区ID As Long, _
    Optional bln显示退费 As Boolean = False, Optional bln显示零费 As Boolean = False, Optional bln发生时间 As Boolean = True, _
    Optional lng主页ID As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印一日清单
    '参数:int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    '编制:刘兴洪
    '日期:2010-10-29 17:16:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If int场合 = 1 Then
        With frmDailyPrint
            .mlng病人ID = lng病人ID
            .mlng病区ID = lng病区ID
            .mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p一日清单)
            .Show 1, frmMain
        End With
        Exit Sub
    End If
    If Not frmDailyListAsk Is Nothing Then Unload frmDailyListAsk
    frmDailyListAsk.mlngModul = 1141    '仍然以一日清单模块的参数为准
    frmDailyListAsk.mbytInFun = 1
    frmDailyListAsk.mlng病人ID = lng病人ID
    frmDailyListAsk.mlngPageID = lng主页ID
    frmDailyListAsk.Show vbModal, frmMain
    If frmDailyListAsk.mblnAskOk Then
        ReportOpen gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1141", frmMain, "病人ID=" & lng病人ID, _
            "开始时间=" & Format(frmDailyListAsk.mdatBegin, "YYYY-MM-DD HH:MM:SS"), _
            "结束时间=" & Format(frmDailyListAsk.mdatEnd, "YYYY-MM-DD HH:MM:SS"), _
            "显示退费=" & IIf(bln显示退费, "1", "0"), _
            "显示零费用=" & IIf(bln显示零费, "1", "0"), _
            "病人病区=" & lng病区ID, _
            "主页ID=" & frmDailyListAsk.mlngPageID, _
            "费用时间=" & IIf(bln发生时间, "发生时间", "登记时间"), 1
    End If
End Sub

Public Sub zlPrintAccountPage(ByVal frmMain As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印帐页
    '编制:刘兴洪
    '日期:2010-11-01 10:03:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ReportPrintSet gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1139_2", frmMain
End Sub
Public Function zlGetPatiInsure(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保病人信息
    '返回:医保病人信息信息集
    '编制:刘兴洪
    '日期:2010-11-01 10:09:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = _
        " Select A.登记时间, B.险类, E.密码, Nvl(E.医保号, D.信息值) As 医保号,b.病人性质" & vbNewLine & _
        " From 病人信息 A, 病案主页 B, 病案主页从表 D, 医保病人档案 E, 医保病人关联表 F" & vbNewLine & _
        " Where B.病人id = [1] And B.主页id = [2] And A.病人id = B.病人id And B.病人id = D.病人id(+)" & _
        "       And B.主页id = D.主页id(+) And D.信息名(+) = '医保号' And" & vbNewLine & _
        "       A.病人id = F.病人id(+) And F.标志(+) = 1 And F.医保号 = E.医保号(+)" & _
        "       And F.险类 = E.险类(+) And F.中心 = E.中心(+)"
    On Error GoTo errH
    Set zlGetPatiInsure = zlDatabase.OpenSQLRecord(strSQL, "获取医保病人信息集", lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlPreBalance(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:预结算功能
    '成功:返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-01 10:08:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int险类 As Integer, str医保号 As String, str密码 As String
    Dim rsTmp As ADODB.Recordset, str结算费用 As String
    Dim blnDateMoved As Boolean, dat登记时间 As Date, bln门诊留观 As Boolean
    
    zlPreBalance = False
    Set rsTmp = zlGetPatiInsure(lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            int险类 = Val(!险类)
            str医保号 = "" & !医保号
            str密码 = "" & !密码
            dat登记时间 = !登记时间
            bln门诊留观 = Val(Nvl(!病人性质)) = 1
        End With
    End If
    
    If int险类 = 0 Then
        MsgBox "读取病人医保相关信息失败!", vbExclamation, gstrSysName
        Exit Function
    End If
    If gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, int险类) Then
        MsgBox "该医保接口不支持结帐设置前预结算!", vbExclamation, gstrSysName
        Exit Function
    End If
    blnDateMoved = zlDatabase.DateMoved(dat登记时间, , , "获取历史信息")
    Screen.MousePointer = 11
    If bln门诊留观 Then
        Set rsTmp = GetVBalance(0, "门诊费用结帐", int险类, lng病人ID, , , , , blnDateMoved)
    Else
        Set rsTmp = GetVBalance(1, "住院费用结帐", int险类, lng病人ID, , , , , blnDateMoved)
    End If
    Screen.MousePointer = 0
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人没有未结帐的保险项目费用!", vbInformation, gstrSysName
    Else
        str结算费用 = gclsInsure.WipeoffMoney(rsTmp, lng病人ID, str医保号, "0", int险类, "|0") '当成中途结算
        MsgBox "预结算成功!" & str结算费用, vbInformation, gstrSysName '可报销金额串:"报销方式;金额;是否允许修改|...."
        zlPreBalance = True
    End If
End Function

Public Sub zlPreBalanceAll(ByVal frmMain As Object, ByVal lng病区ID As Long)
    Dim rsTemp As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim str结算费用 As String, i As Integer, strSQL As String
    Dim lng病人ID As Long, int险类 As Integer, blnDateMoved As Boolean
    Dim str医保号 As String, str密码 As String, str姓名 As String, str登记时间 As Date, bln门诊留观 As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    If lng病区ID = 0 Then
        MsgBox "未选择病区,不能进行批量预结,请选择一个病区!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strSQL = "" & _
    "   Select distinct A.病人ID, B.主页ID,B.住院号,B.出院病床 as 床号, A.登记时间,B.险类, " & vbNewLine & _
    "       E.密码, Nvl(b.姓名, a.姓名) As 姓名,E.医保号,b.病人性质" & vbNewLine & _
    "   From 病人信息 A, 病案主页 B, 医保病人档案 E, 医保病人关联表 F,在院病人 C " & vbNewLine & _
    "   Where A.病人ID = B.病人ID  And Nvl(B.主页ID, 0) <> 0 " & vbNewLine & _
    "               And B.出院日期 is NULL And Nvl(B.状态,0)<>3 And  A.病人ID=C.病人ID  " & _
    "               And A.病人ID = F.病人ID(+) And F.标志(+) = 1 " & vbNewLine & _
    "               And F.医保号 = E.医保号 And F.险类 = E.险类 And F.中心 = E.中心(+)   " & vbNewLine & _
    "             " & vbNewLine & _
      IIf(lng病区ID = 0, " Order by 险类,住院号 Desc", "  And B.当前病区ID =[1] And C.病区ID=[1] Order by 险类,LPAD(床号,10,' ')")
      
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "批量预约功能", lng病区ID)
    If rsTemp.EOF Then
        MsgBox IIf(lng病区ID = 0, "", "当前病区") & "没有发现在院的医保病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("该操作将对" & IIf(lng病区ID = 0, "所有病人", "当前病区") & "中的所有在院医保病人(共有" & rsTemp.RecordCount & "人)进行预结算," & _
        vbCrLf & "这可能会花费较长的时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    With rsTemp
        Do While Not .EOF
            str姓名 = Nvl(!姓名)
            lng病人ID = Val(Nvl(!病人ID))
            int险类 = Val(Nvl(!险类))
            str医保号 = Nvl(!医保号)
            str密码 = Nvl(!密码)
            str登记时间 = Format(!登记时间, "yyyy-mm-dd HH:MM:SS")
            bln门诊留观 = Val(Nvl(!病人性质)) = 1
            
            If Not gclsInsure.GetCapability(support结帐_结帐设置后调用接口, lng病人ID, int险类) Then
                blnDateMoved = zlDatabase.DateMoved(str登记时间, , , "批量预结")
                Call zlCommFun.ShowFlash("正在处理医保病人""" & str姓名 & """ ...", frmMain)
                If Not frmMain Is Nothing Then frmMain.Refresh
                If bln门诊留观 Then
                    Set rsTmp = GetVBalance(0, "门诊费用结帐", int险类, lng病人ID, , , , , blnDateMoved)
                Else
                    Set rsTmp = GetVBalance(1, "住院费用结帐", int险类, lng病人ID, , , , , blnDateMoved)
                End If
                If Not rsTmp Is Nothing Then
                    If Not rsTmp.RecordCount = 0 Then
                        str结算费用 = gclsInsure.WipeoffMoney(rsTmp, lng病人ID, str医保号, "0", int险类, "|0") '当成中途结算
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    Call zlCommFun.StopFlash
    MsgBox "预结成功!", vbInformation + vbOKOnly, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function zlCheckPatiFeeRenewValied(ByVal lng病人ID As Long, _
    lng主页ID As Long, lng病区ID As Long, lng科室ID As Long, _
    ByRef str最后转出时间 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人补费是否超过时限
    '出参:str最后转出时间-最后转出时间(yyyy-mm-dd hh:mm:ss)
    '返回:true-合法补录费用;False-不能补录费用
    '编制:刘兴洪
    '日期:2010-12-10 11:04:04
    '问题:33744
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtDate As Date, strTemp As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select 终止时间,终止原因" & _
    "   From (  Select 终止时间,终止原因 From 病人变动记录  " & _
    "               Where 病人id = [1] and 主页ID=[2] and 病区ID=[3] and 科室ID=[4]  " & _
    "                           And (终止原因 = 3 or 终止原因=15 or 终止原因=10 or 终止原因=1)  " & _
    "               Order By 终止时间 Desc, 开始原因) " & _
    "   Where Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查补费的时限检查", lng病人ID, lng主页ID, lng病区ID, lng科室ID)
    If rsTemp.EOF Then rsTemp.Close: Exit Function
    
    dtDate = CDate(Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS"))
    If dtDate + 1 / 24 * gTy_System_Para.int数据补录时限 < zlDatabase.Currentdate Then
        If gTy_System_Para.int数据补录时限 = 0 Then
            strTemp = IIf(Val(Nvl(rsTemp!终止原因)) = 10, "预出院", IIf(Val(Nvl(rsTemp!终止原因)) = 1, "出院", "已转科或转病区"))
            ShowMsgbox "注意:" & vbCrLf & "    该病人" & strTemp & ",系统设置为不允许进行补录费用操作。"
        Else
            ShowMsgbox "注意:" & vbCrLf & "    该病人转科或转病区已经超过了" & gTy_System_Para.int数据补录时限 & "小时,不能进行补录费用!"
        End If
        rsTemp.Close: Exit Function
        Exit Function
    End If
    rsTemp.Close
    str最后转出时间 = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
    zlCheckPatiFeeRenewValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExecBilling_Mulit(ByVal int场合 As Integer, _
    ByVal frmMain As Object, _
    ByVal lng病区ID As Long, _
    ByVal lng病人ID As Long, bln出院 As Boolean, ByVal bln结清 As Boolean, _
    Optional strUnitIDs As String = "", Optional lng主页ID As Long = 0, _
    Optional bln补费 As Boolean = False, Optional lng科室ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用批量记帐单
    '入参:int场合- 0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS);6-费用查询调用
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-01 11:01:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng部门ID As Long
    Dim lngModule As Long, strPrivs As String
    
    zlExecBilling_Mulit = False
    If InStr(GetInsidePrivs(Enum_Inside_Program.p住院记帐), "所有病区") = 0 Then
        If strUnitIDs = "" Then
            '重新获取操作员的所在病区
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng病区ID & ",") = 0 And bln补费 = False Then
            MsgBox "你没有所有病区的权限，不能对其它病区的病人记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If int场合 = 6 Then
        If lng病区ID = 0 Then
            MsgBox "未选择批量记帐的病区，不能进行批量记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng部门ID = 0
    If bln补费 Then
        If InStr(1, "012", int场合) > 0 Then '    int场合 = 0 - 医生站调用, 1 - 护士站调用, 2 - 医技站调用(PACS / LIS))
            lng部门ID = lng科室ID
        End If
    End If
    If lng病区ID = 0 And lng部门ID = 0 Then
        MsgBox "未选择批量记帐的病区或医技部门，不能进行批量记帐！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    
    
    gbytBilling = 0
    lngModule = Enum_Inside_Program.p住院记帐
    strPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
    
    If frmChargeBat.ShowMe(frmMain, lngModule, strPrivs, 0, lng病区ID, lng部门ID, lng病人ID, bln补费) = False Then Exit Function
    zlExecBilling_Mulit = True
 
End Function

Public Function zlExecBilling(ByVal int场合 As Integer, ByVal frmMain As Object, _
    ByVal lng病区ID As Long, _
    ByVal lng病人ID As Long, bln出院 As Boolean, ByVal bln结清 As Boolean, _
    Optional strUnitIDs As String = "", Optional lng主页ID As Long = 0, _
    Optional bln补费 As Boolean = False, Optional lng科室ID As Long = 0, _
    Optional lng医嘱ID As Long = 0, Optional ByVal bln门诊留观病人 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用指定的记帐单
    '入参:int场合- 0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS);9-费用查询调用
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-01 11:01:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng部门ID As Long, str最后转科时间 As String
    
    zlExecBilling = False
    If InStr(GetInsidePrivs(Enum_Inside_Program.p住院记帐), "所有病区") = 0 Then
        If strUnitIDs = "" Then
            '重新获取操作员的所在病区
            strUnitIDs = GetUserUnits
        End If
        If InStr("," & strUnitIDs & ",", "," & lng病区ID & ",") = 0 And bln补费 = False Then
            MsgBox "你没有所有病区的权限，不能对其它病区的病人记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    lng部门ID = 0
    If bln补费 And int场合 <> 6 Then
        If InStr(1, "012", int场合) > 0 Then '    int场合 = 0 - 医生站调用, 1 - 护士站调用, 2 - 医技站调用(PACS / LIS))
                lng部门ID = lng科室ID
        End If
        '补费检查是否超过时限
        If zlCheckPatiFeeRenewValied(lng病人ID, lng主页ID, lng病区ID, lng科室ID, str最后转科时间) = False Then Exit Function
    End If
    
    '出院病人记帐权限
    If bln出院 Then
        If bln结清 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "出院结清强制记帐") = 0 Then
            MsgBox "该出院(或预出院)病人费用已经结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Function
        ElseIf Not bln结清 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), "出院未结强制记帐") = 0 Then
            MsgBox "该出院(或预出院)病人费用尚未结清,你没有权限对该病人记帐！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '门诊留观病人调用门诊记账
    If bln门诊留观病人 Then
        If Not (gbln门诊留观 And InStr(GetInsidePrivs(Enum_Inside_Program.p记帐操作), ";门诊留观记帐;") > 0) Then
            MsgBox "你没有权限对门诊留观病人进行记帐操作！", vbInformation, gstrSysName
            Exit Function
        End If
        zlExecBilling = ZLShowChargeWindow(frmMain, 2, 0, lng病人ID, lng主页ID, _
            lng部门ID, lng病区ID, bln补费, lng医嘱ID, str最后转科时间)
        Exit Function
    End If
    
    Err.Clear: On Error Resume Next
    
    gbytBilling = 0
    frmCharge.mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐)
    frmCharge.mbytInState = 0
    frmCharge.mbytUseType = 0
    frmCharge.mlngDeptID = lng部门ID
    frmCharge.mlngUnitID = lng病区ID
    frmCharge.mlngModule = 1133
    frmCharge.mlng病人ID = lng病人ID
    frmCharge.mbln补费 = bln补费
    frmCharge.mlng关联医嘱 = lng医嘱ID
    frmCharge.mlng主页ID = lng主页ID
    frmCharge.mstr最后转科时间 = str最后转科时间
    frmCharge.Show IIf(frmMain Is Nothing, 0, 1), frmMain
    If gblnOK Then zlExecBilling = True
End Function

Public Function ZLShowChargeWindow(frmMain As Object, _
    ByVal bytFun As Byte, ByVal bytInState As Byte, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lngDeptID As Long, ByVal lngUnitID As Long, _
    ByVal bln补费 As Boolean, ByVal lng关联医嘱 As Long, _
    ByVal str最后转科时间 As String, Optional ByVal strInNO As String) As Boolean
    '调用门诊费用功能
    '入参：
    '   bytFun 0-收费,1-划价,2-门诊记帐
    '   bytInState 0-执行(或修改),1-浏览,2-调整,3-退费(收费、记帐部份退费),4-重新收费;5-异常单据作废;11-复制单据
    '   lngUnitID As Long '当前记帐病区,为0时表示所有病区
    '   lngDeptID As Long '当前记帐科室,为0时表示所有科室
    '   bln补费 As Boolean '33744
    '   strInNO 传入单据，销账和复制单据时传入（仅门诊记帐时有效）
    Dim strCommon As String, intAtom As Integer, blnOk As Boolean
    
    On Error Resume Next
    If gobjCharge Is Nothing Then
        Set gobjCharge = CreateObject("zl9OutExse.clsOutExse")
        If gobjCharge Is Nothing Then Exit Function
    End If
    
    Err.Clear: On Error GoTo 0
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    blnOk = gobjCharge.Charge(frmMain, gcnOracle, glngSys, gstrDBUser, bytFun, bytInState, lng病人ID, lng主页ID, _
        lngDeptID, lngUnitID, bln补费, lng关联医嘱, str最后转科时间, strInNO)
    Call GlobalDeleteAtom(intAtom)
    ZLShowChargeWindow = blnOk
End Function

Public Function ZlIsOutpatientObserve(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '判断是否为门诊留观病人
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select 1 From 病案主页 Where 病人性质 = 1 And 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否为门诊留观病人", lng病人ID, lng主页ID)
    ZlIsOutpatientObserve = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlWrite_Off_ApplyAndVerfy(ByVal frmMain As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, _
                                            ByVal bln申请 As Boolean, Optional ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行销帐申请或审核
    '入参:bln申请:true销帐申请，false-销帐审核
    '出参:
    '返回:处理成功,返回ture,否则返回False
    '编制:刘兴洪
    '日期:2010-11-01 11:31:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lng病区ID = 0 Then
        MsgBox "请选择病人病区！", vbInformation, gstrSysName
        Exit Function
    End If
    With frmReCharge
        .mbytFun = IIf(bln申请, 0, 1)
        .mbytUseType = 0
        .mlngDeptID = lng病区ID
        .mlngPatientID = lng病人ID
        .mstrPrivs = GetInsidePrivs(Enum_Inside_Program.p住院记帐, True)
        .mstrInNO = strNO
        .Show 1, frmMain
    End With
    If gblnOK Then zlWrite_Off_ApplyAndVerfy = True
End Function
Public Function zlGet收费类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '编制:刘兴洪
    '日期:2010-11-25 14:22:18
    '问题:34260
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    If grs收费类别 Is Nothing Then
        gstrSQL = "Select 编码,类别,系统标志,独立编辑 From 收费类别"
        Set grs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    ElseIf grs收费类别.State <> 1 Then
        gstrSQL = "Select 编码,类别,系统标志,独立编辑 From 收费类别"
        Set grs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    End If
    Set zlGet收费类别 = grs收费类别
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlCheckPatiIsDeath(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否病人已经死亡.
    '入参:
    '出参:
    '返回:已经死亡,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-22 14:32:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 1 From 病案主页  where 病人ID=[1] and 出院方式 like '%死亡%' and RowNum <=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否已经死亡", lng病人ID)
    zlCheckPatiIsDeath = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCheckPatiIsMemo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人备注信息是否存在
    '返回:如果存在,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2010-12-24 09:43:14
    '问题:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select 1 From 病人备注信息 where 病人ID=[1] and nvl(主页ID,0)=[2] and Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人是否存在病人备注信息", lng病人ID, lng主页ID)
    zlCheckPatiIsMemo = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCallPatiMemoWriteAndRead(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef objInPati As Object, Optional blnOnlyReadMemo As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用病人备注修改和显示接口
    '入参:objInPati-病人部件
    '       lng病人ID-病人ID
    '       blnOnlyReadMemo-仅只读,不能编辑(暂不用,以后可能存在调整)
    '出参:
    '返回:调用成功或不存在病人信息部件,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-24 09:50:03
    '问题:34763
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If objInPati Is Nothing Then
        Set objInPati = CreateObject("zl9InPatient.clsInPatient")
    End If
    If objInPati Is Nothing Then zlCallPatiMemoWriteAndRead = True: Exit Function
    Err = 0: On Error GoTo errHandle
    'zlPatiMemoReadAndWrite(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strPrivs As String, Optional ByVal blnEdit As Boolean = False)
    Call objInPati.zlPatiMemoReadAndWrite(frmMain, gcnOracle, lng病人ID, lng主页ID, strPrivs)      ' , Not blnOnlyReadMemo
    zlCallPatiMemoWriteAndRead = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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

Public Sub zlRptControlToVsGrid(ByVal objRpt As ReportControl, ByRef objGrid As VSFlexGrid)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将rptControl的内容装配到网格控件中
    '入参:objRpt-ReportControl
    '     intPrintType-打印类型
    '编制:刘兴洪
    '日期:2011-01-31 13:19:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, objRow As ReportRow
    On Error GoTo errHandle
    objGrid.Cols = 1: objGrid.Rows = 2
    With objRpt
        j = 0
        For i = 0 To .Columns.Count - 1
            If .Columns(i).Visible Then
                objGrid.TextMatrix(0, j) = .Columns(i).Caption
                objGrid.ColKey(j) = Trim(objGrid.TextMatrix(0, j))
                objGrid.ColWidth(j) = .Columns(i).Width * Screen.TwipsPerPixelX
                Select Case .Columns(i).Alignment
                Case xtpAlignmentCenter
                    objGrid.ColAlignment(j) = flexAlignCenterCenter
                Case xtpAlignmentLeft
                    objGrid.ColAlignment(j) = flexAlignCenterCenter
                Case xtpAlignmentRight
                    objGrid.ColAlignment(j) = flexAlignRightCenter
                End Select
                objGrid.FixedAlignment(j) = flexAlignCenterCenter
                objGrid.Cols = objGrid.Cols + 1
                j = j + 1
            End If
        Next
        For Each objRow In .Rows
            If objRow.GroupRow = False Then
                For j = 0 To .Columns.Count - 1
                    If .Columns(j).Visible Then
                        '问题65471,刘尔旋:调整列排序之后,输出列内容与列头不符合的问题
                        objGrid.TextMatrix(objGrid.Rows - 1, objGrid.ColIndex(.Columns(j).Caption)) = objRow.Record(.Columns(j).ItemIndex).Value
                    End If
                Next
                objGrid.Rows = objGrid.Rows + 1
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlGet结帐冲销ID(ByVal lng结帐ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结帐冲销ID
    '入参:lng结帐ID
    '出参:
    '返回:结帐冲销的ID
    '编制:刘兴洪
    '日期:2011-02-10 11:59:36
    '问题:35554
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "结算冲销")
    If rsTemp.EOF Then
        zlGet结帐冲销ID = 0
    Else
        zlGet结帐冲销ID = rsTemp!ID '冲销单据的ID
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetMultiNOs(ByVal strNO As String, Optional lng打印ID As Long, Optional blnNOMoved As Boolean) As String
'功能：根据一张收费单据的NO，返回同一次打印的多张NO
'参数：blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
'返回：格式如"'AAA','BBB','CCC',..."
'      如果指定了"lng打印ID",则返回
'说明：用于多单据收费
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strNos As String
    Dim i As Long
    
    On Error GoTo errH
            
    lng打印ID = 0
    
    '应根据最后一次打印的情况来定
    strSQL = "Select ID,NO From " & IIf(blnNOMoved, "H", "") & "票据打印内容 Where 数据性质=1" & _
        " And ID=(Select Max(ID) From " & IIf(blnNOMoved, "H", "") & "票据打印内容 Where 数据性质=1 And NO=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If Not rsTmp.EOF Then
        lng打印ID = Nvl(rsTmp!ID, 0) '可能没有
        For i = 1 To rsTmp.RecordCount
            strNos = strNos & ",'" & rsTmp!NO & "'"
            rsTmp.MoveNext
        Next
        GetMultiNOs = Mid(strNos, 2)
    Else
        GetMultiNOs = "'" & strNO & "'"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetDelBalanceID(ByVal strNO As String) As Long
'功能：获取退费记录的结帐ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 结帐ID From 门诊费用记录 Where NO=[1] And 记录性质=1 And 记录状态=2 And Rownum < 2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If Not rsTmp.EOF Then GetDelBalanceID = Val("" & rsTmp!结帐ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_GetInvoicePrintFormat(ByVal lngModule As Long, _
    Optional str使用类别 As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '出参:int打印方式-打印方式(0-不打印;1-自动打印;2-提示打印)
    '返回:打印格式(序号)
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
    Dim lngFormat As Long, lngFormat1 As Long
    
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    strShareTypeFormat = Trim(zlDatabase.GetPara("结帐发票格式", glngSys, lngModule, ""))
    '格式:使用类别1,格式1|使用类别2,格式2...
    varData = Split(strShareTypeFormat, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",,", ",")
        lngFormat = Val(varTemp(1))
        If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
        If Trim(varTemp(0)) = str使用类别 Then
            zl_GetInvoicePrintFormat = lngFormat: Exit Function
        End If
    Next
    zl_GetInvoicePrintFormat = lngFormat1
End Function

Public Function zl_GetInvoicePrintMode(ByVal lngModule As Long, _
    Optional str使用类别 As String = "") As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票打印格式
    '出参:int打印方式-打印方式()
    '返回:0-不打印;1-自动打印;2-提示打印
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeFormat As String
     Dim intPrintMode As Long, intPrintMode1 As Long
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    strShareTypeFormat = Trim(zlDatabase.GetPara("病人结帐打印", glngSys, lngModule, ""))
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
Public Function zlisCheckOperatorICU() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前操作员是否为ICU部门的人员
    '返回:是ICU,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-05 23:29:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = _
    " Select 1" & _
    " From  部门性质说明 B,部门人员 C" & _
    " Where  B.部门ID=C.部门ID And B.工作性质='ICU' and C.人员ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前操作员是否ICU科室人员", UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        zlisCheckOperatorICU = True
    End If
    rsTemp.Close
    Set rsTemp = Nothing
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlisCheckDeptICU(ByVal lngDeptID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定科室是否为ICU部门
    '返回:是ICU,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-05 23:29:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    strSQL = _
    " Select 1" & _
    " From  部门性质说明 B " & _
    " Where  B.部门ID=[1] And B.工作性质='ICU'  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前操作员是否ICU科室人员", lngDeptID)
    If rsTemp.RecordCount <> 0 Then
        zlisCheckDeptICU = True
    End If
    rsTemp.Close
    Set rsTemp = Nothing
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsAllowFeeChange(lng病人ID As Long, lng主页ID As Long, _
   Optional int状态 As Integer = -1, Optional str姓名 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许费用变动
    '入参:int状态-(-1表示从数据库中读取审核标志进行判断;>0表示,直接根据该状态进行判断)
    '返回:允许变动返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-05-21 15:44:47
    '问题:49501,51612
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    If gTy_System_Para.byt病人审核方式 = 0 And gTy_System_Para.bln未入科禁止记账 = False Then
        ''保持歉容
        zlIsAllowFeeChange = True: Exit Function
    End If
   
    strSQL = "" & _
    " Select Nvl(审核标志,0) as 审核标志,nvl(状态,0) as 状态" & _
    " From 病案主页 " & _
    " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        
        MsgBox "未找到对应的病人信息" & IIf(str姓名 <> "", "(姓名:" & str姓名 & ")", "") & ",不允许进行记录操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '检查未入科病人不允许记账
    If gTy_System_Para.bln未入科禁止记账 And Val(Nvl(rsTemp!状态)) = 1 Then
        '51612
        MsgBox "病人未入科(" & IIf(str姓名 <> "", "姓名:" & str姓名, "") & "第" & lng主页ID & "次住院) ,不能对该病人进行记账或销账操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '审核相关检查
    If gTy_System_Para.byt病人审核方式 = 0 Then zlIsAllowFeeChange = True: Exit Function
    If int状态 < 0 Then
        int状态 = Val(Nvl(rsTemp!审核标志))
    End If
    '检查相关状态
    If int状态 = 1 Then
        MsgBox "病人" & IIf(str姓名 <> "", ":" & str姓名, "") & "在第" & lng主页ID & "次住院中已经开始审核费用,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    If int状态 = 2 Then
        MsgBox "已经完成了对病人" & IIf(str姓名 <> "", ":" & str姓名, "") & "第" & lng主页ID & "次住院费用的审核,不能对该病人进行费用变动。", vbInformation, gstrSysName
        Exit Function
    End If
    zlIsAllowFeeChange = True
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
    

Public Function GetMzBalanceData(lng病人ID As Long, Optional strDeptIDs As String, _
    Optional strClass As String, Optional dtStartDate As Date, Optional dtEndDate As Date, _
    Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
    Optional bytKind As Byte, Optional strChargeType As String = "", _
    Optional bln发生时间 As Boolean, Optional strTime As String, Optional strChargeTypeNot As String = "", Optional strDiag As String = "") As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定条件的门诊结帐数据
    '入参：lng病人ID-病人ID,
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：""-所有费用(含未设置),"'类型1','类型2',..."
    '      strItem：收据费目串,'西药费','中药费',...
    '      DateBegin,DateEnd：结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnZero：是否读取零费用
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      bytKind：  0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    '      strChargeType:""表示所有费用,否则为指定收费类别的费用;如:5,6,7等  '34260
    '      bln发生时间-是按发生时间统计:true-发生时间;false-按登记时间统计
    '      strTime-门诊留观病人留观次数,0,1,2...(0表示非留观记帐数据)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-04 17:57:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, strCond2 As String
    Dim strTable As String, strWherePage As String, strConO As String
    Dim strDiagCondition As String
    On Error GoTo errHandle
    strWherePage = IIf(strTime = "", "", " And Instr([8],','||Nvl(A.主页ID,0)||',')>0")
         
    strCond = " And A.病人ID=[1]"
    If Not dtStartDate = CDate("0:00:00") Then
        strConO = strCond
        strCond = strCond & " And " & IIf(Not bln发生时间, "A.登记时间", "A.发生时间") & " Between [2] And [3]"
        dtStartDate = CDate(Format(dtStartDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
    
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([4],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([5],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([6],','''||A.收费类别||''',')>0")   '34260
    strCond = strCond & IIf(strChargeTypeNot = "", "", " And Instr([9],','||A.收费类别|| ',')=0")
    
    strConO = strConO & IIf(strDeptIDs = "", "", " And Instr([4],','||A.开单部门ID||',')>0")
    strConO = strConO & IIf(strItem = "", "", " And Instr([5],','''||A.收据费目||''',')>0")
    strConO = strConO & IIf(strChargeType = "", "", " And Instr([6],','''||A.收费类别||''',')>0")   '34260
    strConO = strConO & IIf(strChargeTypeNot = "", "", " And Instr([9],','||A.收费类别|| ',')=0")
    
    If Not (strDiag = "" Or strDiag = "所有诊断") Then
        strDiagCondition = " And Exists (Select 1 From 病人诊断医嘱 K,病人诊断记录 L Where K.医嘱ID = A.医嘱序号 And K.诊断ID = L.ID And 诊断描述 = [10])"
    End If
    
    
    '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    If bytKind = 1 Then
        '仅体检费用
        strCond = strCond & " And A.门诊标志=4"
        strConO = strConO & " And A.门诊标志=4"
    Else
        strCond = strCond & " And A.门诊标志<>2"
        If bytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
        strConO = strConO & " And A.门诊标志<>2"
        If bytKind = 0 Then strConO = strConO & " And A.门诊标志<>4"
    End If
    
    strCond2 = strCond   '已经结过帐的,不管是否上传都要取,所以先把这个条件记录下来,第二个子查询用
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null) "
        strConO = strConO & " And A.是否上传=1  "
    Else
        strCond = strCond & " And A.结帐ID Is Null "
    End If
    
 
     
    '住院,科室,时间,[单据号],项目,费目,婴儿费,[ID],[序号],[记录性质],[记录状态],[执行状态],[A.主页ID],[A.开单部门ID],[登记时间],未结金额,结帐金额,[类型]
    If blnZero Then
        strTable = "" & _
        "   SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,1 as 标志,'门诊' as 住院," & _
        "           A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "           Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目, " & _
        "           A.数次 * A.付数 As 数量, A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额, A.统筹金额," & _
        "           A.费用类型, A.收费类别,A.费别, A.执行部门id, A.开单人, A.保险大类id, A.门诊标志,A.医嘱序号 " & _
        "   From 门诊费用记录 A  " & _
        "   Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWherePage & _
        "   Union  all" & _
        "   SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,1 as 标志,'门诊' as 住院," & _
        "          A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "          Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
        "          A.数次 * A.付数 As 数量, A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额, A.统筹金额," & _
        "          A.费用类型, A.收费类别,A.费别, A.执行部门id,A.开单人, A.保险大类id, A.门诊标志,A.医嘱序号" & _
        "   From 住院费用记录 A  " & _
        "   Where A.记录状态<>0 And (Mod(A.记录性质,10) = 5) And A.记帐费用=1" & strCond
    Else
        strTable = "" & _
        " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,1 as 标志,'门诊' as 住院," & _
        "       A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "       Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目, " & _
        "       A.数次 * A.付数 As 数量, A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额, A.统筹金额," & _
        "       A.费用类型, A.收费类别,A.费别, A.执行部门id, A.开单人, A.保险大类id, A.门诊标志,A.医嘱序号" & _
        " From 门诊费用记录 A," & _
        "      (    Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
        "           From  门诊费用记录 A " & _
        "           Where A.记录状态<>0  And A.记帐费用=1  And Nvl(A.实收金额,0) <> 0  And A.结帐ID Is Null " & strConO & strWherePage & _
        "           Group by A.NO,A.序号,A.记录性质 " & _
        "           Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
        " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond & strWherePage
        
        strTable = strTable & " Union ALL" & _
        " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态, 1 as 标志,'门诊' as 住院," & _
        "        A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "        Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目, " & _
        "        A.数次 * A.付数 As 数量, A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额, A.统筹金额," & _
        "        A.费用类型, A.收费类别,A.费别, A.执行部门id, A.开单人, A.保险大类id, A.门诊标志,A.医嘱序号" & _
        " From   住院费用记录 A ," & _
        "      (    Select A.NO,A.序号,A.记录性质,Nvl(Sum(A.实收金额),0) as 实收金额" & _
        "           From 住院费用记录 A" & _
        "           Where A.记录状态<>0 And (Mod(A.记录性质,10) = 5)  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0 And A.结帐ID Is Null" & strConO & _
        "           Group by A.NO,A.序号,A.记录性质 " & _
        "           Having Nvl(Sum(A.实收金额),0)<>0 ) B " & _
        " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond
    End If
     
    strSQL = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录", 2, "", True, ""), "门诊费用记录")
        
    '住院结帐大改时取掉(原因不明,以前为什么要加入,以待以后查证):
    '   And (Nvl(A.实收金额,0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)
        
        
    strSQL = "" & _
    "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,Max(A.记录状态) As 记录状态,A.执行状态," & _
    "              1 as 标志,'门诊' as 住院,A.主页ID," & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
    "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
    "               avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价,Sum(Nvl(A.应收金额,0)) as 应收金额, Sum(Nvl(A.实收金额,0)) as 实收金额, " & _
    "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额, avg(A.统筹金额) as 统筹金额," & _
    "               A.费用类型, max(A.收费类别) as 收费类别, " & _
    "               max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人, " & _
    "               max( A.保险大类id) as 保险大类id, A.门诊标志,A.医嘱序号 " & _
    "        FROM " & strSQL & " A" & _
    "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 " & _
    "              And Not Exists (Select 1 From 门诊费用记录 C, 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond2 & strWherePage & _
    "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
    "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and  Sum(Nvl(A.结帐金额,0))=0  And Mod(Count(*),2)=0) " & _
    "                     Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
    "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.执行状态, " & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,A.主页ID, " & _
    "               Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型,A.门诊标志,A.医嘱序号" & _
    ""
    
    strSQL = strSQL & " Union ALL " & _
    "        SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,Max(A.记录状态) As 记录状态,A.执行状态," & _
    "               1 as 标志,'门诊' as 住院,A.主页ID," & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
    "               Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
    "               avg(A.数次 * nvl(A.付数,1)) As 数量, avg(A.标准单价) as 标准单价,Sum(Nvl(A.应收金额,0)) as 应收金额, Sum(Nvl(A.实收金额,0)) as 实收金额, " & _
    "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额, avg(A.统筹金额) as 统筹金额," & _
    "               A.费用类型, max(A.收费类别) as 收费类别," & _
    "               max(A.费别) as 费别, max(A.执行部门id) as 执行部门id, max(A.开单人) as 开单人,max( A.保险大类id) as 保险大类id, A.门诊标志,A.医嘱序号 " & _
    "        FROM " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", 2, "", True, ""), "住院费用记录") & " A" & _
    "        Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 And (Mod(A.记录性质,10) = 5)  " & _
    "              And Not Exists (Select 1  From 住院费用记录 C, 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond2 & strWherePage & _
    "        Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
    "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and  Sum(Nvl(A.结帐金额,0))=0  And Mod(Count(*),2)=0) " & _
    "                     Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
    "        Group by A.NO,A.序号,Mod(A.记录性质,10),A.执行状态,A.主页ID, " & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Decode(Nvl(A.婴儿费,0),0,'','√'),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型,A.门诊标志,A.医嘱序号" & _
    ""
 
    strTable = strTable & " Union ALL " & strSQL
    
    strSQL = _
        "Select A.标志,A.住院,Nvl(B.名称,'未知') as 科室,A.时间,A.NO as 单据号 ,Nvl(E.名称,C.名称) as 项目,A.收据费目 as 费目, A.婴儿费,A.ID,A.序号,A.记录性质,A.记录状态,A.执行状态,A.主页ID,A.开单部门ID,A.登记时间," & _
        "      A.数量, A.标准单价 as 价格,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额, " & _
        "       Nvl(A.未结金额,0) 未结金额,Nvl(A.未结金额,0) 结帐金额, A.统筹金额," & _
        "       Nvl(A.费用类型,C.费用类型) as 类型, A.收费类别,M.名称 as 收费类别名," & _
        "       A.费别, A.执行部门id, A.开单人, A.保险大类id,A.收费细目ID,C.计算单位,A.门诊标志,Decode(a.记录状态, 2, 3, 3, 2, 1) As 排序,Max(G.诊断描述) As 诊断" & _
        " From (  " & strTable & ") A,部门表 B,收费项目目录 C,收入项目 D,收费项目别名 E,收费项目类别 M,病人诊断医嘱 F,病人诊断记录 G " & _
        " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID And a.医嘱序号 = f.医嘱id(+) And f.诊断id = g.Id(+) And A.收入项目ID=D.ID" & strDiagCondition & _
                IIf(strClass = "", "", " And Instr([7],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
        "       And A.收费类别=M.编码(+) And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        " Group By A.标志,A.住院,Nvl(B.名称,'未知'),A.时间,A.NO ,Nvl(E.名称,C.名称),A.收据费目, A.婴儿费,A.ID,A.序号,A.记录性质,A.记录状态,A.执行状态,A.主页ID,A.开单部门ID,A.登记时间," & _
        "      A.数量, A.标准单价,nvl(A.应收金额,0),nvl(A.实收金额,0), " & _
        "       Nvl(A.未结金额,0),Nvl(A.未结金额,0), A.统筹金额," & _
        "       Nvl(A.费用类型,C.费用类型), A.收费类别,M.名称," & _
        "       A.费别, A.执行部门id, A.开单人, A.保险大类id,A.收费细目ID,C.计算单位,A.门诊标志,Decode(a.记录状态, 2, 3, 3, 2, 1)" & _
        " Order by A.时间 Desc,A.住院,A.NO Desc,A.记录性质,A.序号"
    
    On Error GoTo errHandle
    Set GetMzBalanceData = zlDatabase.OpenSQLRecord(strSQL, "获取门诊结帐数据", lng病人ID, dtStartDate, dtEndDate, _
                    "," & strDeptIDs & ",", "," & strItem & ",", "," & strChargeType & ",", "," & strClass & ",", _
                    "," & strTime & ",", "," & strChargeTypeNot & ",", strDiag)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetZYBalanceData(lng病人ID As Long, Optional strTime As String, _
    Optional strDeptIDs As String, Optional strClass As String, Optional DateBegin As Date, Optional DateEnd As Date, _
    Optional strBaby As String, Optional strItem As String, _
    Optional blnOnlyYbUpData As Boolean, Optional ByVal blnZero As Boolean, Optional blnDateMoved As Boolean, _
     Optional strChargeType As String = "", Optional bln发生时间 As Boolean, Optional strChargeTypeNot As String = "", Optional strDiag As String = "") As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定病人未结帐金额细目(按每收入项目行)
    '入参：lng病人ID-病人ID,
    '      strTime：住院次数串,"1,2,3"
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：""-所有费用(含未设置),"'类型1','类型2',..."
    '      strItem：收据费目串,'西药费','中药费',...
    '      strBaby：0-所有费用,1-病人费用,2以及上-第bytBaby-1个婴儿费用
    '      DateBegin,DateEnd：结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnZero：是否读取零费用
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      bln发生时间-是否按发生时间统计,true-发生时间，false-登记时间
    '      strChargeType:""表示所有费用,否则为指定收费类别的费用;如:5,6,7等
    '返回：成功=记录集,失败=Nothing
    '编制：刘兴洪
    '日期：2010-03-06 13:21:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim strSQL As String, strCond As String, strCond2 As String, strConO As String
    Dim strTable As String, bytType As Byte '0-门诊,1-住院,2-门诊和住院
    Dim strWherePage As String '住院次数条件
    Dim strWhereMzPage As String
    Dim strDiagCondition As String
        
    strCond = " And A.病人ID=[1]"
    
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
    
    If Not DateBegin = CDate("0:00:00") Then
        strConO = strCond
        strCond = strCond & " And " & IIf(Not bln发生时间, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
        DateBegin = CDate(Format(DateBegin, "yyyy-MM-dd 00:00:00"))
        DateEnd = CDate(Format(DateEnd, "yyyy-MM-dd 23:59:59"))
    End If
    
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([5],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.婴儿费,0) ||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([9],','''||A.收费类别||''',')>0")   '34260
    strCond = strCond & IIf(strChargeTypeNot = "", "", " And Instr([10],','||A.收费类别|| ',')=0")
    strCond = strCond & " And A.门诊标志 In (2,3) "
    
    strConO = strConO & IIf(strDeptIDs = "", "", " And Instr([5],','||A.开单部门ID||',')>0")
    strConO = strConO & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.婴儿费,0) ||',')>0")
    strConO = strConO & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strConO = strConO & IIf(strChargeType = "", "", " And Instr([9],','''||A.收费类别||''',')>0")   '34260
    strConO = strConO & IIf(strChargeTypeNot = "", "", " And Instr([10],','||A.收费类别|| ',')=0")
    strConO = strConO & " And A.门诊标志 In (2,3) "
    
    If Not (strDiag = "" Or strDiag = "所有诊断") Then
        strDiagCondition = " And Exists (Select 1 From 病人诊断医嘱 K,病人诊断记录 L Where K.医嘱ID = A.医嘱序号 And K.诊断ID = L.ID And 诊断描述 = [11])"
    End If
    
    strCond2 = strCond   '已经结过帐的,不管是否上传都要取,所以先把这个条件记录下来,第二个子查询用
    
    
    If blnOnlyYbUpData Then
        strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null) "
        strConO = strConO & " And A.是否上传=1"
    Else
        strCond = strCond & " And A.结帐ID Is Null "
    End If
  
    
    '住院,科室,时间,[单据号],项目,费目,婴儿费,[ID],[序号],[记录性质],[记录状态],[执行状态],[A.主页ID],[A.开单部门ID],[登记时间],未结金额,结帐金额,[类型]
    If blnZero Then
        strTable = "" & _
        "   SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,2 as 标志, '第'||NVL(A.主页ID,0)||'次' as 住院," & _
        "           A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
        "           Nvl(A.婴儿费,0) as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
        "           A.数次 * A.付数 As 数量, A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额, A.统筹金额," & _
        "           A.费用类型,A.收费类别,A.费别,A.执行部门id,A.开单人, A.保险大类id,A.医嘱序号" & _
        "   From 住院费用记录 A " & _
        "   Where A.记录状态<>0 And A.记帐费用=1" & strCond & strWherePage & _
        ""
    Else
        strTable = "" & _
            " SELECT  A.ID,A.NO,A.序号,A.记录性质,A.记录状态,A.执行状态,2 as 标志,'第'||NVL(A.主页ID,0)||'次'  as 住院," & _
            "           A.主页ID,A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
            "           Nvl(A.婴儿费,0) as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
            "           A.数次 * A.付数 As 数量,A.标准单价,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额,Nvl(A.实收金额,0) as 未结金额,A.统筹金额," & _
            "           A.费用类型,A.收费类别,A.费别,A.执行部门id,A.开单人,A.保险大类id,A.医嘱序号" & _
            " From  住院费用记录 A," & _
            "      ( Select A.NO,A.序号,A.记录性质, Nvl(Sum(A.实收金额),0) as 实收金额" & _
            "        From  住院费用记录 A" & _
            "        Where A.记录状态<>0  And A.记帐费用=1 And Nvl(A.实收金额,0)<>0  And A.结帐ID Is Null " & strConO & strWherePage & _
            "        Group by A.NO,A.序号,A.记录性质 Having Nvl(Sum(A.实收金额),0)<>0  ) B " & _
            " Where A.NO=B.NO And A.序号=B.序号 And A.记录性质=B.记录性质 And A.结帐ID Is Null" & strCond & strWherePage
    End If
    
    '住院结帐大改时取掉(原因不明,以前为什么要加入,以待以后查证):
    '   And (Nvl(A.实收金额,0)<>Nvl(A.结帐金额, 0) Or Nvl(A.结帐金额, 0)=0)
    
    strSQL = "" & _
    "   SELECT 0 as ID,A.NO,A.序号,Mod(A.记录性质,10) as 记录性质,Max(A.记录状态) As 记录状态,Nvl(A.执行状态,0) As 执行状态,2 as 标志," & _
    "             '第'||NVL(A.主页ID,0)||'次'  as 住院,A.主页ID," & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 时间,A.登记时间," & _
    "               Nvl(A.婴儿费,0) as 婴儿费,A.收费细目ID,A.收入项目ID,A.收据费目," & _
    "               avg(A.数次 * nvl(A.付数,1)) As 数量,avg(A.标准单价) as 标准单价,Sum(Nvl(A.应收金额,0)) as 应收金额,Sum(Nvl(A.实收金额,0)) as 实收金额," & _
    "               Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 未结金额,avg(A.统筹金额) as 统筹金额," & _
    "               A.费用类型,max(A.收费类别) as 收费类别,max(A.费别) as 费别,max(A.执行部门id) as 执行部门id,max(A.开单人) as 开单人," & _
    "               max( A.保险大类id) as 保险大类id,A.医嘱序号 " & _
    "   FROM 住院费用记录 A " & _
    "   Where A.结帐id Is Not Null And A.记录状态<>0 And A.记帐费用=1 " & _
    "         And Not Exists (Select 1 From 住院费用记录 C, 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond2 & strWherePage & _
    "   Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
    "                    Or (Sum(Nvl(A.实收金额, 0)) = 0 And Sum(Nvl(A.应收金额, 0)) <> 0 and Sum(Nvl(A.结帐金额,0)) =0 And Mod(Count(*),2)=0) " & _
    "                    Or Sum(Nvl(A.结帐金额, 0))=0 And Sum(Nvl(A.应收金额,0))<>0 And Mod(Count(*),2)=0" & _
    "   Group by A.NO,A.序号,Mod(A.记录性质,10),Nvl(A.执行状态,0), '第'||NVL(A.主页ID,0)||'次' ,A.主页ID," & _
    "               A.开单部门ID,To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS'),A.登记时间,Nvl(A.婴儿费,0),A.收费细目ID,A.收入项目ID,A.收据费目,A.费用类型,A.医嘱序号" & _
    ""
    
    strTable = strTable & " Union ALL " & strSQL
    strSQL = _
        "Select A.标志,A.住院,Nvl(B.名称,'未知') as 科室,A.时间,A.NO as 单据号 ,Nvl(E.名称,C.名称) as 项目,A.收据费目 as 费目, A.婴儿费, " & _
        "       A.ID,A.序号,A.记录性质,A.记录状态,A.执行状态,A.主页ID,A.开单部门ID,A.登记时间," & _
        "       A.数量, A.标准单价 as 价格,nvl(A.应收金额,0) as 应收金额,nvl(A.实收金额,0) as 实收金额, " & _
        "       Nvl(A.未结金额,0) 未结金额,Nvl(A.未结金额,0) 结帐金额, A.统筹金额," & _
        "       Nvl(A.费用类型,C.费用类型) as 类型, A.收费类别,M.名称 as 收费类别名," & _
        "       A.费别, A.执行部门id, A.开单人, A.保险大类id,A.收费细目ID,C.计算单位,Decode(a.记录状态, 2, 2, 3, 2, 1) As 排序, Max(G.诊断描述) As 诊断 " & _
        " From (  " & strTable & ") A,部门表 B,收费项目目录 C,收入项目 D,收费项目别名 E,收费项目类别 M,病人诊断医嘱 F,病人诊断记录 G " & _
        " Where A.开单部门ID=B.ID(+) And A.收费细目ID=C.ID And a.医嘱序号 = f.医嘱id(+) And f.诊断id = g.Id(+) And A.收入项目ID=D.ID  And A.收费类别=M.编码(+) " & strDiagCondition & _
        "        And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(strClass = "", "", " And Instr([8],','''||Nvl(A.费用类型,Nvl(C.费用类型,'未知'))||''',')>0") & _
        " Group By A.标志,A.住院,Nvl(B.名称,'未知'),A.时间,A.NO ,Nvl(E.名称,C.名称),A.收据费目, A.婴儿费, " & _
        "       A.ID,A.序号,A.记录性质,A.记录状态,A.执行状态,A.主页ID,A.开单部门ID,A.登记时间," & _
        "       A.数量, A.标准单价,nvl(A.应收金额,0),nvl(A.实收金额,0), " & _
        "       Nvl(A.未结金额,0),Nvl(A.未结金额,0), A.统筹金额," & _
        "       Nvl(A.费用类型,C.费用类型), A.收费类别,M.名称," & _
        "       A.费别, A.执行部门id, A.开单人, A.保险大类id,A.收费细目ID,C.计算单位,Decode(a.记录状态, 2, 2, 3, 2, 1) " & _
        " Order by A.时间 Desc,A.住院,A.NO Desc,A.记录性质,A.序号"
    
    'Mod(Count(*),2)=1是为了区别打折后实收金额为零的费用在结帐后是否作废或再次结帐
    On Error GoTo errH
    Set GetZYBalanceData = zlDatabase.OpenSQLRecord(strSQL, "获取住院结帐记录", lng病人ID, "," & strTime & ",", DateBegin, DateEnd, _
                    "," & strDeptIDs & ",", "," & strBaby & ",", "," & strItem & ",", "," & strClass & ",", "," & strChargeType & ",", "," & strChargeTypeNot & ",", strDiag)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 



Public Function GetMzBalance_Insure(ByVal int险类 As Integer, lng病人ID As Long, _
     Optional dtBeginDate As Date, Optional dtEndDate As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional blnOnly门诊 As Boolean, Optional bytKind As Byte, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, _
     Optional strChargeType As String = "", Optional bln发生时间 As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定病人未结帐细目明细(按收费细目)
    '入参：lng病人ID-病人ID,
    '      dtBeginDate,dtEndDate： 结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      blnOnly门诊：仅门诊记帐费用
    '      bytKind：0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    '      strItem:收据费目串,'西药费','中药费',...
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：所有费用(含未设置),"'类型1','类型2',..."
    '      bln发生时间-是否按发生时间统计,true-发生时间，false-登记时间
    '返回：成功=记录集,失败=Nothing
    '编制：刘兴洪
    '日期:2015-01-06 17:28:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String
    
    On Error GoTo errH
     
    
    blnRelation = gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类)
     
    strCond = " And A.病人ID=[1]"
    If dtBeginDate <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(Not bln发生时间, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
        dtBeginDate = CDate(Format(dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
 
    '刘兴洪:2010-03-06 11:23:52: Or A.结帐ID is Not NULL 这个包含了已结帐的明细,在我的分析来看,是错的,但又没听说医保又问题,因此暂不更改!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null Or A.结帐ID is Not NULL)"
    
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.收费类别||''',')>0")   '34260
    
     '0-仅普通费用,1-仅体检费用,2-普通费用和体检费用
    If bytKind = 1 Then '仅体检费用
        strCond = strCond & " And A.门诊标志=4"
    Else
        strCond = strCond & " And A.门诊标志<>2"
        If bytKind = 0 Then strCond = strCond & " And A.门诊标志<>4"
    End If
    '结算要求：记录性质,记录状态,NO、序号、收费类别、收费细目ID,收费名称、计算单位、开单部门、规格、产地、数量、价格、金额、医生,
    '          发生时间,登记时间,婴儿费,医保项目编码、保险大类ID、保险项目否、是否上传,是否急诊
    '注意：由于结算只能针对有保险项目编码的,所以在与保险支付项目连接时不用(+)
    '   数次为零指按费别重算后产生的打折冲减记录,这类单据明细不上传
    
    '临时更改：结帐作废后该SQL不对,单价暂时改为"金额/数量"
    If blnOnly门诊 Then
        '门诊结帐
        strTable = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A")
        strTable = "" & _
        "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, Max(Decode(A.记录性质,2,A.摘要,Null)) As 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
        "              Avg(Nvl(付数, 0) * 数次) As 数量, Sum(标准单价*decode(sign(A.记录性质-10),1,0,1)) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
        "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
        "       From " & strTable & vbNewLine & _
        "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录", , , , "C"), "门诊费用记录 C") & ", 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond & vbNewLine & _
        "       Group By NO, Mod(记录性质, 10),decode(记录状态,2,2,1), Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 是否急诊," & vbNewLine & _
        "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
        "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 "
        
        strTable = strTable & " Union ALL " & _
        "       Select NO, 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, Max(Decode(A.记录性质,5,A.摘要,Null)) As 摘要, 是否急诊, 开单部门id, 执行部门id," & vbNewLine & _
        "              Avg(Nvl(付数, 0) * 数次) As 数量, Sum(标准单价*decode(sign(A.记录性质-10),1,0,1)) As 单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 实收金额," & vbNewLine & _
        "              Sum(统筹金额) As 统筹金额,费用类型" & vbNewLine & _
        "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & vbNewLine & _
        "       Where 记录状态 <> 0 And 记帐费用 = 1 And A.数次 <> 0 And  Mod(A.记录性质,10) = 5 And A.主页ID Is Null And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", , , , "C"), "住院费用记录 C") & ", 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond & vbNewLine & _
        "       Group By NO, Mod(记录性质, 10), decode(记录状态,2,2,1), Nvl(价格父号, 序号), 病人id, 收费类别, 收据费目, 计算单位, 开单人, 收费细目id, 是否急诊," & vbNewLine & _
        "                开单部门id, 执行部门id,费用类型" & vbNewLine & _
        "       Having Sum(Nvl(实收金额, 0)) - Sum(Nvl(结帐金额, 0)) <> 0 "
            
        strSQL = "" & _
        " Select Sysdate As 结算时间, A.病人id, A.收费类别, A.收据费目, A.计算单位, A.收费细目id," & vbNewLine & _
        "       B.大类id 保险支付大类id, B.是否医保 是否医保, B.项目编码 保险编码, Sum(A.数量) As 数量, Avg(A.单价) As 单价," & vbNewLine & _
        "       Sum(A.实收金额) As 实收金额, Sum(A.统筹金额) As 统筹金额, Max(A.摘要) 摘要, Max(A.是否急诊) 是否急诊," & vbNewLine & _
        "       Max(A.开单部门id) 开单部门id, Max(A.执行部门id) 执行部门id, Max(A.开单人) 开单人,Max(A.费用类型) 费用类型" & vbNewLine & _
        " From ( " & strTable & ") A, 保险支付项目 B, 收费项目目录 C " & vbNewLine & _
        " Where A.收费细目id = C.ID And A.收费细目id = B.收费细目id" & IIf(blnRelation, "(+)", "") & " And B.险类" & IIf(blnRelation, "(+)", "") & " = [5] " & vbNewLine & _
                    IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.费用类型,Nvl(C.费用类型,'无'))||''',')>0") & _
        " Group By A.收费细目id, A.病人id, A.收费类别, A.收据费目, A.计算单位, B.大类id, B.是否医保, B.项目编码" & vbNewLine & _
        " Having Sum(A.实收金额) <> 0"
    Else
        '住院结帐
       strTable = IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录"), "门诊费用记录 A")
        strTable = "" & _
        "       Select Mod(A.记录性质, 10) As 记录性质, decode(A.记录状态,2,2,1) as 标志,max(A.记录状态) as 记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
        "               -1*NULL  as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
        "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
        "              Sum(A.标准单价*decode(sign(A.记录性质-10),1,0,1)) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
        "              A.登记时间, Min(Nvl(A.是否上传, 0)) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, Max(Decode(A.记录性质,2,A.摘要,Null)) As 摘要,A.费用类型" & vbNewLine & _
        "       From " & strTable & " , 收入项目 B" & vbNewLine & _
        "       Where A.记录状态 <> 0 And A.记帐费用 = 1 And A.收入项目id = B.ID And A.数次 <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("门诊费用记录", , , , "C"), "门诊费用记录 C") & ", 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond & vbNewLine & _
        "       Group By Mod(A.记录性质, 10), decode(A.记录状态,2,2,1), A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id," & vbNewLine & _
        "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
        "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),A.费用类型" & vbNewLine & _
        "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　"
            

        strTable = strTable & " Union ALL" & _
        "       Select Mod(A.记录性质, 10) As 记录性质,decode(A.记录状态,2,2,1) as 标志,max(A.记录状态) as 记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
        "               -1*NULL as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
        "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
        "              Sum(A.标准单价*decode(sign(A.记录性质-10),1,0,1)) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
        "              A.登记时间, Min(Nvl(A.是否上传, 0)) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, Max(Decode(A.记录性质,5,A.摘要,Null)) As 摘要,A.费用类型" & vbNewLine & _
        "       From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A") & " , 收入项目 B" & vbNewLine & _
        "       Where A.记录状态 <> 0 And A.记帐费用 = 1 And  Mod(A.记录性质,10) = 5 And A.主页ID Is Null  And A.收入项目id = B.ID And A.数次 <> 0 And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", , , , "C"), "住院费用记录 C") & ", 病人结帐记录 D Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id And Nvl(d.结算状态, 0) = 1) " & strCond & vbNewLine & _
        "       Group By Mod(A.记录性质, 10), decode(A.记录状态,2,2,1), A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id  ," & vbNewLine & _
        "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
        "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),A.费用类型" & vbNewLine & _
        "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　"
        strSQL = "" & _
        "   Select A.记录性质, A.记录状态, A.NO, A.序号, A.门诊标志, A.病人id, A.主页id, A.婴儿费, C.项目编码 As 医保项目编码," & vbNewLine & _
        "       A.保险编码, A.保险大类id, A.收费类别, A.收费细目id, Nvl(E.名称, B.名称) As 收费名称, A.计算单位," & vbNewLine & _
        "       X.名称 As 开单部门, B.规格, B.产地, A.数量, A.标准单价 As 价格, A.金额," & vbNewLine & _
        "       A.医生, A.发生时间, A.登记时间, A.是否上传, A.是否急诊, A.保险项目否, A.摘要,A.费用类型" & vbNewLine & _
        "   From ( " & strTable & ") A, 收费项目目录 B, 保险支付项目 C, 收费项目别名 E,部门表 X" & vbNewLine & _
        "   Where A.收费细目id = B.ID And B.ID = C.收费细目id" & IIf(blnRelation, "(+)", "") & " And C.险类" & IIf(blnRelation, "(+)", "") & " = [5] And A.开单部门id = X.ID " & vbNewLine & _
                IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.费用类型,Nvl(B.费用类型,'无'))||''',')>0") & _
        "      And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1)
    End If
    Set GetMzBalance_Insure = zlDatabase.OpenSQLRecord(strSQL, "获取医保门诊费费数据", lng病人ID, "", dtBeginDate, dtEndDate, int险类, 0, "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetZYBalance_Insure(ByVal int险类 As Integer, lng病人ID As Long, Optional strTime As String, _
     Optional dtBeginDate As Date, Optional dtEndDate As Date, Optional blnOnlyYbUpData As Boolean, _
     Optional blnDateMoved As Boolean, Optional strBaby As String, _
     Optional strItem As String, Optional strDeptIDs As String, Optional strClass As String, _
     Optional strChargeType As String = "", Optional bln发生时间 As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定病人未结帐细目明细(按收费细目)
    '入参：lng病人ID-病人ID,
    '      strTime： 医保病人只能设置住院次数和费用期间 [strTime=住院次数串,"0,1,2,3",0表示门诊]
    '      dtBeginDate,dtEndDate： 结帐费用期间,按登记时间或发生时间,缺省值为CDate("0:00:00")
    '      blnOnlyYbUpData：是否只处理已上传部份
    '      blnDateMoved：病人登记时间是否在转出数据之前
    '      strBaby：0-所有费用,1-病人费用,2以及上-第bytBaby-1个婴儿费用]
    '      strItem:收据费目串,'西药费','中药费',...
    '      strDeptIds：开单科室ID串,"1,2,3",空表示所有
    '      strClass：所有费用(含未设置),"'类型1','类型2',..."
    '      bln发生时间-是否按发生时间统计,true-发生时间，false-登记时间
    '出参：
    '返回：成功=记录集,失败=Nothing
    '编制：刘兴洪
    '日期:2015-01-06 17:31:52
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCond As String, blnRelation As Boolean
    Dim strTable As String, strWherePage As String '住院次数条件
    On Error GoTo errH
    
    blnRelation = gclsInsure.GetCapability(support允许不设置医保项目, lng病人ID, int险类)
    
    strCond = " And A.病人ID=[1]"
    strWherePage = IIf(strTime = "", "", " And Instr([2],','||Nvl(A.主页ID,0)||',')>0")
    
    If dtBeginDate <> CDate("0:00:00") Then
        strCond = strCond & " And " & IIf(Not bln发生时间, "A.登记时间", "A.发生时间") & " Between [3] And [4]"
        dtBeginDate = CDate(Format(dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
 
    '刘兴洪:2010-03-06 11:23:52: Or A.结帐ID is Not NULL 这个包含了已结帐的明细,在我的分析来看,是错的,但又没听说医保又问题,因此暂不更改!
    If blnOnlyYbUpData Then strCond = strCond & " And (A.是否上传=1 And A.结帐ID is Null Or A.结帐ID is Not NULL)"

    strCond = strCond & IIf(strBaby = "", "", " And Instr([6],','|| Nvl(A.婴儿费,0) ||',')>0")
    strCond = strCond & IIf(strItem = "", "", " And Instr([7],','''||A.收据费目||''',')>0")
    strCond = strCond & IIf(strDeptIDs = "", "", " And Instr([8],','||A.开单部门ID||',')>0")
    strCond = strCond & IIf(strChargeType = "", "", " And Instr([10],','''||A.收费类别||''',')>0")   '34260
    strCond = strCond & " And A.门诊标志 In (2,3) "
    '结算要求：记录性质,记录状态,NO、序号、收费类别、收费细目ID,收费名称、计算单位、开单部门、规格、产地、数量、价格、金额、医生,
    '          发生时间,登记时间,婴儿费,医保项目编码、保险大类ID、保险项目否、是否上传,是否急诊
    '注意：由于结算只能针对有保险项目编码的,所以在与保险支付项目连接时不用(+)
    '   数次为零指按费别重算后产生的打折冲减记录,这类单据明细不上传
    
    '临时更改：结帐作废后该SQL不对,单价暂时改为"金额/数量"
 
    strTable = IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录"), "住院费用记录 A")
    strTable = "" & _
    "       Select Mod(A.记录性质, 10) As 记录性质, decode(A.记录状态,2,2,1) as 标志, max( A.记录状态) as 记录状态, A.NO, Nvl(A.价格父号, 序号) As 序号, A.门诊标志, A.病人id," & vbNewLine & _
    "               A.主页id as 主页id, Nvl(A.婴儿费, 0) As 婴儿费, A.开单人 As 医生, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位," & vbNewLine & _
    "              A.保险编码, Nvl(A.保险大类id, 0) As 保险大类id, Avg(Nvl(A.付数, 1) * A.数次) As 数量," & vbNewLine & _
    "              Sum(A.标准单价*decode(sign(A.记录性质-10),1,0,1)) As 标准单价, Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) As 金额, A.发生时间," & vbNewLine & _
    "              A.登记时间, Min(Nvl(A.是否上传, 0)) As 是否上传, Nvl(A.是否急诊, 0) As 是否急诊,Nvl(A.保险项目否, 0) As 保险项目否, Max(Decode(A.记录性质,2,A.摘要,Null)) As 摘要,A.费用类型" & vbNewLine & _
    "       From " & strTable & " , 收入项目 B" & vbNewLine & _
    "       Where A.记录状态 <> 0 And A.记帐费用 = 1  And A.主页ID Is Not Null  And A.收入项目id = B.ID And A.数次 <> 0" & vbNewLine & _
    "             And Not Exists (Select 1 From " & IIf(blnDateMoved, zlGetFullFieldsTable("住院费用记录", , , , "C"), "住院费用记录 C") & ", 病人结帐记录 D" & vbNewLine & _
    "                             Where c.No = a.No And Mod(c.记录性质,10) = Mod(a.记录性质,10) And c.序号 = a.序号 And c.结帐id = d.Id" & vbNewLine & _
    "                                   And Nvl(d.结算状态, 0) = 1) " & strCond & strWherePage & vbNewLine & _
    "       Group By Mod(A.记录性质, 10),decode(A.记录状态,2,2,1), A.NO, Nvl(A.价格父号, 序号), A.门诊标志, A.病人id ,A.主页id ," & vbNewLine & _
    "                Nvl(A.婴儿费, 0), A.开单人, A.开单部门id, A.收费类别, A.收费细目id, A.计算单位, A.保险编码," & vbNewLine & _
    "                Nvl(A.保险大类id, 0), A.发生时间, A.登记时间, Nvl(A.是否急诊, 0), Nvl(A.保险项目否, 0),A.费用类型" & vbNewLine & _
    "       Having Sum(Nvl(A.实收金额, 0)) - Sum(Nvl(A.结帐金额, 0)) <> 0　"
    
    strSQL = "" & _
    " Select A.记录性质, A.记录状态, A.NO, A.序号, A.门诊标志, A.病人id, A.主页id, A.婴儿费, C.项目编码 As 医保项目编码," & vbNewLine & _
    "       A.保险编码, A.保险大类id, A.收费类别, A.收费细目id, Nvl(E.名称, B.名称) As 收费名称, A.计算单位," & vbNewLine & _
    "       X.名称 As 开单部门, B.规格, B.产地, A.数量, A.标准单价 As 价格, A.金额," & vbNewLine & _
    "       A.医生, A.发生时间, A.登记时间, A.是否上传, A.是否急诊, A.保险项目否, A.摘要,A.费用类型" & vbNewLine & _
    " From ( " & strTable & ") A, 收费项目目录 B, 保险支付项目 C, 收费项目别名 E,部门表 X" & vbNewLine & _
    " Where A.收费细目id = B.ID And B.ID = C.收费细目id" & IIf(blnRelation, "(+)", "") & " And C.险类" & IIf(blnRelation, "(+)", "") & " = [5] And A.开单部门id = X.ID " & vbNewLine & _
        IIf(strClass = "", "", " And Instr([9],','''||Nvl(A.费用类型,Nvl(B.费用类型,'无'))||''',')>0") & _
    "      And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = " & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1)
    Set GetZYBalance_Insure = zlDatabase.OpenSQLRecord(strSQL, "获取医保住院费用信息集", lng病人ID, "," & strTime & ",", dtBeginDate, dtEndDate, int险类, "," & strBaby & ",", "," & strItem & ",", "," & strDeptIDs & ",", "," & strClass & ",", "," & strChargeType & ",")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetFromIDToBalanceData(ByVal lng结帐ID As Long, ByVal blnNOMoved As Boolean, _
    ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结帐ID来获取结算数据
    '入参:lng结帐ID-结帐ID
    '     blnNoMoved-是否已经转移到后备表中
    '出参:rsOutBalance-结帐数据
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2015-01-08 15:32:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String
    On Error GoTo errHandle
    
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡;6-误差费
     strSQL = "" & _
       "   Select  A.ID, " & _
       "        Case when Mod(A.记录性质,10)=1 then 1  " & _
       "             when nvl(M.性质,0)=3 or nvl(M.性质,0)=4  then 2 " & _
       "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
       "             when J.结算方式 is not null   then  4 " & _
       "             when nvl(M.性质,0)=9 then 6 " & _
       "             else 0 end as 类型, " & _
       "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交,A.摘要, " & _
       "        A.卡类别ID,A.结算卡序号, " & _
       "        A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
       "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
       "        Decode(C.卡号密文,NULL,0,1) as  是否密文," & _
       "        C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志, " & _
       "        decode(nvl(M.性质,0),3,1,4,1,0) as 医保,0 as 消费卡id,nvl(M.性质,0) as 结算性质" & _
       "   From  病人预交记录 A ,医疗卡类别 C,一卡通目录 J,结算方式 M" & _
       "   Where A.结帐ID= [1] And A.结算方式=M.名称(+) " & _
       "         And A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
       "         And nvl(A.结算卡序号,0)=0"
    
    strSQL = strSQL & " Union ALL " & _
        " Select a.Id," & vbNewLine & _
        "       Case" & vbNewLine & _
        "         When Mod(a.记录性质, 10) = 1 Then" & vbNewLine & _
        "          1" & vbNewLine & _
        "         When Nvl(m.性质, 0) = 3 Or Nvl(m.性质, 0) = 4 Then" & vbNewLine & _
        "          2" & vbNewLine & _
        "         When Nvl(a.卡类别id, 0) <> 0 Then" & vbNewLine & _
        "          3" & vbNewLine & _
        "         When j.结算方式 Is Not Null Then" & vbNewLine & _
        "          4" & vbNewLine & _
        "         When Nvl(m.性质, 0) = 9 Then" & vbNewLine & _
        "          6" & vbNewLine & _
        "         Else" & vbNewLine & _
        "          0" & vbNewLine & _
        "       End As 类型, Mod(a.记录性质, 10) As 记录性质, a.结算方式, a.冲预交, a.摘要, a.卡类别id, a.结算卡序号, a.结算号码, a.卡号, a.交易流水号," & vbNewLine & _
        "       Nvl(c.自制卡, 0) As 自制卡, Nvl(c.是否退现, 0) As 是否退现, Nvl(c.是否全退, 0) As 是否全退, Decode(c.是否密文, Null, 0, 1) As 是否密文," & vbNewLine & _
        "       c.名称 As 卡类别名称, a.交易说明, a.结算序号, a.校对标志, Decode(Nvl(m.性质, 0), 3, 1, 4, 1, 0) As 医保, 0 As 消费卡id," & vbNewLine & _
        "       Nvl(m.性质, 0) As 结算性质" & vbNewLine & _
        "From 病人预交记录 A, 消费卡类别目录 C, 一卡通目录 J, 结算方式 M" & vbNewLine & _
        "Where a.结帐id = [1] And a.结算方式 = m.名称(+) And a.结算方式 = j.结算方式(+) And a.结算卡序号 = c.编号 And Nvl(a.卡类别id, 0) = 0 And Mod(a.记录性质,10) =1"

          
    strSQL = strSQL & " Union ALL " & _
       "   Select A.ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要,A.卡类别ID,A.结算卡序号," & _
       "        A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
       "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
       "        nvl(M.是否密文,0) as  是否密文," & _
       "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,M1.性质 as 结算性质" & _
       "   From 病人预交记录 A ,病人卡结算记录 B, " & _
       "        消费卡类别目录 M,结算方式 M1 " & _
       "   Where  a.Id = b.结算Id " & _
       "        And a.结算卡序号 = m.编号 And A.结算方式=M1.名称(+) " & _
       "        And A.结帐ID = [1] and Mod(A.记录性质,10)<>1 "
       
      strSQL = "" & _
      "   Select A.类型,a.记录性质,a.结算方式,a.摘要,a.卡类别ID,a.卡类别名称,a.自制卡,a.结算卡序号,a.结算号码,a.卡号,a.交易流水号,a. 交易说明,a.结算序号,a.校对标志,a.医保,a.消费卡id," & _
      "         max(A.是否密文) as 是否密文,max(A.是否全退) as 是否全退,max(a.是否退现) as 是否退现, nvl(sum(a.冲预交),0) as 冲预交,Max(A.结算性质) as 性质" & _
      "   From (" & strSQL & ") A " & _
      "   Group by A.类型,a.记录性质,a.结算方式,a.摘要,a.卡类别ID,a.卡类别名称,a.自制卡,a.结算卡序号,a.结算号码,a.卡号,a.交易流水号,a. 交易说明,a.结算序号,a. 校对标志,a.医保,a.消费卡id" & _
      "   Order by 类型"
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "病人预交记录", "H病人预交记录")
        strSQL = Replace(strSQL, "病人卡结算记录", "H病人卡结算记录")
    End If
    
    Set rsOutBalance = zlDatabase.OpenSQLRecord(strSQL, "获取结帐数据", lng结帐ID)
    zlGetFromIDToBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function zlGetFormerBalanceID(strNO As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取原结帐单据的ID
    '返回:获取成功,返回原结帐ID,否则返回0
    '编制:刘兴洪
    '日期:2015-01-26 09:51:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ID From 病人结帐记录 Where 记录状态 IN(1,3) And NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取原始结帐ID", strNO)
    If Not rsTmp.EOF Then zlGetFormerBalanceID = rsTmp!ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlFromIDGetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取收费结算信息
    '入参:bytType-查找类型:0-根据结帐ID查找;1-根据单据号来获取结算方式
    '     strValue-要查找的值(为0时,结帐ID, 2时为结帐单据号)
    '     blnDel-作废结算:true-查作废结算;false-非作废结算
    '返回:收费结算的相关信息集
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    '编制:刘兴洪
    '日期:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    On Error GoTo errHandle
    
    strTable = IIf(blnHistory, "H", "") & "病人预交记录"
    Select Case bytType
    Case 0  '0-根据结帐ID查找
        strWhere = " And  A.结帐ID= [1]"
    Case 1 '根据单据号来获取结算数据
        strTable1 = "" & _
        "   Select distinct ID  " & _
        "   From 病人结帐记录 M " & _
        "   Where m.no=[2]  And 记录状态 in (1,3) And nvl(M.结算状态,0)<>1"
        strTable1 = ",(" & strTable1 & ") Q1"
        
        If blnHistory Then strTable1 = Replace(strTable1, "病人结帐记录", "H病人结帐记录")
        strWhere = " And A.结帐ID=Q1.ID"
    Case Else
        Exit Function
    End Select
    
    If blnDel Then
        '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        strSQL = "" & _
        "   Select  A.ID,decode(A.记录状态,2,A.结帐ID,NULL) as 结帐ID," & _
        "        Case when Mod(A.记录性质,10)=1 then 1  " & _
        "             when B.名称 is not null then  2 " & _
        "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
        "             when J.结算方式 is not null   then  4 " & _
        "             else 0 end as 类型, " & _
        "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
        "        decode(A.记录状态,2,A.摘要,NULL) as 摘要,decode(A.记录状态,2,1,0) as 退费," & _
        "        A.卡类别ID,A.结算卡序号, " & _
        "        decode(A.记录状态,2,A.结算号码,NULL) as 结算号码,decode(A.记录状态,2,A.卡号,NULL) as 卡号, " & _
        "        decode(A.记录状态,2,A.交易流水号,NULL) as 交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
        "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
        "        Decode(C.卡号密文,NULL,0,1) as  是否密文," & _
        "        C.名称 as 卡类别名称,decode(A.记录状态,2,A.交易说明,NULL) as 交易说明,A.结算序号,decode(A.记录状态,2,A.校对标志,0) as 校对标志, " & _
        "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
        "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
        "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
        "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
        "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
        "         And nvl(A.结算卡序号,0)=0 " & strWhere
        strSQL = strSQL & " Union ALL " & _
        "   Select A.ID,decode(A.记录状态,2,A.结帐ID,NULL) as 结帐ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要," & _
        "        decode(A.记录状态,2,1,0) as 退费,A.卡类别ID,A.结算卡序号," & _
        "        decode(A.记录状态,2,A.结算号码,NULL) as 结算号码,decode(A.记录状态,2,B.卡号,NULL) as 卡号, " & _
        "        decode(A.记录状态,2,B.交易流水号,NULL) as 交易流水号,nvl(M.自制卡,0) as 自制卡, " & _
        "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
        "        nvl(M.是否密文,0) as  是否密文," & _
        "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,nvl(q.性质,1) as 结算性质" & _
        "   From  " & strTable & " A ,病人卡结算记录 B, " & _
        "        消费卡类别目录 M ,结算方式 q " & strTable1 & _
        "   Where  a.Id = b.结算Id  And a.结算卡序号 = m.编号  " & _
        "         and Mod(A.记录性质,10)<>1 and A.结算方式=q.名称(+) " & strWhere
        
        strSQL = "" & _
        "   Select /*+ Rule */ max(结帐id) as 结帐id,类型,max(退费) as 退费,记录性质,结算方式,Max(摘要) as 摘要,卡类别ID,卡类别名称,max(自制卡) as 自制卡,结算卡序号, " & _
        "         max(结算号码) as 结算号码,max(卡号) as 卡号,max(交易流水号) as 交易流水号, max(交易说明) as 交易说明, " & _
        "         结算序号,max(校对标志) as 校对标志,医保,消费卡id,结算性质," & _
        "         max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
        "   From (" & strSQL & ") " & _
        "   Group by 类型, 记录性质,结算方式,卡类别ID,卡类别名称,结算卡序号,结算序号,医保,消费卡id,结算性质 having  sum(冲预交) <>0"
        Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue)
        Exit Function
    End If
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    strSQL = "" & _
    "   Select /*+ Rule */ A.ID,A.结帐ID," & _
    "        Case when Mod(A.记录性质,10)=1 then 1  " & _
    "             when B.名称 is not null then  2 " & _
    "             when nvl(A.卡类别ID,0)<>0  then  3 " & _
    "             when J.结算方式 is not null   then  4 " & _
    "             else 0 end as 类型, " & _
    "        Mod(A.记录性质,10) as 记录性质,A.结算方式,A.冲预交," & _
    "        A.摘要,decode(A.记录状态,2,1,0) as 退费," & _
    "        A.卡类别ID,A.结算卡序号, " & _
    "        A.结算号码,A.卡号,A.交易流水号,nvl(C.是否自制,0) as 自制卡, " & _
    "        nvl(C.是否退现,0) as 是否退现,nvl(C.是否全退,0) as 是否全退, " & _
    "        Decode(C.卡号密文,NULL,0,1) as  是否密文," & _
    "        C.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志, " & _
    "        decode(B.名称,Null,0,1) as 医保,0 as 消费卡id,nvl(q.性质,1) as 结算性质" & _
    "   From " & strTable & " A ,医疗卡类别 C,一卡通目录 J,结算方式 q," & _
    "        (Select 名称 From 结算方式 where 性质 in (3,4)) B " & strTable1 & _
    "   Where A.结算方式=J.结算方式(+) And A.卡类别ID=C.ID(+) " & _
    "         And A.结算方式=B.名称(+) and A.结算方式=q.名称(+) " & _
    "         And nvl(A.结算卡序号,0)=0 " & strWhere
       
    strSQL = strSQL & " Union ALL " & _
    "   Select /*+ Rule */ A.ID,A.结帐ID,5 as  类型,Mod(A.记录性质,10) as 记录性质,A.结算方式,-1*nvl(b.应收金额,0) as 冲预交,A.摘要," & _
    "        decode(A.记录状态,2,1,0) as 退费,A.卡类别ID,A.结算卡序号," & _
    "        A.结算号码,B.卡号,B.交易流水号,nvl( M.自制卡,0) as 自制卡, " & _
    "        nvl( M.是否退现,0) as 是否退现,nvl(M.是否全退,0) as 是否全退, " & _
    "        nvl(M.是否密文,0) as  是否密文," & _
    "        M.名称 as 卡类别名称,A.交易说明,A.结算序号,A.校对标志,0 as 医保,B.消费卡id,nvl(q.性质,1) as 结算性质" & _
    "   From  " & strTable & " A ,病人卡结算记录 B, " & _
    "        消费卡类别目录 M ,结算方式 q " & strTable1 & _
    "   Where  a.Id = b.结算Id  And a.结算卡序号 = m.编号  " & _
    "         and Mod(A.记录性质,10)<>1 and A.结算方式=q.名称(+) " & strWhere
    gstrSQL = "" & _
    "   Select  结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质," & _
    "         max(是否密文) as 是否密文,max(是否全退) as 是否全退,max(是否退现) as 是否退现 , nvl(sum(冲预交),0) as 冲预交" & _
    "   From (" & gstrSQL & ") " & _
    "   Group by 结帐ID,类型,退费,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id,结算性质"
    Set zlFromIDGetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "获取收费结算方式", Val(strValue), strValue)
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetPatiRsByUnit(ByVal lng病区ID As Long, ByVal lng病人ID As Long, _
    bln取适用病人 As Boolean, ByVal bln取剩余款 As Boolean, _
    Optional ByVal bln含预出院 As Boolean = False, _
    Optional ByVal int提取范围 As Integer = -1, _
    Optional ByVal bln包含出院病人 As Boolean = False, _
    Optional ByVal strOutBeginDate As String = "", _
    Optional ByVal strOutEndDate As String = "") As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前病区的病人信息数据
    '入参:lng病区ID-病区ID
    '     lng病人ID-包含此病人ID
    '     bln取适用病人-True时:返回的病人信息集中是否包含"适用病人"信息,即通过"zl_PatiWarnScheme"函数返回值
    '                   False时:返回NULL
    '     bln取剩余款-true时:返回的病人信息集中是否包含"剩余款",即:预交余额-费用余额+预结算结果)
    '                 False时:返回NULL
    '     bln含预出院-是否包含预出院病人
    '     int提取范围:-1:所有病人包括婴儿
    '                 0-只包含病人
    '                 1.只包含婴儿
    '     bln包含出院病人-是否包含出院病人
    '     strOutBeginDate:出院开始时间(bln包含出院病人＝true时有效),格式:yyyy-mm-dd hh24:mi:ss
    '     strOutEndDate:出院结束时间(bln包含出院病人＝true时有效),格式:yyyy-mm-dd hh24:mi:ss
    '返回:成功,返回病人信息集返回true,否则返回Nothing
    '编制:刘兴洪
    '日期:2015-07-08 10:59:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intBedLen As Integer, strTable As String
    Dim dtOutBegin As Date, dtOutEnd As Date
    Dim strWithTable As String
    Dim strFields As String '字段
    Dim strWhere As String  '条件
    
    On Error GoTo errH
    
    intBedLen = GetMaxBedLen(lng病区ID, False)
    strFields = ""
    
    
    If bln取剩余款 Then
        strFields = strFields & "," & _
        "  Nvl(E.预交余额,0)-Nvl(E.费用余额,0)+ " & _
        "  Decode(B.险类,Null,0,(Select Nvl(Sum(金额),0) From 保险模拟结算 F Where B.病人ID=F.病人ID And B.主页ID=F.主页ID)) as 剩余款"
    Else
        strFields = strFields & ",NULL as 剩余款"
    End If
    strFields = strFields & "," & IIf(bln取适用病人, "zl_PatiWarnScheme(A.病人ID,B.主页ID)", "NULL") & " as 适用病人"
    
    If int提取范围 = 1 Then '只包含婴儿
        strWhere = " And Exists(select 1 from 病人新生儿记录 Z Where z.病人id=b.病人ID And z.主页id=b.主页ID)"
    End If
    
    
    
    strTable = "" & _
    " Select a.病人ID" & vbNewLine & _
    " From 病人信息 A, 病案主页 B,在院病人 R" & vbNewLine & _
    " Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人ID=R.病人ID " & _
    "       And A.当前病区ID=R.病区ID And (R.病区id = [1] or B.婴儿病区ID = [1]) " & _
            IIf(bln含预出院, "", " And Nvl(b.状态,0)<>3") & vbNewLine & _
    " Union" & vbNewLine & _
    " Select 0+[2] as 病人ID From Dual"
    
    dtOutBegin = CDate("1991-01-01")
    dtOutEnd = CDate("1991-01-01")
    If bln包含出院病人 Then
         '出院病人时间范围
        If strOutBeginDate = "" Then
            strOutBeginDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00")
            strOutEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD 23:59:59")
        End If
        dtOutBegin = CDate(strOutBeginDate)
        dtOutEnd = CDate(strOutEndDate)
        strTable = strTable & _
        " Union" & vbNewLine & _
        " Select a.病人id" & vbNewLine & _
        " From 病人信息 A, 病案主页 B" & vbNewLine & _
        " Where a.病人id = b.病人id And a.主页id = b.主页id  " & _
        "       And (b.当前病区id + 0 = [1] Or b.婴儿病区id + 0 = [1]) " & _
        "       And B.出院日期 Between [3] And [4]"
    End If
        
        
    strWithTable = "" & _
    " With T病人信息 as ( " & _
    "       Select A.病人ID,B.主页ID,nvl(B.姓名,A.姓名) as 姓名,B.住院号,LPAD(B.出院病床," & intBedLen & ",' ') as 床号," & _
    "           Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额," & _
    "           zl_PatiDayCharge(A.病人ID) as 当日额," & _
    "           E.预交余额,E.费用余额,B.住院医师,nvl(B.费别,A.费别) as 费别,D.名称 as 护理等级,C.名称 as 科室," & _
    "           c.id as 科室id,B.入院日期,B.出院日期,B.病人类型,nvl(B.性别,A.性别) as 性别," & _
    "           nvl(B.年龄,A.年龄) as 年龄,b.审核标志,B.婴儿科室ID,B.婴儿病区ID,nvl(B.险类,0) as 险类,B.状态," & _
    "           nvl(b.医疗付款方式,A.医疗付款方式) as 医疗付款方式,B.病人性质 as 病人性质,B.住院医师 As 开单人,A.当前科室ID As 开单科室ID,M.名称 as 开单科室名称" & strFields & _
    "       From 病人信息 A,病案主页 B,部门表 C,部门表 M,收费项目目录 D,病人余额 E,(" & strTable & ") F" & _
    "       Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.护理等级ID=D.ID(+)" & _
    "           And B.出院科室ID=C.ID And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
    "           And A.当前科室ID=M.id(+) And A.病人ID=E.病人ID(+) And E.性质(+)=1 And E.类型(+) = 2 And A.病人ID=F.病人ID " & strWhere & _
    "       Order by 床号)"
    
    strSQL = "" & vbCrLf & _
    " Select A.病人id, A.主页id, A.姓名, A.住院号, A.床号, A.担保额,A.当日额,A.预交余额,A.费用余额, A.剩余款,  A.适用病人, A.险类,b.名称 as 保险类别, " & _
    "       A.住院医师, A.费别, A.护理等级, A.科室,A.科室id," & _
    "       A.入院日期, A.出院日期, A.病人类型, A.性别,A.年龄, A.审核标志,A.婴儿科室ID,A.婴儿病区ID,Null as 婴儿姓名,Null as 婴儿序号," & _
    "       A.状态,A.医疗付款方式,A.病人性质,A.开单人,A.开单科室ID,A.开单科室名称" & _
    " From T病人信息 A,保险类别 B " & _
    " Where A.险类=B.序号(+)"
    
    If int提取范围 = 1 Or int提取范围 = -1 Then '包含婴儿
        strSQL = IIf(int提取范围 = 1, "", strSQL & " Union  ALL ") & vbCrLf & _
        " Select a.病人id,a.主页id,a.姓名,a.住院号,a.床号,a.担保额,A.当日额,a.预交余额,a.费用余额,a.剩余款, a.适用病人,a.险类,C.名称 as 保险类别," & _
        "        a.住院医师,a.费别,a.护理等级,a.科室,a.科室id," & _
        "        a.入院日期, a.出院日期, a.病人类型, a.性别,a.年龄, a.审核标志,a.婴儿科室ID,a.婴儿病区ID,b.婴儿姓名,B.序号 AS 婴儿序号," & _
        "        A.状态,A.医疗付款方式,A.病人性质,A.开单人,A.开单科室ID,A.开单科室名称" & _
        " From T病人信息 A,病人新生儿记录 B,保险类别 C" & _
        " Where A.病人id=b.病人id and A.主页ID=b.主页id and  A.险类=C.序号(+)"
    End If
    strSQL = strWithTable & vbCrLf & strSQL & vbCrLf
    strSQL = "" & _
    " Select * " & vbCrLf & _
    " From (" & strSQL & ")" & vbCrLf & _
    " Order by 床号,NVL(婴儿序号,0)"
    Set GetPatiRsByUnit = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息列表", lng病区ID, lng病人ID, dtOutBegin, dtOutEnd)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal bytFunc As Byte, ByVal lng结帐ID As Long, lng病人ID As Long) As Boolean
    '功能:将住院信息写入卡中
    '入参：
    '    frmMain - 调用窗体
    '    lngModul - 模块号
    '    strPrivs - 权限串
    '    objSquareCard - 医疗卡对象
    '    bytFun - 0:门诊，1:住院
    Dim strExpend As String, lng结算序号 As Long
    
    If lng病人ID = 0 Or lng结帐ID = 0 Then Exit Function
    Err = 0: On Error GoTo errH:
    '问题:56615
    If bytFunc = 0 Then
        If InStr(strPrivs, ";门诊信息写卡;") = 0 Then Exit Function
        'Public Function zlMzInforWriteToCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, _
            ByVal lng病人ID As Long, _
            ByVal lngBalanceID As Long, _
            Optional ByRef strExpend As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:写门诊信息接口
            '    frmMain Object  In  调用的主窗体
            '    lngModule   Long    In  调用的模块号
            '    lngCardTypeID   Long    In  传入写卡类别ID:
            '           1)传入刷卡的类别ID
            '           2)传入零时,需要选择某个卡类别ID
            '    lng病人ID   Long    In  病人ID
            '    lngBalanceID    Long    In  结算序号(某次结算的序号)
            '    strExpend   String  In/Out  XML,暂无,待以后扩展
            ' 函数返回    True:调用成功,False:调用失败
            '调用时机:
            '         医疗卡类别.是否写卡=1才调用
        Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng病人ID, lng结帐ID, strExpend)
        '门诊结算数据没有结算序号，所以直接传结帐ID
    Else
        If InStr(strPrivs, ";住院信息写卡;") = 0 Then Exit Function
        'Public Function zlZyInforWriteToCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, _
            ByVal lng病人ID As Long, _
            ByVal lngBalanceID As Long, _
            Optional ByRef strExpend As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:写住院信息接口
            '    frmMain Object  In  调用的主窗体
            '    lngModule   Long    In  调用的模块号
            '    lngCardTypeID   Long    In  传入写卡类别ID:
            '           1)传入刷卡的类别ID
            '           2)传入零时,需要选择某个卡类别ID
            '    lng病人ID   Long    In  病人ID
            '    lngBalanceID    Long    In  结帐ID(可以不传入)
            '    strExpend   String  In/Out  XML,暂无,待以后扩展
            ' 函数返回    True:调用成功,False:调用失败
            '调用时机:
            '        医疗卡类别.是否写卡=1才调用
        Call objSquareCard.zlZyInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng病人ID, lng结帐ID, strExpend)
    End If
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function CreatePlugIn(ByVal lngModule As Long, _
    Optional ByVal int场合 As Integer) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugIn = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    If gobjPlugIn Is Nothing Then
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    End If
    If gobjPlugIn Is Nothing Then Exit Function
    
    Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngModule, int场合)
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "Initialize")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    CreatePlugIn = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    objErr.Clear
End Sub

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
    ByVal intTYPE As Integer, ByVal intMode As Integer, _
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
    If objPlugIn.CheckChargeItem(lngSys, lngModule, intTYPE, intMode, rsDetail, strExpend) = False Then
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

Public Sub CreatePublicExpenseObject(ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
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



Public Function CreatePublicExpenseBillOperation() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共费用部件
    '入参:
    '编制:
    '日期:
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpenseBillOperation Is Nothing Then
        Set gobjPublicExpenseBillOperation = CreateObject("zlPublicExpense.clsBillOperation")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        CreatePublicExpenseBillOperation = True
        Exit Function
    End If
    If gobjPublicExpenseBillOperation Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If gobjPublicExpenseBillOperation.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Function
    End If
    CreatePublicExpenseBillOperation = True
End Function

Public Function zlShowMsgBox(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intTYPE As Integer) As VbMsgBoxResult
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示消息框
    '入参:frmMain-调用的主窗体
    '     strInfo=提示信息,需要自已处理换行,可用"^"表示回车,">"表示缩进
    '     intType=消息框类型=0(缺省)=MsgBox类型,1-皮试类型
    '     blnNoAsk="intType=0"时有效，表示是否只显示一个确定按钮,不以询问方式显示是和否。
    '返回:
    '    intType=0：vbIgnore=是且不再提示,vbCancel=否且不再提示,vbYes=是,vbNo=否
    '    intType=1：vbYes=阳性,vbNo=阴性,vbCancel=取消
    '编制:刘兴洪
    '日期:2017-11-08 11:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjPublicExpense Is Nothing Then
        Call CreatePublicExpenseObject(glngModul)
    End If
    If gobjPublicExpense Is Nothing Then GoTo GoMsgbox:

    Err = 0: On Error Resume Next
    zlShowMsgBox = gobjPublicExpense.zlShowMsgBox(frmMain, strInfo, blnNoAsk, intTYPE)
    If Err.Number = 438 Then GoTo GoMsgbox
    If Err <> 0 Then zlShowMsgBox = vbCancel
    Err = 0: On Error GoTo 0
    Exit Function
GoMsgbox:
    '直接使用Msgbox消费框
    If blnNoAsk Then
        MsgBox strInfo, vbInformation + vbOKOnly, gstrSysName
        zlShowMsgBox = vbOK: Exit Function
    End If
    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
         zlShowMsgBox = vbIgnore
    Else
         zlShowMsgBox = vbCancel
    End If
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

Public Sub zlShowThreeSwapErrInfor(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方转账检查与代扣业务出错提示
    '编制:冉俊明
    '时间:2014-12-2
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "调用转帐检查交易失败！"
        Else
            strValue = vbCrLf & "调用转帐交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function zlGetTimeDataFromTimes(ByVal str主页Ids As String, ByRef int主页ID As Integer, intInsure As Integer, _
    Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据住院次数信息，获取主页ID,险类及险类名称
    '入参:str主页IDs:格式:主页ID|险类|险类名称
    '出参:int主页ID
    '     intInsure-险类
    '     strInsureName-险类名称
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-13 11:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    On Error GoTo errHandle
    varTemp = Split(str主页Ids & "||||", "|")
    int主页ID = Val(varTemp(0))
    intInsure = Val(varTemp(1))
    strInsureName = Trim(varTemp(2))
    zlGetTimeDataFromTimes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlGetAllTims(ByVal str主页Ids As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据选择的主页IDs,来获取只包含主页ID的串
    '入参:str主页IDs:格式:主页ID|险类|险类名称,主页ID1|险类1|险类名称1,....
    '出参:
    '返回:只返回所涉及的住院次数
    '编制:刘兴洪
    '日期:2017-11-13 11:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, int主页ID As Integer, intInsure As Integer
    Dim strAllTims As String, i As Long
    
    On Error GoTo errHandle
    
    
    varTemp = Split(str主页Ids, ",")
    For i = 0 To UBound(varTemp)
        Call zlGetTimeDataFromTimes(varTemp(i), int主页ID, intInsure)
        strAllTims = strAllTims & "," & int主页ID
    Next
    If strAllTims <> "" Then strAllTims = Mid(strAllTims, 2)
    zlGetAllTims = strAllTims
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function TruncStringEx(ByVal strValue As String, Optional blnReverse As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:字符特殊处理
    '入参:strValue-字符值
    '     blnReverse-颠倒
    '返回:格式化的串
    '编制:刘兴洪
    '日期:2017-11-13 09:53:05
    '说明:此过程为临时出来，有空后，不应该这么处理
    '    blnReverse=False
    '         1.将","替换成"～，～"
    '         2.将"|"替换成"～｜～"
    '    blnReverse=true
    '         1.将"～，～"替换成","
    '         2.将"～｜～"替换成"|"
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If blnReverse Then
        strValue = Replace(strValue, "～，～", ",")
        strValue = Replace(strValue, "～｜～", "|")
    Else
        strValue = Replace(strValue, ",", "～，～")
        strValue = Replace(strValue, "|", "～｜～")
    End If
    TruncStringEx = strValue
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ZlShowBillFormat(ByVal bytInvoiceKind As Byte, lblFormat As Label, ByVal intFormat As Integer)
    '功能：显示票据格式名称
    '入参：
    '   bytInvoiceKind - 0-住院医疗费收据,1-门诊医疗费收据
    '   lblFormat - 显示票据格式的标签对象
    '   intFormat - 票据格式序号
    '返回：票据格式的名称
    Dim strFormatName As String
    
    On Error GoTo errHandler
    strFormatName = ZlGetBillFormat(bytInvoiceKind, intFormat)
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

Public Function ZlGetBillFormat(ByVal bytInvoiceKind As Byte, ByVal intFormat As Integer) As String
    '功能：获取票据格式名称
    '入参：
    '   bytInvoiceKind - 0-住院医疗费收据,1-门诊医疗费收据
    '   intFormat - 票据格式序号
    '返回：票据格式的名称
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strRptName As String
    
    On Error GoTo errHandler
    If bytInvoiceKind = 0 Then
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1137"
    Else
        strRptName = "ZL" & glngSys \ 100 & "_BILL_1137_2"
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
    '     rsSaveItems=当前保存的项目集，字段(字段 :病人ID，主页ID,单据序号, 序号,价格父号,收费细目ID，收入项目id，付数 ，数次，标准单价，应收金额 ，
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

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng病人ID, ByVal lng主页ID As Long, ByVal bln门诊 As Boolean, _
                                    ByVal int记录性质 As Integer, ByVal strNos As String)
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


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef rsItems As ADODB.Recordset, Optional blnBill As Boolean) As Boolean
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
    Dim int序号 As Integer
    
    On Error GoTo errHandle
    
    Set rsItems = Nothing
    
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
    int序号 = 0
    For Each objBillDetail In objBills.Details
        If objBillDetail.数次 <> 0 Then
            int价格父号 = 0
            For Each objBillIncome In objBillDetail.InComes
              int序号 = int序号 + 1 '当前记录序号
               rsItems.AddNew
               rsItems!病人ID = IIf(blnBill, objBills.Details(int序号).病人ID, objBills.病人ID)
               rsItems!主页ID = IIf(blnBill, objBills.Details(int序号).主页ID, objBills.主页ID)
               rsItems!单据序号 = 1
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
               rsItems!开单部门ID = objBills.开单部门ID
               rsItems!开单人 = objBills.开单人
               rsItems.Update
              If int价格父号 = 0 Then int价格父号 = int序号
            Next     '每一行收费项目
        End If
    Next
    
    zlGetSaveDataItems_Plugin = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlCallReturnCashCheckInterface(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, ByVal str卡号 As String, ByVal strBalances As String, ByVal dblMoney As Double, _
    ByVal str交易流水号 As String, ByVal str交易说明 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用"zlReturnCashCheck"接口
    '入参:
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-09 14:21:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String
    If gobjSquare Is Nothing Then Call CreateSquareCardObject(frmMain, lngModule)
    If gobjSquare Is Nothing Then Exit Function
    If gobjSquare.objSquareCard Is Nothing Then Exit Function
    Err = 0: On Error GoTo errHandle
    With gobjSquare.objSquareCard
        If .zlReturnCashCheck(frmMain, lngModule, lngCardTypeID, str卡号, strBalances, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then
            MsgBox "接口检查退现失败，无法退现！", vbInformation, gstrSysName
           Exit Function
        End If
    End With
    zlCallReturnCashCheckInterface = True
    Exit Function
errHandle:
    If Err.Number = 438 Then
        MsgBox "缺失一卡通“zlReturnCashCheck”接口，不允退现，请与系统管理员联系!", vbOKOnly, gstrSysName
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 Public Sub zlCloseSquareCardObject()
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
         Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub
Public Function zlCloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    zlCloseWindows = blnChildren And (Forms.Count = 0)
End Function
Public Function zlReleaseResources() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:释放资源
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-13 10:30:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '进程数为0时，才放资源
    If glngInstanceCount > 0 Then Exit Function
    
    Call zlCloseSquareCardObject  '释放CardSquare对象
    
    Call zlCloseWindows   '关闭窗体
    
    Err = 0: On Error Resume Next
    If Not gclsInsure Is Nothing Then Set gclsInsure = Nothing
    If Not gobjBillPrint Is Nothing Then Set gobjBillPrint = Nothing
    If Not gobjTax Is Nothing Then Set gobjTax = Nothing
    If Not grsABCNum Is Nothing Then Set grsABCNum = Nothing
    If Not gobjPati Is Nothing Then Set gobjPati = Nothing
    If Not grs收费类别 Is Nothing Then Set grs收费类别 = Nothing
    If Not gobjPlugIn Is Nothing Then Set gobjPlugIn = Nothing
    If Not gobjPublicDrug Is Nothing Then Set gobjPublicDrug = Nothing
    If Not gobjPublicExpense Is Nothing Then Set gobjPublicExpense = Nothing
    If Not gobjPublicExpenseBillOperation Is Nothing Then Set gobjPublicExpenseBillOperation = Nothing
    If Not grs医疗付款方式 Is Nothing Then Set grs医疗付款方式 = Nothing
    If Not grsOneCard Is Nothing Then Set grsOneCard = Nothing
    If Not grsSquareType Is Nothing Then Set grsSquareType = Nothing
    If Not gobjXml Is Nothing Then Set gobjXml = Nothing
    If Not gobjKernel Is Nothing Then Set gobjKernel = Nothing
    If Not gfrmMain Is Nothing Then Set gfrmMain = Nothing
    zlReleaseResources = True
End Function

Public Function zlSelectChargePatiFromInputName(ByVal frmMain As Object, ByVal strPrivsOpt As String, ByRef strInput As String, ByVal bln所有病区 As Boolean, ByVal strUnitIDs As String, _
    ByVal intOutDay As Integer, ByRef lng病人ID_Out As Long, Optional strErrMsg_out As String, Optional lngHwnd As Long, Optional lngHeight As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入的病人信息，获取满足条件的病人信息
    '入参:frmMain-调用的主窗体
    '     strPrivsOpt-记帐操作的相关权限
    '     strInput-输入的值
    '     intOutDay-查找出院病人天数
    '     strUnitIDs-查找的病区IDs
    '     bln所有病区-是否查找所有病区,如果查找所有病区, strUnitIDs将不启作用
    '出参:lng病人ID_Out-接口返回true时，返回病人ID,否则返回0
    '     strErrMsg_out-接口返回False时，返回的错误信息,否则返回空
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-08 18:04:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String, rsOutSel As ADODB.Recordset
    Dim blnCancel As Boolean, vRect As RECT
    
    On Error GoTo errHandle
    
    strErrMsg_out = ""
    'a.是否具有强制记帐权限
    strWhere = ""
    If InStr(strPrivsOpt, ";出院未结强制记帐;") > 0 And InStr(strPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And nvl(A.在院,0)=1"
        Else
            strWhere = " And (nvl(A.在院,0)=1 Or  B.出院日期>Trunc(Sysdate)-" & intOutDay & ")"
        End If
    ElseIf InStr(strPrivsOpt, ";出院未结强制记帐;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And ((nvl(A.在院,0)=1 And B.状态<>3) Or (Nvl(X.费用余额,0)<>0 And nvl(A.在院,0)=1  And B.状态=3) )"
        Else
            strWhere = " And ((nvl(A.在院,0)=1 And B.状态<>3) Or (Nvl(X.费用余额,0)<>0 And ((nvl(a.在院,0)=1 And B.状态=3) Or (B.出院日期>Trunc(Sysdate)-" & intOutDay & "))))"
        End If
    ElseIf InStr(strPrivsOpt, ";出院结清强制记帐;") > 0 Then
        If intOutDay = 0 Then
            strWhere = " And ((nvl(A.在院,0)=1  And B.状态<>3) Or (Nvl(X.费用余额,0)=0 And nvl(A.在院,0)=1  And B.状态=3))"
        Else
            strWhere = " And ((nvl(A.在院,0)=1  And B.状态<>3) Or (Nvl(X.费用余额,0)=0 And ((nvl(A.在院,0)=1 And B.状态=3) Or (B.出院日期>Trunc(Sysdate)-" & intOutDay & "))))"
        End If
    Else
        '没有权限对出院和预出院病人结帐
        strWhere = " And Nvl(A.在院,0)=1 And Nvl(B.状态,0)<>3 "
    End If
    
    
    'b.是否可以记所有病区病人
    If Not bln所有病区 Then
        If InStr(1, strUnitIDs, ",") = 0 Then
            strWhere = strWhere & " And B.当前病区ID+0=[3]"
        Else
            strWhere = strWhere & " And B.当前病区ID+0 IN(Select Column_Value From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
    
    'c.是否留观病人记帐权限
    If (InStr(strPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观) And (InStr(strPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观) Then
        strWhere = strWhere & " And Nvl(B.病人性质,0) IN(0,1,2)"
    ElseIf InStr(strPrivsOpt, ";门诊留观记帐;") > 0 And gbln门诊留观 Then
        strWhere = strWhere & " And Nvl(B.病人性质,0) IN(0,1)"
    ElseIf InStr(strPrivsOpt, ";住院留观记帐;") > 0 And gbln住院留观 Then
        strWhere = strWhere & " And Nvl(B.病人性质,0) IN(0,2)"
    Else
        strWhere = strWhere & " And Nvl(B.病人性质,0)=0"
    End If
        
    strSQL = _
    " Select Rownum as ID, A.病人ID,Decode(nvl(A.在院,0),1,'√','') as 在院, nvl(B.姓名,A.姓名) as 姓名,nvl(b.性别,A.性别) as 性别,nvl(b.年龄,A.年龄) as 年龄,B.费别,B.住院医师,B.医疗付款方式, " & _
    "       to_Char(B.入院日期,'yyyy-mm-dd') as 入院日期,to_Char(B.出院日期,'yyyy-mm-dd') 出院日期," & _
    "       A.住院号,B.出院病床 as 床号,X.费用余额,C.名称 as 当前病区,B.病人类型,B.备注" & _
    " From 病人信息 A,病案主页 B,病人余额 X,部门表 C" & _
    " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID and B.当前病区ID=C.ID(+)" & strWhere & _
    "       And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) and X.性质(+)=1 and X.类型(+)=2 And A.停用时间 is NULL  " & _
    "       And A.姓名 like [1] " & _
    " Order by 在院 Desc,出院日期"
    
    If lngHwnd = 0 Then
        Set rsOutSel = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "病人选择器", False, "", "请选择病人", False, False, False, 0, 0, 0, blnCancel, False, True, strInput & "%", "", Val(strUnitIDs), strUnitIDs, "bytSize=1")
    Else
        vRect = zlControl.GetControlRect(lngHwnd)
        Set rsOutSel = zlDatabase.ShowSQLSelect(frmMain, strSQL, 0, "病人选择器", False, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, lngHeight, blnCancel, False, True, strInput & "%", "", Val(strUnitIDs), strUnitIDs, "bytSize=1")
    End If
    If blnCancel Then Exit Function
    
    If Not rsOutSel Is Nothing Then
        If rsOutSel.State = 1 Then
            If rsOutSel.EOF = False Then
                lng病人ID_Out = Val(rsOutSel!病人ID)
                Set rsOutSel = Nothing
                zlSelectChargePatiFromInputName = True: Exit Function
            End If
        End If
    End If
    strErrMsg_out = "未找到姓名符合『" & strInput & "』的病人,请检查是否输入正确!"
    Set rsOutSel = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


