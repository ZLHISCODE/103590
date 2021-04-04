Attribute VB_Name = "mdl江苏省"
Option Explicit

'函数返回值为0时表示成功,非0表示失败

'===============================================================================================================
'功能：本地医保数据库生成医保使用的门诊收费流水或住院流水号,用于关联HIS费用记录和医保费用记录,
'      一次收费对应一个号码,一次住院对应一个号码
'入口参数：处方类型（0门诊/1住院）
'出口参数：收费流水号
'===============================================================================================================
Public Declare Function FGetRecCode Lib "HInterface.dll" (ByVal intRecType As Long, _
    ByVal strRecCode As String) As Long
    
'===============================================================================================================
'功能：读取参保人的基本信息和帐户、支出信息
'入口参数：类型（0门诊/1住院）,收费流水号,个人证号
'出口参数：0个人ID,1卡号,2姓名,3区属代码,4身份证号码,5单位ID,6单位名称,7性别(男/女),8人员类别,9出生日期,10民族,
'          11本人身份,12行政职务,13门诊特殊病种(脱机返回'未知'/联机未申请时返回'未申请'/其他返回已申请的特殊病种),
'          14其它(保留),15本年累计住院次数,16帐户总收,17帐户总支,18支出版本号,19本年统筹支付累计,20本年大病基金支付累计,
'          21本年公务员补充/企业补充支付累计,22本年普通门诊费用累计,23本年普通门诊三个范围内费用累计,
'          24本年特殊门诊三个范围内费用累计,25本年比照住院三个范围内费用累计,26本年普通住院费用累计,
'          27本年普通住院三个范围内费用累计,28本年家庭病床住院三个范围内费用累计,29其他1,30其他2,
'          31本年储蓄帐户支付累计,32本年其它基金支付累计,33本年现金支付累计,34帐户余额
'===============================================================================================================
Public Declare Function FGetCardInfo Lib "HInterface.dll" (ByVal intRecType As Long, ByVal strRecCode As String, _
    ByVal strVoucherID As String, intInsID As Long, ByVal strCardID As String, ByVal STRNAME As String, _
    ByVal strAreaCode As String, ByVal strQueryID As String, ByVal strUnitID As String, ByVal strUnitName As String, _
    ByVal strSex As String, ByVal strKind As String, ByVal strBirthday As String, ByVal strNational As String, _
    ByVal strIndustry As String, ByVal strDuty As String, ByVal strChronic As String, ByVal strOthers1 As String, _
    sngInHosNum As Single, sngAccIn As Single, sngAccOut As Single, sngFeeNO As Single, _
    sngPubPay As Single, sngHelpPay As Single, sngSupplyPay As Single, sngOutpatSum As Single, _
    sngOutpatGen1 As Single, sngOutpatGen2 As Single, sngOutpatGen3 As Single, _
    sngInpatSum As Single, sngInpatGen1 As Single, sngInpatGen2 As Single, _
    sngOther1 As Single, sngOther2 As Single, sngBankAccPay As Single, sngOtrPay As Single, _
    sngCashPay As Single, sngAccLeft As Single) As Long

'===============================================================================================================
'功能：门诊挂号
'入口参数：收费流水号,挂号类别（专家号，普通号...）,科室名称,挂号费项目编码,挂号费项目名称,挂号费金额,
'          诊疗费项目编码,诊疗费项目名称,诊疗费金额,费别,操作员,日期,交易模式(T输卡号/F刷卡)
'出口参数：统筹支付,帐户支付,现金支付
'===============================================================================================================
Public Declare Function FOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strRegName As String, _
    ByVal strDepartName As String, ByVal strRegFeeCode As String, ByVal strRegFeeName As String, ByVal sngRegFee As Single, _
    ByVal strDiagFeeCode As String, ByVal strDiagFeeName As String, ByVal sngDiagFee As Single, ByVal strFeeType As String, _
    ByVal strOpCode As String, ByVal strRegDate As String, ByVal pRegMode As String, sngPubPay As Single, _
    sngAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：取消挂号，退还挂号费
'入口参数：门诊号,操作员代码
'出口参数：无
'===============================================================================================================
Public Declare Function FCancleOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Long

'===============================================================================================================
'功能：过该接口向医保系统HT.INS传入参保病人发生在医院的所有费用明细，包括药品、诊疗项目、服务设施收费等
'入口参数：类型(0门诊/1住院),收费流水号（住院、门诊不同）,项目类型('0'非药品/'1'药品),项目编码(HIS编码),明细编码,
'          项目名称,单位,规格、剂型等,费别码,处方药标志,数量,应售单价,实售单价,每次用量,使用频次,用法,执行天数,
'          收费员工号,科室编码,处方医生工号,发生日期
'出口参数：个人支付比例,个人自付金额,超标准费用
'费别码传入编码：由于医院的全部药品、项目必须对照后才能使用，费别码可以不传
'处方药标志传入编码：0非处方药；1处方药；2收费项目
'===============================================================================================================
Public Declare Function FWriteFeeDetail Lib "HInterface.dll" (ByVal intRecType As Long, _
    ByVal strRecCode As String, ByVal strItmFlag As String, ByVal strItmCode As String, ByVal strAliasCode As String, _
    ByVal strItmName As String, ByVal strItmUnit As String, ByVal strItmDesc As String, ByVal strFeeCode As String, _
    ByVal strOTCCode As String, ByVal sngQuantity As Single, ByVal sngPharPrice As Single, ByVal sngFactPrice As Single, _
    ByVal sngDosage As Single, ByVal strFrequency As String, ByVal strUsage As String, ByVal sngDays As Single, _
    ByVal strOpCode As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strRecDate As String, _
    sngRate As Single, sngSelfFee As Single, sngDeduct As Single) As Long

'===============================================================================================================
'功能：门诊费用录入完毕后的试结算；试算时上传费用记录
'入口参数：收费流水号,操作员工,是否使用帐户(是/否),科室编码,医生工号,医疗方式,医疗类别,('A'),报销代码,保留1,保留2,备注
'出口参数：0结账流水号,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,8大病救助基金支付,
'          9大病救助基金自付,10公务员补充支付/企业补充支付,11公务员补充支付/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
'===============================================================================================================
Public Declare Function FTryOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strMedMode As String, _
    ByVal strRecClass As String, ByVal strICDMode As String, ByVal strICD As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：门诊费用录入完毕后结算，在调用此函数前必须调用门诊试结算函数，函数返回费用构成情况，结果记录入库
'入口参数：收费流水号,操作员工,是否使用帐户(是/否),科室编码,医生工号,医疗方式,医疗类别,('A'),报销代码,保留1,保留2,备注
'出口参数：0结账流水号,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,8大病救助基金支付,
'          9大病救助基金自付,10公务员补充支付/企业补充支付,11公务员补充支付/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
'===============================================================================================================
Public Declare Function FOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strMedMode As String, _
    ByVal strRecClass As String, ByVal strICDMode As String, ByVal strICD As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：取消门诊收费，需要传入收费流水号
'入口参数：收费流水号,结账流水号,操作员工号
'出口参数：0总费用,1三个范围内费用,2自付费用,3自费费用,4起付标准,5统筹支付,6统筹自付,7大病救助基金支付,
'          8大病救助基金自付,9公务员/企业补充支付,10公务员/企业补充自付,11其它基金支付,12个人医疗帐户支付,
'          13个人储蓄帐户支付,14现金支付
'===============================================================================================================
Public Declare Function FCancelOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strBillCode As String, _
    ByVal strOpCode As String, sngSumFee As Single, sngGenFee As Single, sngFirstPay As Single, _
    sngSelfFee As Single, sngPayLevel As Single, sngPubPay As Single, sngPubSelf As Single, _
    sngHelpPay As Single, sngHelpSelf As Single, sngSupplyPay As Single, sngSupplySelf As Single, _
    sngOtrPay As Single, sngMedAccPay As Single, sngBankAccPay As Single, _
    sngCashPay As Single) As Long

'===============================================================================================================
'功能：参保病人入院时的入院登记，将登记信息记录入库
'入口参数：住院号（收费流水号）,医疗方式,医疗类别,操作员编码,入院日期,ICD编码规则('A'),入院诊断(ICD10编码),
'          科室代码,病区代码,入院医生代码
'出口参数：本年内累计住院次数
'===============================================================================================================
Public Declare Function FInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strMedMode As String, _
    ByVal strMedClass As String, ByVal strRegOpCode As String, ByVal strBegDate As String, ByVal strICDMode As String, _
    ByVal strICD As String, ByVal strDepCode As String, ByVal strSecCode As String, ByVal strRegDoc As String, _
    sngInHosNum As Single) As Long

'===============================================================================================================
'功能：参保病人登记住院后不打算住院，做取消登记处理
'入口参数：住院号（收费流水号）,操作员编码
'出口参数：无
'===============================================================================================================
Public Declare Function FCancelInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Long

'===============================================================================================================
'功能：参保病人住院中登记信息变更，如转科等，调用该函数修改登记信息
'入口参数：住院号,医疗方式,医疗类别,操作员编码,就诊开始日期,ICD编码规则('A'),入院诊断(ICD10编码),科室代码,
'          病区代码,入院医生代码
'出口参数：无
'===============================================================================================================
Public Declare Function FChgInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strMedMode As String, _
    ByVal strMedClass As String, ByVal strRegOpCode As String, ByVal strBegDate As String, ByVal strICDMode As String, _
    ByVal strICD As String, ByVal strDepCode As String, ByVal strSecCode As String, _
    ByVal strRegDoc As String) As Long

'===============================================================================================================
'功能：参保病人出院的操作，类似入院登记，并不是结账处理
'入口参数：住院号,操作员代码,出院日期,出院原因,ICD编码规则('A'),出院诊断(ICD10编码),出院医生代码
'出口参数：无
'出院原因传入编码：1治愈；2好转；3未愈；4死亡；5转院；6转外；9其他
'===============================================================================================================
Public Declare Function FInpatLeave Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOutOpCode As String, _
    ByVal strEndDate As String, ByVal strOutCause As String, ByVal strICDMode As String, ByVal strICD As String, _
    ByVal strOutDoc As String) As Long

'===============================================================================================================
'功能：医院对在院病人的费用进行预结算，方便收取住院押金，
'      此情况下可以直接调用住院费用试算：返回试算结果供参考，信息不入医保本地库
'入口参数：住院号,操作员工号,是否使用帐户(是/否),结算方式,报销代码,保留1,保留2,备注
'出口参数：0（'UnKnown'）,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,
'          8大病救助基金支付,9大病救助基金自付,10公务员/企业补充支付,11公务员/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
'结算方式传入编码：0正常结算；1中途结算
'===============================================================================================================
Public Declare Function FTryInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal intLiquiMode As String, ByVal strRefundID As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：医院对在院病人的费用进行预结算，方便收取住院押金，
'      此情况下可以直接调用住院费用试算：返回试算结果供参考，信息不入医保本地库
'入口参数：住院号,操作员工号,是否使用帐户(是/否),结算方式,报销代码,保留1,保留2,备注
'出口参数：0结账流水号,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,
'          8大病救助基金支付,9大病救助基金自付,10公务员/企业补充支付,11公务员/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
'结算方式传入编码：0正常结算；1中途结算
'===============================================================================================================
Public Declare Function FInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal intLiquiMode As String, ByVal strRefundID As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：作废住院结帐信息，参保病人返回在院状态
'入口参数：住院号,结账流水号,操作员工号
'出口参数：0总费用,1三个范围内费用,2自付费用,3自费费用,4起付标准,5统筹支付,6统筹支付,7大病救助基金支付,
'          8大病救助基金支付,9公务员/企业补充支付,10公务员/企业补充支付,11其它基金支付,12个人医疗帐户支付,
'          13个人储蓄帐户支付,14现金支付
'===============================================================================================================
Public Declare Function FCancelInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, _
    ByVal strBillCode As String, ByVal strOpCode As String, sngSumFee As Single, sngGenFee As Single, _
    sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, sngPubPay As Single, _
    sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, sngSupplyPay As Single, _
    sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'功能：1、在医院门诊或住院结算时，若发现医保中心返回的帐户、统筹和自付各项分担之和不等医院HIS系统中的总费用，
'         需要调用此函数清除医保本地库和中心的数据，重新上传；
'      2、医院HIS系统需要重新把所有该病人的数据调用 '费用录入FWriteFeeDetail'函数传入本地医保数据库，随后，
'         可以调用上传函数上传，也可由试算接口函数自动上传；
'入口参数：2,收费流水号
'出口参数：无
'===============================================================================================================
Public Declare Function FSynData Lib "HInterface.dll" (ByVal intType As Long, ByVal strRecCode As String) As Long

'===============================================================================================================
'功能：调用此函数将医保本地库HT.HIS中尚未上传的数据传输至医保中心
'入口参数：2,收费流水号（若为*表示上传所有）
'出口参数：无
'===============================================================================================================
Public Declare Function FUpLoad Lib "HInterface.dll" (ByVal intType As Long, ByVal strRecCode As String) As Long

'===============================================================================================================
'功能：该函数为通用的数据导入函数，用于从医院HIS导入数据到医保本地的数据库
'入口参数：2,收费流水号（若为*表示上传所有）
'出口参数：类型(1科室/2操作员/3医生/4药品字典/5诊疗项目),导入信息,导入信息,导入信息,导入信息,备注,操作状态(I)
'A、科室
'   piType： 1
'   psInfo1： 科室编码
'   psInfo2： 科室名称
'B、操作员：
'   piType： 2
'   psInfo1： 操作员工号
'   psInfo2： 姓名
'C、医生:
'   piType： 3
'   psInfo1： 医生编码
'   psInfo2： 医生姓名
'   psInfo3：隶属科室编码
'   psInfo4：职称(主任医师/副主任医师/主治医师/…)
'D、药品字典
'   piType： 4
'   psInfo1： 药品自编码
'   psInfo2：市医保编码
'   psInfo3： 名称(商品名)
'   psInfo4：零售价
'   psRemark：剂型|计量单位|规格|产地|批号
'   （说明：psInfo4应与计量单位对应
'           psRemark参数中所有的数据项用|分隔，为空时填空字符串）
'E、诊疗项目：
'   piType ： 5
'   psInfo1：项目自编码
'   psInfo2：市医保编码
'   psInfo3：名称
'   psInfo4：零售价
'   psRemark：单位
'===============================================================================================================
Public Declare Function FImpInfo Lib "HInterface.dll" (ByVal intType As Long, ByVal strInfo1 As String, _
    ByVal strInfo2 As String, ByVal strInfo3 As String, ByVal strInfo4 As String, ByVal strRemark As String, _
    ByVal strOpStaus As String) As Long

'===============================================================================================================
'功能：用于导出医保本地库数据，如医保药品字典、对照表等；传入参数为表名（另外约定）和导出文件名（为TXT文本，
'      空格分隔）；文件名包含路径，路径不存在时，自动保存在接口动态库所在路径
'入口参数：表名,文件名
'出口参数：无
'===============================================================================================================
Public Declare Function FExpInfo Lib "HInterface.dll" (ByVal strTable As String, ByVal strFile As String) As Long

'===============================================================================================================
'功能：门诊试算完成后，如果需要取消本次试算并重新录入费用明细，则调用此函数
'入口参数：收费流水号,操作员工号
'出口参数：无
'===============================================================================================================
Public Declare Function FCancelTryOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, _
    ByVal strOpCode As String) As Long
    
'以上为动态链接库函数定义部分

Public gstrRecCode As String             '收费流水号
Public gblnReadCard As Boolean           '是否使用读卡器
Public gint医疗方式 As Long           '1普通门诊；2普通住院；3特殊门诊；4紧急抢救；5急诊；
Private intReturn As Long

'以下为医保接口函数据部分

Public Function 医保初始化_江苏() As Boolean
    医保初始化_江苏 = True
End Function

Public Function 身份标识_江苏(Optional bytType As Byte = 0, Optional lng病人ID As Long = 0) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify江苏
    Dim strPatiInfo As String, cur余额 As Currency, str就诊编号 As String
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strPatiInfo = frmIDentified.GetPatient(bytType, lng病人ID)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
                lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_江苏)

        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        '写入就诊编号
        If bytType = 1 Then
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_江苏 & ",'顺序号','''" & gstrRecCode & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_江苏")
        ElseIf bytType = 3 Or bytType = 0 Then
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_江苏 & ",'退休证号','''" & gstrRecCode & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_江苏")
        End If
        Unload frmIDentified
    Else
        身份标识_江苏 = ""
        MsgBox "未提取病人信息。", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_江苏 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_江苏 = ""
End Function

Public Function 挂号结算_江苏(ByVal lng结帐ID As Long, cur金额 As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strCode1 As String, strName1 As String, sngAcc1 As Single, lng病人ID As Long
    Dim strCode2 As String, strName2 As String, sngAcc2 As Single
    Dim sng统筹支付 As Single, sng个帐支付 As Single, sng现金支付 As Single, cur个帐余额 As Currency
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    On Error GoTo errHandle:
    gstrSQL = "select * from 收费细目 where 名称 = '门诊诊疗费'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    strCode2 = rsTemp!编码
    strName2 = rsTemp!名称
    
    '获取病人挂号信息
    gstrSQL = "select a.no,a.id,a.病人id,a.姓名,a.病人科室id,c.名称 as 科室名称,a.实收金额,a.操作员姓名,a.费别,to_char(a.发生时间,'yyyy-mm-dd HH24:MI:SS') as 时间,a.收据费目,b.名称,b.编码 from 门诊费用记录 a,收费细目 b,部门表 c where 结帐id=[1] and b.id=a.收费细目id and a.病人科室id=c.id order by 收据费目"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng病人ID = rsTemp!病人ID
    strCode1 = rsTemp!编码
    strName1 = rsTemp!名称
    sngAcc1 = rsTemp!实收金额
    rsTemp.MoveNext
'    strCode2 = rsTemp!编码
'    strName2 = rsTemp!名称
    sngAcc2 = rsTemp!实收金额
    '进行医保挂号
    intReturn = FOutpatReg(gstrRecCode, "普通号", rsTemp!科室名称, strCode1, strName1, sngAcc1, strCode2, strName2, _
        sngAcc2, rsTemp!费别, UserInfo.编号, rsTemp!时间, IIf(gblnReadCard, "F", "T"), sng统筹支付, sng个帐支付, sng现金支付)
        
    If intReturn <> 0 Then
        MsgBox "进行医保门诊挂号失败，未获取错误信息。", vbInformation, gstrSysName
        挂号结算_江苏 = False
        Exit Function
    End If
'    rsTemp.MoveFirst
'    gstrSQL = "zl_病人记帐记录_上传(" & rsTemp!ID & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
'    rsTemp.MoveNext
'    gstrSQL = "zl_病人费用记录_上传(" & rsTemp!ID & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '如果个帐支付金额大于个帐余额，则将多出部分以现金方式支付
'    cur个帐余额 = 个人余额_江苏(rsTemp!病人ID)
'    If cur个帐余额 < sng个帐支付 Then
'        sng现金支付 = sng现金支付 + sng个帐支付 - cur个帐余额
'        sng个帐支付 = cur个帐余额
'    End If
    Call Get帐户信息(TYPE_江苏, lng病人ID, Year(Date), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_江苏 & "," & Year(Date) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + sng个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + sng统筹支付 & "," & int住院次数累计 + 1 & ",0,0,0,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '保险结算记录
    '====================================注意===================================='
    '结算后需要把gstrRecCode、strBillCode和lng结帐ID对应并保存,这个值存放在备注中'
    '============================================================================'
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_江苏 & "," & _
            lng病人ID & "," & Year(Date) & "," & _
            "0" & "," & cur帐户支出累计 + sng个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + sng统筹支付 & "," & int住院次数累计 + 1 & ",NULL,NULL,NULL," & _
            sng统筹支付 + sng现金支付 + sng个帐支付 & "," & sng现金支付 & ",NULL,NULL," & sng统筹支付 & ",NULL,NULL," & _
            sng个帐支付 & ",NULL,NULL,NULL," & gstrRecCode & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    cur金额 = sng统筹支付
    挂号结算_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    挂号结算_江苏 = False
End Function

Public Function 挂号结算冲销_江苏(ByVal lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
'===============================================================================================================
'功能：取消挂号，退还挂号费
'入口参数：门诊号,操作员代码
'出口参数：无
'===============================================================================================================
'Public Declare Function FCancleOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Integer
    gstrSQL = "Select 退休证号 From 保险帐户 Where 病人id In (Select 病人id From 门诊费用记录 Where 结帐id=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到冲销记录的原数据，冲销不能继续执行。", vbInformation, gstrSysName
        挂号结算冲销_江苏 = False
        Exit Function
    End If
    
    intReturn = FCancleOutpatReg(rsTemp!退休证号, UserInfo.姓名)
    If intReturn <> 0 Then
        MsgBox "医保挂号退费时发生错误，未获得错误信息。", vbInformation, gstrSysName
        挂号结算冲销_江苏 = False
        Exit Function
    End If
    
    挂号结算冲销_江苏 = True
End Function

Public Function 个人余额_江苏(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_江苏)
    
    If rsTemp.EOF Then
        个人余额_江苏 = 100000
    Else
        个人余额_江苏 = IIf(rsTemp("帐户余额") = 0, 100000, rsTemp("帐户余额"))
    End If
End Function

Public Function 门诊虚拟结算_江苏(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim cur个帐 As Currency, cur统筹 As Currency, cur余额 As Currency, strBillCode As String
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset, sngArrInfo(20) As Single
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有发生费用，不能进行预结算。", vbInformation, gstrSysName
        门诊虚拟结算_江苏 = False
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    cur个帐 = 0: cur统筹 = 0
    gstrSQL = "Select * from 保险帐户 where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保预结算", lng病人ID)
    cur余额 = rsTemp!帐户余额
    
    intReturn = FCancelTryOutpatBalance(gstrRecCode, UserInfo.编号)
    
    '传递费用明细
    If 费用明细传递_江苏(0, rs明细, 1) = False Then Exit Function
    
    '调用预结算函数进行门诊预结算
    gstrSQL = "select a.姓名,a.编号,a.id,c.编码 as 部门id,c.名称 from 人员表 a,部门人员 b,部门表 c where a.id=b.人员id and a.姓名=[1] and c.id=b.部门id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rs明细!开单人))
    intReturn = FTryOutpatBalance(gstrRecCode, UserInfo.编号, "是", rsTemp!部门ID, rsTemp!编号, CStr(gint医疗方式), _
        IIf(gint医疗方式 = 3, "12", "11"), "A", "", 0, 0, "", strBillCode, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
'    If intReturn <> 0 Then
'        MsgBox "在进行医保门诊预结算时发生错误，未取得错误信息。", vbInformation, gstrSysName
'        门诊虚拟结算_江苏 = False
'        Exit Function
'    End If
'
    cur个帐 = CCur(sngArrInfo(13) + sngArrInfo(12))
    cur统筹 = CCur(sngArrInfo(0) - sngArrInfo(14)) - cur个帐
    
    '如果报销额大于帐户余额，则允许从帐户中支付的最大额为帐户余额
'    If cur个帐 > cur余额 Then cur个帐 = cur余额
    
'    MsgBox str报销明细, vbInformation, "报销明细"
    
    str结算方式 = "省个人帐户;" & cur个帐 & ";0"
    str结算方式 = str结算方式 & "|" & "省统筹基金;" & cur统筹 & ";0"
    门诊虚拟结算_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function 门诊结算_江苏(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim cur统筹 As Currency, cur余额 As Currency, strBillCode As String, datCurr As Date
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset, sngArrInfo(20) As Single
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur发生费用 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    gstrSQL = "Select 病人id From 门诊费用记录 Where 结帐id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng病人ID = rsTemp(0)
    gstrSQL = "Select * from 保险帐户 where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    cur余额 = rsTemp!帐户余额
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select 开单部门id,开单人,b.编码,c.编号 from 门诊费用记录 a,部门表 b,人员表 c where b.id=a.开单部门id and c.姓名=a.开单人 and a.结帐id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    '调用预结算函数进行门诊预结算
    strBillCode = Space(7)
    intReturn = FOutpatBalance(gstrRecCode, UserInfo.编号, "是", rsTemp!编码, rsTemp!编号, CStr(gint医疗方式), _
        IIf(gint医疗方式 = 3, "12", "11"), "A", "", 0, 0, "", strBillCode, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "在进行医保门诊结算时发生错误，未取得错误信息。"
        门诊结算_江苏 = False
        Exit Function
    End If
    
    '获取个人帐户支付和个人现金支付
    cur个人帐户 = CCur(sngArrInfo(13) + sngArrInfo(12))
    cur统筹 = CCur(sngArrInfo(0) - sngArrInfo(14)) - cur个人帐户
    
'出口参数：0结账流水号,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,8大病救助基金支付,
'          9大病救助基金自付,10公务员补充支付/企业补充支付,11公务员补充支付/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
    
    cur全自付 = CCur(sngArrInfo(13)) + CCur(sngArrInfo(14))
    cur发生费用 = CCur(sngArrInfo(0))
    '帐户年度信息
    Call Get帐户信息(TYPE_江苏, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_江苏 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur统筹 & "," & int住院次数累计 + 1 & "," & sngArrInfo(4) & "," & _
            sngArrInfo(4) & ",0,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '保险结算记录
    '====================================注意===================================='
    '结算后需要把gstrRecCode、strBillCode和lng结帐ID对应并保存,这个值存放在备注中'
    '============================================================================'
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_江苏 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur统筹 & "," & int住院次数累计 + 1 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & ",NULL,NULL,NULL,NULL,NULL," & _
            cur个人帐户 & ",NULL,NULL,NULL,'" & strBillCode & ";" & gstrRecCode & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    门诊结算_江苏 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 费用明细传递_江苏(lng结帐ID As Long, Optional rs明细IN As ADODB.Recordset = Nothing, Optional int门诊标志 As Integer = 1) As Boolean
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str就诊编号 As String, str医生工号 As String, objSystem As New FileSystemObject, objStream As TextStream
    Dim str科室编号 As String, str科室名称 As String, lng科室ID As Long
    Dim strTemp As String, sngRate As Single, sngSelfFee As Single, sngDeduct As Single
    Dim sng数量 As Single, sng单价 As Single
    Dim sng实收金额 As Single
    
    On Error GoTo errHandle
    
    Set objStream = objSystem.OpenTextFile("C:\Trans.LOG", ForAppending, True, TristateFalse)
    If rs明细IN Is Nothing Then
        gstrSQL = "Select * From " & IIf(int门诊标志 = 1, "门诊费用记录", "住院费用记录") & " Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    Else
        Set rs明细 = rs明细IN.Clone
    End If
    If rs明细.EOF = True Then
'        MsgBox "没有需要上传的收费记录", vbExclamation, gstrSysName
        If int门诊标志 = 1 Then
            费用明细传递_江苏 = False
        Else
            费用明细传递_江苏 = True
        End If
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    If int门诊标志 = 2 Then
        gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Else
        gstrSQL = "Select nvl(退休证号,0) as 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_江苏)
    str就诊编号 = rsTemp!顺序号: gstrRecCode = rsTemp!顺序号
    objStream.WriteBlankLines 1
    While Not rs明细.EOF
'        If IsNull(rs明细!是否上传) Or rs明细!是否上传 = 0 Then
'0病人ID
'1收费类别
'2收据费目
'3计算单位
'4开单人
'5收费细目ID
'6数量
'7单价
'8实收金额
'9统筹金额
'10保险支付大类ID
'11是否医保
'12摘要
'13是否急诊
            On Error Resume Next
            Err = 0
            strTemp = Nvl(rs明细!开单人)
            If Err.Number <> 0 Then
                strTemp = rs明细!医生
                sng实收金额 = rs明细!金额
            Else
                sng实收金额 = rs明细!实收金额
            End If
            Err.Clear
            On Error GoTo errHandle
            gstrSQL = "select b.编号,b.姓名,c.编码,c.名称 from 部门人员 a,(select id,编号,姓名 from 人员表 Where 姓名=[1]) b,(select id,编码,名称 from 部门表) c where a.部门id=c.id and a.人员id=b.id"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, strTemp)
            If Not rsTemp.EOF Then
                str医生工号 = rsTemp!编号
                str科室编号 = rsTemp!编码
                str科室名称 = rsTemp!名称
            Else
                str医生工号 = ""
                str科室编号 = ""
                str科室名称 = ""
            End If
'            gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
            gstrSQL = "select a.费用类型,A.名称,C.项目编码 as 编码,A.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格" & _
                  " from 收费细目 A,药品目录 B,保险支付项目 C where A.id = C.收费细目id and A.id=B.药品id(+) and A.id =[1] And C.险类=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID), TYPE_江苏)
            strTemp = IIf(rs明细!收费类别 = 5 Or rs明细!收费类别 = 6 Or rs明细!收费类别 = 7, "1", "0")
            
            If int门诊标志 <> 1 Then
                sng数量 = rs明细!数次 * rs明细!付数
                sng单价 = rs明细!标准单价
            Else
                sng数量 = rs明细!数量
                sng单价 = rs明细!单价
            End If
            
            '上传明细
'入口参数：类型(0门诊/1住院),收费流水号（住院、门诊不同）,项目类型('0'非药品/'1'药品),项目编码(HIS编码),明细编码,
'          项目名称,单位,规格、剂型等,费别码,处方药标志,数量,应售单价,实售单价,每次用量,使用频次,用法,执行天数,
'          收费员工号,科室编码,处方医生工号,发生日期
            objStream.WriteLine "FWriteFeeDetail(" & IIf(int门诊标志 = 2, 1, 0) & ",""" & str就诊编号 & """,""" & _
                strTemp & """,""" & rsTemp!编码 & """,""" & rsTemp!编码 & """,""" & rsTemp!名称 & """,""" & Nvl(rsTemp!计算单位) & """,""" & Nvl(rsTemp!规格) & ""","""",""" & _
                IIf(strTemp = "0", "2", IIf(rsTemp!费用类型 = "甲类药" Or rsTemp!费用类型 = "乙类药", "1", "0")) & """," & _
                sng数量 & "," & sng单价 & "," & sng实收金额 / sng数量 & ",0,"""","""",0,""" & _
                UserInfo.编号 & """,""" & str科室编号 & """,""" & str医生工号 & """,""" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM:SS") & """," & _
                sngRate & "," & sngSelfFee & "," & sngDeduct & ")"
            
            If int门诊标志 = 1 Then
                intReturn = FWriteFeeDetail(IIf(int门诊标志 = 2, 1, 0), str就诊编号, _
                    strTemp, rsTemp!编码, rsTemp!编码, rsTemp!名称, Nvl(rsTemp!计算单位), Nvl(rsTemp!规格), " ", _
                    IIf(strTemp = "0", "2", IIf(rsTemp!费用类型 = "甲类药" Or rsTemp!费用类型 = "乙类药", "1", "0")), _
                    sng数量, sng单价, sng实收金额 / sng数量, 0, " ", " ", 0, _
                    UserInfo.编号, str科室编号, str医生工号, Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM:SS"), _
                    sngRate, sngSelfFee, sngDeduct)
            Else
                intReturn = FWriteFeeDetail(IIf(int门诊标志 = 2, 1, 0), str就诊编号, _
                    strTemp, rsTemp!编码, rsTemp!编码, rsTemp!名称, Nvl(rsTemp!计算单位), Nvl(rsTemp!规格), " ", _
                    IIf(strTemp = "0", "2", IIf(rsTemp!费用类型 = "甲类药" Or rsTemp!费用类型 = "乙类药", "1", "0")), _
                    sng数量, sng单价, sng实收金额 / sng数量, 0, " ", " ", 0, _
                    UserInfo.编号, str科室编号, str医生工号, Format(rs明细!发生时间, "yyyy-MM-dd HH:MM:SS"), _
                    sngRate, sngSelfFee, sngDeduct)
            End If
            If intReturn <> 0 Then
                MsgBox "在进行数据传递时发生错误，未取得错误信息。", vbInformation, gstrSysName
                费用明细传递_江苏 = False
                objStream.Close
                Exit Function
            End If
            
            If int门诊标志 <> 1 Then
                WriteInfo "NO:" & rs明细!NO & "      序号:" & rs明细!序号
                gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            End If
'        End If
        rs明细.MoveNext
    Wend
    If int门诊标志 = 2 Then
        intReturn = FUpLoad(2, gstrRecCode)
        If intReturn <> 0 Then
            MsgBox "在进行数据传递时发生错误。", vbInformation, gstrSysName
            费用明细传递_江苏 = False
            objStream.Close
            Exit Function
        End If
    End If
    objStream.Close
    费用明细传递_江苏 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    objStream.Close
End Function

Public Function 门诊结算冲销_江苏(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 门诊费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    lng冲销ID = rsTemp("结帐ID")
    
    '提取在结帐时保存的收费流水号和结帐流水号
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_江苏, lng结帐ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        门诊结算冲销_江苏 = False
        Exit Function
    End If
    strRecCode = Mid(rsTemp!备注, InStr(rsTemp!备注, ";") + 1)
    strBillCode = Left(rsTemp!备注, InStr(rsTemp!备注, ";") - 1)
    '调用接口数冲销
    
'入口参数：收费流水号,结账流水号,操作员工号
    intReturn = FCancelOutpatBalance(strRecCode, strBillCode, UserInfo.编号, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "进行门诊结算冲销时发生错误，未获得错误信息。", vbInformation, gstrSysName
        门诊结算冲销_江苏 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_江苏, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_江苏 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_江苏 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",Null,Null,Null,'" & strBillCode & ";" & strRecCode & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    门诊结算冲销_江苏 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_江苏(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date
    Dim lng病种ID As Long, sngInHosNum As Single
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.门诊医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记_江苏 = False
        Exit Function
    End If
    
    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, True) '入院诊断
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    '获取住院医师代码
    gstrSQL = "Select ID,编号,姓名,简码,个人简介,接受培训 from 人员表 Where 姓名=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!门诊医师))
    
    '进行医保登记
    intReturn = FInpatReg(gstrRecCode, "2", IIf(gint医疗方式 = 3, "22", "21"), UserInfo.编号, _
        Format(rsTemp!入院日期, "yyyy-MM-dd HH:MM:SS"), "A", strInNote, rsTemp!科室编码, " ", _
        IIf(rsTmp.EOF, " ", rsTmp!编号), sngInHosNum)
    If intReturn <> 0 Then
        MsgBox "进行医保入院登记时发生错误，未能取得错误信息。", vbInformation, gstrSysName
        入院登记_江苏 = False
        Exit Function
    End If
     
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_江苏 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_江苏 = False
End Function

Public Function 转科转床_江苏(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将转科转床信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date
    Dim lng病种ID As Long, sngInHosNum As Single
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.顺序号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        转科转床_江苏 = False
        Exit Function
    End If
    
    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, True) '入院诊断
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    '获取住院医师代码
    gstrSQL = "Select ID,编号,姓名,简码,个人简介,接受培训 from 人员表 Where 姓名=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!住院医师))
    
    '进行医保登记
    intReturn = FChgInpatReg(rsTemp!顺序号, "2", IIf(gint医疗方式 = 3, "22", "21"), UserInfo.编号, _
        Format(rsTemp!入院日期, "yyyy-MM-dd HH24:MI:SS"), "A", strInNote, rsTemp!科室编码, rsTemp!科室编码, _
        rsTmp!编号)
    If intReturn <> 0 Then
        MsgBox "进行医保入院病人信息变动时发生错误，未能取得错误信息。", vbInformation, gstrSysName
        转科转床_江苏 = False
        Exit Function
    End If
     
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_江苏 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    转科转床_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    转科转床_江苏 = False
End Function

Public Function 入院登记撤消_江苏(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '获取病人相关信息
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If rsTemp.EOF Then
        MsgBox "不能找到病人的相关信息。", vbInformation, gstrSysName
        入院登记撤消_江苏 = False
        Exit Function
    End If
    
    '调用接口进行撤消登记
    intReturn = FCancelInpatReg(rsTemp!顺序号, UserInfo.编号)
    If intReturn <> 0 Then
        MsgBox "撤消入院登记时发生错误，未获取错误信息。", vbInformation, gstrSysName
        入院登记撤消_江苏 = False
        Exit Function
    End If
    入院登记撤消_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    入院登记撤消_江苏 = False
End Function

Public Function 住院结算冲销_江苏(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, StrInput As String, sngArrInfo(20) As Single
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String, lng病人ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, strTemp As String
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, rstTemp As String
    Dim cur票据总金额 As Currency, lng主页ID As Long
    Dim datCurr As Date, cur个人帐户 As Currency
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额,主页id From 住院费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng主页ID = rsTemp!主页ID
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_江苏)
    str就诊编号 = Nvl(rsTemp!顺序号, "0")
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_江苏, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        住院结算冲销_江苏 = False
        Exit Function
    End If
    cur个人帐户 = rsTemp!个人帐户支付
    strTemp = rsTemp!备注
    '调用接口数冲销
'入口参数：住院号,结账流水号,操作员工号
'出口参数：0总费用,1三个范围内费用,2自付费用,3自费费用,4起付标准,5统筹支付,6统筹支付,7大病救助基金支付,
'          8大病救助基金支付,9公务员/企业补充支付,10公务员/企业补充支付,11其它基金支付,12个人医疗帐户支付,
'          13个人储蓄帐户支付,14现金支付
    intReturn = FCancelInpatBalance(Mid(strTemp, InStr(strTemp, ";") + 1), Left(strTemp, InStr(strTemp, ";") - 1), _
        UserInfo.编号, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), _
        sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), _
        sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "住院结算冲销时发生错误。", vbInformation, gstrSysName
        住院结算冲销_江苏 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(TYPE_江苏, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_江苏 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_江苏 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 - 1 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算冲销_江苏 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算_江苏(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur起付线 As Currency, cur基本统筹限额 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur全自付 As Currency, cur先自付 As Currency
    
    Dim cur个人帐户支付 As Currency, cur个人现金支付 As Currency
    Dim cur统筹支付 As Currency, cur医保支付 As Currency, cur补充医保 As Currency
    Dim strBillCode As String, sngArrInfo(20) As Single
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From 住院费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    
    If rs明细.EOF = True Then
        Err.Raise 9000, gstrSysName, "没有填写收费记录", vbExclamation, gstrSysName
        住院结算_江苏 = False
        Exit Function
    End If
    lng病人ID = rs明细("病人ID")
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,帐户余额 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_江苏)
    str就诊编号 = rsTemp!顺序号
    cur余额 = rsTemp!帐户余额
    
    datCurr = zlDatabase.Currentdate
'入口参数：住院号,操作员工号,是否使用帐户(是/否),结算方式,报销代码,保留1,保留2,备注
'出口参数：0（'UnKnown'）,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,
'          8大病救助基金支付,9大病救助基金自付,10公务员/企业补充支付,11公务员/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
    strBillCode = Space(7)
    intReturn = FInpatBalance(str就诊编号, UserInfo.编号, "否", 0, "IA01", 0, 0, "", strBillCode, sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "住院病人预结算时发生错误，未获得错误信息。", vbInformation, gstrSysName
        住院结算_江苏 = False
        Exit Function
    End If

    '获取个人帐户支付和个人现金支付
    cur个人帐户支付 = CCur(sngArrInfo(13))
    cur个人现金支付 = CCur(sngArrInfo(14))
    cur补充医保 = CCur(sngArrInfo(7))
    cur医保支付 = CCur(sngArrInfo(9))
    cur统筹支付 = CCur(sngArrInfo(5))
    
    '帐户年度信息
    Call Get帐户信息(TYPE_江苏, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_江苏 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur补充医保 + cur医保支付 + cur统筹支付 & "," & int住院次数累计 + 1 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_江苏 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个人帐户支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur补充医保 + cur医保支付 + cur统筹支付 & "," & int住院次数累计 + 1 & _
            "," & cur补充医保 + cur医保支付 + cur统筹支付 & ",NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL,NULL,NULL,NULL," & _
            cur个人帐户支付 & ",NULL,NULL,NULL,'" & strBillCode & ";" & str就诊编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    住院结算_江苏 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院虚拟结算_江苏(rs费用明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim cur个人帐户支付 As Currency, cur个人现金支付 As Currency
    Dim cur统筹支付 As Currency, cur医保支付 As Currency, cur补充医保 As Currency
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, str同步 As String
    Dim datCurr As Date, str就诊编号 As String, strBillCode As String
    Dim curCount As Currency, sngArrInfo(20) As Single, cur余额 As Currency
    
    On Error Resume Next
    Kill "C:\Trans.LOG"
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始住院预结算"
    Set rs明细 = rs费用明细.Clone
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs明细.EOF
        curCount = curCount + rs明细!金额
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    str同步 = ""
reTrans:
    WriteInfo "开始传递明细"
    If 记帐传输_江苏("", 2, str同步, lng病人ID) = False Then Exit Function
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,帐户余额 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_江苏)
    str就诊编号 = rsTemp!顺序号
    cur余额 = rsTemp!帐户余额
    
    datCurr = zlDatabase.Currentdate
    
    
    
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 '以下增加出院函数调用功能2007-09-28陈云
    Dim strTemp As String, strInNote As String
    Dim rsTmp As New ADODB.Recordset
    Dim lng主页ID As Long, 主页ID As Long


    gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)

    gstrRecCode = rsTemp!顺序号


    If rsTemp.EOF Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.编号)

    ElseIf rsTemp(0) = 0 Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.编号)


    Else
        gstrSQL = "select max(主页id) 主页id from 病案主页 where 病人id=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取主页id", lng病人ID, lng主页ID)
        lng主页ID = rsTemp!主页ID
        gstrSQL = "select A.出院日期,D.名称 as 出院科室,D.编码 as 科室编码,A.出院病床,A.住院医师," & _
                "A.出院方式,C.顺序号 from 病案主页 A,保险帐户 C,部门表 D Where A.病人ID=C.病人ID " & _
                "And A.出院科室ID = D.ID And A.主页ID = [2] And A.病人ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)


        '获取出院诊断（病种编码）
        strInNote = 获取入出院诊断(lng病人ID, lng主页ID, False, True, True) '出院诊断
        If strInNote <> "" Then
            strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
        End If

        '获取住院医师代码
        gstrSQL = "Select ID,编号,姓名,简码,个人简介,接受培训 from 人员表 Where 姓名=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!住院医师))

        Select Case rsTemp!出院方式
            Case "正常"
                strTemp = "1"
            Case "转院"
                strTemp = "5"
            Case "死亡"
                strTemp = "4"
            Case "好转"
                strTemp = "2"
            Case "未愈"
                strTemp = "3"
            Case "转外"
                strTemp = "6"
            Case Else
                strTemp = "9"
        End Select

    '入口参数：住院号,操作员代码,出院日期,出院原因,ICD编码规则('A'),出院诊断(ICD10编码),出院医生代码
        intReturn = FInpatLeave(rsTemp!顺序号, UserInfo.编号, Format(Nvl(rsTemp!出院日期, Date), "yyyy-MM-dd HH:MM:SS"), _
            strTemp, "A", strInNote, IIf(rsTmp.EOF, " ", rsTmp!编号))

End If
''
    
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'入口参数：住院号,操作员工号,是否使用帐户(是/否),结算方式,报销代码,保留1,保留2,备注
'出口参数：0（'UnKnown'）,1总费用,2三个范围内费用,3自付费用,4自费费用,5起付标准,6统筹支付,7统筹自付,
'          8大病救助基金支付,9大病救助基金自付,10公务员/企业补充支付,11公务员/企业补充自付,12其它基金支付,
'          13个人医疗帐户支付,14个人储蓄帐户支付,15现金支付
    strBillCode = Space(7)
    intReturn = FTryInpatBalance(str就诊编号, UserInfo.编号, "否", 0, "IA01", 0, 0, " ", strBillCode, sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        MsgBox "住院病人预结算时发生错误。", vbInformation, gstrSysName
        住院虚拟结算_江苏 = ""
        Exit Function
    End If

    '获取个人帐户支付和个人现金支付
    cur个人帐户支付 = CCur(sngArrInfo(13) + sngArrInfo(12))
    cur个人现金支付 = CCur(sngArrInfo(14))
    cur补充医保 = CCur(sngArrInfo(7))
    cur医保支付 = CCur(sngArrInfo(9))
    cur统筹支付 = CCur(sngArrInfo(5))
'    If curCount <> CCur(sngArrInfo(0)) Then
'        MsgBox "请注意：医保返回结算金额与当前单据金额不符" & vbCrLf, vbInformation, gstrSysName
'    End If
    WriteInfo "预结算返回:" & CCur(sngArrInfo(0)) & "    医院:" & curCount
    If CCur(sngArrInfo(0)) <> curCount Then
        If MsgBox("请注意：医保返回结算金额与当前单据金额不符" & vbCrLf & "　　院方金额：" & curCount & _
            "　　　中心返回：" & CCur(sngArrInfo(0)) & vbCrLf & "是否需要进行数据同步？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intReturn = FSynData(2, str就诊编号)                '取消住院试结算
            WriteInfo "数据同步"
            str同步 = "1"
            GoTo reTrans
        End If
    End If
    
    住院虚拟结算_江苏 = "省个人帐户;" & cur个人帐户支付 & ";0" '不允许修改个人帐户
    If cur统筹支付 <> 0 Then
        住院虚拟结算_江苏 = 住院虚拟结算_江苏 & "|省统筹基金;" & cur统筹支付 & ";0" '不允许修改统筹支付
    End If
    If cur补充医保 <> 0 Then
        住院虚拟结算_江苏 = 住院虚拟结算_江苏 & "|省大病统筹;" & cur补充医保 & ";0"
    End If
    If cur医保支付 <> 0 Then
        住院虚拟结算_江苏 = 住院虚拟结算_江苏 & "|省公务员/企业补充支付;" & cur医保支付 & ";0"
    End If
    WriteInfo "完成预结"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    住院虚拟结算_江苏 = ""
End Function

Public Function 出院登记_江苏(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim strTemp As String, rsTemp As New ADODB.Recordset, datCurr As Date, strInNote As String
    Dim rsTmp As New ADODB.Recordset
    
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If rsTemp.EOF Then
        MsgBox "不能找到病人的相关信息。", vbInformation, gstrSysName
        出院登记_江苏 = False
        Exit Function
    End If
    gstrRecCode = rsTemp!顺序号
    
    gstrSQL = "Select Sum(实收金额) From 住院费用记录 Where 病人id=[1] And 主页id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.编号)
        If intReturn <> 0 Then
'            MsgBox "撤消入院登记时发生错误，未获取错误信息。", vbInformation, gstrSysName
            出院登记_江苏 = False
            Exit Function
        End If
    ElseIf rsTemp(0) = 0 Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.编号)
        If intReturn <> 0 Then
'            MsgBox "撤消入院登记时发生错误，未获取错误信息。", vbInformation, gstrSysName
            出院登记_江苏 = False
            Exit Function
        End If
    Else
        gstrSQL = "select A.出院日期,D.名称 as 出院科室,D.编码 as 科室编码,A.出院病床,A.住院医师," & _
                "A.出院方式,C.顺序号 from 病案主页 A,保险帐户 C,部门表 D Where A.病人ID=C.病人ID " & _
                "And A.出院科室ID = D.ID And A.主页ID = [2] And A.病人ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
        If rsTemp.EOF Then
            MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
            出院登记_江苏 = False
            Exit Function
        End If
        
        '获取出院诊断（病种编码）
        strInNote = 获取入出院诊断(lng病人ID, lng主页ID, False, True, True) '出院诊断
        If strInNote <> "" Then
            strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
        End If
        
        '获取住院医师代码
        gstrSQL = "Select ID,编号,姓名,简码,个人简介,接受培训 from 人员表 Where 姓名=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!住院医师))
        
        Select Case rsTemp!出院方式
            Case "正常"
                strTemp = "1"
            Case "转院"
                strTemp = "5"
            Case "死亡"
                strTemp = "4"
            Case "好转"
                strTemp = "2"
            Case "未愈"
                strTemp = "3"
            Case "转外"
                strTemp = "6"
            Case Else
                strTemp = "9"
        End Select
    
    '入口参数：住院号,操作员代码,出院日期,出院原因,ICD编码规则('A'),出院诊断(ICD10编码),出院医生代码
        intReturn = FInpatLeave(rsTemp!顺序号, UserInfo.编号, Format(Nvl(rsTemp!出院日期, Date), "yyyy-MM-dd HH:MM:SS"), _
            strTemp, "A", strInNote, IIf(rsTmp.EOF, " ", rsTmp!编号))
        If intReturn <> 0 Then
            MsgBox "进行医保病人出院登记时发生错误，未能获取错误信息。", vbInformation, gstrSysName
            出院登记_江苏 = False
            Exit Function
        End If
        
    End If
    
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_江苏 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_江苏 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_江苏 = False
End Function

Public Function 医保设置_江苏() As Boolean
    医保设置_江苏 = frmSet江苏.ShowME(TYPE_江苏)
End Function

Private Function Get病人ID(str医保号 As String, str医保中心编码 As String) As String
'功能：通过医保中心号码和医保号求出病人ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 病人ID from 保险帐户 where 险类 = [1] and 医保号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_江苏, str医保号)
    If Not rsTmp.BOF Then
        Get病人ID = CStr(rsTmp("病人ID"))
    Else
        Get病人ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get病人ID = ""
End Function

Public Function 记帐传输_江苏(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng主页ID As Long
    
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    
    If str消息 = "" Then
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B" & _
                  " Where A.门诊标志=2 And Nvl(A.实收金额,0)<>0 And A.记录状态<>0 And Nvl(A.是否上传,0)=0 And nvl(A.附加标志,0)<>9 " & _
                  " and A.病人id=[1] And A.主页id=[2]" & _
                  " and A.病人ID=B.病人ID And B.险类=[3]" & _
                  " order by A.NO,A.序号"
    Else
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B" & _
                  " Where A.门诊标志=2 And Nvl(A.实收金额,0)<>0 And A.记录状态<>0 And nvl(A.附加标志,0)<>9 " & _
                  " and A.病人id=[1] And A.主页id=[2]" & _
                  " and A.病人ID=B.病人ID And B.险类=[3]" & _
                  " order by A.NO,A.序号"
    End If
    WriteInfo "提取病人费用记录:" & gstrSQL
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSQL, lng病人ID, lng主页ID, TYPE_江苏)
    If Not rsTemp.EOF Then
        WriteInfo "上传记录:" & rsTemp.RecordCount & "条"
        记帐传输_江苏 = 费用明细传递_江苏(0, rsTemp, 2)
    Else
        记帐传输_江苏 = True
        Exit Function
    End If
    If 记帐传输_江苏 = True And rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            gstrSQL = "zl_病人记帐记录_上传 ('" & rsTemp("ID") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTemp.MoveNext
        Wend
    End If
End Function


