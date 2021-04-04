Attribute VB_Name = "mdl云南"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

'仅用于云南医保的内部门诊变量
Private mblnInit As Boolean         '是否已初始化
Private mstr顺序号 As String        '存放顺序号,仅用于门诊,住院存放于保险帐户中
Private mstr医保号 As String        '存放医保号,仅用于门诊
Private mcur帐户余额 As Double      '存放个人帐户余额,如果要用,仅用于门诊(身份验证返回)
Public mbln公务员 As Boolean       '存放公务员标志
Private mlng病人ID As Long          '存放病人ID，仅用于特殊门诊
Private mstr明细事务号 As String    '存放事务控制号，仅用于处理门诊费用明细撤消

Private mstrAverageFeeType As String
Private mstrTsyybz As String    '保险参数中的平均付费类别与特殊医院标志，每次初始化时更新

Private mstrErr As String * 4

'###医保接口函数原型，需要改写为API方式
'以下几点需注意：
'（1）字符串参数不论传入还是传出，都加上ByVal关键字；
'（2）传出的字符串参数在调用前必须初始化；
'（3）数值参数对于传入的情况是要加上ByVal关键字的，但传出的一定不能加
'（4）对于浮点参数，对应类型是Double
'（5）千万别入结构的域

'====================================================================================
'1 费用明细传递
'输入：顺序号（就诊登记号）、数据批号、收费大类编码、收费项目编码、项目名称、数量、价格（单价）、产地、规格、用法用量、经办人、科室名称、事务控制号、医生姓名；
'输出：自付比例、自付金额、允许报销金额，错误代码；

Private Declare Sub yh_feedetailtrans Lib "Hisint" Alias "int_feedetailtrans" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Charge_Category As String, _
    ByVal Charge_Item As String, ByVal Charge_Name As String, ByVal Count As Double, ByVal Price As Double, ByVal Pr_Area As String, _
    ByVal Standard As String, ByVal Usage_Dosage As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal Transaction_No As String, _
    ByVal Doctor_Name As String, ByVal Charge_Time As String, Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'2 费用结算(省市医保都调用一个函数)
'输入：顺序号（就诊登记号）、经办人、科室名称、事务控制号；
'输出：全自付金额、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、
'       医疗照顾人员的自费部分、医疗照顾人员的统筹部分、公务员统筹支付部分、人员状态、初始化机构名称、特殊挂钩支付部分、费用类型,血透次数,错误代码；
Private Declare Sub yh_feebalance Lib "Hisint" Alias "int_feebalance" _
    (ByVal Serial_No As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal SickSortCode As String, ByVal Transaction_No As String, _
    Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    OfficialPay As Double, ByVal ryzt As String, ByVal initinstitution As String, tsggzfbf As Double, _
    ByVal feebalancetype As String, xtcs As Double, ByVal error_code As String)
    
'3、费用明细更改（备注，可用来完成退费操作）
'输入：顺序号（就诊登记号）、数据批号、新的数量、新的价格、事务控制号；
'输出：自付比例、自付金额、允许报销金额、错误代码；
Private Declare Sub yh_recedefeedetail Lib "Hisint" Alias "int_recedefeedetail" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Count As Double, ByVal Price As Double, _
     ByVal Transaction_No As String, Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'4 入院登记
'输入：卡介质类型、医生姓名、医院编码、经办人、科室名称、病历号、住院号、是否特种病、特种病编码、入院时间、入院诊断、事务控制号；
'输出：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、初始化机构名称、单位编码、单位名称、错误编码；
'注：特种病编码可以为空
'2007-10-15修改:
'int_admit函数增加两个输出参数：人员状态：akc021；人群类别：akc300
'如果参保人员享受特殊病待遇，在进行特殊病门诊、住院业务办理时，必须传入：特种病标志、特种病编码，特种病标志和特种病编码同职工特慢病门诊
'详细内容见: 参数说明
'本交易用于HIS在办理病人就诊登记时，触发接口程序从中心数据库中获取人员状态、住院基数、住院限额等将来用于费用分割、费用结算的数据，同时从接口得到病人姓名、性别、年龄等用于HIS办理入院登记时要使用的IC卡中的基本信息。
'输入参数：卡介质类型、医院编码、经办人、科室名称、病历号、住院号、是否特种病、特种病编码、入院时间、入院诊断、事务控制号；
'输出参数：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、人员状态、人群类别、初始化机构名称、错误编码；
'2008-3-10修改:
'int_admit函数增加输出参数：yck002（特殊人群标志 0为否，1为是）
'输入参数：卡介质类型、医院编码、经办人、科室名称、病历号、住院号、是否特种病、特种病编码、入院时间、入院诊断、事务控制号；
'输出参数：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、人员状态、人群类别、特殊人群标志、初始化机构名称、错误编码；
Private Declare Sub yh_admit Lib "Hisint" Alias "int_admit" _
    (ByVal card_mode As String, ByVal doctorname As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_Name As String, _
    ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, ByVal specialsick_no As String, _
    ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, ByVal Serial_No As String, ByVal CARD_NO As String, _
    ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal akc021 As String, ByVal akc300 As String, ByVal yck002 As String, ByVal initinstitution As String, ByVal dwbm As String, ByVal dwmc As String, ByVal error_code As String)

'年终结转入院
'2007-10-15修改:
'int_kndadmit()函数增加两个输出参数：人员状态：akc021；人群类别：akc300
'详细内容见: 参数说明
'本交易适用于通过跨年度结算出院的病人，新年度办理入院时使用。
'输入参数：医生姓名，个人编号，医院编码，经办人，科室名称，病历号，住院号，是否特种病，特种病编码，入院时间，入院诊断，事务控制号
'输出参数：顺序号，IC卡号，姓名，性别，出生日期，人员状态，人群类别，初始化机构代码，单位编码，错误编码
'2008-3-10修改:
'int_admit函数增加输出参数：ykc002（特殊人群标志 0为否，1为是）
Private Declare Sub yh_kndadmit Lib "Hisint" Alias "int_kndadmit" _
    (ByVal doctorname As String, ByVal Personal_No As String, ByVal Hospital_No As String, ByVal Arranger As String, _
    ByVal Section_Name As String, ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, _
    ByVal specialsick_no As String, ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, _
    ByVal Serial_No As String, ByVal CARD_NO As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal akc021 As String, ByVal akc300 As String, ByVal ykc002 As String, ByVal initinstitution As String, ByVal dwbm As String, ByVal dwmc As String, ByVal error_code As String)

'5 IC卡支付
'输入：卡介质类型、顺序号（就诊登记号）、经办人、支付原因,支付金额；
'输出：初始化机构名称、错误代码；
Private Declare Sub yh_cardpay Lib "Hisint" Alias "int_cardpay" _
    (ByVal card_mode As String, ByVal Serial_No As String, ByVal Arranger As String, ByVal Pay_reason As String, ByVal Pay_amount As Double, _
     ByVal initinstitution As String, ByVal error_code As String)


'6 虚拟结算
'输入、输出参数、使用场合和时间与费用结算相同。
'输入：顺序号（就诊登记号）、预结算标志、结算编号、事务控制号；
'输出：全自付金额、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、
'       医疗照顾人员的自费部分、医疗照顾人员的统筹部分、公务员统筹支付、人员状态、初始化机构名称、特殊挂钩支付部分、错误代码；
'注意：预结算标志          0 表示虚拟结算，在医保中心没有任何记录；1  表示预结算，可以作为中途结算使用
'      医疗照顾人员金额    如果不为空，那只有这两个字段有效。

Private Declare Sub yh_virtualbalance Lib "Hisint" Alias "int_virtualbalance" _
    (ByVal Serial_No As String, ByVal ForeBalance_Flag As String, ByVal balance_no As String, ByVal Transaction_No As String, _
    Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    OfficialPay As Double, ByVal ryzt As String, ByVal initinstitution As String, tsggzfbf As Double, ByVal error_code As String)

'7 门诊身份识别
'输入：卡介质类型、医生姓名、医院编码、经办人、科室名称、病历号、门诊号、就医时间；
'输出：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、初始化机构名称、卡余额、错误编码；
'2007-10-15修改:
'int_outpatientidentify()函数增加两个输出参数：人员状态：akc021；人群类别：akc300
'详细内容见: 参数说明
'本交易的目的是病人在门诊结算前从IC卡中读出基本信息给HIS或者从市医保中心数据库中获取病人的基本信息，必要时从中心数据库取得人员状态，并将卡基本信息返回HIS。
'输入：卡介质类型、医院编码、经办人、科室、病历号、门诊号、就医时间、就诊诊断、事务控制号；
'输出：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、人员状态、人群类别、初始化机构名称、卡余额、错误编码；
Private Declare Sub yh_outpatientidentify Lib "Hisint" Alias "int_outpatientidentify" _
    (ByVal card_mode As String, ByVal doctorname As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_No As String, _
    ByVal anamnesis_No As String, ByVal outpatient_No As String, ByVal hospitalize_time As String, _
    ByVal admit_diagnose As String, ByVal Transaction_No As String, ByVal Serial_No As String, _
    ByVal CARD_NO As String, ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal akc021 As String, ByVal akc300 As String, ByVal initinstitution As String, accountremain As Double, ByVal officesign As String, ByVal error_code As String)

'8 IC卡基本信息查询
'输入：卡介质类型；
'输出: 余额、卡号、姓名、性别、身份证号、年龄、错误代码
'2007-10-15修改:
'int_cardinfo()函数，增加AKC300人群类别输出参数：1、职工；2、居民。
'本交易用于HIS查询个人基本信息?
'输入：卡介质类型；
'输出：余额、卡号、姓名、性别、身份证号、年龄、人群类别，错误代码
Private Declare Sub yh_cardinfo Lib "Hisint" Alias "int_cardinfo" _
    (ByVal Code_Mode As String, Amount As Double, ByVal CARD_NO As String, ByVal Name As String, _
    ByVal Sex As String, ByVal Identify As String, age As Double, ByVal akc300 As String, ByVal error_code As String)

'9 密码更改
'输入: 卡介质类型
'输出: 错误代码
Private Declare Sub yh_changepassword Lib "Hisint" Alias "int_changepassword" _
    (ByVal Code_Mode As String, ByVal error_code As String)

'10    个人帐户支出查询
'输入：顺序号；
'输出：已支付总额，错误代码
Private Declare Sub yh_accountpay Lib "Hisint" Alias "int_accountpay" _
    (ByVal Serial_No As String, Amount As Double, ByVal error_code As String)

'11    门诊帐户支付
'输入：卡介质类型、医院编码、科室名称、经办人、支付原因、费用总额、帐户支付额；
'输出：初始化机构名称、顺序号、错误代码；
Private Declare Sub yh_outpay Lib "Hisint" Alias "outpay" _
    (ByVal card_mode As String, ByVal Hospital_No As String, ByVal Section_No As String, ByVal Arranger As String, ByVal payreason As String, _
    ByVal Amount As Double, ByVal accountpay As Double, ByVal initinstitution As String, ByVal Serial_No As String, ByVal error_code As String)

'12    初始化
'输入: 无
'输出: 错误代码
Private Declare Sub yh_init_yns Lib "Hisint" Alias "init" _
    (ByVal Errcode As String)

'13    断开连接
'输入：无
'输出: 无
Public Declare Sub yh_quit Lib "Hisint" Alias "quit" ()

'14 IC卡圈存
'输入：无
'输出: 错误代码
Private Declare Sub yh_loadcard Lib "Hisint" Alias "int_loadcard" (ByVal error_code As String)
    
'15 数据传输
'输入：无
'输出: 错误代码
Private Declare Sub yh_datatrans Lib "Hisint" Alias "int_datatrans" (ByVal error_code As String)


'16 事务控制
'输入：交易类别，就诊顺序号，事务控制号，事务控制类型；
'输出: 错误代码
Private Declare Sub yh_transaction Lib "Hisint" Alias "int_transaction" _
    (ByVal Trade_Sort As String, ByVal Serial_No As String, ByVal Transaction_No As String, ByVal Affirm_Mode As String, ByVal error_code As String)

'17 获取事务控制号
'输入：无；
'输出: 事务控制号
Private Declare Sub yh_gettranssequence Lib "Hisint" Alias "int_gettranssequence" (ByVal Transaction_No As String)

'18    待遇变更分段费用查询
'输入参数：顺序号；
'输出参数：分段标准、分段序号、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、专项补助款支付额、错误代码；
Private Declare Sub yh_SubsecFee Lib "Hisint" Alias "int_SubsecFee" _
    (ByVal Serial_No As String, ByVal Standard_Subsec As String, ByVal Subsec_No As String, _
      Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, _
      Basepay As Double, outpay As Double, Preqpay As Double, Preqselfpay As Double, _
      SubsidyPay As Double, ByVal error_code As String)

'19 退费处理
'输入参数：顺序号，回退标志，结算编号，事务控制号；
'输出参数: 错误码
Private Declare Sub yh_recedefeebalance Lib "Hisint" Alias "int_recedefeebalance" _
    (ByVal Serial_No As String, ByVal return_flag As String, ByVal balance_no As String, ByVal Transaction_No As String, _
        ByVal error_code As String)

'删除所有未执行结算或预结算前的费用明细。如果数据只是做了虚拟结算，仍会被删除
Private Declare Sub yh_rollbackdetail Lib "Hisint" Alias "int_rollbackdetail" _
    (ByVal Serial_No As String, ByVal error_code As String)

'查询某次结算后病人统筹累计,基本统筹支付限额，大病统筹支付限额等信息
'输入参数：顺序号；
'输出参数: 起付线，统筹累计，基本统筹支付限额，大病统筹支付限额，基数累计，起付线信息，审批标志编码，用药限制，错误代码
Private Declare Sub yh_RyspInfo Lib "Hisint" Alias "int_RyspInfo" _
   (ByVal series_no As String, qfx As Double, tclj As Double, dczfxe As Double, _
    dbxe As Double, jslj As Double, ByVal qfxinfo As String, ByVal spbzbm As String, ByVal yyxz As String, ByVal error_code As String)

'本交易作用是出院办理时，修改出院诊断、出院时间时调用。
'输入：顺序号、出院原因、出院时间、出院诊断、出院经办人、出院科室、出院床位；
'输出：错误编码；
Private Declare Sub yh_ReLeaveHosInfo Lib "Hisint" Alias "int_ReLeaveHosInfo" _
   (ByVal series_no As String, ByVal Cyyy As String, ByVal Cysj As String, ByVal Cyzd As String, _
   ByVal Cyjbr As String, ByVal Cyks As String, ByVal Cycw As String, ByVal error_code As String)

'仅应用于市医保
'病种范围解析函数int_sicksortchk（），该函数为新增函数，HIS接口在调用int_feebalance（）函数进行正式结算前必须调用该函数
'（即int_sicksortchk（）后紧接着调用int_feebalance（）函数）。该函数实现功能是完成费用明细对应的单病种范围的解析。
'把所有明细对应的单病种反给前台。病种编码之间通过'$'分隔。如0101$0102$0103$0104。HIS前台只有在该函数返回的病种范围内进行病种选择，
'如果为空则表明该患者不能进行单病种结算。目前只有特殊病门诊和住院结算调用。如果HIS不严格控制病种选择范围，
'如果传到中心的病种与病种消费明细不匹配，中心审核发现收费明细与病种结算收费项目不符情况，后果由HIS开发商和医院承担。
Private Declare Sub yh_sicksortchk Lib "Hisint" Alias "int_sicksortchk" _
    (ByVal Serial_No As String, ByRef sicksorts As String, ByRef error_code As String)

Private Declare Sub yh_init_kms Lib "Hisint" Alias "init" _
    (ByVal HospNO As String, ByVal AverageFeeType As String, ByVal Tsyybz As String, ByVal Errcode As String)

'昆明市医保才有：预警函数
Private Declare Sub yh_AlertInfo_kms Lib "Hisint" Alias "int_alertinfo" _
    (ByVal SerialNO As String, ByRef ErrorCode As String, ByRef ErrMsg As String)

Public Const gint昆明市 As Integer = 31

'以下结构体用于纪录虚拟结算结果，以便在结算时核对
Private Type typBalance
    cur个人帐户 As Double
    cur医保基金 As Double
    cur大病统筹 As Double
    cur公务员补助 As Double
    cur特殊补助 As Double
End Type
Private pre_Balance As typBalance

Public Function 医保初始化_云南(ByVal intinsure As Integer) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim strAverageFeeType As String, strTsyybz As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mblnInit Then
        医保初始化_云南 = True
        Exit Function
    End If
    
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & intinsure
    Call OpenRecordset(rsTemp, "获取医院编码")
    gstr医院编码 = Nvl(rsTemp!医院编码, "")
    
    mstrErr = Space(4)
    If intinsure <> gint昆明市 Then
        Call yh_init_yns(mstrErr)
    Else
        Call yh_init_kms(gstr医院编码, mstrAverageFeeType, mstrTsyybz, mstrErr)
    End If
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbExclamation, gstrSysName
    Else
        mblnInit = True
        医保初始化_云南 = True
    End If
    
    '将返回的费用类别与特殊医院标志保存到保险帐户中，这两个标志都将在结算时原样复制到保险结算记录中
    gstrSQL = "zl_保险参数_Insert(" & intinsure & ",0,'平均付费类别','''" & mstrAverageFeeType & "''',10)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新平均付费类别")
    gstrSQL = "zl_保险参数_Insert(" & intinsure & ",0,'特殊医院标志','''" & mstrTsyybz & "''',11)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新特殊医院标志")
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_云南() As Boolean
    Call yh_quit
    mblnInit = False
End Function

Public Function 身份标识_云南(Optional bytType As Byte, Optional lng病人ID As Long, Optional ByVal intinsure As Integer) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str卡号 As String, STR姓名 As String, str性别 As String
    Dim str身份证号 As String, str出生日期 As String, lng年龄 As Double, str单位编码 As String, str单位名称 As String
    Dim str初始化机构 As String, str事务号 As String, str人群类别 As String, str人员状态 As String, str特殊人群标志 As String
    Dim str就诊诊断 As String, str事务控制号 As String, str公务员 As String, str医保号 As String
    Dim str历史顺序号 As String
    
    Dim strArranger As String
    Dim strSection As String
    Dim strPatiNo As String
    
    Dim str卡类型 As String, lng病种ID As Long, str疾病编码 As String
    Dim rsTemp As New ADODB.Recordset
    Dim ybhcf As New ADODB.Recordset    '用于记录更新卡号
    Dim dat当前 As Date
    Dim strIdentify As String, str附加 As String
    '---------特殊门诊使用--------
    Dim str诊断编码 As String, str诊断名称 As String, int病种类别 As Integer
    '-----------------------------
    
    On Error GoTo errHandle
    '初始化几个全局的变量
    mstr医保号 = Space(20)
    mstr顺序号 = Space(19)
    mcur帐户余额 = 0
    
    str卡号 = Space(18)
    STR姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    str人群类别 = Space(3)
    str人员状态 = Space(3)
    str出生日期 = Space(10)
    str初始化机构 = Space(4)
    str就诊诊断 = Space(56)
    str事务控制号 = Space(18)
    str公务员 = Space(4)
    dat当前 = zlDatabase.Currentdate
    
    If frmIdentify云南.GetIdentifyMode(intinsure, bytType, str卡类型, lng病种ID, str疾病编码) = False Then
        Exit Function
    End If
    DoEvents
        
    '门诊身份证验
    '返回的本次交易的顺序号放在:mstr顺序号,在交易时使用
    '返回的余额存放在mcur帐户余额中，在取余额时使用
    
    '读取IC卡信息
    strArranger = LeftDB(UserInfo.姓名, 8)
    strSection = LeftDB(UserInfo.部门, 24)
    strPatiNo = LeftDB(UserInfo.编号, 12)
    
    Screen.MousePointer = vbHourglass
    mstrErr = Space(4)
    '获取事务控制号 gzh
    str事务控制号 = Get事务号()
    If str事务控制号 = "" Then Exit Function
    If bytType = 0 Then
        '适用：昆明市、云南省；普通门诊才调OutPatientidentifhy，特殊门诊调CardInfo
        If lng病种ID = 0 Then
            Call yh_outpatientidentify(str卡类型, strArranger, gstr医院编码, strArranger, strSection, strPatiNo, _
                strPatiNo, Format(dat当前, "yyyy-MM-dd"), str就诊诊断, str事务控制号, mstr顺序号, str卡号, _
                mstr医保号, STR姓名, str性别, str出生日期, str身份证号, str人员状态, str人群类别, str初始化机构, mcur帐户余额, str公务员, mstrErr)
        Else
            Call yh_cardinfo(str卡类型, mcur帐户余额, str卡号, STR姓名, str性别, str身份证号, lng年龄, str人群类别, mstrErr)
        End If
    Else
        Call yh_cardinfo(str卡类型, mcur帐户余额, str卡号, STR姓名, str性别, str身份证号, lng年龄, str人群类别, mstrErr)
    End If
    If mstrErr <> "0000" Then
        Screen.MousePointer = vbDefault
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr顺序号 = TrimStr(mstr顺序号)
    str卡号 = TrimStr(str卡号)
    STR姓名 = TrimStr(STR姓名)
    str身份证号 = TrimStr(str身份证号)
    str人群类别 = TrimStr(str人群类别)
   ' str人员状态 = TrimStr(str人员状态)

    If bytType = 0 And lng病种ID = 0 Then
        '只有普通门诊才能得到医保号，其它交易调用CardInfo函数，无法得到医保号
        mstr医保号 = TrimStr(mstr医保号)
    Else
        '因为住院未返回医保号，只有从数据库中取，如果没取到，则将卡号做为医保号保存，在入院时再更新
        gstrSQL = "Select 医保号 From 保险帐户 Where 险类=" & intinsure & " And 卡号='" & str卡号 & "'"
        Call OpenRecordset(rsTemp, "获取原医保号")
        If Not rsTemp.EOF Then
            mstr医保号 = Nvl(rsTemp!医保号)
        End If
        If Trim(mstr医保号) = "" Then
            mstr医保号 = str卡号
        Else
            mstr医保号 = Mid(mstr医保号, 2)
        End If
    End If
    str医保号 = TrimStr(mstr医保号)
    mbln公务员 = (TrimStr(str公务员) = "1")
    
    If bytType = 0 And lng病种ID = 0 Then
        '只有普通门诊通过调用outpatientidentify接口得到顺序号
        If mstr顺序号 = "" Then
            MsgBox "未能从前置服务器获得顺序号,请重试或检查卡。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If str卡号 = "" Then
        MsgBox "未能从卡中读取卡号,请重试或检查卡。", vbInformation, gstrSysName
        Exit Function
    End If
    If mstr医保号 = "" Then
        MsgBox "未能从卡中读取医保号,请重试或检查卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln公务员 Then
        '如果是公务员，需要调用yh_RyspInfo获取审批信息
        Dim cur起付线 As Double, cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double, cur基数累计 As Double
        Dim str起付线信息 As String, str审批标志编码 As String, str用药限制 As String
        Call yh_RyspInfo(mstr顺序号, cur起付线, cur统筹累计, cur基本统筹限额, cur大额统筹限额, cur基数累计, str起付线信息, str审批标志编码, str用药限制, mstrErr)
        cur起付线 = strVal(cur起付线)
        cur统筹累计 = strVal(cur统筹累计)
        cur基本统筹限额 = strVal(cur基本统筹限额)
        cur大额统筹限额 = strVal(cur大额统筹限额)
        cur基数累计 = strVal(cur基数累计)
        str起付线信息 = strVal(str起付线信息)
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'起付线','''" & cur起付线 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新公务员的起付线")
        gstrSQL = "zl_病人审批信息_insert(1," & lng病人ID & "," & intinsure & "," & Year(dat当前) & ",'" & _
        mstr顺序号 & "'," & cur起付线 & "," & cur统筹累计 & "," & _
        cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur基数累计 & ",'" & str起付线信息 & "','" & str审批标志编码 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病人审批信息")
    End If
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    mstr医保号 = str卡类型 & Left(mstr医保号, 19)
    strIdentify = str卡号 & ";" & mstr医保号 & ";;" & TrimStr(STR姓名) & ";" & TrimStr(str性别) & ";" & TrimStr(str出生日期) & ";" & TrimStr(str身份证号) & ";"
    strIdentify = Replace(strIdentify, " ", "")
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    ';8中心;9.顺序号;10人员身份(在职、退休、学生儿童、中学生、大学生、成年人);11帐户余额;12当前状态;13病种ID;14在职(0,1,2);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计;23就诊类型 (1、急诊门诊)
        
    '不区分门诊与住院，那就不能使用新的顺序号。而不能用以前的
    gstrSQL = "select 顺序号 from 保险帐户 where 险类=" & intinsure & " and 卡号='" & str卡号 & "'"
    Call OpenRecordset(rsTemp, "云南医保")
    If rsTemp.RecordCount > 0 Then
        str历史顺序号 = Nvl(rsTemp("顺序号"))
    End If
    If bytType = 2 Then mstr顺序号 = str历史顺序号
    
    If IsDate(str出生日期) = True Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), dat当前)
    End If
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";" & str历史顺序号             '9.顺序号
    str附加 = str附加 & ";"                             '10人员身份
    str附加 = str附加 & ";" & mcur帐户余额              '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";" & IIf(lng病种ID <> 0, lng病种ID, "") '13病种ID
    str附加 = str附加 & ";1"                            '14在职(1,2,3)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & mcur帐户余额              '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";"                             '23就诊类型 (1、急诊门诊)
    
    '-------------------------------------------------------------------------------
    '处理卡号医保号重复情况(2007-6-8)
    On Error Resume Next
    Err = 0
    gstrSQL = "update 保险帐户 set 卡号='" & str卡号 & "' where 病人ID in(select 病人ID from 保险帐户 where substr(医保号,2)='" & str医保号 & "') "
    Call OpenRecordset(rsTemp, "更新卡号,医保号")
    If Err <> 0 Then
        gstrSQL = "update 医保病人档案 set 卡号='" & str卡号 & "' where 病人ID in(select 病人ID from 保险帐户 where substr(医保号,2)='" & str医保号 & "') "
        Call OpenRecordset(rsTemp, "提交")
    End If
    On Error GoTo errHandle
    
    '----------------------------------------
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID, intinsure)
    If lng病人ID = 0 Then Exit Function '未建立正确的保险帐户
    
    If bytType = 0 And lng病种ID > 0 Then
        '如果是特殊病、慢性病门诊，同时进行就诊登记
        
        '再次初始化变量
        mstr医保号 = Space(20)
        str卡号 = Space(18)
        STR姓名 = Space(60)
        str性别 = Space(3)
        str身份证号 = Space(20)
        str人员状态 = Space(3)
        str人群类别 = Space(3)
        str特殊人群标志 = Space(3)
        str出生日期 = Space(10)
        str初始化机构 = Space(4)
        mstr顺序号 = Space(19)
        
        str事务号 = Get事务号
        If str事务号 = "" Then
            Exit Function
        End If
        
        '取该病种的类别，如果是慢特病就传1
        gstrSQL = "Select Nvl(类别,0) 类别 From 保险病种 Where ID=" & lng病种ID
        Call OpenRecordset(rsTemp, "取病种类别")
        int病种类别 = rsTemp!类别
        
        '只有门诊慢特病需要获取病人的门诊诊断
        If int病种类别 <> 0 Then
            Call frm诊断信息.ShowME(lng病人ID, str诊断编码, str诊断名称, True)
        End If
        If str诊断名称 = "" Then str诊断名称 = "普通"
        'str诊断名称 = "唇恶性肿瘤" '测试时用
        '0092-特殊群体门诊和0094-住院都必须设置成普通病
        mstrErr = Space(4)
        Call yh_admit(str卡类型, LeftDB(UserInfo.姓名, 8), gstr医院编码, LeftDB(UserInfo.姓名, 8), "门诊", _
            LeftDB(lng病人ID, 12), LeftDB(lng病人ID, 12), IIf(Val(rsTemp!类别) = 0, "0", "1"), LeftDB(str疾病编码, 8), _
            Format(dat当前, "yyyy-MM-dd HH:mm:ss"), str诊断名称, str事务号, mstr顺序号, str卡号, _
            mstr医保号, STR姓名, str性别, str出生日期, str身份证号, str人员状态, str人群类别, str特殊人群标志, str初始化机构, str单位编码, str单位名称, mstrErr)
        
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            '医保数据库回滚
            Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
            Exit Function
        End If
        mstr顺序号 = TrimStr(mstr顺序号) '1、用于门诊预算
        If mstr顺序号 = "" Then
            MsgBox "不能得到正确的入院登记顺序号。", vbInformation, gstrSysName
            Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
            Exit Function
        End If
        Call yh_transaction("0", mstr顺序号, str事务号, "1", mstrErr)
        
        str特殊人群标志 = TrimStr(str特殊人群标志)
        If str特殊人群标志 = 1 Then
            On Error Resume Next
            Err = 0
            gstrSQL = "update 保险帐户 set 特殊人群标志=" & str特殊人群标志 & "  where 病人ID=" & lng病人ID
            Call OpenRecordset(rsTemp, "更新人员身份")
            If Err <> 0 Then
                gstrSQL = "update 医保病人档案 set 特殊人群标志=" & str特殊人群标志 & "  where 病人ID=" & lng病人ID
                Call OpenRecordset(rsTemp, "更新人员身份")
            End If
            On Error GoTo errHandle
        End If
        
        '特殊群体门诊,住院必须要调用yh_RyspInfo获取审批信息
        Call yh_RyspInfo(mstr顺序号, cur起付线, cur统筹累计, cur基本统筹限额, cur大额统筹限额, cur基数累计, str起付线信息, str审批标志编码, str用药限制, mstrErr)
        cur起付线 = strVal(cur起付线)
        cur统筹累计 = strVal(cur统筹累计)
        cur基本统筹限额 = strVal(cur基本统筹限额)
        cur大额统筹限额 = strVal(cur大额统筹限额)
        cur基数累计 = strVal(cur基数累计)
        str起付线信息 = strVal(str起付线信息)
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'起付线','''" & cur起付线 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新起付线")
        gstrSQL = "zl_病人审批信息_insert(1," & lng病人ID & "," & intinsure & "," & Year(dat当前) & ",'" & _
        mstr顺序号 & "'," & cur起付线 & "," & cur统筹累计 & "," & _
        cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur基数累计 & "," & str起付线信息 & ",'" & str审批标志编码 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "病人审批信息")
        mstr医保号 = str卡类型 & Left(TrimStr(mstr医保号), 19) '2、用于门诊预算
        str卡号 = TrimStr(str卡号)
        
        '可能已存在记录
        gstrSQL = " Select 病人ID,顺序号 from 保险帐户 Where 险类=" & intinsure & " And 医保号='" & mstr医保号 & "'"
        Call OpenRecordset(rsTemp, "判断是否已存在")
        If rsTemp.RecordCount = 0 Then
            '强制把登记顺序号、及新的医保号填入（保险帐户中的顺序号只保存住院，门诊使用mstr顺序号）
            gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & intinsure & _
                        ",'" & str卡号 & "','" & mstr医保号 & "','" & str历史顺序号 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
        Else
            '因以前没有记录，所以没有找到顺序号，因此只更新卡号即可
            gstrSQL = "ZL_保险帐户_修改医保号(" & rsTemp!病人ID & "," & intinsure & ",'" & str卡号 & "','" & mstr医保号 & "','" & Nvl(rsTemp!顺序号) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
            lng病人ID = rsTemp!病人ID
        End If
    End If
    '得到费用明细传递的事务控制号，以便于多次重试
    If bytType = 0 Then
        mstr明细事务号 = Get事务号 '3、用于门诊结算
        If mstr明细事务号 = "" Then
            Exit Function
        End If
    End If
    
    On Error Resume Next
    Err = 0
    If str人员状态 <> "" Then
        str人员状态 = TrimStr(str人员状态)
        '更新保险帐户中的人员身份:
        'str人员状态 = Decode(str人员状态, "11", "在职", "21", "退休", "61", "学生儿童", "62", "中学生", "63", "大学生", "64", "成年人")
        gstrSQL = "update 保险帐户 set 人员身份=" & str人员状态 & "  where 病人ID=" & lng病人ID
        Call OpenRecordset(rsTemp, "更新人员身份")
        If Err <> 0 Then
            gstrSQL = "update 医保病人档案 set 人员身份=" & str人员状态 & "  where 病人ID=" & lng病人ID
            Call OpenRecordset(rsTemp, "更新人员身份")
        End If
    End If
    
    '更新保险帐户中的在职:
    'str人群类别 = Decode(str人群类别, 1, "城镇职工", 2, "城镇居民", 3, "离休")
    Err = 0
    gstrSQL = "update 保险帐户 set 在职=" & str人群类别 & " where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "更新在职")
    If Err <> 0 Then
        gstrSQL = "update 医保病人档案 set 在职=" & str人群类别 & "  where 病人ID=" & lng病人ID
        Call OpenRecordset(rsTemp, "更新在职")
    End If
    
    On Error GoTo errHandle
    '返回格式:中间插入病人ID
    mlng病人ID = lng病人ID '4、用于门诊预算
    身份标识_云南 = strIdentify & ";" & lng病人ID & str附加
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_云南(lng病人ID As Long, strSelfNo As String, ByVal bytPlace As Byte, ByVal intinsure As Integer) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'      表示调用位置：10-门诊,20-入院,30-预交,40-结算
'返回: 返回个人帐户余额的金额
    Dim cur余额 As Currency
    On Error GoTo errHandle
    
    If Not (strSelfNo = mstr医保号 And (bytPlace = 10 Or bytPlace = 20)) Then
        Call Get卡余额(strSelfNo, cur余额, intinsure)
        mcur帐户余额 = cur余额
    End If
    '直接利用上次身份识别时得到的数据返回
    个人余额_云南 = mcur帐户余额
    
    '更新保险帐户中的帐户余额
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'帐户余额','" & mcur帐户余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_云南(rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str数据批号 As String, strTemp As String
    
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String, str规格 As String, str产地 As String
    Dim cur发生费用 As Currency, dbl金额 As Double, dbl数量 As Double, str疾病编码 As String
    Dim str发生时间 As String, str病种类别 As String, str大类 As String
    Dim rsTemp As New ADODB.Recordset
    
    If rs明细.EOF = True Then
        MsgBox "请输入费用明细再进行医保预算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rs明细("病人ID") <> mlng病人ID Then
        MsgBox "该病人未通过身份验证，不能进行预结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With pre_Balance
        .cur大病统筹 = 0
        .cur个人帐户 = 0
        .cur公务员补助 = 0
        .cur特殊补助 = 0
        .cur医保基金 = 0
    End With
    
    '只有特殊门诊才使用本函数
    On Error GoTo errHandle
    '判断该病人是否属于特殊门诊
    gstrSQL = "select nvl(A.病种ID,0) 病种ID,Nvl(B.类别,0) 类别 from 保险帐户 A,保险病种 B where A.病人ID=" & mlng病人ID & " And A.病种ID=B.ID(+) and A.险类=" & intinsure
    Call OpenRecordset(rsTemp, "医保接口")
    If rsTemp.EOF Then
        '非公务员提示 gzh
        If mbln公务员 = False Then
        '有特殊病的病人需要预算
            MsgBox "该病人不需要进行预算。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        str病种类别 = rsTemp!类别
    End If
    
    '删除前置服务器的所有未结明细
    mstrErr = Space(4)
    Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
            
    '费用明细传递
    strTemp = rs明细("病人ID") & "_" & Format(zlDatabase.Currentdate, "ddHHmmss")
    Do Until rs明细.EOF
        gstrSQL = "select A.名称,A.编码,A.类别,A.计算单位,B.项目编码,B.附注" & _
                    " ,Decode(Sign(Instr(A.规格,'┆')),0,A.规格,Substr(A.规格,1,Instr(A.规格,'┆')-1)) as 规格" & _
                    " ,Decode(Sign(Instr(A.规格,'┆')),0,A.规格,Substr(A.规格,Instr(A.规格,'┆')+1)) as 产地" & _
                    " from 收费细目 A,保险支付项目 B where A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID and B.险类=" & intinsure
        Call OpenRecordset(rsTemp, "门诊预算")
        If rsTemp.EOF = True Then
            MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
            Exit Function
        End If
        str大类 = Nvl(rsTemp!附注)
        If str病种类别 = "1" Then
            If str大类 <> "01" And str大类 <> "02" And str大类 <> "90" Then
                MsgBox "慢病医保病人只能使用药品。", vbInformation, gstrSysName
                mstrErr = Space(4)
                Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
                Exit Function
            End If
        End If
        
        str医生 = LeftDB(UserInfo.姓名, 8)
        str规格 = LeftDB(IIf(IsNull(rsTemp("规格")), "无规格", rsTemp("规格")), 30)
        str产地 = LeftDB(IIf(IsNull(rsTemp("产地")), " ", rsTemp("产地")), 30)
        str科室 = LeftDB(UserInfo.部门, 24)
        '不能传递负数，传0的目的是为了删除已经上传但被冲销的费用记录
        dbl数量 = Val(IIf(rs明细("数量") > 0, rs明细("数量"), 0))
        If Nvl(rs明细!实收金额, 0) > 0 Then
            dbl金额 = Round(rs明细!实收金额 / rs明细!数量, 3)
        Else
            dbl金额 = 0
        End If
        str发生时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        str数据批号 = ToVarchar(strTemp & "_" & rs明细.AbsolutePosition, 18)
        
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr顺序号, str数据批号, str大类, rsTemp("项目编码"), _
            rsTemp("名称"), dbl数量, dbl金额, str产地, str规格, " ", str医生, str科室, mstr明细事务号, str医生, str发生时间, _
            cur自付比例, cur自付金额, cur报销金额, mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            '医保数据库回滚
            Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
            Exit Function
        End If
        
        cur发生费用 = cur发生费用 + rs明细("实收金额")
        rs明细.MoveNext
    Loop
        
    '虚拟结算
    Dim str结算标志 As String, cur病人自费 As Double, cur余额 As Currency
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    Dim str结算事务号 As String, str人员状态 As String, cur特殊挂钩支付部分 As Double
    
    '用于门诊预算
    str结算事务号 = Get事务号
    If str结算事务号 = "" Then
        Exit Function
    End If
    
    str初始化机构 = Space(4)
    mstrErr = Space(4)
    str结算标志 = "0" '虚拟结算
    Call yh_virtualbalance(mstr顺序号, str结算标志, "", str结算事务号, cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, _
        cur超限自付, cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, str人员状态, str初始化机构, cur特殊挂钩支付部分, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    
    cur余额 = 个人余额_云南(mlng病人ID, mstr医保号, 10, intinsure)
    If cur特殊人员统筹 > 0 Then
        cur病人自费 = cur特殊人员自付
    Else
        cur病人自费 = cur全自付 + cur挂钩自付 + cur基数自付 + cur统筹自付 + cur大病自付 + cur超限自付 - cur公务员统筹
    End If
    cur余额 = IIf(cur余额 > cur病人自费, cur病人自费, cur余额) '取两者的小值
        
    str结算方式 = "个人帐户;" & cur余额 & ";1" '允许修改
    
    If cur统筹支付 <> 0 Then
        str结算方式 = str结算方式 & "|医保基金;" & cur统筹支付 & ";0"
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "|大病统筹;" & cur大病统筹 & ";0"
    End If
    If cur公务员统筹 <> 0 Then
        str结算方式 = str结算方式 & "|公务员补助;" & cur公务员统筹 & ";0"
    End If
    If cur特殊人员统筹 > 0 Then
        str结算方式 = str结算方式 & "|特殊补助;" & cur特殊人员统筹 & ";0"
    End If
    
    With pre_Balance
        .cur大病统筹 = cur大病统筹
        .cur医保基金 = cur统筹支付
        .cur公务员补助 = cur公务员统筹
        .cur特殊补助 = cur特殊人员统筹
        .cur个人帐户 = cur余额
    End With
    
    门诊虚拟结算_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_云南(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, ByVal intinsure As Integer, ByRef strAdvance As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, lng病种ID As Long
    Dim i As Long, curDate As Date, cur发生费用 As Currency, lng病人ID As Long
    Dim str卡类型 As String
    Dim str结算事务号 As String   '事务控制号
    Dim str初始化机构 As String
    
    '单病种相关
    Dim strSicks As String          '银海返回的病种列表，注：门诊特殊病目前只能选择血透
    Dim strFeeBalanceType As String, dblXTCS As Double      '结算函数的返回出参：支付费用类别、血透次数
    Dim strSickSel As String        '操作员选择的病种编码
    Dim bln单病种 As Boolean
    
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String
    Dim str规格 As String, str产地 As String
    Dim str大类 As String, str顺序号 As String
    
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double
    Dim cur特殊人员统筹 As Double, cur特殊人员自付 As Double, cur公务员统筹 As Double, str人员状态 As String, cur特殊挂钩支付部分 As Double
    Dim str发生时间 As String, str出院诊断 As String, str出院原因 As String, str出院经办人 As String, str出院科室 As String, str出院床号 As String
    Dim strErrMsg As String, str特殊人群标志 As String
    Dim blnReverse As Boolean   '校正数据标志
    Dim str结算方式 As String
    On Error GoTo errHandle
    
    Call DebugTool("门诊收费")
    '此时所有收费细目必然有对应的医保编码
    gstrSQL = "Select A.ID,A.病人ID,A.NO,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),3) as 实际价格,A.结帐金额," & _
            "   A.收费类别,D.项目编码 as 收费项目,D.附注,B.名称 as 项目名称," & _
            "   decode(Instr(B.规格,'┆'),0,B.规格,substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "   decode(Instr(B.规格,'┆'),0,'',substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "   C.名称 as 科室名称" & _
            " From (Select * From 门诊费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0) A,收费细目 B,部门表 C,保险支付项目 D " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID And D.险类=" & intinsure & _
            " Order by A.ID"
    Call OpenRecordset(rs明细, "云南医保")
    
    If rs明细.EOF = True Then
        Err.Raise 9000, gstrSysName, "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    lng病人ID = rs明细("病人ID")
    
    '判断本次是否为门诊特殊病就诊，条件：病种ID<>0
    gstrSQL = " Select Nvl(病种ID,0) AS 病种ID From 保险帐户 " & _
              " Where 险类=" & intinsure & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "判断本次是否为门诊特殊病就诊")
    lng病种ID = Val(rsTemp("病种ID"))
    bln单病种 = (lng病种ID <> 0)
    '判断是否为特殊人群,条件特殊人群标志=1
    gstrSQL = "select nvl(特殊人群标志,0) 特殊人群标志 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & intinsure
    Call OpenRecordset(rsTemp, "判断本次是否为门诊特殊人群标志就诊")
    str特殊人群标志 = Nvl(rsTemp!特殊人群标志, 0)
    '一、费用明细传递
    '顺序号采用身份验证时返回的值:mstr顺序号
    Call DebugTool("费用明细传递")
    str医生 = LeftDB(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 8)
    str科室 = LeftDB(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 24)
    
    '普通门诊由于没有预算，所以还需要传输费用明细
    '删除前置服务器的所有未结明细（由于前一次确定时明细传输成功，但结算失败时）
    Call DebugTool("删除所有明细")
    mstrErr = Space(4)
    Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
    
    Do Until rs明细.EOF
        str规格 = LeftDB(IIf(IsNull(rs明细("规格")), "无规格", rs明细("规格")), 30)
        str产地 = LeftDB(IIf(IsNull(rs明细("产地")), " ", rs明细("产地")), 30)
        str科室 = LeftDB(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 24)
        cur发生费用 = cur发生费用 + rs明细("结帐金额")
        str发生时间 = Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")
        str大类 = Nvl(rs明细!附注)
        
        Call DebugTool("明细上传")
        mstrErr = Space(4)
        Call yh_feedetailtrans(mstr顺序号, rs明细("ID"), str大类, rs明细("收费项目"), LeftDB(rs明细("项目名称"), 24), _
            rs明细("数量"), Round(rs明细("实际价格"), 3), str产地, str规格, " ", str医生, str科室, mstr明细事务号, str医生, str发生时间, _
            cur自付比例, cur自付金额, cur报销金额, mstrErr)
           mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Call DebugTool("上传发生错误")
            MsgBox GetErrInfo(mstrErr, intinsure)
            '医保数据库回滚
            Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
            Exit Function
        End If
        gstrSQL = "ZL_病人记帐记录_上传(" & rs明细("ID") & "," & cur报销金额 & ",'" & cur自付比例 & "|" & cur自付金额 & "|" & cur报销金额 & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        rs明细.MoveNext
    Loop
    If lng病种ID <> 0 Then cur发生费用 = g结算数据.发生费用金额        '该处是应收金额，与预算保持一致
        
    '调预警函数，目前市医保只有门诊才可能预警，省医保不需调用此函数
    If intinsure = gint昆明市 Then
        mstrErr = Space(4)
        strErrMsg = Space(255)
        Call yh_AlertInfo_kms(mstr顺序号, mstrErr, strErrMsg)
           mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            If MsgBox("医保返回如下信息预警，继续结算请点“是”，取消结算请点“否”" & vbCrLf & _
            strErrMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    '二、写IC卡
    Call DebugTool("写IC卡")
    str卡类型 = Left(strSelfNo, 1)
    str初始化机构 = Space(4)
    If CDbl(cur个人帐户) <> 0 Then
        mstrErr = Space(4)
        Call yh_cardpay(str卡类型, mstr顺序号, str医生, "门诊收费", CDbl(cur个人帐户), str初始化机构, mstrErr)
    End If
    mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Call DebugTool("写卡错误")
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        '医保数据库回滚
        Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
        Exit Function
    End If
    
    '仅适用于市医保
    '说明：门诊单病种收费现只存在门诊血透，如果HIS前台传来的明细int_sicksortchk函数解析出多个单病种，
    'HIS前台也只能选择其中的血透单病种，不能录入其他单病种。如果返回的病种无血透（单病种编码1301）则，不能做单病种结算。
    If intinsure = gint昆明市 And bln单病种 Then
        mstrErr = Space(4)
        strSicks = Space(100)
        
        Call yh_sicksortchk(mstr顺序号, strSicks, mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, "结算失败但个人帐户已扣款：" & cur个人帐户 & "元，请与HIS商联系！" & vbCrLf & "详细错误：" & GetErrInfo(mstrErr, intinsure)
            Exit Function
        End If
        
        '有多个单病种，需要操作员选择
        strSicks = Trim(strSicks)
        '弹出病种供操作员选择
        strSickSel = frm昆明市特殊病种选择.ShowSelect(strSicks)
    End If
    
    '三、费用结算
    Call DebugTool("费用结算")
    str结算事务号 = Get事务号
    If str结算事务号 = "" Then
        Exit Function
    End If
    
    str初始化机构 = Space(4)
    mstrErr = Space(4)
    Call yh_feebalance(mstr顺序号, str医生, str科室, strSickSel, str结算事务号, _
        cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, cur超限自付, cur大病统筹, _
        cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, str人员状态, str初始化机构, cur特殊挂钩支付部分, _
        strFeeBalanceType, dblXTCS, mstrErr)
    mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str结算事务号, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr顺序号, str结算事务号, "1", mstrErr)
    
    '如果不符则校正结果结算
    If Not (pre_Balance.cur大病统筹 = cur大病统筹 And pre_Balance.cur特殊补助 = cur特殊人员统筹 And pre_Balance.cur医保基金 = cur统筹支付 _
            And pre_Balance.cur个人帐户 = cur个人帐户 And pre_Balance.cur公务员补助 = cur公务员统筹) Then
        blnReverse = True
        str结算方式 = "个人帐户|" & cur个人帐户
        str结算方式 = str结算方式 & "||医保基金|" & cur统筹支付
        str结算方式 = str结算方式 & "||大病统筹|" & cur大病统筹
        str结算方式 = str结算方式 & "||公务员补助|" & cur公务员统筹
        str结算方式 = str结算方式 & "||特殊补助|" & cur特殊人员统筹
        
        #If gverControl < 2 Then
            gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
        #Else
            strAdvance = str结算方式
            gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
    End If
    
    '四、保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    '定义 cur统筹累计 变量的目的是为了调用API，类型兼容
    Dim cur起付线 As Double, cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double
    Dim int住院次数累计 As Integer, cur基数累计 As Double, str起付线信息 As String, str审批标志编码 As String, str用药限制 As String
    curDate = zlDatabase.Currentdate
            
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    mstrErr = Space(4)
    Call yh_RyspInfo(mstr顺序号, cur起付线, cur统筹累计, cur基本统筹限额, cur大额统筹限额, cur基数累计, str起付线信息, str审批标志编码, str用药限制, mstrErr)
    cur统筹报销累计 = strVal(cur统筹累计)
    cur起付线 = strVal(cur起付线)
    cur基本统筹限额 = strVal(cur基本统筹限额)
    cur大额统筹限额 = strVal(cur大额统筹限额)
    cur基数累计 = strVal(cur基数累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur基数自付 & "," & cur基数累计 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")

    '将cur特殊人员自付改为起付线
    '备注中保存内容：特病列表|本次结算选择的病种编码|结算函数返回的费用结算类别|血透次数|本次初始化函数返回的平均付费类别|特殊医院标志
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & Get病种编码(lng病种ID) & "," & cur基数自付 & "," & _
        cur发生费用 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & mstr顺序号 & "'" & _
        ",NULL,NULL,'" & strSicks & "|" & strSickSel & "|" & strFeeBalanceType & "|" & dblXTCS & "|" & mstrAverageFeeType & "|" & mstrTsyybz & "','" & str特殊人群标志 & "'," & IIf(blnReverse, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")

    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '五、特殊门诊要进行出院信息修改
    If lng病种ID <> 0 Then
        If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
        str发生时间 = Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")
        str出院诊断 = lng病种ID
        str出院经办人 = LeftDB(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 8)
        str出院床号 = "1"
        str出院原因 = "9"
        str出院科室 = LeftDB(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 24)
        mstrErr = Space(4)
        Call yh_ReLeaveHosInfo(mstr顺序号, str出院原因, str发生时间, str出院诊断, str出院经办人, str出院科室, str出院床号, mstrErr)
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        'Exit Function 如果有意外错误出现，因为已经结算成功继续执行
        End If
    End If
    门诊结算_云南 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_云南(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, str退费事务号 As String
    Dim lng冲销ID As Long, str顺序号 As String, lng疾病编码 As Double
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim curDate As Date
    Dim str特殊人群标志 As String
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额  From 门诊费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call OpenRecordset(rsTemp, "云南医保")
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "云南医保")
    
    lng冲销ID = rsTemp("结帐ID")
    rsTemp.Close
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & intinsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "云南医保")
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    str顺序号 = rsTemp("支付顺序号")
    str特殊人群标志 = Trim(Nvl(rsTemp("特殊人群标志"), 0))
    lng疾病编码 = IIf(IsNull(rsTemp("封顶线")), 0, rsTemp("封顶线"))
    
    If Is卡正确(lng病人ID, intinsure) = False Then
        Exit Function
    End If
    
    str退费事务号 = Get事务号
    If str退费事务号 = "" Then
        Exit Function
    End If
    
    '3-表示普通门诊的个人账户退费处理；2-表示特殊门诊的个人账户预统筹基金的退费
    mstrErr = Space(4)
    Call yh_recedefeebalance(str顺序号, IIf(lng疾病编码 > 0, 2, 3), "", str退费事务号, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str退费事务号, "0", mstrErr)
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & _
        rsTemp("实际起付线") * -1 & "," & cur票据总金额 * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & "," & rsTemp("大病自付金额") * -1 & "," & rsTemp("超限自付金额") * -1 & "," & _
        cur个人帐户 * -1 & ",'" & str顺序号 & "',null,null,null,'" & str特殊人群标志 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")

    门诊结算冲销_云南 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人帐户转预交_云南(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim str卡类型 As String
    Dim str初始化机构 As String
    Dim str医生 As String
    
    On Error GoTo errHandle
    
    If Is卡正确(lng病人ID, intinsure) = False Then Exit Function
    
    str初始化机构 = Space(4)
    str卡类型 = Left(strSelfNo, 1)
    
    mstrErr = Space(4)
    str医生 = LeftDB(UserInfo.姓名, 8)
    If cur个人帐户 <> 0 Then Call yh_cardpay(str卡类型, str顺序号, LeftDB(UserInfo.姓名, 8), "预交款", cur个人帐户, str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    
    '---------------------------------------------------------------------------------------------
    '填写结算表
    curDate = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur个人帐户 & ",0,0,0,0,0,0," & _
        cur个人帐户 & ",'" & str顺序号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    个人帐户转预交_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_云南(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String, ByVal intinsure As Integer) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false

    Dim rsTemp As New ADODB.Recordset
    Dim rsTemp1 As New ADODB.Recordset
    Dim curDate As Date
    Dim str卡类型 As String
    Dim str卡号 As String
    Dim STR姓名 As String
    Dim str性别 As String
    Dim str出生日期 As String
    Dim str身份证号 As String
    Dim str人员状态 As String
    Dim str人群类别 As String
    Dim str特殊人群标志 As String
    Dim str初始化机构 As String, str单位编码 As String, str单位名称 As String
    Dim str事务号 As String   '事务控制号
    Dim blnTrans As Boolean
    Dim bln病种 As Boolean
    Dim str病种ID As String
    Dim lng疾病ID As Long, str疾病编码 As String
    '-----------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    '定义 cur统筹累计 变量的目的是为了调用API，类型兼容
    Dim cur起付线 As Double, cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double
    Dim int住院次数累计 As Integer, cur基数累计 As Double, str起付线信息 As String, str审批标志编码 As String, str用药限制 As String
    On Error GoTo errHandle
    mstr顺序号 = Space(19)
    str医保号 = Space(20)
    str事务号 = Space(18)
    str卡号 = Space(18)
    STR姓名 = Space(60)
    str性别 = Space(3)
    str出生日期 = Space(10)
    str身份证号 = Space(20)
    str人员状态 = Space(3)
    str人群类别 = Space(3)
    str特殊人群标志 = Space(3)
    str初始化机构 = Space(4)
    curDate = zlDatabase.Currentdate
    
    '注意：此时不能读保险帐户，因为尚未取到医保号，而是需要返回医保号
    gstrSQL = "Select A.入院日期,A.入院病床,B.名称 as 入院科室,C.住院号,A.登记时间,D.医保号,E.ID AS 病种ID,E.编码 as 病种编码,E.类别 as 病种类别 " & _
            " From 病案主页 A,部门表 B,病人信息 C,保险帐户 D,保险病种 E " & _
            " Where A.入院科室ID=B.ID And A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & _
            " And A.病人ID=C.病人ID And A.病人ID=D.病人ID and D.险类=" & intinsure & " and D.病种ID=E.ID(+)"
    Call OpenRecordset(rsTemp, "云南医保")
    If rsTemp.EOF = True Then
        MsgBox "没有发现此病人的信息！", vbExclamation, gstrSysName
        Exit Function
    End If
    lng疾病ID = Nvl(rsTemp!病种ID, 0)
    str疾病编码 = Nvl(rsTemp!病种编码)
    
    If IsNull(rsTemp("医保号")) = False Then
        str卡类型 = Left(rsTemp("医保号"), 1)
    Else
        If frmIdentify云南.GetIdentifyMode(intinsure, 1, str卡类型, lng疾病ID, str疾病编码) = False Then Exit Function
    End If
    
    '入院登记
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
'    '判断病种是否为0
'    gstrSQL = "select nvl(病种ID,0) 病种ID,医保号 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & intInsure
'    Call OpenRecordset(rsTemp, "医保接口")
'    'bln病种 = Nvl(rsTemp!病种ID, 0)
'    str病种ID = Nvl(rsTemp!病种ID, 0)
'    str医保号 = rsTemp!医保号
'    str医保号 = Right(str医保号, 10)
'    If str病种ID <> "0" Then
'       gstrSQL = "select 编码,名称 from 保险病种 where ID=" & str病种ID & " and 险类=" & intInsure
'       Call OpenRecordset(rsTemp, "医保接口")
'       str疾病编码 = Trim(rsTemp!编码)
'    End If
    '0092-特殊群体门诊和0094-住院都必须设置成普通病
    mstrErr = Space(4)
    If str疾病编码 = "0093" Then    '只有市医保才支持年终结转住院，要求病种编码传为零
        Call yh_kndadmit(LeftDB(UserInfo.姓名, 8), Mid(mstr医保号, 2), gstr医院编码, LeftDB(UserInfo.姓名, 8), LeftDB(rsTemp("入院科室"), 8), _
            LeftDB(lng病人ID, 12), LeftDB(rsTemp("住院号"), 12), IIf(rsTemp("病种类别") <> "0", "1", "0"), "", _
            Format(rsTemp!入院日期, "yyyy-01-01 01:01:01"), LeftDB(获取入出院诊断(lng病人ID, lng主页ID, True, False), 50), str事务号, mstr顺序号, str卡号, _
            STR姓名, str性别, str出生日期, str人员状态, str人群类别, str特殊人群标志, str初始化机构, str单位编码, str单位名称, mstrErr)
    Else
        Call yh_admit(str卡类型, LeftDB(UserInfo.姓名, 8), gstr医院编码, LeftDB(UserInfo.姓名, 8), LeftDB(rsTemp("入院科室"), 8), _
            LeftDB(lng病人ID, 12), LeftDB(rsTemp("住院号"), 12), IIf(rsTemp("病种类别") <> "0", "1", "0"), LeftDB(IIf(IsNull(rsTemp("病种编码")), "0", rsTemp("病种编码")), 8), _
            Format(rsTemp!入院日期, "yyyy-MM-dd hh:mm:ss"), LeftDB(获取入出院诊断(lng病人ID, lng主页ID, True, False), 50), str事务号, mstr顺序号, str卡号, _
            str医保号, STR姓名, str性别, str出生日期, str身份证号, str人员状态, str人群类别, str特殊人群标志, str初始化机构, str单位编码, str单位名称, mstrErr)
    End If
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    blnTrans = True
    str特殊人群标志 = TrimStr(str特殊人群标志)
    str人员状态 = TrimStr(str人员状态)
    str卡号 = TrimStr(str卡号)
    If str疾病编码 = "0093" Then
        str医保号 = mstr医保号
    Else
        str医保号 = str卡类型 & Left(TrimStr(str医保号), 19)
    End If
    
    On Error Resume Next
    Err = 0
    gstrSQL = "update 保险帐户 set 特殊人群标志='" & str特殊人群标志 & "',人员身份='" & str人员状态 & "',卡号='" & str卡号 & "'  where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "更新人员身份")
    If Err <> 0 Then
        gstrSQL = "update 医保病人档案 set 特殊人群标志='" & str特殊人群标志 & "',人员身份='" & str人员状态 & "',卡号='" & str卡号 & "'  where 病人ID=" & lng病人ID
        Call OpenRecordset(rsTemp, "更新人员身份")
    End If
    On Error GoTo errHandle
    
    mstr顺序号 = TrimStr(mstr顺序号)
    If mstr顺序号 = "" Then
        MsgBox "不能得到正确的入院登记顺序号。", vbInformation, gstrSysName
        Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    '帐户年度信息获取年度
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        '住院必须要调用yh_RyspInfo获取审批信息
    Call yh_RyspInfo(mstr顺序号, cur起付线, cur统筹累计, cur基本统筹限额, cur大额统筹限额, cur基数累计, str起付线信息, str审批标志编码, str用药限制, mstrErr)
        cur起付线 = strVal(cur起付线)
        cur统筹累计 = strVal(cur统筹累计)
        cur基本统筹限额 = strVal(cur基本统筹限额)
        cur大额统筹限额 = strVal(cur大额统筹限额)
        cur基数累计 = strVal(cur基数累计)
        str起付线信息 = strVal(str起付线信息)
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'起付线','''" & cur起付线 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "再一次住院病人更新起付线")
    gstrSQL = "zl_病人审批信息_insert(2," & lng病人ID & "," & intinsure & "," & Year(curDate) & ",'" & _
        mstr顺序号 & "'," & cur起付线 & "," & cur统筹累计 & "," & _
        cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur基数累计 & "," & str起付线信息 & ",'" & str审批标志编码 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "病人审批信息")
    
    '强制把登记顺序号、及新的医保号填入
    gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & intinsure & _
                ",'" & str卡号 & "','" & str医保号 & "','" & mstr顺序号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    Call yh_transaction("0", mstr顺序号, str事务号, "1", mstrErr)
    
    入院登记_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
End Function

Public Function 出院登记_云南(lng病人ID As Long, lng主页ID As Long, str顺序号 As String, ByVal intinsure As Integer, _
    Optional ByVal bln结帐出院 As Boolean = False, Optional ByVal bln撤销入院 As Boolean = False) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim str事务号 As String   '事务控制号
    Dim strMsg As String
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double, str人员状态 As String, cur特殊挂钩支付部分 As Double
    
    '单病种相关
    Dim strSicks As String          '银海返回的病种列表，注：门诊特殊病目前只能选择血透
    Dim strFeeBalanceType As String, dblXTCS As Double      '结算函数的返回出参：支付费用类别、血透次数
    Dim strSickSel As String        '操作员选择的病种编码
    Dim str备注 As String
    
    Dim blnTrans As Boolean
    Dim rsInfo As New ADODB.Recordset
    Dim str出院原因 As String, str出院时间 As String, str出院诊断 As String
    Dim str出院经办人 As String, str出院科室 As String, str出院床号 As String
    Dim str特殊人群标志 As String
    '出院方式:1-正常;2-转院;3-死亡；对应医保的出院原因：0、正常出院；1、死亡；2、转院；3、审批未住院（中途取消）；9、其他
    Dim rsTemp As New ADODB.Recordset
    Dim rstemp特殊 As New ADODB.Recordset
    
    On Error GoTo errHandle
    '如果存在未结费用，则仅办理HIS出院；否则同时办理医保及HIS出院
    Call DebugTool("进入出院登记接口")
    If bln撤销入院 Or Not 存在未结费用(lng病人ID, lng主页ID) Then
        str初始化机构 = Space(4)
        
        str事务号 = Get事务号
        If str事务号 = "" Then
            
        End If
        mstr顺序号 = str顺序号
        
        '仅适用于市医保
        '说明：门诊单病种收费现只存在门诊血透，如果HIS前台传来的明细int_sicksortchk函数解析出多个单病种，
        'HIS前台也只能选择其中的血透单病种，不能录入其他单病种。如果返回的病种无血透（单病种编码1301）则，不能做单病种结算。
        '住院解析出多个单病种，操作员只能选择其中一个进行结算.
        If intinsure = gint昆明市 Then
            mstrErr = Space(4)
            strSicks = Space(100)
            Call yh_sicksortchk(mstr顺序号, strSicks, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                Exit Function
            End If
            
            '弹出病种供操作员选择
            strSickSel = frm昆明市特殊病种选择.ShowSelect(strSicks)
        End If
        '出院登记是通过调用结算交易完成。此时假设病人的费用已经全部结清
        Call DebugTool("调用医保出院接口")
        mstrErr = Space(4)
        Call yh_feebalance(mstr顺序号, LeftDB(UserInfo.姓名, 8), LeftDB(UserInfo.部门, 24), strSickSel, str事务号, cur全自付, _
            cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, cur超限自付, _
            cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, _
            str人员状态, str初始化机构, cur特殊挂钩支付部分, strFeeBalanceType, dblXTCS, mstrErr)
           mstrErr = TrimStr(mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            '医保数据库回滚
            Call yh_transaction("2", mstr顺序号, str事务号, "0", mstrErr)
            Exit Function
        End If
        '提交医保前置机数据库
        Call yh_transaction("2", mstr顺序号, str事务号, "1", mstrErr)
    
        '更新保险结算记录中的备注字段
        '结算函数自动调用本函数，因此肯定有权限执行zl_保险结算记录_Insert()的
        gstrSQL = "Select nvl(特殊人群标志,0) as 特殊人群标志 From 保险帐户 Where 险类=" & intinsure & " And 病人ID=" & lng病人ID
        Call OpenRecordset(rstemp特殊, "读取医保病人的特殊类别")
        str特殊人群标志 = Nvl(rstemp特殊!特殊人群标志, 0)
        str备注 = strSicks & "|" & strSickSel & "|" & strFeeBalanceType & "|" & dblXTCS & "|" & mstrAverageFeeType & "|" & mstrTsyybz
        gstrSQL = " Select 性质,记录ID,险类,病人ID,年度,帐户累计增加,帐户累计支出,累计进入统筹,累计统筹报销,住院次数," & _
                  " 起付线,封顶线,实际起付线,发生费用金额,全自付金额,首先自付金额,进入统筹金额,统筹报销金额," & _
                  " 大病自付金额,超限自付金额,个人帐户支付,支付顺序号,主页ID From 保险结算记录" & _
                  " Where 记录ID=(Select Max(记录ID) From 保险结算记录 Where 性质=2 And 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & ")"
        Call OpenRecordset(rsTemp, "取最后一次结算记录数据")
        If rsTemp.RecordCount <> 0 Then '更改零费用出院(2007-08-30)
        gstrSQL = "zl_保险结算记录_insert(2," & rsTemp!记录ID & "," & rsTemp!险类 & "," & rsTemp!病人ID & "," & _
            rsTemp!年度 & "," & rsTemp!帐户累计增加 & "," & rsTemp!帐户累计支出 & "," & rsTemp!累计进入统筹 & "," & _
            rsTemp!累计统筹报销 & "," & rsTemp!住院次数 & "," & rsTemp!起付线 & "," & rsTemp!封顶线 & "," & rsTemp!实际起付线 & "," & _
            rsTemp!发生费用金额 & "," & rsTemp!全自付金额 & "," & rsTemp!首先自付金额 & "," & _
            rsTemp!进入统筹金额 & "," & rsTemp!统筹报销金额 & "," & rsTemp!大病自付金额 & "," & rsTemp!超限自付金额 & "," & _
            rsTemp!个人帐户支付 & ",'" & rsTemp!支付顺序号 & "'," & rsTemp!主页ID & ",NULL,'" & str备注 & "','" & str特殊人群标志 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
        End If
        blnTrans = True
        '更新出院诊断（省市医保都支持年终结转，传4）
        mstrErr = Space(4)
        gstrSQL = "select decode(出院方式,'正常',0,'转院',2,'死亡',1,'年终结转',4,9) 出院方式 From 病案主页 " & _
                " Where 病人ID = " & lng病人ID & " And 主页ID = " & lng主页ID
        Call OpenRecordset(rsInfo, "出院方式")
        str出院原因 = rsInfo!出院方式
        
        gstrSQL = "select b.名称 出院科室,终止时间,操作员姓名  " & _
                 " from 病人变动记录 A,部门表 B  " & _
                 " where 病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID & " and 终止原因=1 " & _
                 " and A.科室ID=B.ID"
        Call DebugTool("提取病人出院时间的SQL：" & gstrSQL)
        Call OpenRecordset(rsInfo, "出院情况")
        If rsInfo.RecordCount <> 0 Then
            str出院时间 = Format(rsInfo!终止时间, "yyyy-MM-dd HH:mm:ss")
            str出院科室 = LeftDB(rsInfo!出院科室, 20)
            str出院床号 = "10"
            str出院经办人 = LeftDB(rsInfo!操作员姓名, 20)
            str出院诊断 = LeftDB(获取入出院诊断(lng病人ID, lng主页ID, False, False), 100)
        Else
            '撤销入院也调这个函数，可能病人没有出院信息
            str出院时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            str出院科室 = LeftDB("撤销入院", 20)
            str出院床号 = "10"
            str出院经办人 = LeftDB(gstrUserName, 20)
            str出院诊断 = "撤销入院"
        End If
        mstrErr = Space(4)
        Call yh_ReLeaveHosInfo(mstr顺序号, str出院原因, str出院时间, str出院诊断, str出院经办人, str出院科室, str出院床号, mstrErr)
        Call DebugTool("病人ID=" & lng病人ID & "|主页ID=" & lng主页ID & "|出院时间=" & str出院时间)
    Else
        strMsg = "还存在未结费用,不能办理医保出院！"
        If Not bln结帐出院 Then
            strMsg = strMsg & "本次仅办理HIS出院"
        Else
            strMsg = strMsg & "请在保险帐户中为该病人办理补充出院登记"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
        If bln结帐出院 Then Exit Function
    End If
    '改变病人状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    出院登记_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
End Function

Public Function 出院登记撤消_云南(lng病人ID As Long, lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim str事务号 As String   '事务控制号
    Dim str顺序号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '如果存在未结费用，则仅办理HIS出院；否则同时办理医保及HIS出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '出院登记是通过调用结算交易完成。此时假设病人的费用已经全部结清
        gstrSQL = "Select 支付顺序号 From 保险结算记录 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
        Call OpenRecordset(rsTemp, "撤消出院")
        If rsTemp.EOF = True Then
            MsgBox "该病人未做过医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        
        str顺序号 = Nvl(rsTemp("支付顺序号"), "")
        mstrErr = Space(4)
        Call yh_recedefeebalance(str顺序号, "1", "", String(18, "1"), mstrErr) '目前都是用预结算在处理
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    出院登记撤消_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_云南(rsExse As Recordset, ByVal lng病人ID As Long, ByVal intinsure As Integer) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim str事务号 As String   '事务控制号
    Dim cn上传 As New ADODB.Connection, str数据批号 As String
    Dim cur个人帐户 As Currency, cur自付总额 As Currency
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String, str规格 As String, str产地 As String
    Dim cur发生费用 As Currency, dbl金额 As Double, dbl数量 As Double
    Dim str发生时间 As String, str本次入院时间 As String, str大类 As String, str登记时间 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    With g结算数据
        .病人ID = rsExse("病人ID")
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & rsExse("病人ID")
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
    End With
    
    '取本次入院时间
    gstrSQL = " Select To_Char(入院日期,'yyyy-MM-dd hh24:mi:ss') 入院时间 From 病案主页" & _
              " Where 病人ID=" & lng病人ID & " And 主页ID=" & g结算数据.主页ID
    Call OpenRecordset(rsTemp, "读取本次入院时间")
    str本次入院时间 = rsTemp!入院时间
    
    '打开另外一个连接串，以达到不受当前连接事务的控制
    Set cn上传 = GetNewConnection
    
    '顺序号取入院登记验证返回的
    gstrSQL = "Select 医保号,顺序号 From 保险帐户 " & _
              "Where 顺序号 is Not NULL And 病人ID=" & lng病人ID & " And 险类=" & intinsure
    Call OpenRecordset(rsTemp, "虚拟结算")
    
    If rsTemp.EOF Then
        MsgBox "未发现该病人的住院交易顺序号,不能执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    mstr医保号 = rsTemp("医保号")
    mstr顺序号 = rsTemp("顺序号")
    
    '删除前置服务器的所有未结明细
    mstrErr = Space(4)
    Call yh_rollbackdetail(mstr顺序号, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
            
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    
    '为了避免负记录在前，而正记录在后。不能有效冲销
    rsExse.Sort = "NO,序号 asc,数量 Desc"
    
    '费用明细传递
    Do Until rsExse.EOF
        '昆明医保全部重新传
        str医生 = LeftDB(IIf(IsNull(rsExse("医生")), UserInfo.姓名, rsExse("医生")), 8)
        str规格 = LeftDB(IIf(IsNull(rsExse("规格")), "无规格", rsExse("规格")), 30)
        str产地 = LeftDB(IIf(IsNull(rsExse("产地")), "", rsExse("产地")), 30)
        str科室 = LeftDB(IIf(IsNull(rsExse("开单部门")), UserInfo.部门, rsExse("开单部门")), 24)
        '不能传递负数
        If rsExse("记录状态") = 1 And rsExse("数量") < 0 Then
            MsgBox "医保不支持直接录入负数，只能选择原有单据进行冲销。", vbInformation, gstrSysName
            Exit Function
        End If
        '传0的目的是为了删除已经上传但被冲销的费用记录
        dbl数量 = Val(IIf(rsExse("数量") > 0, rsExse("数量"), 0))
        dbl金额 = Val(IIf(rsExse("价格") > 0, rsExse("价格"), 0))
        str发生时间 = Format(rsExse("发生时间"), "yyyy-MM-dd HH:mm:ss")
        str登记时间 = Format(rsExse("登记时间"), "yyyy-MM-dd HH:mm:ss")
        str大类 = Get大类编码(rsExse!收费细目ID, intinsure)
        mstrErr = Space(4)
        
        '为了让负记录能正确找到正记录，所以数据批号中不包含记录状态
        str数据批号 = rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") '& "_" & rsExse("记录状态")
        
        '如果登记时间小于本次住院时间则不上传
        If str发生时间 >= str本次入院时间 Then
            Call DebugTool("上传明细串：" & mstr顺序号 & "," & str数据批号 & "," & str大类 & "," & rsExse("医保项目编码") & "," & _
                rsExse("收费名称") & "," & dbl数量 & "," & dbl金额 & "," & str产地 & "," & str规格 & "," & "" & "," & str医生 & "," & str科室 & "," & str事务号 & "," & str医生 & "," & str登记时间 & "," & _
                cur自付比例 & "," & cur自付金额 & "," & cur报销金额)
            Call yh_feedetailtrans(mstr顺序号, str数据批号, str大类, rsExse("医保项目编码"), _
                rsExse("收费名称"), dbl数量, dbl金额, str产地, str规格, "", str医生, str科室, str事务号, str医生, str登记时间, _
                cur自付比例, cur自付金额, cur报销金额, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                '医保数据库回滚
                Call yh_transaction("1", mstr顺序号, str事务号, "0", mstrErr)
                Exit Function
            End If
            cur发生费用 = cur发生费用 + rsExse("金额")
            
            '提取该费用记录的ID
            gstrSQL = "Select ID From 住院费用记录 " & _
                " Where NO='" & rsExse!NO & "' And 记录性质=" & rsExse!记录性质 & " And 记录状态=" & rsExse!记录状态 & " And 序号=" & rsExse!序号
            Call OpenRecordset(rsTemp, "提取该费用记录的ID")
            
            If rsTemp.RecordCount <> 0 Then
                gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & "," & cur报销金额 & ",'" & cur自付比例 & "|" & cur自付金额 & "|" & cur报销金额 & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        '处理9版本中自动计算没有时分秒问题(2007-12-06)
        Else
            Call DebugTool("上传明细串：" & mstr顺序号 & "," & str数据批号 & "," & str大类 & "," & rsExse("医保项目编码") & "," & _
                rsExse("收费名称") & "," & dbl数量 & "," & dbl金额 & "," & str产地 & "," & str规格 & "," & "" & "," & str医生 & "," & str科室 & "," & str事务号 & "," & str医生 & "," & str登记时间 & "," & _
                cur自付比例 & "," & cur自付金额 & "," & cur报销金额)
            Call yh_feedetailtrans(mstr顺序号, str数据批号, str大类, rsExse("医保项目编码"), _
                rsExse("收费名称"), dbl数量, dbl金额, str产地, str规格, "", str医生, str科室, str事务号, str医生, str登记时间, _
                cur自付比例, cur自付金额, cur报销金额, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
                '医保数据库回滚
                Call yh_transaction("1", mstr顺序号, str事务号, "0", mstrErr)
                Exit Function
            End If
            cur发生费用 = cur发生费用 + rsExse("金额")

            '提取该费用记录的ID
            gstrSQL = "Select ID From 住院费用记录 " & _
                " Where NO='" & rsExse!NO & "' And 记录性质=" & rsExse!记录性质 & " And 记录状态=" & rsExse!记录状态 & " And 序号=" & rsExse!序号
            Call OpenRecordset(rsTemp, "提取该费用记录的ID")

            If rsTemp.RecordCount <> 0 Then
                gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & "," & cur报销金额 & ",'" & cur自付比例 & "|" & cur自付金额 & "|" & cur报销金额 & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
        rsExse.MoveNext
    Loop
        
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    '虚拟结算
    Dim str结算标志 As String
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double, str人员状态 As String, cur特殊挂钩部分 As Double
    
    str初始化机构 = Space(4)
    mstrErr = Space(4)
    str结算标志 = "0" '虚拟结算
    Call yh_virtualbalance(mstr顺序号, str结算标志, "", str事务号, cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, _
        cur超限自付, cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, str人员状态, str初始化机构, cur特殊挂钩部分, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '保存临时数据，为结算操作做准备
    'Modified By ZYB 20030812
'    gstrSQL = "Select 帐户余额 From 保险帐户 Where 险类=" & intInsure & " And 病人ID=" & lng病人ID
'    Call OpenRecordset(rsTemp, "读取医保病人的医保号")
'    cur个人帐户 = rsTemp!帐户余额
    cur个人帐户 = 个人余额_云南(lng病人ID, mstr医保号, 10, intinsure)
    If cur特殊人员统筹 > 0 Then
        cur自付总额 = cur特殊人员自付
    Else
        cur自付总额 = cur发生费用 - (cur统筹支付 + cur大病统筹 + cur公务员统筹 + cur特殊人员统筹)
    End If
    cur个人帐户 = IIf(CDbl(Format(cur个人帐户, "#####0.00")) >= CDbl(Format(cur自付总额, "#####0.00")), cur自付总额, cur个人帐户)
    If Not 医保病人已经出院(lng病人ID) Then cur个人帐户 = 0
    
    With g结算数据
        .病人ID = lng病人ID
        .发生费用金额 = cur发生费用
    End With
    
    住院虚拟结算_云南 = "个人帐户;" & cur个人帐户 & ";1" '允许修改
    住院虚拟结算_云南 = 住院虚拟结算_云南 & "|医保基金;" & cur统筹支付 & ";0"
    住院虚拟结算_云南 = 住院虚拟结算_云南 & "|大病统筹;" & cur大病统筹 & ";0"
    住院虚拟结算_云南 = 住院虚拟结算_云南 & "|公务员补助;" & cur公务员统筹 & ";0"
    住院虚拟结算_云南 = 住院虚拟结算_云南 & "|特殊补助;" & cur特殊人员统筹 & ";0"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_云南(lng结帐ID As Long, str顺序号 As String, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim str事务号 As String   '事务控制号
    Dim str卡类型 As String, str医生 As String
    Dim str结算标志 As String, strSelfNo As String
    Dim cur个人帐户 As Currency
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String, str人员状态 As String, cur特殊挂钩部分 As Double
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date, lng病种ID As Long, rsTemp As New ADODB.Recordset
    
    Dim str卡号 As String, STR姓名 As String, str性别 As String, str身份证号 As String, lng年龄 As Double, str人群类别 As String
    str初始化机构 = Space(4)
    On Error GoTo errHandle
    '检查病人是否已经出院，如果处于在院状态则直接退出系统(20071101)
    gstrSQL = "select a.出院日期 " & _
              " from 病案主页 a,病人信息 b " & _
              " Where a.病人ID = b.病人ID And a.主页ID = b.住院次数 " & _
              " and a.出院日期 is null and a.病人id = " & lng病人ID
    Call OpenRecordset(rsTemp, "判断出院日期")
    If rsTemp.RecordCount <> 0 Then
        Err.Raise 9000, gstrSysName, "病人未出院，不能办理结算，请先办理出院！"
        Exit Function
    End If
    '取入院登记验证所返回的顺序号
    mstr顺序号 = str顺序号
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    '费用结算:结帐。为了达到中途结帐的目的，没有使用结算函数
    '先读取医保号
    gstrSQL = "Select 医保号 From 保险帐户 Where 险类=" & intinsure & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "读取医保病人的医保号")
    strSelfNo = rsTemp!医保号
    '读取本次个人帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=" & intinsure & _
        " And A.结算方式 in ('个人帐户') And A.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "获取本次个人帐户支付额")
    cur个人帐户 = 0
    If Not rsTemp.EOF Then
        cur个人帐户 = rsTemp!个人帐户
    End If
    
    mstrErr = Space(4)
    str结算标志 = "1"   '结算
    Call yh_virtualbalance(mstr顺序号, str结算标志, lng结帐ID, str事务号, cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, _
        cur超限自付, cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, str人员状态, str初始化机构, cur特殊挂钩部分, mstrErr)
    mstrErr = TrimStr(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr顺序号, str事务号, "1", mstrErr)
    '填写结算表
    curDate = zlDatabase.Currentdate
    '读出该病人本次结算的病种信息
    gstrSQL = "Select nvl(病种ID,0) 病种ID From 保险帐户 A Where A.险类=" & intinsure & " and A.病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "保险结算")
    If rsTemp.EOF = False Then
        lng病种ID = rsTemp("病种ID")
    End If
    '写IC卡（由于银海不在一个事务中处理，因此卡写失败时，仍然继续处理）
    str卡类型 = Left(strSelfNo, 1)
    str初始化机构 = Space(4)
    str医生 = LeftDB(UserInfo.姓名, 8)
    If CDbl(cur个人帐户) <> 0 Then
        '单病种要求：下卡前必须调用cardinfo()
'        str卡号 = Space(18)
'        str姓名 = Space(60)
'        str性别 = Space(3)
'        str身份证号 = Space(20)
'        mstrErr = Space(4)
'        Call yh_cardinfo(str卡类型, mcur帐户余额, str卡号, str姓名, str性别, str身份证号, lng年龄, str人群类别, mstrErr)
'       If Trim(mstrErr) <> "0000" Then
'           MsgBox "本次结算该病人不用卡支付!!!!", vbOKOnly
'       Else
        mstrErr = Space(4)
        Call yh_cardpay(str卡类型, mstr顺序号, str医生, "住院结算", CDbl(cur个人帐户), str初始化机构, mstrErr)
        mstrErr = TrimStr(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure) & "（如果下卡失败,请补缴现金￥" & Format(cur个人帐户, "#####0.00") & "）"
            cur个人帐户 = 0
        End If
       'End If
    End If
    
    '临时提取起付线
    Dim cur起付线 As Double
    gstrSQL = "select nvl(起付线,0) as 起付线 from 保险帐户 where 顺序号='" & mstr顺序号 & "' and 险类=" & intinsure & " and  病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "提取起付线")
    cur起付线 = strVal(rsTemp!起付线)
    '更新保险结算记录:
    If mstrErr <> "0000" Then
       cur个人帐户 = 0
       '特殊人员自付"cur特殊人员自付"改为"cur基数自付"
      gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & Get病种编码(lng病种ID) & "," & cur基数自付 & "," & _
        g结算数据.发生费用金额 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & mstr顺序号 & "'," & g结算数据.主页ID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    Else
      gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & Get病种编码(lng病种ID) & "," & cur基数自付 & "," & _
        g结算数据.发生费用金额 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & mstr顺序号 & "'," & g结算数据.主页ID & ")"
      Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    End If
    '帐户年度信息
    Call Get帐户信息(intinsure, lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
    '定义 cur统筹累计 变量的目的是为了调用API，类型兼容
    Dim cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double
    Dim cur基数累计 As Double, str起付线信息 As String, str审批标志编码 As String, str用药限制 As String
    '昆明市医保支持查询支付累计
    mstrErr = Space(4)
    Call yh_RyspInfo(mstr顺序号, cur起付线, cur统筹累计, cur基本统筹限额, cur大额统筹限额, cur基数累计, str起付线信息, str审批标志编码, str用药限制, mstrErr)
    cur统筹报销累计 = cur统筹累计
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & cur基数累计 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    '病人审批信息
'    gstrSQL = "zl_病人审批信息_insert(" & lng病人ID & "," & intInsure & "," & Year(curDate) & "," & _
'        mstr顺序号 & "," & cur起付线 & "," & cur统筹累计 & "," & _
'        cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur基数累计 & "," & str起付线信息 & ",'" & str审批标志编码 & "',null)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "病人审批信息")
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    住院结算_云南 = True
    '判断是否需要调用出院结算（如果HIS已出院且不存在未结费用）
    Dim lng主页ID As Long
    '取得主页ID
    gstrSQL = "Select Nvl(住院次数,0) 主页ID From 病人信息 Where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "取主页ID")
    lng主页ID = rsTemp!主页ID
    
    If Not 存在未结费用(lng病人ID, lng主页ID) And 医保病人已经出院(lng病人ID) Then
        gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
                " C.卡号,C.医保号,C.密码,C.顺序号 " & _
                " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
                " Where A.病人ID=D.病人ID And A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & _
                " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=" & intinsure
        Call OpenRecordset(rsTemp, "取顺序号")
    
        If rsTemp.EOF Then
            Err.Raise 9000 + vbExclamation, gstrSysName, "没有此病人或此病人不是医保病人，无法办理出院手续！（请在医保帐户中办理补充出院手续）"
            Exit Function
        End If
        If IsNull(rsTemp!顺序号) Then
            Err.Raise 9000, gstrSysName, "未发现该病人的住院交易顺序号,无法办理出院手续！（请在医保帐户中办理补充出院手续）"
            Exit Function
        End If
        
        Call 出院登记_云南(lng病人ID, lng主页ID, rsTemp!顺序号, intinsure, True)
    End If
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_云南(lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str顺序号 As String, cur个人帐户 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim curDate As Date
    Dim str特殊人群标志 As String
    
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=" & intinsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    str顺序号 = rsTemp("支付顺序号")
    str特殊人群标志 = Nvl(rsTemp("特殊人群标志"), 0)
'    mstrErr = Space(4)
'    Call yh_recedefeebalance(str顺序号, "1", lng结帐ID, String(18, "1"), mstrErr) '1表示作费结算
'         mstrErr = Trim(mstrErr)
'    If mstrErr <> "0000" Then
'       Err.Raise 9000,gstrSysName, GetErrInfo(mstrErr, intInsure)
'        Exit Function
'    End If
    mstrErr = Space(4)
    Call yh_recedefeebalance(str顺序号, "0", lng结帐ID, String(18, "1"), mstrErr) '目前都是用预结算在处理,表示0作费预结算
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        Err.Raise 9000, gstrSysName, GetErrInfo(mstrErr, intinsure)
        Exit Function
    End If
    
    '将个人帐户支付额回退到卡号（银海是本次住院卡支付全部回退）
    cur个人帐户 = Nvl(rsTemp("个人帐户支付"), 0)
    If CDbl(cur个人帐户) <> 0 Then
        mstrErr = Space(4)
        Call yh_recedefeebalance(str顺序号, "4", lng结帐ID, String(18, "1"), mstrErr) '目前都是用预结算在处理
        mstrErr = Trim(mstrErr)
        If mstrErr <> "0000" Then
            Err.Raise 9000, gstrSysName, "个人帐户回退失败但结算回退成功，请与HIS商联系！" & vbCrLf & "详细错误：" & GetErrInfo(mstrErr, intinsure)
        End If
    End If
    
    '帐户年度信息
    Call Get帐户信息(intinsure, rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & intinsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    '封顶线保存有疾病编码，所以不取反
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & intinsure & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & _
        rsTemp("实际起付线") * -1 & "," & rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & "," & rsTemp("大病自付金额") * -1 & "," & rsTemp("超限自付金额") * -1 & "," & _
        cur个人帐户 * -1 & ",'" & str顺序号 & "'," & rsTemp("主页ID") & ",null,null,'" & str特殊人群标志 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "云南医保")

    住院结算冲销_云南 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 错误信息_云南(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Private Function LeftDB(ByVal strText As String, ByVal lngLength As Long)
'功能：按数据库的长度计算方式得到字符串的实际可用子串
    LeftDB = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, lngLength), vbUnicode)
End Function
Private Function strVal(ByVal strText As String)
'将字符型换为数字型
       strVal = Val(strText)
End Function

Private Function Get事务号() As String
    Dim str事务号 As String
    
    On Error GoTo errHandle
    
    str事务号 = Space(18)
    Call yh_gettranssequence(str事务号) '这里费用传递和结算是两个事务号
    str事务号 = TrimStr(str事务号)
    If str事务号 = "" Then
        MsgBox "获取事务控制号失败。", vbInformation, gstrSysName
    End If
    
    Get事务号 = str事务号
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Is卡正确(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'功能：判断读卡器的卡是否就是要操作的病人的
    Dim rsTemp As New ADODB.Recordset
    Dim str卡号_库 As String, str卡号 As String, str卡类型 As String
    
    Dim cur余额 As Double, STR姓名 As String, str性别 As String
    Dim str身份证号 As String, lng年龄 As Double, str人群类别 As String
    
    On Error GoTo errHandle
    
    gstrSQL = "select 卡号,医保号 from 保险帐户 where 险类=" & intinsure & " and 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "云南医保")
    
    str卡号_库 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
    str卡类型 = Left(rsTemp("医保号"), 1)
    
    str卡号 = Space(20)
    STR姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str卡类型, cur余额, str卡号, STR姓名, str性别, str身份证号, lng年龄, str人群类别, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    str卡号 = TrimStr(str卡号)
    
    If str卡号 <> str卡号_库 Then
        MsgBox "刷卡器中的卡不是当前病人的，请插入正确的IC卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Is卡正确 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get卡余额(ByVal str医保号 As String, 卡余额 As Currency, ByVal intinsure As Integer) As Boolean
'功能：得到卡余额
    Dim cur余额 As Double, STR姓名 As String, str性别 As String, str卡号 As String
    Dim str身份证号 As String, lng年龄 As Double, str卡类型 As String, str人群类别 As String
    
    str卡类型 = Left(str医保号, 1)
    
    str卡号 = Space(20)
    STR姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    
    mstrErr = Space(4)
    Call yh_cardinfo(str卡类型, cur余额, str卡号, STR姓名, str性别, str身份证号, lng年龄, str人群类别, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    卡余额 = cur余额
    Get卡余额 = True
End Function

Private Function Get病种编码(ByVal lng病种ID As Long) As String
'功能：判断读卡器的卡是否就是要操作的病人的
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select 编码 from 保险病种 where ID=" & lng病种ID
    Call OpenRecordset(rsTemp, "云南医保")
    
    If rsTemp.EOF = False Then
        Get病种编码 = Val(rsTemp("编码")) '为了保存在封顶线字段，所以必须是数字
        If Val(Get病种编码) = 0 Then Get病种编码 = "9000" '特批特种病也为0000，所以强制改为9000
    Else
        Get病种编码 = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 撤消急诊登记_云南(ByVal str顺序号 As String, ByVal intinsure As Integer) As Boolean
'功能：撤消急诊登记
    Dim rsTemp As New ADODB.Recordset
    Dim str事务号 As String   '事务控制号
    Dim str医生 As String, str科室 As String
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double, str人员状态 As String, cur特殊挂钩支付部分 As Double
    
    '单病种相关(省市医保调用一个函数)
    Dim strSicks As String          '银海返回的病种列表，注：门诊特殊病目前只能选择血透
    Dim strFeeBalanceType As String, dblXTCS As Double      '结算函数的返回出参：支付费用类别、血透次数
    Dim strSickSel As String        '操作员选择的病种编码
    Dim str备注 As String
    
    On Error GoTo errHandle
    str初始化机构 = Space(4)
    
    gstrSQL = "Select 支付顺序号 from 保险结算记录 where 支付顺序号='" & str顺序号 & "'"
    Call OpenRecordset(rsTemp, "云南医保")
    
    If rsTemp.EOF = False Then
        MsgBox "该病人的急诊交易已经成功完成，不能撤消，只能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '删除前置服务器的所有未结明细
    mstrErr = Space(4)
    Call yh_rollbackdetail(str顺序号, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        Exit Function
    End If
    
    '出院登记是通过调用结算交易完成。零费用结算
    str事务号 = Get事务号
    If str事务号 = "" Then
        
    End If
    
    mstrErr = Space(4)
    Call yh_feebalance(mstr顺序号, str医生, str科室, strSickSel, str事务号, _
            cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, cur超限自付, cur大病统筹, _
            cur大病自付, cur特殊人员自付, cur特殊人员统筹, cur公务员统筹, str人员状态, str初始化机构, cur特殊挂钩支付部分, _
            strFeeBalanceType, dblXTCS, mstrErr)
       mstrErr = Trim(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr, intinsure), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("2", str顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    
    撤消急诊登记_云南 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get大类编码(ByVal str项目编码 As String, ByVal intinsure As Integer) As String
    Dim rsClass As New ADODB.Recordset
    '获取某个医保项目的大类编码
    gstrSQL = "Select 附注 From 保险支付项目 Where 险类=" & intinsure & " And 收费细目ID=" & str项目编码
    Call OpenRecordset(rsClass, "获取某个医保项目的大类编码")
    If rsClass.RecordCount = 0 Then
        MsgBox "医保编码为：" & str项目编码 & "的项目在保险项目表中不存在！", vbInformation, gstrSysName
        Exit Function
    End If
    Get大类编码 = Nvl(rsClass!附注)
End Function
