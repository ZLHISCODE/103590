Attribute VB_Name = "mdl宁海"
Option Explicit
'浙江宁海医保接口补充说明
'不提供实时上传接口，要求医院每天下班前执行病人费用预结算完成，因为对方接口存在缺陷
'科室不存在对照

'传入参数：$$交易体$$
'用户提出交易申请时，通过外部传入数据包提供参数。一般交易传入交易体前4个域固定为：是否有社保卡（0无，1有）~ IC卡数据（有卡时一般无意义，无卡时传医保号）~现金支付方式（1现金、2电子钱包、3银行借记卡）~银行卡信息。指定交易没有某数据域时相应值为空，没有社保卡时，需要读写卡的交易不可执行。
'
'传出参数：$$交易状态~错误信息~交易结果信息$$
'其中交易状态：0成功，>0成功但是有警告提示，<0失败。
'如果用户交易处理成功，返回给用户时状态为"交易成功"，错误信息一般为空，交易体实际格式为"$$0~~交易结果信息$$"。如果有警告信息时，其警告内容在错误信息中，" 交易体格式一般为""$$x~警告信息~交易结果信息$$""。另外，交易结果信息中前3个域固定：写医保卡结果(写医保卡结果：0表示不写或写卡成功，其他表示写卡错误信息)~扣银行卡结果(0表示不扣或扣成功)~写卡后IC卡数据。没有数据域相应值为空。
'如果交易失败，一般只有出错信息，没有交易结果信息，格式为"$$-1~错误信息~"，当然特务交易需要也可以既有错误信息又有交易结果信息。
'错误信息外部格式为："产品号%%函数名.错误点%%错误号%%错误原因"。其中"错误原因"为出错问题描述。结构如下图所示：
'
'例1：$$-1~3333%%f_UserBargaingApply.3%%-3%%非法数据~$$
'例2：$$0~~交易结果信息$$
'一般交易失败时，需要医院系统将错误原因显示给操作用户，供其查找错误原因，如果需要将错误信息提供给接口开发者查找原因，则需要提供完整地返回参数。

'----------特殊字符说明----------
'   字符    字符说明           类型说明
'   $$         双美元符号      分隔符，分隔交易数据包
'   ~          单波浪线        分隔交易包中不同域
'   %%         双百分比符号    分隔交易包不同域间元素
'   '          单引号          系统字符串分隔符，不可使用

'----------IC卡数据布局----------
'3.4.    IC卡数据格式
'根据IC卡结构，以方便接口使用为前提，定义如下IC卡参数串。无论读写卡，卡控制模块与接口模块和中心模块都使用此卡数据格式。基本规范为：日期长度10位（格式yyyy-mm-dd）；金额长度12位（格式000000000.00），长度不足时前面补0，数字类型长度不足时也前面补0；字符类型长度不足时后面空格添满长度。国标码格式类型在保存上全部按字符规则（实际可能是数字）。
'卡数据统一成一个文件，文件中有两段若干个字节保留字段：
'序号    字节    数据元　    格式    长度
'（字节）
'1   1-14    医疗证号（内部卡号）    字符    14
'2   15-24   基本医疗保险个人帐号    字符    10
'3   25-28   医疗人员类别（参保状态）    数字    4
'4   29-32   医保待遇（政策）类别，结算依据  数字    4
'5   33  特殊病标志  字符    1
'6   34  公务员参保标志  字符    1
'7   35-52   公民身份号  GB11643 18
'8   53-62   姓名    字符    10
'9   63-66   性别    GB2261  4
'10  67-68   民族    GB3304  2
'11  69-74   出生地  GB/T 2260   6
'12  75-84   出生日期    日期　  10
'13  85-124  单位名称    字符    40
'14  125-134 医保年度结转时间    日期　  10
'15  135-146 上年结转余额    金额　  12
'16  147-156 医保信息更新时间    日期　  10
'17  157-168 个帐当年拨付额度 /当年实际拨付  金额    12
'18  169-170 自负比例递减比例    数字    2
'19  171-180 保留字段1   字符    10
'20  181-194 锁卡状态，全为0表示未锁，否则为指定交易写卡失败 数字    14
'21  195-206 基本医疗保险个人账户余额    金额    12
'22  207-218 个帐当年使用累计金额    金额    12
'23  219-230 个帐历年使用累计金额    金额    12
'24  231-242 年度个人自负累计金额（普通门诊）    金额    12
'25  243-254 当年医保总累计金额  金额    12
'26  255-266 当年大病累计金额    金额    12
'27  267-278 当年门诊统筹支付累计金额    金额    12
'28  279-290 当年住院统筹支付累计金额    金额    12
'29  291-302 住院超过起付线累计金额  金额    12
'30  303-314 起付线后自负累计金额    金额    12
'31  315-326 当年门诊特殊病累计金额  金额    12
'32  327-338 当年公务员累计金额  金额　  12
'33  339-341 当年门诊统筹次数器  数字　  3
'34  342-344 当年住院统筹次数器  数字　  3
'35  345-347 当年医保累加次数器  数字　  3
'36  348 当前住院标志    字符    1
'37  349-352 结算次数    数字　  4
'38  353-368 保留字段2   字符    16
'
'其中:
'当年账户余额 = 个帐当年拨付额度 / 实付 - 个帐当年使用累计金额
'历年账户余额 = 上年结转余额 - 个帐历年使用累计金额
'保留字段为加密字符串，其内容可能不是标准可视ASCII码，保存到数据库时需要注意。
'
'读卡方式:
'1：读取第一个文件，由于只有一个文件，因此读取一个文件和读取所有文件含义相同
'2：读取第二个文件，医院端无法使用
'10: 读取所有文件
'
'
'IC卡格式用例（）：
'"1111111111111122222222220033001111555555555555555555测试者    男  11浙江  1977-05-25测试者单位                              2002-02-05000001000.002002-02-05000002000.0005          00000000000000000003000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.00000000000.0000000000000000000000000.00000000000.00000000000.00000000000.00"


'返回值定义：0-正常;>0-存在警告;<0-失败，错误信息在strReturnMsg中
Private Declare Function LHYB_Init Lib "BargaingApply" Alias "f_UserBargaingInit" _
    (ByVal StrInput As String, ByVal strReturnMsg As String, ByVal strOutput As String) As Integer
Private Declare Function LHYB_Close Lib "BargaingApply" Alias "f_UserBargaingClose" _
    (ByVal StrInput As String, ByVal strReturnMsg As String, ByVal strOutput As String) As Integer
Private Declare Function LHYB_Business Lib "BargaingApply" Alias "f_UserBargaingApply" _
    (ByVal intCode As Integer, ByVal dblSequence As Double, ByVal StrInput As String, _
    ByVal strReturnMsg As String, ByVal strOutput As String) As Integer

Type IC_Struct
    IC卡数据                As String           '保存病人IC卡内完整数据
    医疗证号                As String
    帐号                    As String
    人员类别                As String
    待遇类别                As String
    特殊病                  As Byte
    公务员                  As Byte
    身份号                  As String
    姓名                    As String
    性别                    As String
    民族                    As String
    出生地                  As String
    出生日期                As String
    单位名称                As String
    结转时间                As String
    结转余额                As Double
    更新时间                As String
    当年实际拨付            As Double
    自负比例                As Double
    卡状态                  As String
    个人账户余额            As Double
    个帐当年使用累计        As Double
    个帐历年使用累计        As Double
    个人自负累计金额        As Double
    总累计金额              As Double
    大病累计金额            As Double
    门诊统筹累计金额        As Double
    住院统筹累计金额        As Double
    住院超过起付线累计金额  As Double
    起付线后自负累计金额    As Double
    门诊特殊病累计金额      As Double
    公务员累计金额          As Double
    门诊统筹次数器          As Double
    住院统筹次数器          As Double
    医保累加次数器          As Double
    住院标志                As Byte
    结算次数                As Integer

'以下内容为附加内容
    mstr医院编码 As String
    mstr医院等级 As String
    mstr业务类型 As String
    mlng疾病ID As Long
    mstr交易流水号 As String
    mstr门诊单据号 As String
    mstr处方号 As String
    mdbl非医保金额 As Double
    mstr结算入口参数串 As String
End Type
Public IC_Data_宁海 As IC_Struct

Private mintFunc As Long   '功能号
Private mstrFunc As String  '函数名
Private mstrInput As String '入参
Private mstrOutput As String '输出串
Private mstrMsg As String   '返回的信息

Public gcn宁海 As New ADODB.Connection

Private mblnInit As Boolean

Public Enum Function_宁海
    InitInsure = 0                '打开
    EndInsure = 1               '关闭
    ReadIC = 22             '读卡
    GetSequence = 23        '获取交易流水号
    PreRegist = 27          '挂号预结算
    Regist = 28             '挂号结算
    RegistDel = 31          '挂号作废/门诊作废
    PreClinic = 29          '门诊预结算
    clinic = 30             '门诊结算
    ClinicDel = 31          '门诊结算作废
    Comein = 32             '入院登记
    ComeIndel = 40          '取消入院登记
    ChargeDetail = 33       '住院记帐/医嘱执行
    PreSettle = 34          '住院预结算
    Settle = 36             '住院结算
    SettleDel = 37          '因不支持中途结算，因此，取消出院的同时即作废结算
    ModifyPatient = 38      '修改在院病人信息
    出院病历 = 35           '填写出院病历(相当于出院)
    医保转自费病人 = 39     '医保病人转为自费病人
    转诊转院申请 = 41
    转诊转院查询 = 42
    RequestBusiness = 43    '查询交易结果
    Decide = 49             '交易确认
    ConfigPara = 52         '接口参数配置
End Enum

Private Const strSplit As String = "%%"
Private Const strField As String = "~"

Public Function ReadIC_宁海(Optional ByVal str医保号 As String = "") As Boolean
    Dim arrReturn
    Dim intReturn As Integer
    Dim strTest As String, strBit As String
    
    '完成IC卡的读操作，返回数据填写到IC_Data_宁海
    Call Interface_Prepare_宁海(ReadIC, IIf(str医保号 <> "", "0", "1") & "~" & str医保号 & "~~~10", "")
    intReturn = Interface_Exec_宁海()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    '按接口要求进行检查
    '交易结果~错误信息 + 更新卡结果 + 空 + 读出的IC卡数据 + 身份验证结果（15位，每位代表一个含意）+ 空
    '身份验证结果含义（位数从左边开始，判断优先级：987126534，各位数字为0表示对应身份验证正常）：
    '第1位: 人员黑名单冻结 第2位: 卡被冻结或作废
    '第3位: 当年账户被冻结 第4位: 往年帐户被冻结
    '第5位: 住院冻结 第7位: 内部保留位
    '第6位：需要圈存或结转(0成功或不需要圈存，1需要圈存但没有圈存，2失败)
    '第7位：卡数据需要更新(0成功或不需要更新，1卡被锁且获得了数据但没有更新，2中心有被医院取消的结算数据但卡没有更新3卡被本医院锁但更新失败，4卡被其他医院锁但由于无法连接中心或中心没有数据更新失败，5其他原因造成需要更新但更新失败)
    '第8位：上次结算写卡失败被锁，不可进行任何处理，必须先解锁（0：不 1：是）
    '第9位：参保人员是否有效（0：正常参保 1：没参保或参保状态无效）
    '其它位: 内部保留
    arrReturn = Split(mstrMsg, "~")
    strTest = arrReturn(5)
    If Mid(strTest, 9, 1) <> 0 Then
        MsgBox "该病人没有参保或参保状态无效！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 8, 1) <> 0 Then
        MsgBox "上次结算写卡失败被锁，不可进行任何处理，必须先解锁！", vbInformation, gstrSysName
        Exit Function
    End If
    strBit = Mid(strTest, 7, 1)
    If strBit <> 0 Then
        If strBit = 1 Then
            MsgBox "卡被锁且获得了数据但没有更新！", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 2 Then
            MsgBox "中心有被医院取消的结算数据但卡没有更新！", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 3 Then
            MsgBox "卡被本医院锁但更新失败！", vbInformation, gstrSysName
            Exit Function
        ElseIf strBit = 4 Then
            MsgBox "卡被其他医院锁但由于无法连接中心或中心没有数据更新失败！", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "其他原因造成需要更新但更新失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Mid(strTest, 1, 1) <> 0 Then
        MsgBox "人员黑名单冻结！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 2, 1) <> 0 Then
        MsgBox "卡被冻结或作废！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 6, 1) <> 0 Then
        MsgBox "需要圈存或结转！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 5, 1) <> 0 Then
        MsgBox "住院冻结！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 3, 1) <> 0 Then
        MsgBox "当年账户被冻结！", vbInformation, gstrSysName
        Exit Function
    End If
    If Mid(strTest, 4, 1) <> 0 Then
        MsgBox "往年账户被冻结！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '将数据填写到结构体中
    IC_Data_宁海.IC卡数据 = arrReturn(4)
    If Not ExchangeICData(arrReturn(4), str医保号) Then Exit Function
    
    ReadIC_宁海 = True
End Function

Private Function ExchangeICData(ByVal strBuffer As String, Optional ByVal str医保号 As String) As Boolean
    On Error GoTo errHand
    With IC_Data_宁海
        If str医保号 = "" Then
            .医疗证号 = Getsubstr(strBuffer, 1, 1, 14)
        Else
            .医疗证号 = str医保号
        End If
        .帐号 = Getsubstr(strBuffer, 1, 171, 10)        '取的保留字段1，此处和接口文档不符
        .人员类别 = Getsubstr(strBuffer, 1, 25, 4)
        .待遇类别 = Getsubstr(strBuffer, 1, 29, 4)
        .特殊病 = Getsubstr(strBuffer, 1, 33, 1)
        .公务员 = Getsubstr(strBuffer, 1, 34, 1)
        .身份号 = Getsubstr(strBuffer, 1, 35, 18)
        .姓名 = Getsubstr(strBuffer, 1, 53, 10)
        .性别 = Getsubstr(strBuffer, 1, 63, 4)
        .民族 = Getsubstr(strBuffer, 1, 67, 2)
        .出生地 = Getsubstr(strBuffer, 1, 69, 6)
        .出生日期 = Getsubstr(strBuffer, 1, 75, 10)
        .单位名称 = Getsubstr(strBuffer, 1, 85, 40)
        .结转时间 = Getsubstr(strBuffer, 1, 125, 10)
        .结转余额 = Val(Getsubstr(strBuffer, 1, 135, 12))
        .更新时间 = Getsubstr(strBuffer, 1, 147, 10)
        .当年实际拨付 = Val(Getsubstr(strBuffer, 1, 157, 12))
        .自负比例 = Val(Getsubstr(strBuffer, 1, 169, 2))
        .卡状态 = Getsubstr(strBuffer, 1, 181, 14)
        .个人账户余额 = Val(Getsubstr(strBuffer, 1, 195, 12))
        .个帐当年使用累计 = Val(Getsubstr(strBuffer, 1, 207, 12))
        .个帐历年使用累计 = Val(Getsubstr(strBuffer, 1, 219, 12))
        .个人自负累计金额 = Val(Getsubstr(strBuffer, 1, 231, 12))
        .总累计金额 = Val(Getsubstr(strBuffer, 1, 243, 12))
        .大病累计金额 = Val(Getsubstr(strBuffer, 1, 255, 12))
        .门诊统筹累计金额 = Val(Getsubstr(strBuffer, 1, 267, 12))
        .住院统筹累计金额 = Val(Getsubstr(strBuffer, 1, 279, 12))
        .住院超过起付线累计金额 = Val(Getsubstr(strBuffer, 1, 291, 12))
        .起付线后自负累计金额 = Val(Getsubstr(strBuffer, 1, 303, 12))
        .门诊特殊病累计金额 = Val(Getsubstr(strBuffer, 1, 315, 12))
        .公务员累计金额 = Val(Getsubstr(strBuffer, 1, 327, 12))
        .门诊统筹次数器 = Val(Getsubstr(strBuffer, 1, 339, 3))
        .住院统筹次数器 = Val(Getsubstr(strBuffer, 1, 342, 3))
        .医保累加次数器 = Val(Getsubstr(strBuffer, 1, 345, 3))
        .住院标志 = Val(Getsubstr(strBuffer, 1, 348, 1))
        .结算次数 = Val(Getsubstr(strBuffer, 1, 349, 4))
    End With
    
    ExchangeICData = True
    Exit Function
errHand:
End Function

Public Function Getsubstr(ByVal strBuffer As String, ByVal intBase As Integer, ByVal intStart As Integer, ByVal intLen As Integer) As String
    Dim intMAX As Integer   '整个串的实际长度
    '获取子串，该串开始位置减去基数就得到真实的起始位置
    intStart = intStart - intBase + 1
    Getsubstr = Trim(StrConv(MidB(StrConv(strBuffer, vbFromUnicode), intStart, intLen), vbUnicode))
End Function

Public Function Init_宁海() As Boolean
    Dim intReturn As Integer
    Call Interface_Prepare_宁海(InitInsure, "", "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    Init_宁海 = True
End Function

Public Sub Interface_Prepare_宁海(ByVal intFunc As Integer, ByVal StrInput As String, ByVal strOutput As String, Optional ByVal str交易流水号 As String = "")
    '做接口调用前的准备工作
    mintFunc = intFunc
    mstrInput = StrInput
    mstrOutput = strOutput
    IC_Data_宁海.mstr交易流水号 = Trim(TruncZero(str交易流水号))
    Call DebugTool("函数:" & mintFunc & ";入参:" & mstrInput)
    
    Select Case mintFunc
    Case InitInsure
        mstrFunc = "IntInsure"
    Case EndInsure
        mstrFunc = "EndInsure"
    Case ReadIC
        mstrFunc = "ReadIC"
    Case GetSequence
        mstrFunc = "GetSequence"
    Case PreRegist
        mstrFunc = "PreRegist"
    Case Regist
        mstrFunc = "Regist"
    Case RegistDel
        mstrFunc = "RegistDel"
    Case PreClinic
        mstrFunc = "PreClinic"
    Case clinic
        mstrFunc = "Clinic"
    Case ClinicDel
        mstrFunc = "ClinicDel"
    Case Comein
        mstrFunc = "ComeIn"
    Case ComeIndel
        mstrFunc = "ComeInDel"
    Case ChargeDetail
        mstrFunc = "ChargeDetail"
    Case PreSettle
        mstrFunc = "PreSettle"
    Case Settle
        mstrFunc = "Settle"
    Case SettleDel
        mstrFunc = "SettleDel"
    Case ModifyPatient
        mstrFunc = "ModifyPatient"
    Case RequestBusiness
        mstrFunc = "RequestBusiness"
    Case Decide
        mstrFunc = "Decide"
    Case ConfigPara
        mstrFunc = "ConfigPara"
    Case 转诊转院申请
        mstrFunc = "转诊转院申请"
    Case 转诊转院查询
        mstrFunc = "转诊转院查询"
    End Select
    
End Sub

Public Function Interface_Exec_宁海() As Integer
    '执行接口指定功能
    Dim intReturn  As Integer
    Dim dbl交易流水号 As Double
    On Error GoTo errHand
    
    mstrInput = "$$" & mstrInput & "$$"
    mstrMsg = "$$" & String(3000, " ") & "$$"
    dbl交易流水号 = CDbl(Val(IC_Data_宁海.mstr交易流水号))
    Select Case mintFunc
    Case InitInsure
        intReturn = LHYB_Init(mstrInput, mstrMsg, mstrOutput)
    Case EndInsure
        intReturn = LHYB_Close(mstrInput, mstrMsg, mstrOutput)
        Interface_Exec_宁海 = (intReturn >= 0)
        Exit Function
    Case Else
        intReturn = LHYB_Business(mintFunc, dbl交易流水号, mstrInput, mstrMsg, mstrOutput)
    End Select
    
    Call DebugTool("输入信息:" & mstrMsg)
    mstrMsg = Replace(mstrMsg, "$$", "")
    Interface_Exec_宁海 = intReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Interface_Exec_宁海 = -1
    mstrMsg = "-1~未知错误！~~~~~~~~~~~~"
End Function

Public Function Interface_Analyse_宁海() As Boolean
    '分析接口返回的数据
    
End Function

Public Function 医保初始化_宁海(Optional ByVal blnTest As Boolean = False) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim strServer As String, strUser As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If mblnInit = False Then
        If Not blnTest Then '如果是测试，则说明是保险参数设置处调用
            '读出连接医保服务器的配置
            gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_宁海)
            
            Do Until rsTemp.EOF
                Select Case rsTemp("参数名")
                    Case "医保用户名"
                        strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保服务器"
                        strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保用户密码"
                        strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                End Select
                rsTemp.MoveNext
            Loop
            
            If OraDataOpen(gcn宁海, strServer, strUser, strPass, False) = False Then
                MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If Not Init_宁海() Then Exit Function
        
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_宁海)
        IC_Data_宁海.mstr医院编码 = Nvl(rsTemp!医院编码)
        '取医院等级
        If IC_Data_宁海.mstr医院编码 <> "" Then
            gstrSQL = "Select YYDJ From SIM_YLJG Where YYBH='" & IC_Data_宁海.mstr医院编码 & "'"
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn宁海
            IC_Data_宁海.mstr医院等级 = Nvl(rsTemp!YYDJ)
        End If
        
        If Not blnTest Then mblnInit = True
    End If
    
    医保初始化_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保设置_宁海() As Boolean
    医保设置_宁海 = frmSet宁海.参数设置
End Function

Public Function 医保终止_宁海() As Boolean
    Call Interface_Prepare_宁海(EndInsure, "", "")
    Call Interface_Exec_宁海
    
    mblnInit = False
    医保终止_宁海 = True
End Function

Public Function 身份标识_宁海(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '完成医保病人身份的识别
    身份标识_宁海 = frmIdentify宁海.GetPatient(bytType, lng病人ID)
End Function

Private Function Get住院号(ByVal lng病人ID As Long, Optional ByVal blnNew As Boolean = False) As String
    Dim strText As String
    Dim str就诊时间 As String
    Dim str当前时间 As String
    Dim strSequence As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDO As Integer, intCOUNT As Integer, intPos As Integer
    '将当前时间、就诊时间进行处理，转换为唯一的流水号标识
    '编程思路：将年、月、日、时、分、秒都转换为一个字母的形式表示，因为一共只有12位
    intCOUNT = 6
    intPos = 1
    str当前时间 = Format(zlDatabase.Currentdate, "yyMMddHHmmss")
    
    '获取该病人的就诊时间
    gstrSQL = "Select 退休证号 From 保险帐户" & _
        " Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取该病人的就诊时间", TYPE_宁海, lng病人ID)
    '肯定不会为空
    If Not blnNew Then
        Get住院号 = Nvl(rsTemp!退休证号)
    Else
        Get住院号 = CStr(zlDatabase.GetNextID("部门表"))
    End If
End Function

Private Function Get自付比例(ByVal int项目类型 As Integer, ByVal str医保编码 As String, Optional ByVal str就诊类型 As String = "11") As Double
    Dim rsTemp As New ADODB.Recordset
    '返回指定项目的自负比例与单价限额
    '函数入参说明：
'    (sLbbz in hi_zymx.lbbz%type,   -- 1药品2诊疗
'     sXmbh in hi_zymx.xmbh%type,   -- 项目编号
'     iDylb in sio_ybdyzb.dylb%type,-- 待遇类别
'     sYydj in sio_ybfdjs.yydj%type,-- 医院等级
'     sJzlx in sio_jzlx.jzlx%type,  -- 就诊类型
'     iDfff in number,              -- 1单方2复方
'     nJbbm in sim_jbda.jbbm%type   -- 疾病编码
'     ) return number is
    gstrSQL = "Select orafGetzfbl(" & int项目类型 & ",'" & str医保编码 & "','" & IC_Data_宁海.待遇类别 & "'," & _
        "'" & IC_Data_宁海.mstr医院等级 & "','" & str就诊类型 & "',1," & IC_Data_宁海.mlng疾病ID & ") from dual"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcn宁海
    Get自付比例 = Nvl(rsTemp.Fields(0).Value, 0)
End Function

Private Function Get门诊流水号() As String
    '返回门诊流水号
    '10位数字，规则：取部门表的序列后十位
    Get门诊流水号 = Right(CStr(zlDatabase.GetNextID("部门表")), 10)
End Function

Public Function 门诊虚拟结算_宁海(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim strBill As String                   '单据头
    Dim strDetail As String                 '明细串
    Dim strErrInfo As String
    Dim lngPatient As Long                  '病人ID
    Dim intReturn As Integer
    Dim strBalance As String                '接口返回的结算信息
    Dim arrBalance
    Dim str业务类型 As String               '保险帐户中的业务类型，为1表示急诊
    Dim strDepart As String, strDoctor As String
    Dim lngDisease As Long, strDiseaseCode As String, strDiseaseName As String
    Dim dbl费用总额 As Double, dbl非医保总额 As Double, dbl自负比例 As Double
    Dim int明细数 As Integer, int单复方 As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    
    Const int费用总额 As Integer = 0
    Const int自费总额 As Integer = 1
    Const int自理费用 As Integer = 2
    Const int统筹基金 As Integer = 3
    Const int往年帐户 As Integer = 4
    Const int当年帐户 As Integer = 5
    Const int大病救助 As Integer = 6
    Const int公务员补助 As Integer = 7
    Const int单位支付 As Integer = 8
    Const int个人自付 As Integer = 9
    
    On Error GoTo errHand
    
    IC_Data_宁海.mstr门诊单据号 = Get门诊流水号
    IC_Data_宁海.mstr处方号 = Format(zlDatabase.Currentdate(), "yyyyMMddHHmmss")
'    入口参数 (Data)
'    是否有医保卡 + IC信息 + 空~空 + 本次结算单据张数 + 医保收费项目列表Clinic
'    Clinic结构体（[]表示可以重复）：Clinic = [Bill(单据)] + [Prescription（明细）]。
'    门诊挂号和收费时，需要通过此字符串结构将费用数据传递到结算函数中。单据号BillID是单据唯一的标志，在HIS不能重复。其中疾病、科室未知时填写"0"，非医保总额指"不在医保目录范围内项目"的总额，不包括医保目录范围内项目中自负比例部分。保存的字符串结构按如下格式组装（中间以%%分隔，日期类型转换为yyyy.mm.dd格式）：
'    Bill = 单据号(N10) + 门诊号(N10) +处方号码(VC15) +就诊日期(Dt) +收费类型(N1:0门诊挂号，1门诊收费，2急诊收费) +科室名称(VC20) +医生姓名(VC10)+疾病编号(VC12)+疾病名称(VC50)+疾病描述(VC255) + 此单据中非医保项目总额N(12,2) + 此单据中收费明细个数(count)Integer（不包括非医保纪录，明细纪录条数应等于Count值）；
'    Prescription = 单据号码(N10)+药品诊疗类型(N1:1药品，2诊疗)+项目编号(N10)+项目医院端名称(VC80)+ 医院端规格(VC20) + 单复方标志(N1:0非草药(不存在单复方)，1草药单方，2草药复方) + 单价N(14,4) + 数量N(14,4)+自付比率N(5,4)。
    With rs明细
        '取费用总额
        Do While Not .EOF
            lngPatient = !病人ID
            dbl费用总额 = dbl费用总额 + Nvl(!实收金额, 0)
            
            '判断是否医保项目
            gstrSQL = "Select B.名称,B.规格,A.项目编码 AS 医保编码 " & _
                    "From 保险支付项目 A,收费细目 B " & _
                    "Where B.ID=[1] And A.收费细目ID=B.ID And A.险类=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保项目编号、项目名称及规格", CLng(!收费细目ID), TYPE_宁海)
            If rsTemp.RecordCount = 0 Then
                dbl非医保总额 = dbl非医保总额 + Nvl(!实收金额, 0)
            Else
                If Nvl(!实收金额, 0) <> 0 Then int明细数 = int明细数 + 1
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        '只传入医保项目明细
        Do While Not .EOF
            If Nvl(!是否医保, 0) <> 0 And Nvl(!实收金额, 0) <> 0 Then
                '决定单复方标志
                int单复方 = IIf(!收费类别 = "7", 1, 0)
                
                '提取医保项目编号、项目名称及规格
                        
                ''''陈东 20041228
               'gstrSQL = "Select B.名称,B.规格,A.项目编码 AS 医保编码 " & _
               '         "From 保险支付项目 A,收费细目 B " & _
               '         "Where B.ID=" & !收费细目ID & " And A.收费细目ID=B.ID And A.险类=" & TYPE_宁海
               
                gstrSQL = "Select C.名称 as 大类,B.名称,B.规格,A.项目编码 AS 医保编码 " & _
                        "From 保险支付项目 A,收费细目 B,保险支付大类 C " & _
                        "Where B.ID=[1] And A.收费细目ID=B.ID  And A.大类ID=C.ID And A.险类=[2]"
                
                ''''
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保项目编号、项目名称及规格", CLng(!收费细目ID), TYPE_宁海)
                
                If rsTemp.RecordCount <> 0 Then
                    '提取自负比例
                    
                    ' 陈东  20041228
                    'dbl自负比例 = Get自付比例(IIf(InStr(1, "5,6,7", !收费类别) > 0, 1, 2), rsTemp!医保编码, "11")
                    dbl自负比例 = Get自付比例(IIf(rsTemp!大类 = "药品", 1, 2), rsTemp!医保编码, IC_Data_宁海.mstr业务类型)
                    
                    
                    If dbl自负比例 < 0 Then
                        MsgBox "项目[" & rsTemp!名称 & "]的自负比例获取错误！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    
                    ' 陈东  20041228
                    'strDetail = strDetail & strSplit & IC_Data_宁海.mstr门诊单据号 & strSplit & IIf(InStr(1, "5,6,7", !收费类别) <> 0, 1, 2) & strSplit & _
                    '    rsTemp!医保编码 & strSplit & ToVarchar(rsTemp!名称, 80) & strSplit & ToVarchar(Nvl(rsTemp!规格), 20) & strSplit & _
                   '    int单复方 & strSplit & Format(!单价, "#0.0000") & strSplit & _
                   '     Format(!数量, "#0.0000") & strSplit & dbl自负比例
                        
                    strDetail = strDetail & strSplit & IC_Data_宁海.mstr门诊单据号 & strSplit & IIf(rsTemp!大类 = "药品", 1, 2) & strSplit & _
                        rsTemp!医保编码 & strSplit & ToVarchar(rsTemp!名称, 80) & strSplit & ToVarchar(Nvl(rsTemp!规格), 20) & strSplit & _
                       int单复方 & strSplit & Format(!单价, "#0.0000") & strSplit & _
                        Format(!数量, "#0.0000") & strSplit & dbl自负比例
                        
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        If strDetail <> "" Then strDetail = Mid(strDetail, 3)
    End With
    
    '提取该病人的疾病信息
    lngDisease = IC_Data_宁海.mlng疾病ID
    
    strDiseaseCode = "0"
    strDiseaseName = "未知"
    If lngDisease <> 0 Then
        gstrSQL = "Select JBBZDM,JBMC From SIM_JBDA Where JBBM=" & lngDisease
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn宁海
        If rsTemp.RecordCount <> 0 Then
            strDiseaseCode = lngDisease
            strDiseaseName = rsTemp!JBMC
        End If
    End If
    
    '提取该单据的开单科室
    strDoctor = Trim(Nvl(rs明细!开单人))
    If strDoctor <> "" Then
        gstrSQL = "SELECT C.名称 AS 开单科室 " & _
                 " FROM 部门人员 A,部门性质说明 B,部门表 C " & _
                 " WHERE A.人员ID= " & _
                 "     (SELECT ID FROM 人员表 WHERE 姓名=[1]) " & _
                 " AND A.部门ID=B.部门ID AND A.部门ID=C.ID AND B.工作性质='临床' AND 服务对象 IN (1,3) " & _
                 " AND ROWNUM<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取开单科室", CStr(rs明细!开单人))
        strDepart = "0"
        If rsTemp.RecordCount <> 0 Then strDepart = Nvl(rsTemp!开单科室)
    End If
    
'    strBill = IC_Data_宁海.mstr门诊单据号 & strSplit & IC_Data_宁海.mstr门诊单据号 & strSplit & _
'        IC_Data_宁海.mstr处方号 & strSplit & Format(zlDatabase.Currentdate, "yyyy.MM.dd") & strSplit & _
'        IIf(str业务类型 = "1", "2", "1") & strSplit & strDepart & strSplit & strDoctor & strSplit & _
'        strDiseaseCode & strSplit & strDiseaseName & strSplit & strSplit & _
'        Format(dbl非医保总额, "#0.00") & strSplit & int明细数
    'Modified by ZYB 2006-04-12，固定传“1”表示门诊收费，将 IIf(str业务类型 = "1", "2", "1") 替换为 "1"
    strBill = IC_Data_宁海.mstr门诊单据号 & strSplit & IC_Data_宁海.mstr门诊单据号 & strSplit & _
        IC_Data_宁海.mstr处方号 & strSplit & Format(zlDatabase.Currentdate, "yyyy.MM.dd") & strSplit & _
        "1" & strSplit & strDepart & strSplit & strDoctor & strSplit & _
        strDiseaseCode & strSplit & strDiseaseName & strSplit & strSplit & _
        Format(dbl非医保总额, "#0.00") & strSplit & int明细数
    
    'Modified by ZYB 2006-04-12，根据宁海县医保中心2006-04-06下发的文件要求修改，入参最后增加就诊类型
    StrInput = "1" & strField & IC_Data_宁海.IC卡数据 & strField & strField & strField & "1" & strField & strBill & strSplit & strDetail & strField & IC_Data_宁海.mstr业务类型
    Call Interface_Prepare_宁海(PreClinic, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
'        If (Trim(Split(mstrMsg, "~")(7)) <> "") Then
'            strErrInfo = strErrInfo & vbCrLf & "详细信息：" & vbCrLf & Trim(Split(mstrMsg, "~")(7))
'        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    
    '取结算信息
    strBalance = Split(mstrMsg, "~")(6)
    '结算结果(中间用%%分隔)：①费用总额+②自费总额(非医保，只能现金)+③自理总额(目录内自负比例部分)+
    '④统筹基金支付+⑤往年帐户支付+⑥当年帐户支付+⑦大病救助支付+⑧公务员补助支付+⑨单位支付 + '
    '⑩个人自负 (账户不足现金) + 个人现金支付 + 本次门诊特病结算前费用累计
    '费用总额①=②+③+④+⑤+⑥+⑦+⑧+⑨+⑩，现金支付=②+③+⑨+⑩
    arrBalance = Split(strBalance, "%%")
    str结算方式 = "个人帐户;" & Val(arrBalance(int往年帐户)) + Val(arrBalance(int当年帐户)) & ";0"
    If Val(arrBalance(int统筹基金)) <> 0 Then str结算方式 = str结算方式 & "|统筹基金;" & Val(arrBalance(int统筹基金)) & ";0"
    If Val(arrBalance(int公务员补助)) <> 0 Then str结算方式 = str结算方式 & "|公务员补助;" & Val(arrBalance(int公务员补助)) & ";0"
    If Val(arrBalance(int大病救助)) <> 0 Then str结算方式 = str结算方式 & "|大病救助;" & Val(arrBalance(int大病救助)) & ";0"
    
    IC_Data_宁海.mstr结算入口参数串 = "1" & strField & strBill & strSplit & strDetail
    门诊虚拟结算_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_宁海(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim intReturn As Integer
    Dim lng病人ID As Long
    Dim blnTrans As Boolean
    Dim str交易流水号 As String
    Dim strErrInfo As String
    Dim strBalance As String
    Dim arrBalance
    Dim StrInput As String, strOutput As String
    Dim dbl统筹基金 As Double, dbl大病补助 As Double, dbl公务员补助 As Double, dbl单位支付 As Double
    Dim dbl当年帐户 As Double, dbl往年帐户 As Double, dbl当年帐户_余额 As Double, dbl往年帐户_余额 As Double
    Dim dbl费用总额 As Double, dbl现金支付 As Double    '现金支付包含单位支付
    Dim rsTemp As New ADODB.Recordset
    
    Const int费用总额 As Integer = 0
    Const int自费总额 As Integer = 1
    Const int自理费用 As Integer = 2
    Const int统筹基金 As Integer = 3
    Const int往年帐户 As Integer = 4
    Const int当年帐户 As Integer = 5
    Const int大病救助 As Integer = 6
    Const int公务员补助 As Integer = 7
    Const int单位支付 As Integer = 8
    Const int个人自付 As Integer = 9
    
    On Error GoTo errHand
    '取病人ID
    gstrSQL = "Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    '取余额
    gstrSQL = "Select 往年帐户余额,本年帐户余额 From 保险帐户 Where 险类=[2] And 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取余额", lng病人ID, TYPE_宁海)
    dbl当年帐户_余额 = Nvl(rsTemp!本年帐户余额, 0)
    dbl往年帐户_余额 = Nvl(rsTemp!往年帐户余额, 0)
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & clinic, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '门诊结算
    '缺省现金支付
    'Modified by ZYB 2006-04-12，根据宁海县医保中心2006-04-06下发的文件要求修改，入参最后增加就诊类型
    StrInput = "1~~1~~" & IC_Data_宁海.mstr结算入口参数串 & strField & UserInfo.姓名 & strField & IC_Data_宁海.mstr业务类型
    Call Interface_Prepare_宁海(clinic, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
'        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
'            strErrInfo = strErrInfo & vbCrLf & "详细信息：" & vbCrLf & "写卡失败，请根据写卡错误原因重新读卡或换机器重新读一遍卡以自动同步卡数据！"
'        End If
        Err.Raise 9000, gstrSysName, strErrInfo
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    '取结算信息
    strBalance = Split(mstrMsg, "~")(6)
    '结算结果(中间用%%分隔)：①费用总额+②自费总额(非医保，只能现金)+③自理总额(目录内自负比例部分)+
    '④统筹基金支付+⑤往年帐户支付+⑥当年帐户支付+⑦大病救助支付+⑧公务员补助支付+⑨单位支付 + '
    '⑩个人自负 (账户不足现金) + 个人现金支付 + 本次门诊特病结算前费用累计
    '费用总额①=②+③+④+⑤+⑥+⑦+⑧+⑨+⑩，现金支付=②+③+⑨+⑩
    arrBalance = Split(strBalance, "%%")
    dbl费用总额 = Val(arrBalance(int费用总额))
    dbl当年帐户 = Val(arrBalance(int当年帐户))
    dbl往年帐户 = Val(arrBalance(int往年帐户))
    dbl统筹基金 = Val(arrBalance(int统筹基金))
    dbl公务员补助 = Val(arrBalance(int公务员补助))
    dbl大病补助 = Val(arrBalance(int大病救助))
    dbl单位支付 = Val(arrBalance(int单位支付))
    dbl现金支付 = dbl费用总额 - dbl统筹基金 - dbl大病补助 - dbl公务员补助 - dbl当年帐户 - dbl往年帐户
    
    '大病自付=大病补助;超限自付=公务员补助
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_宁海 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0," & dbl往年帐户_余额 & "," & dbl当年帐户_余额 & "," & dbl往年帐户 & "," & dbl当年帐户 & "," & dbl费用总额 & "," & dbl现金支付 & ",0," & _
        dbl统筹基金 & "," & dbl统筹基金 & "," & dbl大病补助 & "," & dbl公务员补助 & "," & dbl当年帐户 + dbl往年帐户 & ",'" & IC_Data_宁海.mstr交易流水号 & "|" & IC_Data_宁海.待遇类别 & "|" & IC_Data_宁海.mstr业务类型 & "',NULL,NULL,'" & Replace(Split(mstrMsg, "~")(6), "'", "") & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '交易确认
    '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
    '缺省现金支付
    StrInput = "~~~~" & clinic & strField & IC_Data_宁海.mstr交易流水号 & strField & "0" & strField & "HIS成功！"
    Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "警告：本次交易确认失败，请记录下本次交易流水号，并通知系统管理员使用工具包再次确认该交易" & _
        vbCrLf & "交易流水号：" & str交易流水号
    End If
    门诊结算_宁海 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '交易确认
        '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
        '缺省现金支付
        StrInput = "~~~~" & clinic & strField & str交易流水号 & strField & "-1" & strField & "医保成功，而HIS失败！"
        Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    End If
End Function

Public Function 门诊结算冲销_宁海(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng冲销ID As Long
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim blnTrans As Boolean
    Dim str交易流水号 As String, str原交易流水号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & ClinicDel, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '取冲销记录的结帐ID
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    '取原交易流水号
    gstrSQL = "Select 支付顺序号 From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取原交易流水号", lng结帐ID)
    str原交易流水号 = Split(rsTemp!支付顺序号, "|")(0)
    
    '调用作废交易
    '是否有医保卡 + IC信息(可以传空)+ ~空~空 + 要作废的门诊/挂号结算交易号
    Call Interface_Prepare_宁海(ClinicDel, "1~~~~" & str原交易流水号, "", str交易流水号)
    intReturn = Interface_Exec_宁海()
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        Err.Raise 9000, gstrSysName, strErrInfo
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    '提取原结算记录，做为产生本次结算记录的依据
    gstrSQL = "Select * From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原结算记录，做为产生本次结算记录的依据", lng结帐ID)
    
    '保存结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_宁海 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",0," & _
        -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & _
        -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & str交易流水号 & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '交易确认
    '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
    '缺省现金支付
    StrInput = "~~~~" & ClinicDel & strField & str交易流水号 & strField & "0" & strField & "HIS成功！"
    Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "警告：本次交易确认失败，请记录下本次交易流水号，并通知系统管理员使用工具包再次确认该交易" & _
            vbCrLf & "交易流水号：" & str交易流水号
    End If
    门诊结算冲销_宁海 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '交易确认
        '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
        '缺省现金支付
        StrInput = "~~~~" & ClinicDel & strField & str交易流水号 & strField & "-1" & strField & "医保成功，而HIS失败！"
        Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    End If
End Function

Public Function 入院登记_宁海(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim str住院号 As String, str交易流水号 As String
    Dim str入院日期 As String, str医生 As String, str入院诊断 As String, str疾病编号 As String
    Dim lng病种ID As Long, lng科室ID As Long, str科室名称 As String, str科室编号 As String, str床号 As String
    Dim StrInput As String, strOutput As String
    Dim blnTrans As Boolean
    Dim str无卡病人 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & Comein, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    ',D.工作性质  ',临床部门 D  ' And B.ID=D.部门ID(+)
    '提取入院日期、医生、入院诊断、入院疾病、入院科室名称、医保科室编号、床号
    gstrSQL = " Select to_char(A.入院日期,'yyyy-MM-dd') 入院日期,B.ID AS 入院科室ID,B.编码 as 科室编码," & _
              " B.名称 科室,A.住院医师 医生,A.入院病床,C.病种ID " & _
              " From 病案主页 A,部门表 B,保险帐户 C " & _
              " Where A.病人ID=[1] And A.主页ID=[2]" & _
              " And A.入院科室ID=B.ID And A.病人ID=C.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院基本信息", lng病人ID, lng主页ID)
    str入院日期 = Format(rsTemp!入院日期, "yyyy.MM.dd")
    str医生 = Nvl(rsTemp!医生)
    str科室名称 = Nvl(rsTemp!科室)
    lng科室ID = Nvl(rsTemp!入院科室ID, 0)
    str科室编号 = Nvl(rsTemp!科室编码)
    str床号 = Nvl(rsTemp!入院病床)
    lng病种ID = Nvl(rsTemp!病种ID, 0)
    '取入院诊断
    str入院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, True, True)
'    '取医保端疾病标准编码
'    gstrSQL = "Select JBBZDM From SIM_JBDA Where JBBM=" & lng病种ID
'    If rsTemp.State = 1 Then rsTemp.Close
'    rsTemp.Open gstrSQL, gcn宁海
'    If rsTemp.RecordCount <> 0 Then
'        str疾病编号 = Nvl(rsTemp!JBBZDM, "0")
'    Else
'        str疾病编号 = "0"
'    End If
    
    '住院确认成功后住院号被使用，住院确认失败后住院号被作废，住院号都不能再使用
    str住院号 = Get住院号(lng病人ID, True)
    str无卡病人 = IS无卡病人(lng病人ID)
    '是否有医保卡 + IC信息(可以传空)+ ~空~空 + 住院号 + 入院日期 + 入院诊断医生姓名 + 入院病情描述 + 入院疾病序号 + 入院科室名称 + 入院科室ID + 床号。
    StrInput = Split(str无卡病人, "|")(0) & strField & Split(str无卡病人, "|")(1) & strField & strField & strField & str住院号 & strField & _
        str入院日期 & strField & str医生 & strField & str入院诊断 & strField & lng病种ID & strField & _
        str科室名称 & strField & lng科室ID & strField & str床号
    Call Interface_Prepare_宁海(Comein, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海()
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "详细信息：" & vbCrLf & "写卡失败，请根据写卡错误原因重新读卡或换机器重新读一遍卡以自动同步卡数据！"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    blnTrans = True
    
    '交易确认
    '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
    '缺省现金支付
    StrInput = "~~~~" & Comein & strField & str交易流水号 & strField & "0" & strField & "HIS成功！"
    Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then MsgBox "警告：本次交易确认失败，请记录下本次交易流水号，并通知系统管理员使用工具包再次确认该交易" & _
        vbCrLf & "交易流水号：" & str交易流水号, vbInformation, gstrSysName
        
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_宁海 & ",'顺序号','''" & str交易流水号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新顺序号")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_宁海 & ",'退休证号','''" & str住院号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新住院号")
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_宁海 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    入院登记_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        '交易确认
        '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
        '缺省现金支付
        StrInput = "~~~~" & Comein & strField & str交易流水号 & strField & "-1" & strField & "医保成功，而HIS失败！"
        Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    End If
End Function

Public Function 入院登记撤销_宁海(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim str交易流水号 As String, str住院号 As String, str无卡病人 As String
    Dim StrInput As String, strOutput As String
    Dim blnAllow As Boolean '是否允许撤销入院
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '未结费用为零，且未进行结算过，才允许取消入院登记
    
    blnAllow = True
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '判断该病人是否结算过，没有结算过的病人费用为零，说明需要调用就诊登记撤销
        gstrSQL = "Select 1 From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And Nvl(结帐ID,0)<>0 and Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否进行过费用结算", lng病人ID, lng主页ID)
        If Not rsTemp.EOF Then
            blnAllow = False
        End If
    Else
        blnAllow = False
    End If
    
    If Not blnAllow Then
        MsgBox "该病人存在未结费用或已进行过住院结算，不允许撤销入院！" & vbCrLf & _
        "（只有零费用，且未进行过结算的病人，才允许撤销医保入院）", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & ComeIndel, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    str住院号 = Get住院号(lng病人ID)
    str无卡病人 = IS无卡病人(lng病人ID)
    
    '是否有医保卡 + IC信息(可以传空)+ 空 + 空 + 要注销的住院号
    StrInput = Split(str无卡病人, "|")(0) & strField & Split(str无卡病人, "|")(1) & strField & strField & strField & str住院号
    Call Interface_Prepare_宁海(ComeIndel, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海()
    '交易结果~错误信息+写医保卡结果 + 空 + 写卡后IC卡数据
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "详细信息：" & vbCrLf & "写卡失败，请根据写卡错误原因重新读卡或换机器重新读一遍卡以自动同步卡数据！"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_宁海 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    入院登记撤销_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保转普通病人_宁海(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim intReturn As Integer
    Dim strErrInfo As String
    Dim str交易流水号 As String, str住院号 As String, str无卡病人 As String
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '未结费用为零，且未进行结算过，才允许取消入院登记
    
    '按应军杰的要求屏蔽掉
    '判断该病人是否结算过，没有结算过的病人费用为零，说明需要调用就诊登记撤销
'    gstrSQL = "Select 1 From 病人费用记录 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & " And Nvl(结帐ID,0)<>0 and Rownum<2"
'    Call OpenRecordset(rsTemp, "判断是否进行过费用结算")
'    If Not rsTemp.EOF Then
'        MsgBox "该病人已进行过住院结算，不允许转为普通病人！", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & 医保转自费病人, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    str住院号 = Get住院号(lng病人ID)
    str无卡病人 = IS无卡病人(lng病人ID)
    
    '是否有医保卡 + IC信息(可以传空)+ 空 + 空 + 要注销的住院号
    StrInput = Split(str无卡病人, "|")(0) & strField & Split(str无卡病人, "|")(1) & strField & strField & strField & str住院号
    Call Interface_Prepare_宁海(医保转自费病人, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海()
    '交易结果~错误信息+写医保卡结果 + 空 + 写卡后IC卡数据
    If intReturn <> 0 Then
        strErrInfo = "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If Val(Split(mstrMsg, "~")(2)) <> 0 Then
            strErrInfo = strErrInfo & vbCrLf & "详细信息：" & vbCrLf & "写卡失败，请根据写卡错误原因重新读卡或换机器重新读一遍卡以自动同步卡数据！"
        End If
        MsgBox strErrInfo, vbInformation, gstrSysName
        If intReturn < 0 Then Exit Function
    End If
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "撤消医保入院")
    MsgBox "操作成功，该医保病人已经转为普通病人！", vbInformation, gstrSysName
    
    医保转普通病人_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_宁海(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim intReturn As Integer
    Dim lng明细ID As Long
    Dim lng主页ID As Long
    Dim dbl自负比例 As Double
    Dim blnUpload As Boolean
    Dim strBalance As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str医保编码 As String, str医保名称 As String, str医保规格 As String, str医保单位 As String
    Dim str住院流水号 As String, str住院号 As String, str医保证号 As String, str交易流水号 As String
    Dim int记录数 As Integer, dbl非医保金额 As Double
    Dim str无卡病人 As String
    Dim dbl费用总额 As Double, dbl费用总额_YB As Double, dbl自费总额 As Double, dbl自理总额 As Double, dbl统筹基金 As Double
    Dim dbl帐户支付 As Double, dbl大病补助 As Double, dbl公务员补助 As Double, dbl单位支付 As Double, dbl起付线 As Double
    
    Const int费用总额 As Integer = 0
    Const int统筹基金 As Integer = 3
    Const int往年帐户 As Integer = 4
    Const int当年帐户 As Integer = 5
    Const int大病救助 As Integer = 6
    Const int公务员补助 As Integer = 7
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim gcn上传 As New ADODB.Connection
    
    On Error GoTo errHand
    
    '打开连接
    Set gcn上传 = GetNewConnection
    
    '先提取该病人的住院流水号
    gstrSQL = "Select A.业务类型,A.医保号,A.顺序号,A.IC,A.病种ID,B.住院次数 主页ID From 保险帐户 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "先提取该病人的住院流水号", lng病人ID)
    str医保证号 = rsTemp!医保号
    str住院流水号 = rsTemp!顺序号
    lng主页ID = rsTemp!主页ID
    IC_Data_宁海.IC卡数据 = Nvl(rsTemp!ic)
    IC_Data_宁海.mlng疾病ID = Nvl(rsTemp!病种ID, 0)
    str住院号 = Get住院号(lng病人ID)
    'Modified by ZYB 2006-04-12，根据宁海县医保中心2006-04-06下发的文件要求修改，入参最后增加就诊类型
    IC_Data_宁海.mstr业务类型 = Nvl(rsTemp!业务类型, "21")
    
    Call ExchangeICData(IC_Data_宁海.IC卡数据)
    
    '提取本次费用明细
    '陈东 20041228
    'gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
    '          "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格,A.开单人 AS 医生,A.登记时间 " & _
    '          "  From 病人费用记录 A,收费细目 B,保险支付项目 C " & _
    '          "  where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & " and A.记帐费用=1 And A.操作员姓名 is not null AND A.实收金额 IS NOT NULL " & _
    '          "        and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= " & TYPE_宁海 & _
    '          "  Order by A.病人ID,A.发生时间"
    gstrSQL = "Select D.名称 as 大类,A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格,A.开单人 AS 医生,A.登记时间 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,保险支付大类 D " & _
              "  where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND Nvl(A.实收金额,0)<>0 " & _
              "        and nvl(A.是否上传,0)=0 And C.大类ID=D.ID And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "  Order by A.病人ID,A.发生时间"
    
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", lng病人ID, lng主页ID, TYPE_宁海)
    
    '先删除所有明细
    gstrSQL = "Delete Hi_zymx_temp Where JYH='" & str住院流水号 & "'"
    gcn宁海.Execute gstrSQL
    
    With rs明细
        Do While Not .EOF
            str医保编码 = Nvl(!医保项目编码)
            
            If str医保编码 <> "" Then
                '获取自负比例
                lng明细ID = !ID
                ''陈东 20041228
                'dbl自负比例 = Get自付比例(IIf(InStr(1, "5,6,7", !收费类别) > 0, 1, 2), str医保编码, "21")
                dbl自负比例 = Get自付比例(IIf(rs明细!大类 = "药品", 1, 2), str医保编码, IC_Data_宁海.mstr业务类型)
                
                '再提取医院项目的规格及单位
                gstrSQL = "Select 名称,规格,计算单位 From 收费细目 Where ID=[1]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "再提取医院项目的规格及单位", CLng(!收费细目ID))
                str医保名称 = ToVarchar(Nvl(rsItem!名称), 40)
                str医保规格 = ToVarchar(Nvl(rsItem!规格), 30)
                str医保单位 = ToVarchar(Nvl(rsItem!计算单位), 8)
                
            '    系统标识: hi_zymx_temp
            '    序号    字段标识    字段名称    类型    长度    小数    允许空  缺省值  备注
            '    YYBH    医院编号    VARCHAR2    6       N
            '    JYH     交易号  NUMBER  20      N       医院填写住院登记交易号，结算时正式表中自动改为结算交易号
            '    JYH2    提交交易号  NUMBER  20      N   0   医院无需填写，提交时自动填写
            '    MXXH    明细序号    NUMBER  20      N       每个住院号费用明细序号都不重复
            '    JZLX    就诊类型    CHAR    2       N       11门诊，21住院
            '    YNBH    院内编号    VARCHAR2    12      N
            '    TYBZ    退药标志    CHAR    1       N       退费时为1，且数量为负数
            '    GRNM    个人社保编号    VARCHAR2    18      N
            '    LBBZ    类别标志    CHAR    1       N       1:药品  2:诊疗
            '    XMBH    项目编号    VARCHAR2    10      N
            '    XMMC    项目名称    VARCHAR2    40      N
            '    XMGG    项目规格    VARCHAR2    30
            '    XMDW    项目单位    VARCHAR2    8
            '    YZRQ    医嘱日期    DATE
            '    SSXM    医生姓名    VARCHAR2    20
            '    XMDJ    项目单价    NUMBER  12  2   N   0
            '    XMSL    项目数量    NUMBER  12  4   N   0   退费时为负数
            '    XMTS    项目贴数    NUMBER  6   2   N   0   始终为1
            '    XMJE    项目金额    NUMBER  10  4   N   0
            '    ZFBL    自负比例    NUMBER  5   4   N   0
            '    ZFJE    自负金额    NUMBER  12  2   N   0   没有意义，汇总结算时才获得
                
                gstrSQL = "Insert Into Hi_zymx_temp(" & _
                          "YYBH,JYH,MXXH,JZLX,YNBH,TYBZ,GRNM,LBBZ,XMBH,XMMC," & _
                          "XMGG,XMDW,YZRQ,SSXM,XMDJ,XMSL,XMTS,XMJE,ZFBL,ZFJE)" & _
                          "Values (" & _
                          "'" & IC_Data_宁海.mstr医院编码 & "'," & str住院流水号 & "," & lng明细ID & "," & _
                          "'21','" & str住院号 & "'," & IIf(!数量 < 0, 1, 0) & ",'" & str医保证号 & "'," & _
                          "" & IIf(rs明细!大类 = "药品", 1, 2) & ",'" & str医保编码 & "'," & _
                          "'" & str医保名称 & "','" & str医保规格 & "','" & str医保单位 & "'," & _
                          "To_Date('" & Format(!登记时间, "yyyy-MM-dd") & "','yyyy-MM-dd')," & _
                          "'" & Nvl(!医生) & "'," & Format(!价格, "#0.00") & "," & Format(!数量, "#0.0000") & ",1," & _
                          "" & Format(!金额, "#0.00") & "," & dbl自负比例 & "," & Round(Format(!金额, "#0.00") * dbl自负比例, 2) & ")"
                gcn宁海.Execute gstrSQL
                
                '取流水号
                Call Interface_Prepare_宁海(GetSequence, "~~~~" & ChargeDetail, "")
                intReturn = Interface_Exec_宁海
                If intReturn <> 0 Then
                    MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
                    If (intReturn < 0) Then Exit Function
                End If
                str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
                
                '上传明细
                '入参
                '空~空~空~空 + 住院号个数 + 住院号列表（用%%分隔）
'                    返回参数(包括"交易结果~错误信息~交易结果信息")
'                    交易结果~错误信息+空~空~空 + 无法保存的住院号列表 + 无法保存的医嘱流水号列表（此字段不用，为空） + 无法保存的费用流水号列表（列表间%%分隔）
'                    无法保存的住院号列表 = 住院号个数%%[住院号]
'                    无法保存的费用流水号列表= 不能保存的原因（1主健重复，2自负比例错误）%%记录条数%%[明细序号%%正确的自负比例（若不能保存的原因为1主健重复，则此域为空）]
'                    注: 住院返回的无法保存明细列表和门诊返回的不可报原因列表不同?
'                    函数返回值
'                    0函数调用成功，且本次费用全部通过校验提交
'                    -1函数调用失败
'                    -2至少有一个住院号由于不在住院中而校验失败，此时只读取无法保存的住院号列表，无法保存的费用明细列表域为空，即出口参数为：空~空~空+无法保存的住院号列表+空。
'                    -3至少有一条费用明细不能保存，此时只要读取无法保存的费用明细列表，无法保存的住院号列表域为空，即出口参数为：空~空~空+空+无法保存的费用明细列表。
                StrInput = "~~~~1~" & str住院号
                Call Interface_Prepare_宁海(ChargeDetail, StrInput, strOutput, str交易流水号)
                intReturn = Interface_Exec_宁海()
                If intReturn <> 0 Then
                    Select Case intReturn
                    Case -1
                        strErrInfo = "函数调用失败！"
                    Case -2
                        strErrInfo = "由于当前病人不在住院中，校验失败！"
                    Case -3
                        If Val(Split(Split(mstrMsg, strField)(7), "%%")(0)) = 1 Then
                            strErrInfo = "主键重复！"
                        Else
                            strErrInfo = "自付比例错误！"
                        End If
                    End Select
                    MsgBox strErrInfo, vbInformation, gstrSysName
                    If (intReturn < 0) Then Exit Function
                End If
                gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                gcn上传.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    
    '统计医保明细数及非医保项目总额
    With rsExse
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl费用总额 = dbl费用总额 + Nvl(!金额, 0)
            If Nvl(!医保项目编码) = "" Then
                dbl非医保金额 = dbl非医保金额 + Nvl(!金额, 0)
            Else
                int记录数 = int记录数 + 1
            End If
            .MoveNext
        Loop
    End With
    
    '准备进行预结算
    str无卡病人 = IS无卡病人(lng病人ID)
    '是否有医保卡 + IC信息 + 空~空  +住院号+本次结算明细条数+ 非医保项目总额
    StrInput = Split(str无卡病人, "|")(0) & strField & IIf(Split(str无卡病人, "|")(0) = 1, IC_Data_宁海.IC卡数据, Split(str无卡病人, "|")(1)) & strField & strField & strField & str住院号 & strField & int记录数 & strField & dbl非医保金额 & strField & IC_Data_宁海.mstr业务类型
    IC_Data_宁海.mstr结算入口参数串 = int记录数 & strField & dbl非医保金额
    Call Interface_Prepare_宁海(PreSettle, StrInput, strOutput)
    intReturn = Interface_Exec_宁海()
    If intReturn < 0 Then
        MsgBox "[" & mstrFunc & "]接口返回错误信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        Exit Function
    End If
    
    '分解结算信息
    '交易结果~错误信息+空~空~空 +
    '结算结果（费用总额%%自费总额%%自理总额%%统筹基金支付%%往年帐户支付%%当年帐户支付
    '%%大病支付%%公务员补助支付%%单位支付%%个人现金支付%%起付标准%%分段信息） + IC卡写卡前后信息(His不使用)
    strBalance = Split(mstrMsg, strField)(5)
    dbl费用总额_YB = Val(Split(strBalance, "%%")(int费用总额))
    dbl统筹基金 = Val(Split(strBalance, "%%")(int统筹基金))
    dbl帐户支付 = Val(Split(strBalance, "%%")(int当年帐户)) + Val(Split(strBalance, "%%")(int往年帐户))
    dbl大病补助 = Val(Split(strBalance, "%%")(int大病救助))
    dbl公务员补助 = Val(Split(strBalance, "%%")(int公务员补助))
    IC_Data_宁海.mdbl非医保金额 = dbl非医保金额
    
    If Format(dbl费用总额 - dbl非医保金额, "#0.00") <> Format(dbl费用总额_YB, "#0.00") Then
        MsgBox "HIS费用总额不等于医保费用总额！" & vbCrLf & _
        "医院：" & Format(dbl费用总额 - dbl非医保金额, "#0.00") & Space(10) & "医保：" & Format(dbl费用总额_YB, "#0.00"), vbInformation, gstrSysName
    End If
    
    住院虚拟结算_宁海 = "个人帐户;" & dbl帐户支付 & ";0"
    住院虚拟结算_宁海 = 住院虚拟结算_宁海 & "|统筹基金;" & dbl统筹基金 & ";0"
    住院虚拟结算_宁海 = 住院虚拟结算_宁海 & "|公务员补助;" & dbl公务员补助 & ";0"
    住院虚拟结算_宁海 = 住院虚拟结算_宁海 & "|大病救助;" & dbl大病补助 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_宁海(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng主页ID As Long
    Dim blnTrans As Boolean
    Dim str交易流水号 As String, str住院号 As String, strBalance As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str无卡病人 As String
    Dim dbl费用总额 As Double, dbl自费总额 As Double, dbl自理总额 As Double, dbl统筹基金 As Double, dbl现金支付 As Double
    Dim dbl帐户支付 As Double, dbl大病补助 As Double, dbl公务员补助 As Double, dbl单位支付 As Double, dbl起付线 As Double
    Dim dbl当年帐户 As Double, dbl往年帐户 As Double, dbl当年帐户_余额 As Double, dbl往年帐户_余额 As Double
    
    Const int费用总额 As Integer = 0
    Const int统筹基金 As Integer = 3
    Const int往年帐户 As Integer = 4
    Const int当年帐户 As Integer = 5
    Const int大病救助 As Integer = 6
    Const int公务员补助 As Integer = 7
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    'If MsgBox("医保接口不支持中途结算（一次住院只能进行一次结算），你确定要结算吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    '取主页ID
    gstrSQL = "Select 住院次数 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    lng主页ID = rsTemp!住院次数
    str住院号 = Get住院号(lng病人ID)
    
    '取余额
    gstrSQL = "Select 往年帐户余额,本年帐户余额 From 保险帐户 Where 险类=[2] And 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取余额", lng病人ID, TYPE_宁海)
    dbl当年帐户_余额 = Nvl(rsTemp!本年帐户余额, 0)
    dbl往年帐户_余额 = Nvl(rsTemp!往年帐户余额, 0)
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & Settle, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '进行住院结算
'    入口参数 (Data)
'    是否有医保卡 IC卡数据(可以传空) + 现金支付方式 + 空 + 住院号 + 本次结算明细条数 + 非医保项目总额 + 操作员姓名
'    返回参数(包括"交易结果~错误信息~交易结果信息")
'    交易结果~错误信息+写医保卡结果　+ 扣银行账户结果 + 写卡后IC卡数据 + 结算结果(参考住院预结算)
    str无卡病人 = IS无卡病人(lng病人ID)
    StrInput = Split(str无卡病人, "|")(0) & strField & IIf(Split(str无卡病人, "|")(0) = 1, IC_Data_宁海.IC卡数据, Split(str无卡病人, "|")(1)) & strField & "1" & strField & strField & _
    str住院号 & strField & IC_Data_宁海.mstr结算入口参数串 & strField & UserInfo.姓名 & strField & IC_Data_宁海.mstr业务类型
    Call Interface_Prepare_宁海(Settle, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "[" & mstrFunc & "]接口返回错误信息：" & vbCrLf & Split(mstrMsg, "~")(1)
        Exit Function
    End If
    blnTrans = True
    
    '分解结算信息
    '交易结果~错误信息+空~空~空 +
    '结算结果（费用总额%%自费总额%%自理总额%%统筹基金支付%%往年帐户支付%%当年帐户支付
    '%%大病支付%%公务员补助支付%%单位支付%%个人现金支付%%起付标准%%分段信息） + IC卡写卡前后信息(His不使用)
    strBalance = Split(mstrMsg, strField)(5)
    dbl费用总额 = Val(Split(strBalance, "%%")(int费用总额))
    dbl统筹基金 = Val(Split(strBalance, "%%")(int统筹基金))
    dbl帐户支付 = Val(Split(strBalance, "%%")(int当年帐户)) + Val(Split(strBalance, "%%")(int往年帐户))
    dbl大病补助 = Val(Split(strBalance, "%%")(int大病救助))
    dbl公务员补助 = Val(Split(strBalance, "%%")(int公务员补助))
    dbl往年帐户 = Val(Split(strBalance, "%%")(int往年帐户))
    dbl当年帐户 = Val(Split(strBalance, "%%")(int当年帐户))
    dbl现金支付 = dbl费用总额 - dbl统筹基金 - dbl帐户支付 - dbl大病补助 - dbl公务员补助 + IC_Data_宁海.mdbl非医保金额
    
    '写保险结算记录
    '大病自付=大病补助;超限自付=公务员补助
    strBalance = "'" & TruncZero(strBalance) & "'"
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_宁海 & "," & lng病人ID & "," & _
        Year(zlDatabase.Currentdate()) & ",0,0,0,0," & dbl往年帐户_余额 & "," & dbl当年帐户_余额 & "," & dbl往年帐户 & "," & dbl当年帐户 & "," & dbl费用总额 & "," & dbl现金支付 & ",0," & _
        dbl统筹基金 & "," & dbl统筹基金 & "," & dbl大病补助 & "," & dbl公务员补助 & "," & dbl帐户支付 & ",'" & str交易流水号 & "|" & IC_Data_宁海.待遇类别 & "|" & IC_Data_宁海.mstr业务类型 & "'," & lng主页ID & ",NULL," & strBalance & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "写保险结算记录")
    
    '交易确认
    '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
    '缺省现金支付
    StrInput = "~~~~" & Settle & strField & str交易流水号 & strField & "0" & strField & "HIS成功！"
    Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then
        Err.Raise 9000, gstrSysName, "警告：本次交易确认失败，请记录下本次交易流水号，并通知系统管理员使用工具包再次确认该交易" & _
            vbCrLf & "交易流水号：" & str交易流水号
    End If
    住院结算_宁海 = True
    
    Call 出院登记_宁海(lng病人ID, lng主页ID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then
        '交易确认
        '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
        '缺省现金支付
        StrInput = "~~~~" & Settle & strField & str交易流水号 & strField & "-1" & strField & "医保成功，而HIS失败！"
        Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    End If
End Function

Public Function 住院结算冲销_宁海(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    '撤销出院登记的同时，医保就自动完成了住院结算冲销，因此，本接口不做任何处理
    MsgBox "医保病人撤销出院时，该病人在医保中心的出院结算单同时作废，本次结算作废仅处理HIS端的费用！", vbInformation, gstrSysName
    If MsgBox("你确定要结算作废吗？（如不清楚，可以咨询系统管理员）" & vbCrLf & "正常处理流程：先办理撤销出院登记，再进行住院结算作废，然后再次结算，再次办理出院登记！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    住院结算冲销_宁海 = True
End Function

Public Function 住院信息变动_宁海(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '在院信息变动或疾病选择调用此接口
    Dim intReturn As Integer
    Dim StrInput As String, strOutput As String, strErrInfo As String
    
    Dim str交易流水号 As String, str住院号 As String
    Dim str床号 As String, str科室名称 As String, lng科室ID As Long
    Dim str出院诊断 As String, str出院日期 As String, str医生 As String, str疾病编号 As String, lng病种ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '空~空~空~空+ 住院号 + 变动时间(yyyy.mm.dd) + 病人床号+诊断医生姓名+诊断描述+疾病编号（疾病库中序号）+科室名称+科室ID(不确定时填0)"（中间用~分隔），仅传变动内容，未变动的可以为空串。例如：科室为"内科（201）"：变动信息为"~~~~内科~201"
    '返回参数(包括"交易结果~错误信息~交易结果信息")
    '交易结果~错误信息+空~空~空
    
    gstrSQL = "Select A.当前科室ID,B.名称,A.当前床号 From 病人信息 A,部门表 B" & _
        " Where A.病人ID=[1] And A.当前科室ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取科室名称及床号", lng病人ID)
    lng科室ID = Nvl(rsTemp!当前科室ID, 0)
    str科室名称 = Nvl(rsTemp!名称)
    str床号 = Nvl(rsTemp!当前床号)
    str住院号 = Get住院号(lng病人ID)
    str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, True)
    '没有出院诊断时，以入院诊断为准
    If Trim(str出院诊断) = "" Then str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, True, True)
    gstrSQL = "Select 住院医师,出院日期 From 病案主页 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生姓名及出院日期", lng病人ID, lng主页ID)
    If Nvl(rsTemp!出院日期) <> "" Then
        str出院日期 = Format(rsTemp!出院日期, "yyyy.MM.dd")
    End If
    str医生 = Nvl(rsTemp!住院医师)
    gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病种ID", lng病人ID)
    lng病种ID = rsTemp!病种ID
'    If lng病种ID <> 0 Then
'        '提取前置机中的疾病信息
'        gstrSQL = "Select JBBZDM From SIM_JBDA Where JBBM=" & lng病种ID
'        If rsTemp.State = 1 Then rsTemp.Close
'        rsTemp.Open gstrSQL, gcn宁海
'        If rsTemp.RecordCount <> 0 Then str疾病编号 = rsTemp!JBBZDM
'    End If
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & ModifyPatient, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    '调用住院信息变动
    StrInput = "~~~~" & str住院号 & strField & Format(zlDatabase.Currentdate(), "yyyy.MM.dd") & strField & _
        str床号 & strField & str医生 & strField & str出院诊断 & strField & lng病种ID & strField & str科室名称 & strField & lng科室ID
    Call Interface_Prepare_宁海(ModifyPatient, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    
    住院信息变动_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_宁海(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim intReturn As Integer
    Dim lng病种ID As Long
    Dim str交易流水号 As String
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str医生 As String, str住院号 As String, str出院诊断 As String, str疾病编号 As String, str出院日期 As String
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '入口参数 (Data)
    '空~空~空~空 + 住院号 + 出院诊断医生+出院诊断说明+出院确诊疾病编号+出院日期(yyyy.mm.dd)
    '返回参数(包括"交易结果~错误信息~交易结果信息")
    '交易结果~错误信息+空~空~空
    
    '必须先结算，后出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        '取流水号
        Call Interface_Prepare_宁海(GetSequence, "~~~~" & 出院病历, "")
        intReturn = Interface_Exec_宁海
        If intReturn <> 0 Then
            MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
            If (intReturn < 0) Then Exit Function
        End If
        str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
        
        '准备调用出院病历填写接口
        str住院号 = Get住院号(lng病人ID)
        str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, True)
        gstrSQL = "Select 住院医师,出院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医生姓名及出院日期", lng病人ID, lng主页ID)
        str出院日期 = Format(rsTemp!出院日期, "yyyy.MM.dd")
        str医生 = Nvl(rsTemp!住院医师)
        gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病种ID", lng病人ID)
        lng病种ID = rsTemp!病种ID
         If lng病种ID <> 0 Then
            '提取前置机中的疾病信息
            gstrSQL = "Select JBMC From SIM_JBDA Where JBBM=" & lng病种ID
            If rsTemp.State = 1 Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn宁海
             If rsTemp.RecordCount <> 0 Then str疾病编号 = rsTemp!JBMC '疾病名称
        End If
        
        StrInput = "~~~~" & str住院号 & strField & str疾病编号 & strField & lng病种ID & strField & str出院日期
        Call Interface_Prepare_宁海(出院病历, StrInput, strOutput, str交易流水号)
        If intReturn <> 0 Then
            MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
            If (intReturn < 0) Then Exit Function
        End If
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_宁海 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    出院登记_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_宁海(lng病人ID As Long, lng主页ID As Long) As Boolean
    '调医保的出院结算撤销功能实现，同时撤销出院，医保恢复到在院状态
    Dim intReturn As Integer
    Dim blnTrans As Boolean
    Dim StrInput As String, strOutput As String, strErrInfo As String
    Dim str无卡病人 As String
    Dim str交易流水号 As String, str结算流水号 As String, str住院号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '入口参数 (Data)
    '是否有医保卡 IC信息(可以传空) + 空 + 空 + 要作废的住院结算交易号 + 要作废的住院号
    '说明：只有"要作废的住院结算交易号"对应的结算单据中住院号是"要作废的住院号"时才可以退费，避免错误退费。
    '返回参数(包括"交易结果~错误信息~交易结果信息")
    '交易结果~错误信息+写医保卡结果　+ 空 + 写卡后IC卡数据+是否为重复退费（0正常退费，1结算单在医保接口已经被退费过）+ 结算结果（参见住院预结算）
    '各值为负表示退费，若重复退费则置参数中重复退费为1，且返回成功。
    
    '取流水号
    Call Interface_Prepare_宁海(GetSequence, "~~~~" & SettleDel, "")
    intReturn = Interface_Exec_宁海
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    str交易流水号 = Trim(TruncZero(Split(mstrMsg, "~")(5)))
    
    str住院号 = Get住院号(lng病人ID)
    '提取结算流水号
    gstrSQL = "Select 支付顺序号 From 保险结算记录" & _
        " Where 记录ID=(Select Max(记录ID) From 保险结算记录 Where 性质=2 And 主页ID=[2] And 病人ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取结算流水号", lng病人ID, lng主页ID)
    str结算流水号 = Split(rsTemp!支付顺序号, "|")(0)
    
    str无卡病人 = IS无卡病人(lng病人ID)
'    入口参数 (Data)
'    是否有医保卡 IC信息(可以传空) + 空 + 空 + 要作废的住院结算交易号 + 要作废的住院号
    StrInput = Split(str无卡病人, "|")(0) & strField & Split(str无卡病人, "|")(1) & strField & strField & strField & str结算流水号 & strField & str住院号
    Call Interface_Prepare_宁海(SettleDel, StrInput, strOutput, str交易流水号)
    intReturn = Interface_Exec_宁海()
    If intReturn <> 0 Then
        MsgBox "[" & mstrFunc & "]接口返回" & IIf(intReturn > 0, "警告", "错误") & "信息：" & vbCrLf & Split(mstrMsg, "~")(1), vbInformation, gstrSysName
        If (intReturn < 0) Then Exit Function
    End If
    blnTrans = True
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_宁海 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "入院登记")
    
    '交易确认
    '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
    '缺省现金支付
    StrInput = "~~~~" & SettleDel & strField & str交易流水号 & strField & "0" & strField & "HIS成功！"
    Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    intReturn = Interface_Exec_宁海
    If intReturn < 0 Then MsgBox "警告：本次交易确认失败，请记录下本次交易流水号，并通知系统管理员使用工具包再次确认该交易" & _
        vbCrLf & "交易流水号：" & str交易流水号, vbInformation, gstrSysName
    
    出院登记撤销_宁海 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        '交易确认
        '空~空~空~空~交易类型~医保交易流水号~HIS事务结果~附加信息
        '缺省现金支付
        StrInput = "~~~~" & SettleDel & strField & str交易流水号 & strField & "-1" & strField & "医保成功，而HIS失败！"
        Call Interface_Prepare_宁海(Decide, StrInput, strOutput)
    End If
End Function

Public Function 个人余额_宁海(ByVal lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Nvl(帐户余额,0) From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取帐户余额", lng病人ID, TYPE_宁海)
    个人余额_宁海 = rsTemp.Fields(0).Value
End Function

Public Function IS无卡病人(ByVal lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '是否无卡病人，无卡返回0，有卡返回1，返回格式：卡标志|医保号
    gstrSQL = "Select Nvl(灰度级,0) AS 卡标志,医保号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否无卡病人", lng病人ID, TYPE_宁海)
    If rsTemp!卡标志 = 0 Then
        '有卡病人，医保号返回空
        IS无卡病人 = "1|"
    Else
        IS无卡病人 = "0|" & rsTemp!医保号
    End If
End Function


