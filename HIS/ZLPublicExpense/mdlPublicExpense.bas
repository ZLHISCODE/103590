Attribute VB_Name = "mdlPublicExpense"
Option Explicit
Public gstrSysName As String                '系统名称
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String    '产品名称
Public gstrSQL As String
Public glngSys As Long
Public glngMainModule As Long '调用者的模块号
Public gstrMainPrivs As String '调用者的相关权限
Public gblnOK As Boolean
Public gclsInsure As New clsInsure '医保部件
Public gstrDBUser As String '所有者
Public gcnOracle As ADODB.Connection
Public gcolPrivs As Collection              '记录内部模块的权限
Public gobjSquare As Object '卡结算部件
Public gobjPlugIn As Object '外挂功能

'挂号用参数
Public gstrRooms As String
Public glngModul As Long
Public gbytState As Byte
Public gstrDocs As String
Public gstrDeptIDs As String
Public gstrPrivs As String
Public gblnBill挂号 As Boolean
Public gbytRegistMode As Byte
Public gdatRegistTime As Date

Public grs医疗付款方式 As ADODB.Recordset
Public grsOneCard As ADODB.Recordset

Private Type TY_Decimal_Precision '小数精度
    byt_Bit As Byte '小数位数:表示核算到小数点后第多少位。
    strFormt_VB As String   'VB格式化:0.0000;...
    strFormt_ORA As String  'Oracle格式化:999990.00000...
End Type

Private Type ty_SysPara
    bln报警包含划价费用  As Boolean
    byt票据分配规则 As Byte   '票据分配规则:0-根据实际打印分配票号;1-根据系统预定规则分配;2-根据用户自定义规则分配
    Money_Decimal As TY_Decimal_Precision  '费用金额小数格式
    Price_Decimal As TY_Decimal_Precision  '费用单价小数格式
    bln从项汇总折扣  As Boolean
    byt药品名称显示 As Byte '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    byt输入药品显示 As Byte '0-按输入匹配显示，1-固定显示通用名和商品名

    byt病人审核方式 As Byte '病人审核方式:0-未审核不允许结帐，缺省为0;1-审核时不许调整费用和医嘱（包含医嘱调整和费用调整）
    bln未入科禁止记账  As Boolean
    bln卫材执行发料 As Boolean '执行之后卫材自动发料
    bln执行后审核 As Boolean
    bln执行前先结算 As Boolean '一卡通执行前先收费或记帐审核
    bln开单后立即结算 As Boolean '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
    int医保对码 As Integer '是否对住院医保病人的项目对码情况进行检查:0-不检查,1-检查并提醒,2-检查并禁止
    dblMaxMoney As Double   '最大金额检查
    bytBillOpt As Byte '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
    dbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
    bln简码匹配方式切换 As Boolean '允许在窗口界面的工具栏切换简码匹配方式切换

    int住院自动发料 As Integer   '住院记帐完成后是否自动发料:0-不自动发料，1-自动发料，2-本科室开单时自动发料
    bln门诊自动发料 As Boolean '门诊记帐完成后是否自动发料
    bln收费后自动发药 As Boolean '
    bln分离发药 As Boolean
    str医保费用类型 As String '医保病人允许的费用类型
    str公费费用类型 As String '公费病人允许的费用类型
    strLike As String
    bytCode As Byte
    bln收费类别 As Boolean '是否首先输入类别
    blnFeeKindCode As Boolean '不输类别时,首位当作收费类别简码
    strMatchMode As String '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    blnStock As Boolean '指定药房时是否限定输入药品的库存
    bln门诊留观记帐  As Boolean
    bln住院留观记帐 As Boolean
    bln免挂号模式 As Boolean '是否免挂模式,流程：直接在分诊台取号，然后在接诊时，产生划价单
    byt条码卫材识别控制 As Byte   '是否仅条码识别::1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
    bln不显示无库存卫材 As Boolean
End Type

Public gSysPara As ty_SysPara
Public Enum gEm_BulidIng_SQLType
    EM_Bulid_字符 = 0
    EM_Bulid_数字 = 1
End Enum
Public Const gstrCompentsName = ""
Public Enum Enum_Inside_Program
    p住院记帐操作 = 1150
    p医嘱附费管理 = 1257
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    
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
End Type
Public UserInfo As TYPE_USER_INFO

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

'----------------------------------------------------
'公共对象定义
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gstrNodeNo As String '站点名
Public glngInstanceCount As Long '进程总数
Public glngMax家庭地址 As Long       '家庭地址最大允许录入长度
Public glngMax户口地址 As Long       '户口地址最大允许录入长度
Public glngMax出生地点 As Long       '出生地点最大允许录入长度
Public glngMax联系人地址 As Long    '联系人地址最大允许录入长度

Public Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的权局变量
    '编制:刘兴洪
    '日期:2014-03-20 16:07:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, varTmp As Variant
    Dim strValue As String
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = gobjComlib.GetUnitName
    gbytRegistMode = Val(Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(0))
    If Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    With gSysPara
        .bln报警包含划价费用 = gobjDatabase.GetPara(98, glngSys) = "1"
        With .Money_Decimal '费用金额小数位数
            .byt_Bit = Val(gobjDatabase.GetPara(9, glngSys, , 2))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        With .Price_Decimal  '费用单价小数位数
            .byt_Bit = Val(gobjDatabase.GetPara(157, glngSys, , 5))
            .strFormt_VB = "0." & String(.byt_Bit, "0")
            .strFormt_ORA = "FM" & String(14, "9") & "0." & String(.byt_Bit, "9")
        End With
        '启用标志||NO;执行科室(条数);收据费目(首页汇总,条数);收费细目(条数)
        strTmp = Trim(gobjDatabase.GetPara("票据分配规则", glngSys, 1121, "0||0;0;0;0;0"))
        varTmp = Split(strTmp & "||", "||")
        .byt票据分配规则 = Val(varTmp(0))
        .bln从项汇总折扣 = Val(gobjDatabase.GetPara(93, glngSys)) <> 0
        .byt药品名称显示 = Val(gobjDatabase.GetPara("药品名称显示", , , "2"))
        .byt输入药品显示 = gobjDatabase.GetPara("输入药品显示", , , 0)
        .byt病人审核方式 = Val(gobjDatabase.GetPara(185, glngSys))    '49501
        .bln未入科禁止记账 = Val(gobjDatabase.GetPara(215, glngSys)) = 1 '51612
        .bln门诊留观记帐 = gobjDatabase.GetPara("门诊留观病人记帐", glngSys, 1150) = "1"
        .bln住院留观记帐 = gobjDatabase.GetPara("住院留观病人记帐", glngSys, 1150) = "1"
        
        '33:参数原BUG号为14403（玉溪医院），临床或医技执行登记时，
        ' 该检查、检验、处置项目已经完成，卫生材料已经使用了，
        ' 所以，应该都是先领用到科室的，本科的就不必再走发料的流程，
        ' 所以，本参数没有必要存在，应该都处理为执行后自动对跟踪在用的卫材发料，
        ' 取消执行时自动退料
        
        .bln卫材执行发料 = True ' Val(gobjDatabase.GetPara(33, glngSys)) <> 0
        ' 81参数:该参数至少在10.03以前就存在，未找到BUG号。审核划价单的目的是确认费用，执行之后，如果不确认费用，就还需要人工单独去审核划价单。从业务特性来说，本参数没有必要存在，应该都处理为执行后自动审核划价单。程序相关控制按勾上此参数进行处理
        .bln执行后审核 = True  ' Val(gobjDatabase.GetPara(81, glngSys)) <> 0
        '门诊一卡通,项目执行前必须先收费或先记帐审核
        .bln执行前先结算 = Val(gobjDatabase.GetPara(163, glngSys)) <> 0
        '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
        .bln开单后立即结算 = Val(gobjDatabase.GetPara(232, glngSys)) <> 0
        '医保对码检查
        .int医保对码 = Val(gobjDatabase.GetPara(59, glngSys, , 1))
        '单笔费用最大提醒金额
        .dblMaxMoney = Val(gobjDatabase.GetPara(60, glngSys))
    
        '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
        .bytBillOpt = Val(gobjDatabase.GetPara(23, glngSys))
        '一卡通消费验证
        strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
        If InStr(strValue, "|") = 0 Then strValue = "1|0"
        .dbl预存款消费验卡 = Val(Split(strValue, "|")(0))
        .bln简码匹配方式切换 = Val(gobjDatabase.GetPara("简码匹配方式切换", , , "1")) = 1
        '门诊自动发料
        .bln门诊自动发料 = Val(gobjDatabase.GetPara(92, glngSys)) <> 0
        '住院自动发料
        .int住院自动发料 = Val(gobjDatabase.GetPara(63, glngSys))
        '自动发药退药
        .bln收费后自动发药 = gobjDatabase.GetPara(45, glngSys) = "1"
        '门诊收费与发药分离
        .bln分离发药 = gobjDatabase.GetPara(15, glngSys) = "1"
        '医保费用类型
        .str医保费用类型 = "'" & Replace(gobjDatabase.GetPara(41, glngSys), "|", "','") & "'"
    
        '公费费用类型
        .str公费费用类型 = "'" & Replace(gobjDatabase.GetPara(42, glngSys), "|", "','") & "'"
            
        '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
        .strMatchMode = gobjDatabase.GetPara(44, glngSys, , "00")
        
        .strLike = IIf(gobjDatabase.GetPara("输入匹配") = "0", "%", "")
        .bytCode = Val(gobjDatabase.GetPara("简码方式"))
        '是否要求首先输入类别
        .bln收费类别 = Val(gobjDatabase.GetPara(72, glngSys, , 1)) <> 0
        '当不输类别时,输入费用项目时,首位当作类别简码
        .blnFeeKindCode = Val(gobjDatabase.GetPara(144, glngSys)) <> 0 And Not .bln收费类别
        '指定药房时限制库存
        .blnStock = Val(gobjDatabase.GetPara(18, glngSys)) <> 0
        .bln免挂号模式 = Val(gobjDatabase.GetPara("免挂号模式", glngSys)) = 1
        .byt条码卫材识别控制 = Val(gobjDatabase.GetPara(320, glngSys, , "0"))      '1-必须通过扫码录入或录入条码;0-不控制，可以通过简码等查找
        .bln不显示无库存卫材 = Val(gobjDatabase.GetPara(316, glngSys)) = 1
    End With
    Call InitAddressLength
End Sub

Public Function zlGetFeeFields(Optional strTableName As String = "门诊费用记录", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定表的值
    '入参：strTableName:如:门诊费用记录;住院费用记录;....
    '      blnReadDatabase-从数据库中读取
    '出参：
    '返回：字段集
    '编制：刘兴洪
    '日期：2010-03-10 10:41:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strFileds As String
    
    Err = 0: On Error GoTo Errhand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "门诊费用记录"
        zlGetFeeFields = "" & _
        "Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, " & _
        "姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, " & _
        "加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, " & _
        "发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, " & _
        "保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
        Exit Function
    Case "住院费用记录"
        zlGetFeeFields = "" & _
         " Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, " & _
         " 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, " & _
         " 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, " & _
         " 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, " & _
         " 结帐id , 结帐金额, 保险大类ID, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
         Exit Function
    Case "病人结帐记录"
        zlGetFeeFields = "Id, No, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期, 备注"
        Exit Function
    Case "病人预交记录"
        zlGetFeeFields = "" & _
        " Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, " & _
        " 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补,预交类别,卡类别ID,结算卡序号,卡号,交易流水号,交易说明,合作单位,结算序号,校对标志"
        Exit Function
    Case "人员表"
        zlGetFeeFields = "" & _
        "Id, 编号, 姓名, 简码, 身份证号, 出生日期, 性别, 民族, 工作日期, 办公室电话, 电子邮件, 执业类别, 执业范围, " & _
        "管理职务, 专业技术职务, 聘任技术职务, 学历, 所学专业, 留学时间, 留学渠道, 接受培训, 科研课题, 个人简介, 建档时间, " & _
        "撤档时间, 撤档原因, 别名, 站点"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo Errhand:
    strSql = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取列信息", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!Column_Name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
Errhand:
    zlGetFeeFields = "*"
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = gobjComlib.Nvl(varValue, DefaultValue)
End Function

Public Function GetPatiMoney(ByVal bytType As Byte, ByVal lng病人ID As Long, ByRef objPatiFee As clsPatiFeeinfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人的相关费用信息
    '入参:bytType-0-门诊;1-住院
    '     lng病人ID-病人ID
     '出参:
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 16:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Set objPatiFee = New clsPatiFeeinfor
    On Error GoTo errHandle
    If bytType = 0 Then
        strSql = "" & _
        "   Select Nvl(预交余额,0) 预交余额,Nvl(费用余额,0) 费用余额,0 as 预结费用,0 as 担保额 " & _
        "   From 病人余额 " & _
        "   Where 性质=1 And 类型=1 And 病人ID=[1]" & _
        "   "
    Else
        strSql = "" & _
        "   Select Nvl(预交余额,0) 预交余额,Nvl(费用余额,0) 费用余额,0 as 预结费用 ,0 as 担保额" & _
        "   From 病人余额 " & _
        "   Where 性质=1 And 类型=2 And 病人ID=[1]" & _
        "   Union ALL " & _
        "   Select 0 as 预交余额,0 as 费用余额,Sum(B.金额) as 预结费用 ,0 as 担保额" & _
        "   From 病人信息 A,保险模拟结算 B" & _
        "   Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    End If
    strSql = strSql & "" & _
    "   Union ALL " & _
    "   Select 0 as 预交余额,0 as 费用余额,0 as 预结费用,担保额" & _
    "   From 病人信息 B " & _
    "   Where 病人ID=[1]"
    
    strSql = "" & _
    "   Select Nvl(Sum(预交余额),0) as 预交余额,Nvl(Sum(费用余额),0) as 费用余额,Nvl(Sum(预结费用),0) as 预结费用,Nvl(Sum(担保额),0) as 担保额  " & _
    "   From (" & strSql & ")"
    
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取病人的相关费用金额", lng病人ID)
    If rsTemp.EOF Then GetPatiMoney = True: Exit Function
    With objPatiFee
        .预交余额 = FormatEx(Val(Nvl(rsTemp!预交余额)), 6)
        .未结费用 = FormatEx(Val(Nvl(rsTemp!费用余额)), 6)
        .预结费用 = FormatEx(Val(Nvl(rsTemp!预结费用)), 6)
        .担保额 = FormatEx(Val(Nvl(rsTemp!担保额)), 6)
        .剩余款 = FormatEx(.预交余额 + .预结费用 - .未结费用, 6)
    End With
    GetPatiMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function FromIDsBulidIngSQL(ByVal bytBulidType As gEm_BulidIng_SQLType, _
    ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strAliaName As String, Optional intStartPara As Integer = 1 _
    ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据IDs来获取相关的SQL,如:select ... From str2List Union ALL Selelct ..
    '入参:strValues-值,多个用逗号分离
    '     strAliaName-别名
    '     bytType-0-字符型;1-数字型;
    '     intStartPara-启动的参数
    '出参:varPara-返回的参数值数据组
    '     strBulitSQL-返回的构建的SQL串
    '返回:如果获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-25 17:04:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSql As String
    Dim strPara() As Variant, strTable As String, strColumnName As String
    
    On Error GoTo errHandle
    
    strColumnName = " Column_Value "
    If strAliaName <> "" Then strColumnName = strColumnName & " As " & strAliaName
    
    If bytBulidType = EM_Bulid_字符 Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve strPara(0 To j - 1) As Variant
    
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If gobjCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSql = strSql & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
            ReDim Preserve strPara(0 To j - 1) As Variant
            strPara(j - 1) = Mid(strTemp, 2)
            j = j + 1
            strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSql = strSql & " Union ALL  Select " & strColumnName & " From " & Replace(strTable, "[0]", "[" & j & "]")
        ReDim Preserve strPara(0 To j - 1) As Variant
        strPara(j - 1) = Mid(strTemp, 2)
    End If
    
    varPara = strPara
    If strSql <> "" Then strSql = Mid(strSql, 11)
    strBulitSQL = strSql
    FromIDsBulidIngSQL = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetFeeMoneyFromAdviceIDs(ByVal str医嘱IDs As String, _
    ByRef dblOut应收金额 As Double, ByRef dblOut实收金额 As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱IDs来获取应收和实收金额
    '入参:str医嘱IDs-医嘱ID,多个用逗号分离
    '出参:dblOut应收金额-应收金额
    '     dblOut实收金额-实收金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-25 16:11:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    dblOut应收金额 = 0: dblOut实收金额 = 0
    If str医嘱IDs = "" Then Exit Function
    
    '不能大于4000
    If gobjCommFun.ActualLen(str医嘱IDs) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_数字, str医嘱IDs, varPara, strSql, "医嘱ID") = False Then Exit Function
        strSql = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(应收金额), 0) As 应收金额, Nvl(Sum(实收金额), 0) As 实收金额 " & _
        " From (With 医嘱数据 As (" & strSql & ") " & _
        "        Select Nvl(Sum(a.应收金额), 0) As 应收金额, Nvl(Sum(a.实收金额), 0) As 实收金额 " & _
        "        From 门诊费用记录 A, 医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.应收金额), 0) As 应收金额, Nvl(Sum(a.实收金额), 0) As 实收金额 " & _
        "        From 住院费用记录 A, 医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", varPara)
    Else
        strSql = "" & _
        " Select /*+ RULE */ " & _
        "   Nvl(Sum(应收金额), 0) As 应收金额, Nvl(Sum(实收金额), 0) As 实收金额 " & _
        " From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2list([1]))) " & _
        "        Select Nvl(Sum(a.应收金额), 0) As 应收金额, Nvl(Sum(a.实收金额), 0) As 实收金额 " & _
        "        From 门诊费用记录 A, 医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id " & _
        "        Union All " & _
        "        Select Nvl(Sum(a.应收金额), 0) As 应收金额, Nvl(Sum(a.实收金额), 0) As 实收金额 " & _
        "        From 住院费用记录 A, 医嘱数据 B " & _
        "        Where a.医嘱序号 = b.医嘱id)"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", str医嘱IDs)
    End If
    
    On Error GoTo errHandle
    dblOut应收金额 = FormatEx(Val(Nvl(rsTemp!应收金额)), 6)
    dblOut实收金额 = FormatEx(Val(Nvl(rsTemp!实收金额)), 6)
    
    GetFeeMoneyFromAdviceIDs = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         Set gobjSquare.objSquareCard = Nothing
     End If
     Set gobjSquare = Nothing
     If Err <> 0 Then Err.Clear: Err = 0
End Sub

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


Public Function AdviceIsCharged(ByVal str医嘱IDs As String, _
    ByVal strNos As String, ByRef bytOutChargeStatus As Byte, Optional ByRef strOut未收医嘱IDs As String, _
    Optional ByRef bytOutBillType As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医嘱是否已经收费
    '入参:str医嘱IDs-医嘱ID,多个用逗号分离
    '出参:bytOutChargeStatus-收费状态(0-未收费,1-完全收费;2-部门收费)
    '     strOut未收医嘱IDs-返回未收费或未补审核的医嘱ID
    '     bytOutBillType:返回当前的单据类型(0-不存在任何单据;1-收费单;2-记帐单;3-收费和记帐都有)
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 09:48:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim varPara As Variant
    Dim bytStatus As Byte
    strOut未收医嘱IDs = "": bytOutBillType = 0: bytOutChargeStatus = 0
    If strNos = "" And str医嘱IDs = "" Then Exit Function
    
    If str医嘱IDs <> "" Then
        '不能大于4000
        If gobjCommFun.ActualLen(str医嘱IDs) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_数字, str医嘱IDs, varPara, strSql, "医嘱ID") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */ distinct  记录性质, 记录状态,医嘱序号" & _
            " From (With 医嘱数据 As (" & strSql & ") " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 门诊费用记录 A,医嘱数据 B " & _
            "        Where a.医嘱序号 = b.医嘱id And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 住院费用记录 A, 医嘱数据 B " & _
            "        Where a.医嘱序号 = b.医嘱id And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", varPara)
        Else
            strSql = "" & _
            " Select /*+ RULE */ distinct  记录性质, 记录状态,医嘱序号" & _
            " From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Num2list([1]))) " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 门诊费用记录 A,医嘱数据 B " & _
            "        Where a.医嘱序号 = b.医嘱id And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 住院费用记录 A, 医嘱数据 B " & _
            "        Where a.医嘱序号 = b.医嘱id And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", str医嘱IDs)
        End If
    Else
        '按单据号处理
        '不能大于4000
        If gobjCommFun.ActualLen(strNos) > 4000 Then
            If FromIDsBulidIngSQL(EM_Bulid_字符, strNos, varPara, strSql, "NO") = False Then Exit Function
            strSql = "" & _
            " Select /*+ RULE */ distinct  记录性质, 记录状态,医嘱序号" & _
            " From (With 医嘱数据 As (" & strSql & ") " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 门诊费用记录 A,医嘱数据 B " & _
            "        Where a.NO = b.NO And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 住院费用记录 A, 医嘱数据 B " & _
            "        Where a.NO = b.NO And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", varPara)
        Else
            strSql = "" & _
            " Select /*+ RULE */ distinct  记录性质, 记录状态,医嘱序号" & _
            " From (With 医嘱数据 As (Select Column_Value As 医嘱id From Table(f_Str2list([1]))) " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 门诊费用记录 A,医嘱数据 B " & _
            "        Where a.NO = b.NO And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) " & _
            "        Union All " & _
            "        Select distinct a.记录性质,A.记录状态,A.医嘱序号 " & _
            "        From 住院费用记录 A, 医嘱数据 B " & _
            "        Where a.NO = b.NO And A.记录性质 in (1,2,3) And A.记录状态 IN (0,1,3) )"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", str医嘱IDs)
        End If
        
    End If
    On Error GoTo errHandle
    With rsTemp
        bytStatus = -1
        Do While Not .EOF
             If Val(Nvl(!记录状态)) = 0 Then  '未收费
                If Val(Nvl(!医嘱序号)) <> 0 Then
                    strOut未收医嘱IDs = strOut未收医嘱IDs & "," & Nvl(rsTemp!医嘱序号)
                End If
             End If
             If bytStatus = -1 Then
                If Val(Nvl(!记录状态)) = 0 Then
                    bytStatus = IIf(Val(Nvl(!记录状态)) = 0, 0, 1)
                End If
             ElseIf bytStatus = 0 And (Val(Nvl(!记录状态)) = 1 Or Val(Nvl(!记录状态)) = 3) Then
                bytStatus = 2   '部分收费
             ElseIf bytStatus = 1 And Val(Nvl(!记录状态)) = 0 Then
                bytStatus = 2 '部分收费
             End If
             
             If bytOutBillType = 0 Then
                bytOutBillType = Val(Nvl(!记录性质))
             ElseIf bytOutBillType <> Val(Nvl(!记录性质)) Then
                '两都都有
                bytOutBillType = 3
             End If
            .MoveNext
        Loop
    End With
    bytOutChargeStatus = bytStatus
    AdviceIsCharged = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function BillExistNotBalance(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断收费单据是否存在未收费的
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '出参:
    '返回:单据中存在未收费的,返回true,否则返回False
    '编制:冉俊明
    '日期:2016-08-25 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    '不能大于4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_字符, strNos, varPara, strSql, "NO") = False Then Exit Function
        strSql = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From 门诊费用记录 A,(" & strSql & ") B" & vbNewLine & _
                " Where Mod(a.记录性质, 10) = 1 And a.NO = b.NO And a.记录状态 = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否存在未收费的", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = "Select /*+cardinality(b,10)*/ 1" & vbNewLine & _
                " From 门诊费用记录 A,(Select Column_Value As NO From Table(f_str2list([1]))) B" & vbNewLine & _
                " Where Mod(a.记录性质, 10) = 1 And a.NO = b.NO And a.记录状态 = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否存在未收费的", strNos)
    Else
        strSql = "Select 1" & vbNewLine & _
                " From 门诊费用记录" & vbNewLine & _
                " Where Mod(记录性质, 10) = 1 And NO = [1] And 记录状态 = 0 And Rownum <2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否存在未收费的", strNos)
    End If
    
    If rsTemp.EOF Then
        BillExistNotBalance = False '已全部收费
    Else
        BillExistNotBalance = True '存在未收费
    End If
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBillChargeStatus(ByVal strNos As String, ByRef bytOutStatus As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费单据的计费状态
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '出参:bytOutStatus:0-未收费;1-部分收费/退费;2-全部收费;3-全部退费
    '返回:获取成功,返回true,否则返回False(含未找到数据部分)
    '编制:刘兴洪
    '日期:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '不能大于4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_字符, strNos, varPara, strSql, "NO") = False Then Exit Function
        strSql = "Select /*+cardinality(b,10)*/ Sum(a.数次 * Nvl(a.付数, 1)) As 剩余数量," & vbNewLine & _
                "        Sum(Decode(a.记录性质, 1, 1, 0) * Decode(a.记录状态, 2, 0, 1) * a.数次 * Nvl(a.付数, 1)) As 原始数量," & vbNewLine & _
                "        Sum(Decode(a.记录状态, 0, 1, 0) * a.数次 * Nvl(a.付数, 1)) As 未收数量" & vbNewLine & _
                " From 门诊费用记录 A,(" & strSql & ") B " & _
                " Where Mod(a.记录性质, 10) = 1 And a.价格父号 Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否已经收费", varPara)
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = "Select /*+cardinality(b,10)*/ Sum(a.数次 * Nvl(a.付数, 1)) As 剩余数量," & vbNewLine & _
                "        Sum(Decode(a.记录性质, 1, 1, 0) * Decode(a.记录状态, 2, 0, 1) * a.数次 * Nvl(a.付数, 1)) As 原始数量," & vbNewLine & _
                "        Sum(Decode(a.记录状态, 0, 1, 0) * a.数次 * Nvl(a.付数, 1)) As 未收数量" & vbNewLine & _
                " From 门诊费用记录 A,(Select Column_Value As NO From Table(f_str2list([1]))) B " & _
                " Where Mod(a.记录性质, 10) = 1 And a.价格父号 Is Null And a.NO = b.NO"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否已经收费", strNos)
    Else
        strSql = "Select Sum(数次 * Nvl(付数, 1)) As 剩余数量," & vbNewLine & _
                "        Sum(Decode(记录性质, 1, 1, 0) * Decode(记录状态, 2, 0, 1) * 数次 * Nvl(付数, 1)) As 原始数量," & vbNewLine & _
                "        Sum(Decode(记录状态, 0, 1, 0) * 数次 * Nvl(付数, 1)) As 未收数量" & vbNewLine & _
                " From 门诊费用记录" & vbNewLine & _
                " Where Mod(记录性质, 10) = 1 And 价格父号 Is Null And NO = [1]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否已经收费", strNos)
    End If
    
    If Val(Nvl(rsTemp!原始数量)) = 0 Then Exit Function
    If Val(Nvl(rsTemp!原始数量)) = Val(Nvl(rsTemp!未收数量)) Then
        bytOutStatus = 0 '未收费
    ElseIf Val(Nvl(rsTemp!原始数量)) = Val(Nvl(rsTemp!剩余数量)) And Val(Nvl(rsTemp!未收数量)) = 0 Then
        bytOutStatus = 2 '全部收费
    ElseIf Val(Nvl(rsTemp!剩余数量)) = 0 Then
        bytOutStatus = 3 '全部退费
    Else
        bytOutStatus = 1 '部分收费/退费
    End If
    GetBillChargeStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceStatus(ByVal strNos As String, ByRef bytOutStatus As Byte, _
    Optional bln门诊 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断记帐单是否已经结帐(只针对帐单)
    '入参:strNOs:指定的单据号,允许多个,用逗号分离
    '     bln门诊-门诊记帐单
    '出参:bytOutStatus:0-未结帐;1-部分结帐;2-全部结帐
    '返回:获取成功,返回true,否则返回False(含未找到数据部分)
    '编制:刘兴洪
    '日期:2014-03-26 11:38:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPara As Variant, strSql As String, rsTemp As ADODB.Recordset
    Dim strTable As String
    
    bytOutStatus = 0
    On Error GoTo errHandle
    strTable = IIf(bln门诊, "门诊费用记录", "住院费用记录")
    '不能大于4000
    If gobjCommFun.ActualLen(strNos) > 4000 Then
        If FromIDsBulidIngSQL(EM_Bulid_字符, strNos, varPara, strSql, "NO") = False Then Exit Function
        
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When 未结金额 <> 0 And 结帐金额 <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As 结帐标志" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额, Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额, Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.结帐Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From 门诊费用记录 A, (" & strSql & ") B" & vbNewLine & _
        "       Where a.No = b.No And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.价格父号, a.序号))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号来判断是否已经收费", varPara)
        
    ElseIf InStr(1, strNos, ",") > 0 Then
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When 未结金额 <> 0 And 结帐金额 <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As 结帐标志" & vbNewLine & _
        "From (Select /*+Cardinality(B,10)*/" & vbNewLine & _
        "        a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额, Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额, Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.结帐Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From 门诊费用记录 A, Table(f_Str2list([1])) B" & vbNewLine & _
        "       Where a.No = b.Column_Value And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.价格父号, a.序号))"
        
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据医嘱ID获取相关的费用金额", strNos)
    Else
        strSql = " " & _
        "Select Decode(Nvl(Sum(Nvl((Case" & vbNewLine & _
        "                    When (未结金额 <> 0 And 结帐金额 = 0) Or (未结金额 = 0 And (实收金额 = 0 Or 结帐金额 = 0) And n_Count = 0) Then" & vbNewLine & _
        "                     0" & vbNewLine & _
        "                    When 未结金额 <> 0 And 结帐金额 <> 0 Then" & vbNewLine & _
        "                     1" & vbNewLine & _
        "                    Else" & vbNewLine & _
        "                     2" & vbNewLine & _
        "                  End),0)),0), 0, 0, 2 * Count(1), 2, 1) As 结帐标志" & vbNewLine & _
        "From (Select " & vbNewLine & _
        "        a.No, Nvl(a.价格父号, a.序号) As 序号, Nvl(Sum(Nvl(a.应收金额, 0)), 0) As 应收金额, Nvl(Sum(Nvl(a.实收金额, 0)), 0) As 实收金额," & vbNewLine & _
        "        Nvl(Sum(Nvl(a.结帐金额, 0)), 0) As 结帐金额, Nvl(Sum(Nvl(a.实收金额, 0)) - Sum(Nvl(a.结帐金额, 0)), 0) As 未结金额," & vbNewLine & _
        "        Mod(Sum(Decode(Nvl(a.结帐Id,0),0,0,1)),2) As n_Count" & vbNewLine & _
        "       From 门诊费用记录 A " & vbNewLine & _
        "       Where a.No = [1] And a.记帐费用 = 1 And Mod(a.记录性质, 10) = 2" & vbNewLine & _
        "       Group By a.No, Nvl(a.价格父号, a.序号))"

        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "根据单据号获取记帐单是否已经结帐", strNos)
    End If
    bytOutStatus = Val(Nvl(rsTemp!结帐标志))
    GetBalanceStatus = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceExpenseDetails(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng结帐ID As Long, ByRef rsOutDetails As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定结帐的费用明细数据
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    lng结帐id -结帐ID
    '出参:rsOutDetails-结算数据(费用单号，收费类别、收费名称、收费数量、结帐金额，收费单价、计算单位、执行科室）
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutDetails = Nothing
    blnNOMoved = gobjDatabase.NOMoved("病人结帐记录", "", "ID", lng结帐ID, gstrCompentsName & ":检查结帐是否转储到历史表空间")
    
   strSql = "" & _
    "   Select A.发生时间, A.NO,nvl(价格父号,序号) as 序号,A.收费类别,A.收费细目ID," & _
    "           Avg(Nvl(付数,1)) *Avg(数次) as 数量,A.计算单位,sum(A.结帐金额) as 结帐金额,sum(a.标准单价 ) as 收费单价, " & _
    "           a.执行部门ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "门诊费用记录 A" & _
    "   Where A.结帐ID=[1]" & _
    "   Group by A.发生时间, A.NO,nvl(价格父号,序号),A.收费类别,A.收费细目ID,A.计算单位,a.执行部门ID" & _
    "   Union ALL " & _
    "   Select A.发生时间, A.NO,nvl(价格父号,序号) as 序号,A.收费类别,A.收费细目ID," & _
    "           Avg(Nvl(付数,1)) *Avg(数次) as 数量,A.计算单位,sum(A.结帐金额) as 结帐金额,sum(a.标准单价 ) as 收费单价, " & _
    "           a.执行部门ID" & _
    "   From " & IIf(blnNOMoved, "H", "") & "住院费用记录 A" & _
    "   Where A.结帐ID=[1] " & _
    "   Group by A.发生时间, A.NO,nvl(价格父号,序号),A.收费类别,A.收费细目ID,A.计算单位,a.执行部门ID" & _
    "   "
    strSql = _
    "  Select    A.NO as 费用单号,A.序号,A.收费类别,Nvl(E.名称,D.名称) as 收费名称,A.数量 as 收费数量, " & _
    "             a.结帐金额,a.收费单价 ,A.计算单位,Nvl(B.名称,'未知') as 执行科室 " & _
    " From (" & strSql & ") A,部门表 B,收费项目目录 D,收费项目别名 E" & _
    " Where A.执行部门ID=B.ID(+) And A.收费细目ID=D.ID" & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=3" & _
    " Order by 发生时间 Desc,费用单号 Desc,序号"
    Set rsOutDetails = gobjDatabase.OpenSQLRecord(strSql, gstrCompentsName & ":根据结帐ID获取结帐数据", lng结帐ID)
    GetBalanceExpenseDetails = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function GetBalanceInfor(ByVal frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal lng结帐ID As Long, ByRef rsOutBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定结算数据
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    lng结帐id -结帐ID
    '出参:rsOutDetails-结算数据( 结算方式、结算金额、结算号码,医疗卡类别ID,消费卡,交易流水号,交易说明,刷卡卡号）
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 17:42:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errHandle
    Dim blnNOMoved As Boolean
    
    Set rsOutBalance = Nothing
    blnNOMoved = gobjDatabase.NOMoved("病人结帐记录", "", "ID", lng结帐ID, gstrCompentsName & ":检查结帐是否转储到历史表空间")
    
   strSql = "" & _
    "   Select decode(mod(A.记录性质,10),1,'[冲预交]', A.结算方式) as 结算方式,  " & _
    "       冲预交 as 结算金额,A.结算号码, " & _
    "       A.卡类别ID,A.结算卡序号,decode(nvl(A.结算卡序号,0),0,0,1) as 消费卡, " & _
    "       A.交易流水号,A.交易说明,A.卡号 as 刷卡卡号 " & _
    "   From " & IIf(blnNOMoved, "H", "") & "病人预交记录 A" & _
    "   Where A.结帐ID=[1]"
    Set rsOutBalance = gobjDatabase.OpenSQLRecord(strSql, gstrCompentsName & ":根据结帐ID获取结帐数据", lng结帐ID)
    GetBalanceInfor = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    IncStr = gobjComlib.zlStr.Increase(strVal)
End Function
Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定内部模块编号所具有的权限
    '入参:lngProg-程序号
    '   blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    '出参:
    '返回:返回权限串
    '编制:刘兴洪
    '日期:2014-04-09 11:58:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
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
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.部门ID = Nvl(rsTmp!部门ID, 0)
            UserInfo.部门码 = Nvl(rsTmp!部门码)
            UserInfo.部门名 = Nvl(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = Get专业技术职务(UserInfo.ID)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.用户名
End Function

Public Function Get专业技术职务(ByVal lng人员id As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取当前登录人员的专业技术职务
    '返回:返回指写人员的专业技术职务
    '编制:刘兴洪
    '日期:2014-04-09 13:45:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
 
    strSql = "Select 专业技术职务 From 人员表 Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取人员专业职务", lng人员id)
    
    Get专业技术职务 = "" & rsTmp!专业技术职务
  
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取当前登录人员或指定人员的人员性质
    '返回:返回人员性质,多个用逗号分离
    '编制:刘兴洪
    '日期:2014-04-09 13:46:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    If str姓名 <> "" Then
        strSql = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取人员性质", str姓名)
    Else
        strSql = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取人员性质", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function GetRoom(str号别 As String) As String
'功能：根据号别的分诊方式获取号别的诊室
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
            
    strSql = "Select ID,Nvl(分诊方式,0) as 分诊 From 挂号安排 Where 号码=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", str号别)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp!分诊 = 0 Then Exit Function '不分诊
    
    '处理分诊
    If rsTmp!分诊 = 1 Then
        '指定诊室
        strSql = "Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", CLng(rsTmp!ID))
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 2 Then
        '动态分诊：该个号别当天挂号未诊数最少的诊室   //todo未考虑预约挂号
        strSql = _
            " Select 门诊诊室,Sum(NUM) as NUM From (" & _
                " Select 门诊诊室,0 as NUM From 挂号安排诊室 Where 号表ID=[1]" & _
                " Union ALL" & _
                " Select 诊室,Count(诊室) as NUM From 病人挂号记录" & _
                " Where Nvl(执行状态,0)=0 And 记录性质=1 and 记录状态=1 and  发生时间 Between Trunc(Sysdate) And Sysdate And 号别=[2]" & _
                " And 诊室 IN(Select 门诊诊室 From 挂号安排诊室 Where 号表ID=[1])" & _
                " Group by 诊室)" & _
            " Group by 门诊诊室 Order by Num"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlPublicExpense", CLng(rsTmp!ID), str号别)
        If Not rsTmp.EOF Then GetRoom = rsTmp!门诊诊室
    ElseIf rsTmp!分诊 = 3 Then
        '平均分诊：当前分配=1表示下次应取的当前诊室
        strSql = "Select 号表ID,门诊诊室,当前分配 From 挂号安排诊室 Where 号表ID=" & rsTmp!ID
        Set rsTmp = New ADODB.Recordset
        Call gobjDatabase.OpenRecordset(rsTmp, strSql, "mdlPublicExpense", adOpenDynamic, adLockOptimistic)
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!当前分配), 0, rsTmp!当前分配) = 1 Then
                    GetRoom = rsTmp!门诊诊室
                    rsTmp!当前分配 = 0
                    
                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!当前分配 = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '处理第一次平均分配
            If GetRoom = "" Then
                rsTmp.MoveFirst
                GetRoom = rsTmp!门诊诊室
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!当前分配 = 1
                rsTmp.Update
            End If
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function ReadRegistPrice(ByVal lng项目id As Long, ByVal bln病历 As Boolean, ByVal bln就诊卡 As Boolean, _
    Optional str费别 As String, Optional rsItems As ADODB.Recordset, Optional rsIncomes As ADODB.Recordset, _
    Optional lng病人ID As Long, Optional int险类 As Integer, Optional str号别 As String, Optional bytMode As Integer, _
    Optional lng挂号科室ID As Long = 0, Optional ByVal strPriceGrade As String, Optional strDate As String) As Long
'功能：读取指定挂号项目对应的费用信息到记录集中
'参数：lng项目ID=表示是否读取挂号费用(要读的挂号项目ID)
'      bln病历=表示是否读取病历工本费(可能仅收取病历费)
'      bln就诊卡=表示是否读取就诊卡费用(与挂号费或病历费一起收取)
'      str费别=挂号费别
'      rsItems(Out)=包含挂号项目及从属项目,不能以New方式定义
'      rsInComes(Out)=包含各个项目的收入情况,不能以New方式定义
'返回：读取的项目个数,同时rsItems,rsInCome=Nothing
'说明：主项数次为1,从项按设定数次处理,但为固定
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lng原项ID As Long
    Dim rsFeeTmp As ADODB.Recordset
    Dim strFee As String
    Dim str附加项目ID As String
    Dim strWherePriceGrade As String
    Dim strDateCondition As String
    
    Set rsItems = Nothing
    Set rsIncomes = Nothing
    
    If strDate <> "" Then
        strDateCondition = " [4] "
    Else
        strDateCondition = " Sysdate "
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [3]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [3]" & vbNewLine & _
            "                                   And " & strDateCondition & " Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null "
    End If
    
    '读取挂号项目及从属项目的费用
    If lng项目id <> 0 Then
        strSql = _
            "Select 1 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=[1]" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
        strSql = strSql & " Union ALL " & _
            "Select 2 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " D.从项数次 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=[1]" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    '读取病历工本费对应的费用
    If bln病历 Then
        strSql = strSql & IIf(strSql <> "", " Union ALL ", "") & _
            "Select 3 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
            " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,A.执行科室 as 执行科室类型" & _
            " From 收费项目目录 A,收费价目 B,收入项目 C,收费特定项目 D" & _
            " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.收费细目ID And D.特定项目='病历费'" & _
            " And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    
    If bytMode <> 1 And bytMode <> 10 Then
        strFee = "Select zl_Fun_CustomRegExpenses([1],[2],[3]) As 附加费 From Dual"
        Set rsFeeTmp = gobjDatabase.OpenSQLRecord(strFee, "zl_Fun_CustomRegExpenses", lng病人ID, int险类, str号别)
        If Not rsFeeTmp.EOF Then
            str附加项目ID = Nvl(rsFeeTmp!附加费)
        End If
        
        If str附加项目ID <> "" Then
            If strSql = "" Then
                strSql = " " & _
                    "Select /*+cardinality(D,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                    " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                    " From 收费项目目录 A,收费价目 B,收入项目 C,Table(f_str2list([2])) D " & _
                    " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.Column_Value " & _
                    "       And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                            strWherePriceGrade
            Else
                strSql = strSql & " Union ALL " & _
                    "Select /*+cardinality(D,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                    " 1 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                    " From 收费项目目录 A,收费价目 B,收入项目 C,Table(f_str2list([2])) D " & _
                    " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.Column_Value " & _
                    "       And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                            strWherePriceGrade
            End If
            strSql = strSql & " Union ALL " & _
                "Select /*+cardinality(E,10)*/ 5 as 性质,A.类别,A.ID as 项目ID,A.名称 as 项目名称,A.编码 as 项目编码,A.计算单位,A.屏蔽费别," & _
                " D.从项数次 as 数次,C.ID as 收入项目ID,C.名称 as 收入项目,C.编码 as 收入编码,C.收据费目,B.现价 as 单价,-1 as 执行科室类型" & _
                " From 收费项目目录 A,收费价目 B,收入项目 C,收费从属项目 D,Table(f_str2list([2])) E" & _
                " Where B.收费细目ID=A.ID And B.收入项目ID=C.ID And A.ID=D.从项ID And D.主项ID=E.Column_Value " & _
                "       And " & strDateCondition & " Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-1-1', 'YYYY-MM-DD'))" & vbNewLine & _
                        strWherePriceGrade
        End If
    End If
    
    If strSql = "" Then Exit Function
    
    '按主项,从项,病历顺序排列
    strSql = "Select * From (" & strSql & ") Order by 性质,项目编码,收入编码"
    
    On Error GoTo errH
    If strDate <> "" Then
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng项目id, str附加项目ID, strPriceGrade, CDate(strDate))
    Else
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng项目id, str附加项目ID, strPriceGrade)
    End If
    If Not rsTmp.EOF Then
        '先创建记录集
        Set rsItems = New ADODB.Recordset
        rsItems.Fields.Append "性质", adSmallInt '1-主项,2-从项,3-病历费,4-就诊卡费
        rsItems.Fields.Append "执行科室ID", adBigInt
        rsItems.Fields.Append "类别", adVarChar, 1
        rsItems.Fields.Append "项目ID", adBigInt
        rsItems.Fields.Append "项目名称", adVarChar, 80
        rsItems.Fields.Append "计算单位", adVarChar, 20, adFldIsNullable
        rsItems.Fields.Append "数次", adSingle
        rsItems.Fields.Append "保险项目否", adSmallInt, , adFldIsNullable
        rsItems.Fields.Append "保险大类ID", adBigInt, , adFldIsNullable
        rsItems.Fields.Append "保险编码", adVarChar, 80
        
        rsItems.CursorLocation = adUseClient
        rsItems.LockType = adLockOptimistic
        rsItems.CursorType = adOpenStatic
        rsItems.Open
        
        Set rsIncomes = New ADODB.Recordset
        rsIncomes.Fields.Append "项目ID", adBigInt
        rsIncomes.Fields.Append "收入项目ID", adBigInt
        rsIncomes.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
        rsIncomes.Fields.Append "单价", adSingle
        rsIncomes.Fields.Append "应收", adCurrency
        rsIncomes.Fields.Append "实收", adCurrency
        rsIncomes.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
        rsIncomes.CursorLocation = adUseClient
        rsIncomes.LockType = adLockOptimistic
        rsIncomes.CursorType = adOpenStatic
        rsIncomes.Open
        
        For i = 1 To rsTmp.RecordCount
            '挂号项目部份
            If lng原项ID <> rsTmp!项目ID Then
                rsItems.AddNew
                rsItems!性质 = rsTmp!性质
                 '0-无明确科室,1-病人所在科室,2-病人所在病区,3-开单人所在科室,4-指定科室
                If rsTmp!执行科室类型 = -1 Then
                    rsItems!执行科室ID = lng挂号科室ID      '0-表示挂号科室
                Else
                    rsItems!执行科室ID = Get挂号执行科室ID(rsTmp!项目ID, rsTmp!执行科室类型)
                    If rsItems!执行科室ID = 0 Then rsItems!执行科室ID = lng挂号科室ID
                End If
                
                rsItems!类别 = rsTmp!类别
                rsItems!项目ID = rsTmp!项目ID
                rsItems!项目名称 = rsTmp!项目名称
                rsItems!计算单位 = rsTmp!计算单位
                rsItems!数次 = Format(Nvl(rsTmp!数次, 0), "0.000")
                rsItems.Update
            End If
            lng原项ID = rsTmp!项目ID
            
            '收入项目部份
            rsIncomes.AddNew
            rsIncomes!项目ID = rsTmp!项目ID
            rsIncomes!收入项目ID = rsTmp!收入项目ID
            rsIncomes!收据费目 = rsTmp!收据费目
            rsIncomes!单价 = Format(Nvl(rsTmp!单价, 0), "0.00")
            rsIncomes!应收 = Format(rsItems!数次 * rsIncomes!单价, "0.00")
            If Nvl(rsTmp!屏蔽费别, 0) = 1 Then
                rsIncomes!实收 = rsIncomes!应收
            Else
                rsIncomes!实收 = Format(GetActualMoney(str费别, rsTmp!收入项目ID, rsIncomes!应收, rsTmp!项目ID), "0.00")
            End If
            rsIncomes.Update
            rsTmp.MoveNext
        Next
        ReadRegistPrice = rsItems.RecordCount
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
    Set rsItems = Nothing
    Set rsIncomes = Nothing
End Function

Public Function Get挂号执行科室ID(ByVal lng项目id As Long, ByVal int执行科室类型 As Integer) As Long
'功能：获取挂号附加项目(病历费,就诊卡费)的收费项目的执行科室
'参数：
'返回：如果返回零,表示挂号科室(医生所在科室)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Get挂号执行科室ID = UserInfo.部门ID
    
    Select Case int执行科室类型
        Case 0 '0-无明确科室
        Case 1 '1-病人所在科室
            Get挂号执行科室ID = 0
        Case 2 '2-病人所在病区
            Get挂号执行科室ID = 0
        Case 3 '3-操作员科室
        Case 4 '4-指定科室
            strSql = "Select 执行科室ID From 收费执行科室 Where 收费细目ID=[1] And Nvl(病人来源,1)=1 "
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng项目id)
            
            If Not rsTmp.EOF Then Get挂号执行科室ID = rsTmp!执行科室ID
        Case 5 '院外执行(预留,程序暂未用)
        Case 6 '开单人科室
    End Select
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function ActualMoney(str费别 As String, ByVal lng收入项目ID As Long, ByVal cur应收金额 As Currency, _
    Optional ByVal lng收费细目ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal dbl数量 As Double, Optional ByVal dbl加班加价率 As Double) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据收费细目ID或收入项目ID(前者优先),应收金额,按费别设置的分段比例打折规则计算实收金额；
    '     或对药品按成本加收比例规则计算实收金额
    '入参:str费别=病人费别；如果是按动态费别,传入格式为"病人费别,动态费别1,动态费别2,..."
    '      lng库房ID,dbl数量,对药品类项目按成本价加收打折时才需要传入
    '      dbl数量=包含付数在内的售价数量
    '      dbl加班加价率=小数比率,传入的应收金额已按加班加价计算时需要，用于还原及重算
    '出参:
    '返回:返回：按打折规则和比例计算的实收金额,如果是动态费别,则"str费别"返回最优惠费别(注意如果未打折计算,可能原样返回,也可能返回第一个)
    '编制:刘兴洪
    '日期:2014-04-09 13:54:17
    '说明:
    '   按成本价加收比例打折的两种计算方法(实际是一种)：
    '       1.打折金额 = 成本金额 * (1 + 加收比例)
    '       2.打折金额 = 成本价 * (1 + 加收比例) * 零售数量
    '   相关的计算公式：
    '      成本价 = 药品售价 * (1 - 差价率)
    '      成本金额 = 售价金额 * (1 - 差价率) = 成本价 * 零售数量
    '      有库存金额时:差价率 = 库存差价 / 库存金额,否则:差价率 = 指导差价率
    '      对于分批药品，应每个出库批次分别计算成本价和成本金额
    '      对于时价分批，"药品售价=Nvl(零售价,实际金额/实际数量)"；分批或时价药品库存不足时，不予打折计算。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str费别, lng收费细目ID, lng收入项目ID, cur应收金额 / (1 + dbl加班加价率), dbl数量, lng库房ID)
        
    str费别 = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl加班加价率), gSysPara.Money_Decimal.strFormt_VB)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function


Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
    '入参:vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    '出参:
    '返回:返回格式化的串
    '编制:刘兴洪
    '日期:2014-04-09 14:05:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    FormatEx = gobjComlib.FormatEx(vNumber, intBit, blnShowZero)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:模拟Oracle的Decode函数
    '返回:返回满足条件的值
    '编制:刘兴洪
    '日期:2014-04-09 14:04:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'      strUseType-使用类别
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSql = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录 " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3]  " & _
        "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    '目前的票据号可能和上次不同，所以需要检查范围
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSql = "" & _
    "   Select ID, 前缀文本, 开始号码, 终止号码" & vbNewLine & _
    "   From 票据领用记录" & vbNewLine & _
    "   Where 票种 = [1] And 剩余数量 >= [2] And 领用人 = [3]  " & _
    "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL ) " & _
    "           And 使用方式 = 1" & vbNewLine & _
    "   Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 desc, 开始号码"
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, UserInfo.姓名, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = "" & !前缀文本
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        strSql = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录  " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3] " & _
        "   And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
    '参数：bytKind=票种
    '      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
    '      strBill=要检查范围的票据号
    '说明：
    '    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
    '    2.在检查范围时,长度也在检查范围之内。
    '    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
    '返回：
    '      正常：票据领用ID>0
    '      0=失败
    '      -1:没有自用(用完或未领用)、也没有共用(未设置)
    '      -2:设置的共用已用完
    '      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSql As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '操作员有剩余的自用票据集
    strSql = _
        "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = [1] And 使用方式 = 1 And 剩余数量 > 0 And 领用人 = [2] And (Nvl(使用类别,'LXH')=[3] or  使用类别 is NULL)" & vbNewLine & _
        "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 Desc, 开始号码"
    Set rsSelf = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, UserInfo.姓名, IIf(strUseType = "", "LXH", strUseType))
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        lngReturn = rsSelf!ID
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        strSql = "Select ID,使用方式,剩余数量,前缀文本,开始号码,终止号码 From 票据领用记录 Where 票种=[1]  And (Nvl(使用类别,'LXH')=[3] or  使用类别 is NULL) And ID=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "可用票据批次", bytKind, lng领用ID, IIf(strUseType = "", "LXH", strUseType))
        '问题26352 by 张险华 2009-11-20
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If rsTmp!使用方式 = 2 Then '共用,要先看有没有自用
            If Not rsSelf.EOF Then
                '有自用的，优先
                lngReturn = rsSelf!ID
            Else
                '没有自用取共用
                If rsTmp!剩余数量 = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '自用票据
            If rsTmp!剩余数量 > 0 Then
                '有剩余
                lngReturn = rsTmp!ID
            Else
                '其它有剩余的自用
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
                lngReturn = rsSelf!ID
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!开始号码) And UCase(strBill) <= UCase(rsTmp!终止号码) And Len(strBill) = Len(rsTmp!开始号码)) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
    CheckUsedBill = 0
End Function

Public Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：1.当取不到范围内的有效票据时,返回空由用户输入
'      2.排开已报损的号码
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSql As String, strBill As String
    
    On Error GoTo errH
    
    strSql = "Select 前缀文本,开始号码,终止号码,当前号码" & _
        " From 票据领用记录 Where 剩余数量>0 And ID=[1]"
    Set rsMain = gobjDatabase.OpenSQLRecord(strSql, "取一下票据号", lng领用ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!当前号码) Then
        strBill = UCase(rsMain!开始号码)
    Else
        strBill = UCase(gobjCommFun.IncStr(rsMain!当前号码))
    End If
    
     '问题号:25448
     '刘兴洪:取消了;性质=1 And 原因=5 And 语句:原因是可能存在已经使用了的票据,使用了的,则排除
     '票种: 1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
     '性质:1-发出(原因中1、3、5属该性质)；2-收回(原因中2、4属该性质)
     '原因:1-正常发出票据；2-作废收回废票；3-重打发出票据；4-重打收回票据；5-毁损弃置票据
     
    strSql = "Select Upper(号码) as 号码 From 票据使用明细" & _
        " Where 号码||''>=[1] And 领用ID=[2]" & _
        " Order by 号码"
        
    Set rsDelete = gobjDatabase.OpenSQLRecord(strSql, "取一下票据号", strBill, lng领用ID)
    Do While True
        '检查范围
        If Left(strBill, Len("" & rsMain!前缀文本)) <> UCase("" & rsMain!前缀文本) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!开始号码) And strBill <= UCase(rsMain!终止号码)) Then
            Exit Function
        End If
                
        '排开报损号
        rsDelete.Filter = "号码='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = gobjCommFun.IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
    '入参:strText-日期文本
    '     blnTime=是否处理时间部份
    '出参:
    '返回:返回完整的日期串(yyyy-MM-dd[ HH:mm])
    '编制:刘兴洪
    '日期:2014-04-09 14:03:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = gobjDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function
Public Function NeedName(strList As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:优先判断以回车符分割
    '入参:strList:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
    '     2-分隔符有优先级：回车符(Chr(13)）> - > [] > ()
    '出参:
    '返回: 获取名称
    '编制:刘兴洪
    '日期:2014-04-09 14:03:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    NeedName = gobjComlib.zlStr.NeedName(strList)
    
End Function
Public Function BillExistBalance(ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的收费划价单是否存在已经收费的内容
    '入参:strNO-单据号
    '出参:
    '返回:已收费返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-09 14:12:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select ID From 门诊费用记录 Where 记录性质=1 And 记录状态 IN(1,3) And NO=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "BillExistBalance", strNO)

    BillExistBalance = Not rsTmp.EOF
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function


Public Function ExistIOClass(bytBill As Byte) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否存在指定处方单据类型的入出类别
    '返回:返回入出类别ID
    '编制:刘兴洪
    '日期:2014-04-09 14:17:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 类别ID From 药品单据性质 Where 单据=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", bytBill)
    If Not rsTmp.EOF Then ExistIOClass = Nvl(rsTmp!类别ID, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function GetBillMax序号(ByVal strNO As String, ByVal int记录性质 As Integer, str登记时间 As String, int病人来源 As Integer) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据当前的最大序号+1
    '入参:str登记时间=组合医嘱只生成了部份主费用时，将要新生成的收费划价单(NO相同)的时间与已生成的一致。
    '     int病人来源:1-门诊，2-住院
    '出参:
    '返回:返回当前最大序号+1
    '编制:刘兴洪
    '日期:2014-04-09 14:18:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strTab As String
    
    strTab = IIf(int记录性质 = 1 Or (int记录性质 = 2 And int病人来源 = 1), "门诊费用记录", "住院费用记录")
    On Error GoTo errHandle
    
    str登记时间 = ""
    strSql = "Select Max(序号) as 序号,Max(登记时间) as 时间 From " & strTab & " Where NO=[1] And 记录性质=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", strNO, int记录性质)
    If Not rsTmp.EOF Then
        GetBillMax序号 = Nvl(rsTmp!序号, 0) + 1
        If Not IsNull(rsTmp!时间) Then
            str登记时间 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
        End If
    Else
        GetBillMax序号 = 1
    End If
    Exit Function
    
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function
Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将0零转换为"NULL"串,在生成SQL语句时用
    '入参:blnForceNum=当为Null时，是否强制表示为数字型
    '出参:
    '返回:返回完整的SQL语句
    '编制:刘兴洪
    '日期:2014-04-09 14:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZVal = gobjComlib.ZVal(varValue, blnForceNum)
End Function


Public Function AnalyseComputer() As String
    AnalyseComputer = gobjComlib.OS.ComputerName
End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人当天发生的费用总额
    '返回:返回病人的当日费用总额
    '编制:刘兴洪
    '日期:2014-04-09 14:59:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    strSql = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng病人ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!金额, 0)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function zlPatiCardCheck(ByVal byt调用场合 As Byte, lng病人ID As Long, str卡号 As String, byt刷卡方式 As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查病人刷卡方式
    '入参：byt调用场合: 1-挂号;2-收费
    '         lng病人ID:病人ID(未建档的,传入零)
    '         str卡号;未刷卡时,为空
    '         byt刷卡方式: 1-普能刷卡;2-医保刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-04-27 16:09:08
    '说明：一汽集团的离休病人，使用的医保卡同时也是就诊卡；医院要求必须以医保方式进行
    '          身份验证挂号、收费，而不能以自费方式直接刷卡进行；因此要求在挂号、收费时，离休病人刷卡后如果不是以医保身份验证方式刷的卡，
    '          而是直接刷的卡，就提示并不允许继续。
    '问题:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    strSql = " Select Zl_Paticardcheck([1],[2],[3],[4]) as 提示信息 From Dual "
    ' Zl_Paticardcheck
    '  调用场合_IN NUMBER ,
    '  病人id_In Number,
    '  卡号_In   Varchar2,
    '  刷卡方式_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查病人刷卡方式是否合法", byt调用场合, lng病人ID, str卡号, byt刷卡方式)
    strSql = Nvl(rsTemp!提示信息)
    If strSql <> "" Then
        MsgBox strSql, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
    rsWarn As ADODB.Recordset, ByVal str姓名 As String, ByVal cur剩余款额 As Currency, _
    ByVal cur当日金额 As Currency, ByVal cur记帐金额 As Currency, ByVal cur担保金额 As Currency, _
    ByVal str收费类别 As String, ByVal str类别名称 As String, str已报类别 As String, _
    intWarn As Integer, Optional ByVal bln划价 As Boolean, _
    Optional blnNotCheck类别 As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对病人记帐进行报警提示
    '入参:rsWarn=包含报警参数设置的记录集(该病人病区,并区分好了医保)
    '     str收费类别=当前要检查的类别,用于分类报警
    '     str类别名称=类别名称,用于提示
    '     bln划价=生成划价费用时的报警，类似具有欠费强制记帐权限时的处理
    '     intWarn=是否显示询问性的提示,-1=要显示,0=缺省为否,1-缺省为是
    '     blnNotCheck类别:不对类别进行检查(主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
    '出参:
    '返回:intWarn=本次询问性提示中的选择结果,0=为否,1-为是
    '     0;没有报警,继续
    '     1:报警提示后用户选择继续
    '     2:报警提示后用户选择中断
    '     3:报警提示必须中断
    '     4:强制记帐报警,继续
    '编制:刘兴洪
    '日期:2014-04-09 15:00:33
    '说明:str已报类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim bln已报警 As Boolean, byt标志 As Byte
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant, vMsg As VbMsgBoxResult
    Dim str担保 As String, i As Long
    
    BillingWarn = 0
    
    '报警参数检查:NULL是没有设置,0是设置了的
    If rsWarn.State = 0 Then Exit Function
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str收费类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志1 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str收费类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志2 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str收费类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名称 = "" '所有类别时,不必提示具体的类别
        '刘兴洪 问题:26952 日期:2009-12-25 16:42:54
        '   主要是在针对刚选择病人后，还未输入相关的数据时的首次检查.这情况只能针对限制的类别为所有类别，如果分类别限制的，在这种情况下就不检查,只有再输入内容后才检查!)
        If rsWarn!报警标志3 <> "-" And blnNotCheck类别 Then Exit Function
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    '示例："-" 或 ",ABC,567,DEF"
    '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
    bln已报警 = InStr(str已报类别, str收费类别) > 0 Or str已报类别 Like "-*"
    
    If bln已报警 Then '当intWarn = -1时,也可强行再报警
        If byt标志 = 2 Then
            If str已报类别 Like "-*" Then
                byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
            Else
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If InStr(arrTmp(i), str收费类别) > 0 Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For '取消说明见住院记帐模块
                    End If
                Next
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名称 <> "" Then str类别名称 = """" & str类别名称 & """费用"
    str担保 = IIf(cur担保金额 = 0, "", "(含担保额:" & Format(cur担保金额, "0.00") & ")")
    cur剩余款额 = cur剩余款额 + cur担保金额 - cur记帐金额
    cur当日金额 = cur当日金额 + cur记帐金额
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur剩余款额 < 0 Then
                        byt方式 = 2
                        If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                If vMsg = vbIgnore Then intWarn = 1
                            End If
                            BillingWarn = 3
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽,允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    ElseIf cur剩余款额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 1
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 1
                                End If
                            End If
                        Else
                            If intWarn = -1 Then
                                vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If vMsg = vbCancel Then intWarn = 0
                                    BillingWarn = 2
                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                    If vMsg = vbIgnore Then intWarn = 1
                                    BillingWarn = 4
                                End If
                            Else
                                If intWarn = 0 Then
                                    BillingWarn = 2
                                ElseIf intWarn = 1 Then
                                    BillingWarn = 4
                                End If
                            End If
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur剩余款额 < 0 Then
                            byt方式 = 2
                            If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽," & str类别名称 & "禁止记帐。", frmParent, True)
                                    If vMsg = vbIgnore Then intWarn = 1
                                End If
                                BillingWarn = 3
                            Else
                                If intWarn = -1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & "已经耗尽,允许该病人记帐吗？", frmParent)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If vMsg = vbCancel Then intWarn = 0
                                        BillingWarn = 2
                                    ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                        If vMsg = vbIgnore Then intWarn = 1
                                        BillingWarn = 4
                                    End If
                                Else
                                    If intWarn = 0 Then
                                        BillingWarn = 2
                                    ElseIf intWarn = 1 Then
                                        BillingWarn = 4
                                    End If
                                End If
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur剩余款额 < rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当前剩余款" & str担保 & ":" & Format(cur剩余款额, "0.00") & ",低于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gSysPara.Money_Decimal.strFormt_VB) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 1
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 1
                            End If
                        End If
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gSysPara.Money_Decimal.strFormt_VB) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日金额 > rsWarn!报警值 Then
                    If Not (InStr(";" & strPrivs & ";", ";欠费强制记帐;") > 0 Or bln划价) Then
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gSysPara.Money_Decimal.strFormt_VB) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", frmParent, True)
                            If vMsg = vbIgnore Then intWarn = 1
                        End If
                        BillingWarn = 3
                    Else
                        If intWarn = -1 Then
                            vMsg = frmMsgBox.ShowMsgBox(str姓名 & " 当日费用:" & Format(cur当日金额, gSysPara.Money_Decimal.strFormt_VB) & ",高于" & str类别名称 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",允许该病人记帐吗？", frmParent)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                If vMsg = vbCancel Then intWarn = 0
                                BillingWarn = 2
                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
                                If vMsg = vbIgnore Then intWarn = 1
                                BillingWarn = 4
                            End If
                        Else
                            If intWarn = 0 Then
                                BillingWarn = 2
                            ElseIf intWarn = 1 Then
                                BillingWarn = 4
                            End If
                        End If
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If BillingWarn = 1 Or BillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = "-"
            Else
                str已报类别 = str已报类别 & "," & rsWarn!报警标志3
            End If
        End If
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
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "": bln医保 = False: bln公费 = False
    If grs医疗付款方式 Is Nothing Then
        strSql = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    ElseIf grs医疗付款方式.State <> 1 Then
        strSql = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    End If
    If strSql <> "" Then
        Set grs医疗付款方式 = gobjDatabase.OpenSQLRecord(strSql, "获取医疗付款方式")
    End If
    grs医疗付款方式.Find "名称='" & str医疗付款名称 & "'", , adSearchForward, 1
    If grs医疗付款方式.EOF Then Exit Function
    bln医保 = Val(Nvl(grs医疗付款方式!是否医保)) = 1
    bln公费 = Val(Nvl(grs医疗付款方式!是否公费)) = 1
    zlIsCheckMedicinePayMode = bln医保 Or bln公费
    
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function ShowHelp(ByVal ChmName As String, SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '功能:显示帮助窗体
    '参数:ChmName:CHM格式文件(目前传入的是:App.ProductName)
    '     SHwnd:传入窗口句柄(作为宿主窗口)
    '     htmName:射映在CHM中的htm文件名称
    '编制:刘兴洪
    '日期:2014-05-15 15:49:52
    '-----------------------------------------------------------------------------------------------------------------------------
    ShowHelp = gobjComlib.ShowHelp(ChmName, SHwnd, htmName, Sys)
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复窗体的状态，当左顶边界超出时，则自动设置为0
    '入参:objForm:要恢复的窗体
    '     strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
    '     strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-15 15:53:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
   RestoreWinState = gobjComlib.RestoreWinState(objForm, strProjectName, strUserDef)
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存窗体及其中各种控件的状态
    '入参: objForm:要保存的窗体
    '      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
    '      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    '编制:刘兴洪
    '日期:2014-05-15 15:55:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
   SaveWinState = gobjComlib.SaveWinState(objForm, strProjectName, strUserDef)
End Function
Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function
 



Public Function zlGetDefaultWindow(ByVal str类别 As String, ByVal lng药房ID As Long, _
    ByVal lngModule As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的药房窗口设置
    '入参:str类别-收费类别
    '     lng药房ID-药房ID
    '     lngModule-模块号
    '出参:
    '返回:返回缺省的发药窗口
    '编制:刘兴洪
    '日期:2014-07-23 18:38:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    Dim str西窗 As String, lng西药房 As Long
    Dim str成窗 As String, lng成药房 As Long
    Dim str中窗 As String, lng中药房 As Long
    Select Case str类别
        Case "5"
            str西窗 = gobjDatabase.GetPara("西药房窗口", glngSys, lngModule)
            If lngModule = 1252 Then
                lng西药房 = Val(gobjDatabase.GetPara("门诊缺省西药房", glngSys, lngModule))
            Else
                lng西药房 = Val(gobjDatabase.GetPara("缺省西药房", glngSys, lngModule))
            End If
            If InStr(str西窗, ":") > 0 Then '旧数据没有存药房ID
                 strTmp = str西窗
            ElseIf lng西药房 > 0 And str西窗 <> "" Then
                strTmp = lng西药房 & ":" & str西窗
            End If
        Case "6"
            str成窗 = gobjDatabase.GetPara("成药房窗口", glngSys, lngModule)
            If lngModule = 1252 Then
                lng成药房 = Val(gobjDatabase.GetPara("门诊缺省成药房", glngSys, lngModule))
            Else
                lng成药房 = Val(gobjDatabase.GetPara("缺省成药房", glngSys, lngModule))
            End If
            If InStr(str成窗, ":") > 0 Then
                 strTmp = str成窗
            ElseIf lng成药房 > 0 And str成窗 <> "" Then
                 strTmp = lng成药房 & ":" & str成窗
            End If
        Case "7"
            str中窗 = gobjDatabase.GetPara("中药房窗口", glngSys, lngModule)
            If lngModule = 1252 Then
                lng中药房 = Val(gobjDatabase.GetPara("门诊缺省中药房", glngSys, lngModule))
            Else
                lng中药房 = Val(gobjDatabase.GetPara("缺省中药房", glngSys, lngModule))
            End If
            If InStr(str中窗, ":") > 0 Then
                 strTmp = str中窗
            ElseIf lng中药房 > 0 And str中窗 <> "" Then
                 strTmp = lng中药房 & ":" & str中窗
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
    zlGetDefaultWindow = strTmp
End Function

Public Function zlGet发药窗口(ByVal lngModule As Long, ByVal curDate As Date, ByVal lng药房ID As Long, ByVal str类别 As String, _
    str西窗 As String, str成窗 As String, str中窗 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取药品对应的发药窗口
    '入参:lng药房ID=执行部门ID
    '     curDate=当前时间
    '返回:返回药品对应的发药窗口
    '编制:刘兴洪
    '日期:2014-07-23 18:40:35
    '说明:在同一材质类药房的发药窗口内平均分配
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng西药房 As Long, lng成药房 As Long, lng中药房 As Long
    
    On Error GoTo errH
    
    '指定时固定分配(指定是指没有对应药房上班时指定)
    Select Case str类别
        Case "5"
            lng西药房 = Val(gobjDatabase.GetPara(18, glngSys, lngModule))

            If str西窗 <> "" Then
                zlGet发药窗口 = str西窗
            ElseIf lng西药房 > 0 Then
                zlGet发药窗口 = zlGetDefaultWindow(str类别, lng药房ID, lngModule)
                str西窗 = zlGet发药窗口
            End If
        Case "6"
            lng成药房 = Val(gobjDatabase.GetPara(19, glngSys, lngModule))
            If str成窗 <> "" Then
                zlGet发药窗口 = str成窗
            ElseIf lng成药房 > 0 Then
                zlGet发药窗口 = zlGetDefaultWindow(str类别, lng药房ID, lngModule)
                str成窗 = zlGet发药窗口
            End If
        Case "7"
            lng中药房 = Val(gobjDatabase.GetPara(20, glngSys, lngModule))
            If str中窗 <> "" Then
                zlGet发药窗口 = str中窗
            ElseIf lng中药房 > 0 Then
                zlGet发药窗口 = zlGetDefaultWindow(str类别, lng药房ID, lngModule)
                str中窗 = zlGet发药窗口
            End If
    End Select
    
    
    If zlGet发药窗口 <> "" Then
        strSql = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlOutExse", lng药房ID, zlGet发药窗口)
        If rsTmp.EOF Then zlGet发药窗口 = ""
        Exit Function
    End If
    
    '动态分配上班的非专家窗口,98876
    strSql = "Select Zl_Get发药窗口([1],[2],[3]) As 窗口 From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取发药窗口", lng药房ID, Val(gobjDatabase.GetPara(19, glngSys, , 0)), curDate)
    If Not rsTmp.EOF Then
        zlGet发药窗口 = Nvl(rsTmp!窗口)
    End If
    
    If zlGet发药窗口 <> "" Then
        Select Case str类别
            Case "5"
                str西窗 = zlGet发药窗口
            Case "6"
                str成窗 = zlGet发药窗口
            Case "7"
                str中窗 = zlGet发药窗口
        End Select
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function GetActualMoney(ByVal str费别 As String, ByVal lng收入ID As Long, ByVal cur应收 As Currency, ByVal lng收费细目ID As Long) As Currency
'功能：根据指定的费别和收入项目或收费项目,计算指定金额的实际收款金额
'参数：
'   str费别   ：费别
'   lng收入ID  ：收入项目ID
'   cur应收：应收金额值
'返回：实际应收的金额
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
        
    strSql = "Select 实收比率" & vbNewLine & _
            "From 费别明细" & vbNewLine & _
            "Where 费别 = [1] And 收费细目id = [3] And Abs([4]) Between 应收段首值 And 应收段尾值" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 实收比率" & vbNewLine & _
            "From 费别明细 A" & vbNewLine & _
            "Where 费别 = [1] And 收入项目id = [2] And Abs([4]) Between 应收段首值 And 应收段尾值 And Not Exists" & vbNewLine & _
            " (Select 1 From 费别明细 C Where C.费别 = A.费别 And C.收费细目id = [3])"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, str费别, lng收入ID, lng收费细目ID, cur应收)
    If rsTmp.EOF Then
        GetActualMoney = cur应收
    Else
        GetActualMoney = cur应收 * rsTmp!实收比率 / 100
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function Get未发药品发药窗口(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As String
    '-------------------------------------------------------------------------
    '功能：判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
    '返回：若存在相同执行部门的未发药品，则返回未发药品的发药窗口，否则返回空
    '编制：冉俊明
    '日期：2014-04-09
    '问题：71902
    '说明：
    '   同一个人病人不同时间段多张单据收费，分配同一个发药窗口，方便病人取药
    '-------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSql = "Select 发药窗口" & vbNewLine & _
            "From 未发药品记录" & vbNewLine & _
            "Where 单据 = 8 And 发药窗口 Is Not Null And 病人id = [1] And 库房id = [2]" & vbNewLine & _
            "Order By 已收费 Desc, 填制日期 Desc"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取病人未发药品发药窗口", lng病人ID, lng执行部门ID)
    
    If Not rsTemp.EOF Then
        Get未发药品发药窗口 = Nvl(rsTemp!发药窗口)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function zlGetDrugWindow(ByVal lngModule As Long, ByVal lng药房ID As Long, ByVal str类别 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的发药窗口,如果参数指定了缺省,则以指定为准,否则,如果是划价单,则以第一药品行的窗口为准,否则以已输入相同药品的窗口为准
    '返回:返回发药窗口
    '编制:刘兴洪
    '日期:2014-07-23 18:49:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发药窗口 As String
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Err = 0: On Error GoTo errH:
    str发药窗口 = zlGetDefaultWindow(str类别, lng药房ID, lngModule)
    If str发药窗口 = "" Then Exit Function
    strSql = "Select 编码 From 发药窗口 Where 上班否=1 And 药房ID=[1] And 名称=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取缺省发药窗口", lng药房ID, str发药窗口)
    If rsTmp.EOF Then Exit Function
    zlGetDrugWindow = str发药窗口
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSql = strSql & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSql = strSql & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSql = strSql & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSql = strSql & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSql = strSql & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSql = strSql & "" & IIf(bln预交, 1, 0) & ","
    '退费标志 :1-退费;0-付费
    strSql = strSql & "0,"
    '校对标志
    strSql = strSql & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlAddArray cllPro, strSql
End Function
'
'Public Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
'    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
'    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
'
'
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:保存三方结算数据
'    '入参: bln预交款-是否预交款
'    '       lngID-如果是预交款,则是预交ID,否则结帐ID
'    ' 出参:cllPro-返回SQL集
'    '返回:成功,返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2011-07-19 10:23:30
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
'    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
'
'    Err = 0: On Error GoTo Errhand:
'    '先提交,这样避免风险,再更新相关的交易信息
'    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
'    varData = Split(strExpend, "||")
'    Dim str交易信息 As String, strTemp As String
'    For i = 0 To UBound(varData)
'        If Trim(varData(i)) <> "" Then
'            varTemp = Split(varData(i) & "|", "|")
'            If varTemp(0) <> "" Then
'                strTemp = varTemp(0) & "|" & varTemp(1)
'                If gobjCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
'                    str交易信息 = Mid(str交易信息, 3)
'                    'Zl_三方结算交易_Insert
'                    strSQL = "Zl_三方结算交易_Insert("
'                    '卡类别id_In 病人预交记录.卡类别id%Type,
'                    strSQL = strSQL & "" & lng卡类别ID & ","
'                    '消费卡_In   Number,
'                    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
'                    '卡号_In     病人预交记录.卡号%Type,
'                    strSQL = strSQL & "'" & str卡号 & "',"
'                    '结帐ids_In  Varchar2,
'                    strSQL = strSQL & "'" & strIDs & "',"
'                    '交易信息_In Varchar2:交易项目|交易内容||...
'                    strSQL = strSQL & "'" & str交易信息 & "',"
'                    '预交款缴款_In Number := 0
'                    strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
'                    zlAddArray cllPro, strSQL
'                    str交易信息 = ""
'                End If
'                str交易信息 = str交易信息 & "||" & strTemp
'            End If
'        End If
'    Next
'    If str交易信息 <> "" Then
'        str交易信息 = Mid(str交易信息, 3)
'        'Zl_三方结算交易_Insert
'        strSQL = "Zl_三方结算交易_Insert("
'        '卡类别id_In 病人预交记录.卡类别id%Type,
'        strSQL = strSQL & "" & lng卡类别ID & ","
'        '消费卡_In   Number,
'        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
'        '卡号_In     病人预交记录.卡号%Type,
'        strSQL = strSQL & "'" & str卡号 & "',"
'        '结帐ids_In  Varchar2,
'        strSQL = strSQL & "'" & strIDs & "',"
'        '交易信息_In Varchar2:交易项目|交易内容||...
'        strSQL = strSQL & "'" & str交易信息 & "',"
'        '预交款缴款_In Number := 0
'        strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
'        zlAddArray cllPro, strSQL
'    End If
'    zlAddThreeSwapSQLToCollection = True
'    Exit Function
'Errhand:
'    If gobjComlib.ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Public Function zlFormatNum(ByVal strMoney As String) As String
    strMoney = Replace(strMoney, Chr(44), "")
    zlFormatNum = strMoney
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
        lngColor = gobjDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function GetMoneyInfoRegist(lng病人ID As Long, Optional dblModiMoney As Double, _
    Optional blnInsure As Boolean, _
    Optional int类型 As Integer = -1, _
    Optional bln按类型统计 As Boolean = False, _
    Optional bytModiMoneyType As Byte = 0, _
    Optional ByVal blnFamilyMoney As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的剩余额
    '入参:blnInsure=是否排开医保病人的预结费用
    '       curModiMoney=修改时,原单据的当前病人的费用合计
    '       int类型:类型(0-门诊和住院共用;1-门诊;2-住院),-1表示所有
    '       bytModiMoneyType-修改费用的类别(在按类别统计时有效)
    '       blnFamilyMoney-是否读取家属余额
    '出参:
    '返回:病人剩余额
    '编制:刘兴洪
    '日期:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, bln医保 As Boolean, lng主页Id As Long
    Dim strSql As String
    On Error GoTo errH
    If blnInsure Then
        strSql = "Select A.险类,A.主页ID From 病案主页 A,病人信息 B" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
                " And B.病人ID=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID)
        If Not rsTmp.EOF Then
            bln医保 = Not IsNull(rsTmp!险类)
            lng主页Id = rsTmp!主页ID
        End If
    End If
    strSql = "Select " & IIf(bln按类型统计, "类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
            "       Nvl(费用余额,0) As 费用余额,Nvl(预交余额,0) As 预交余额" & _
            " From 病人余额" & _
            " Where 性质=1 And 病人ID=[1] " & IIf(int类型 = -1, "", " And 类型=[4]")
    '79868,读取病人家属余额
    If blnFamilyMoney Then
        strSql = strSql & " Union All " & _
                " Select " & IIf(bln按类型统计, "a.类型,", "") & IIf(blnFamilyMoney, "1 As 家属,", "") & _
                "       Nvl(a.费用余额, 0) As 费用余额, Nvl(a.预交余额, 0) As 预交余额" & _
                " From 病人余额 A, 病人家属 B" & _
                " Where a.病人id = b.家属id And b.病人id = [1] And a.性质 = 1 " & _
                "       And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) " & _
                IIf(int类型 = -1, "", " And 类型=[4]")
    End If
  
    If dblModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSql = strSql & " Union All " & _
                " Select " & IIf(bln按类型统计, "[4] as 类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
                "       -1*[3] as 费用余额,0 as 预交余额 From Dual"
    End If
    
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If blnInsure And bln医保 Then
        strSql = strSql & " Union All " & _
        " Select  " & IIf(bln按类型统计, "Decode(主页ID,NULL,1,0,1,2) as 类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
        "       -1*Nvl(金额,0) as 费用余额,0 as 预交余额" & _
        " From 保险模拟结算" & _
        " Where 病人ID=[1] And 主页ID=[2] "
    End If
    strSql = "Select " & IIf(bln按类型统计, "类型,", "") & IIf(blnFamilyMoney, "家属,", "") & _
            "       nvl(Sum(费用余额),0) as 费用余额,nvl(Sum(预交余额),0) as 预交余额 " & _
            " From (" & strSql & ")" & vbCrLf & _
            IIf(bln按类型统计 And blnFamilyMoney, " Group by 类型,家属", _
                IIf(bln按类型统计, " Group by 类型", IIf(blnFamilyMoney, " Group by 家属", "")))
    
    Set GetMoneyInfoRegist = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, lng主页Id, dblModiMoney, int类型)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Function ReCalcOld(ByVal DateBir As Date, Optional ByRef cbo年龄单位 As ComboBox, Optional ByVal lng病人ID As Long, Optional ByVal blnSetControl As Boolean = True) As String
'功能:根据出生日期重新计算病人的年龄,重设年龄单位
'参数:blnSetControl是否设置年龄单位控件
'返回:年龄,年龄单位
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTmp As String
 
    strSql = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, DateBir)
    If blnSetControl = False Then
        ReCalcOld = Trim(Nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*岁" Or rsTmp!old Like "*月" Or rsTmp!old Like "*天" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call gobjControl.Cbo.Locate(cbo年龄单位, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo年龄单位.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo年龄单位.ListIndex = 0
            Else
                cbo年龄单位.ListIndex = -1
            End If
        End If
    End If
    If cbo年龄单位.ListIndex = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Function CheckStructAddr(ByVal objCtl As PatiAddress, ByVal lngLen As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查结构化地址控件中的信息录入是否正确
    '入参:objCtl-结构化地址控件，lngLen-限制长度
    '返回:True-输入信息合法
    '编制:李南春
    '日期:2015-12-7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjCommFun.ActualLen(objCtl.Value) > lngLen Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "最多只能输入" & lngLen \ 2 & "个汉字,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.CheckNullValue(, True, False) <> "" Then
        MsgBox "注意:" & vbCrLf & "   " & objCtl.Tag & "的" & objCtl.CheckNullValue & "尚未输入,请检查。", vbInformation + vbOKOnly, gstrSysName
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        Exit Function
    End If
    CheckStructAddr = True
End Function

Public Function zlReadAddrInfo(ByVal objCtrl As PatiAddress, ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
                               ByVal intType As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定的病人地址信息到控件中
    '入参:objCtrl-结构化地址控件,intType -地址类型1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址
    '返回:
    '编制:李南春
    '日期:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    
    strSql = "Select 省,市,县,乡镇,其他 From 病人地址信息 Where 病人ID=[1] and Nvl(主页ID,0)=[2] and 地址类别=[3]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "查询结构化地址", lng病人ID, lng主页Id, intType)
    If rsTmp.RecordCount > 0 Then
        Call objCtrl.LoadStructAdress(Nvl(rsTmp!省), Nvl(rsTmp!市), Nvl(rsTmp!县), Nvl(rsTmp!乡镇), Nvl(rsTmp!其他))
    Else
        objCtrl.Value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlPatiIsReturnVisit(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前病人是否复诊病人
    '入参:lng病人ID-病人ID
    '    lng执行部门ID-挂号科室ID
    '出参:
    '返回:true-复诊,false-初诊病人
    '编制:刘兴洪
    '日期:2017-10-27 15:29:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "Select Zl1_Fun_GetReturnVisit([1],[2]) As 复诊标志 From Dual"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取缺省复诊标志", lng病人ID, lng执行部门ID)
    
    zlPatiIsReturnVisit = Val(Nvl(rsTmp!复诊标志)) = 1
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ShowMsgBox_Custom(ByVal frmMain As Object, ByVal strInfo As String, Optional ByVal blnNoAsk As Boolean, Optional ByVal intType As Integer) As VbMsgBoxResult
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
    Dim frmNewMsg As frmMsgBox
    
    Set frmNewMsg = New frmMsgBox
    ShowMsgBox_Custom = frmNewMsg.ShowMsgBox(strInfo, frmMain, blnNoAsk, intType)
    If Not frmNewMsg Is Nothing Then Unload frmNewMsg: Set frmNewMsg = Nothing
End Function

Public Function SelectWholeItems(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
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
    Dim frmNew As frmWholeSelect
    On Error GoTo errHandle
    Set frmNew = New frmWholeSelect
    SelectWholeItems = frmNew.ShowSelect(frmMain, lngModule, strPrivs, rsOutSel)
    If Not frmNew Is Nothing Then Unload frmNew
    Set frmNew = Nothing
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Function

Public Sub zlChargeSaveAfter_Plugin(ByVal lngModule As Long, ByVal lng病人ID, ByVal lng主页Id As Long, ByVal bln门诊 As Boolean, _
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
    Call gobjPlugIn.ChargeSaveAfter(glngSys, lngModule, lng病人ID, lng主页Id, bln门诊, int记录性质, strNos)
    If Err = 0 Then Exit Sub
    
    '注意，接口不存在时也会进入
    If Err.Number = 438 Then Exit Sub  '接口不存在，认为检查通过
    Call zlPlugInErrH(Err, "ChargeSaveAfter")
    Err = 0: On Error GoTo 0
    
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Sub


Public Function zlGetSaveDataItems_Plugin(ByVal objBills As ExpenseBill, ByRef rsItems As ADODB.Recordset) As Boolean
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
               rsItems!病人ID = objBills.病人ID
               rsItems!主页ID = objBills.主页ID
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
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetOneCard() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一卡通设置记录集
    '返回:返回一卡通设置记录集
    '编制:刘兴洪
    '日期:2014-07-04 10:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo errH
    
    If Not grsOneCard Is Nothing Then
        If grsOneCard.State = 1 Then
            Set GetOneCard = grsOneCard
            Exit Function
        End If
    End If
    strSql = "Select 编号,名称,医院编码,结算方式 From 一卡通目录 Where 启用=1"
    Set grsOneCard = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
    Set GetOneCard = grsOneCard
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function
Public Function zlOldOneCardIsStart(ByVal str结算方式 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查旧一卡通是否启用
    '入参:str结算方式-结算方式
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-01 11:45:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsOne As ADODB.Recordset
    Dim blnSart As Boolean
    On Error GoTo errHandle
    Set rsOne = GetOneCard
    If rsOne Is Nothing Then Exit Function
    
    rsOne.Filter = "结算方式='" & str结算方式 & "'"
    blnSart = Not rsOne.EOF
    rsOne.Filter = 0
    zlOldOneCardIsStart = blnSart
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function



Public Function zlInterfacePrayMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lng挂号结帐ID As Long, ByRef cllTheeSwap As Collection, _
    ByRef cllTheeSwapOther As Collection, dblMoney As Double, ByVal strCardNO As String, lng医疗卡类别ID As Long, bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllTheeSwap-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lng医疗卡类别ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(frmMain, lngModule, lng医疗卡类别ID, bln消费卡, strCardNO, lng挂号结帐ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If lng挂号结帐ID <> 0 Then
        '问题:58322
        'mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
        If Not bln消费卡 Then
            '消费卡已经在插入挂号记录时,已经扣款
            Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, strCardNO, strSwapGlideNO, strSwapMemo, cllTheeSwap)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, strCardNO, strSwapExtendInfor, cllTheeSwapOther)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSql As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSql = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSql = strSql & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSql = strSql & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSql = strSql & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSql = strSql & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSql = strSql & IIf(bln预交款, "1", "0") & ")"
                    zlAddArray cllPro, strSql
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSql = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSql = strSql & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSql = strSql & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSql = strSql & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSql = strSql & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSql = strSql & IIf(bln预交款, "1", "0") & ")"
        zlAddArray cllPro, strSql
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
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
    zlCloseWindows = Forms.Count = 0
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
    If Not gcolPrivs Is Nothing Then Set gcolPrivs = Nothing
    If Not gclsInsure Is Nothing Then Set gclsInsure = Nothing
    If Not gobjPlugIn Is Nothing Then Set gobjPlugIn = Nothing
    If Not gobjComlib Is Nothing Then Set gobjComlib = Nothing
    If Not gobjCommFun Is Nothing Then Set gobjCommFun = Nothing
    If Not gobjControl Is Nothing Then Set gobjControl = Nothing
    If Not gobjInExse Is Nothing Then Set gobjInExse = Nothing
    If Not grs医疗付款方式 Is Nothing Then Set grs医疗付款方式 = Nothing
    If Not grsOneCard Is Nothing Then Set grsOneCard = Nothing
    If Not grs收入项目 Is Nothing Then Set grs收入项目 = Nothing
    zlReleaseResources = True
End Function

Public Function PatiIdentify(ByVal lngModlue As Long, ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal curMoney As Currency, _
    Optional ByVal bln退费 As Boolean = False, Optional ByVal bytDepositShowMode As Byte = 0, Optional ByVal lngDefaultCardTypeID As Long = 0, _
    Optional ByVal blnFamilyMoney As Boolean, Optional ByVal blnOlnyFamilyIDs As Boolean, Optional strFamilyPatiIDs_Out As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行刷卡验证
    '入参:lngModlue-模块号
    '     dblMoney-金额
    '     lng病人ID-病人ID
    '     bln退费-当前是否退费操作
    '     bytDepositShowMode- 预交显示方式(0-余额汇总显示;1-只显示门诊余额;2-只显示住院余额)
    '     lngDefaultCardTypeID-缺省的刷卡类别
    '     blnFamilyMoney-是否读取家属预交余额
    '     blnOlnyFamilyIDs-true:不验卡，只读取家属IDs;False-需要读取卡验卡
    '出参:strFamilyPatiIDs-病人家属ID,多个用逗号分隔，79868
    '返回:身份验证成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-10-24 14:55:59
    '说明:
    '   一、消费验卡输入密码规则（bln退费=false时):
    '       1.不进行刷卡验证,直接返回True
    '       2.门诊消费时，需要进行刷卡验证，同时需要输入密码（无密码时,光标要经过密码框)
    '       3.门诊消费时设置密码的，则必须刷卡验卡且输入密码，无密码时,不需要验卡输入密码
    '       4.表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码(无密码时，光标要经过密码框)
    '  二、退费验卡（bln退费=true时):
    '       1.不进行刷卡控制，直接返回true
    '       2.门诊消费退费时需要刷卡验证,同时需要输入密码（无密码时,光标要经过密码框)
    '       3.门诊消费退费时设置密码的，则必须刷卡验证,无密码时,不需要验卡输入密码
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue  As String, dblBrushCardMoney As Double
    Dim byt消费验卡 As Byte, byt退费验卡 As Byte, blnPassWord As Boolean
    Dim varPara As Variant
    
    On Error GoTo errHandle
    '一卡通消费验证
    strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
    varPara = Split(strValue & "|||", "|")
    byt退费验卡 = Val(varPara(1)) '消费卡退费时是否刷卡验证
    
    dblBrushCardMoney = Val(varPara(0))
    
    If dblBrushCardMoney < 0 Then
        byt消费验卡 = 3 '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证  3-表示病人在消费N元内必须刷卡,不必输入密码即可支付;否则必须输入密码
        dblBrushCardMoney = -1 * dblBrushCardMoney  '刷卡时免密支付额度("gbyt预存款消费验卡"为3时有效)
    Else
        byt消费验卡 = Decode(dblBrushCardMoney, 1, 1, 2, 2, 0)
    End If
    
    If bln退费 Then
        '  byt消费验卡 '预存款退费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
        If byt退费验卡 = 0 Then PatiIdentify = True: Exit Function '不进行刷卡验证,直接返回True
        
        If gobjDatabase.PatiIdentify(frmMain, glngSys, lng病人ID, curMoney, lngModlue, bytDepositShowMode, lngDefaultCardTypeID, , _
                                 blnFamilyMoney, strFamilyPatiIDs_Out, Not blnOlnyFamilyIDs, (byt退费验卡 = 2)) Then Exit Function
        
        PatiIdentify = True: Exit Function
    End If
    If byt消费验卡 = 0 Then PatiIdentify = True: Exit Function '不进行刷卡验证,直接返回True
    
    
    If byt消费验卡 <> 3 Then
        blnPassWord = True
    ElseIf dblBrushCardMoney = 0 Then
        blnPassWord = True
    ElseIf curMoney > dblBrushCardMoney Then
        blnPassWord = True
    ElseIf curMoney = 0 Then '无金额时，无需验证密码
        blnPassWord = False
    Else
        blnPassWord = False
    End If
    
    If gobjDatabase.PatiIdentify(frmMain, glngSys, lng病人ID, curMoney, lngModlue, bytDepositShowMode, lngDefaultCardTypeID, blnPassWord, _
                                   blnFamilyMoney, strFamilyPatiIDs_Out, Not blnOlnyFamilyIDs, (byt消费验卡 = 2)) = False Then Exit Function
    PatiIdentify = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ReserveRegNo(ByVal str号码 As String, ByVal bln严格控制 As Boolean, ByVal bln分时段 As Boolean, _
                            ByVal strTime As String, ByRef lng序号 As Long, _
                            Optional ByVal str备注 As String, Optional ByVal lng记录ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 对当前序号锁号
    ' 入参 :str号码-号别，
    '       bln预约-是否预约锁号
    '       bln严格控制-是否严格控制
    '       bln分时段-是否分时段
    '       lng序号需要锁号的序号
    '       lng记录ID - 出诊排班模式需要传入记录id
    '       str备注 - 区分是哪里在锁号
    ' 出参 :lng序号:如果lng序号被锁定，则取新的序号锁号，并返回新的序号
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/18 15:34
    '---------------------------------------------------------------------------------------
    Dim lngRegLimit As Long, lngLastNo As Long, lngCurrentNo As Long, intTimes As Integer
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    If Not bln严格控制 Then ReserveRegNo = True: Exit Function
    If Not strTime Like "To_Date*" And strTime <> "" Then strTime = "To_Date('" & Format(strTime, "yyyy-MM-dd hh:mm:00") & "','YYYY-MM-DD HH24:MI:SS')"
    '138960:李南春,2019/3/26,不分时段的只取日期
    If Not bln分时段 Then
        strTime = Mid(strTime, InStr(strTime, "'") + 1)
        strTime = Trim(Left(strTime, InStr(strTime, "'") - 1))
        strTime = "To_Date('" & Format(strTime, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    End If
    On Error GoTo errH:
    If bln严格控制 And Not bln分时段 And lng序号 = 0 Then
Retry:
        If lng记录ID <> 0 Then
            strSql = "Select A.限号数,B.序号,Nvl(B.挂号状态,0) as 状态,Nvl(B.是否停诊,0) as 停诊  From 临床出诊记录 A,临床出诊序号控制 B Where A.ID=B.记录ID And A.ID= [1] order by B.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "挂号序号查询", lng记录ID)
            If rsTmp.RecordCount = 0 Then ReserveRegNo = True: Exit Function
            lngRegLimit = Val(Nvl(rsTmp!限号数))
            Do While Not rsTmp.EOF
                If Val(Nvl(rsTmp!状态)) = 0 And Val(rsTmp!停诊) <> 1 Then
                    lngCurrentNo = Val(Nvl(rsTmp!序号))
                    Exit Do
                End If
                lngLastNo = Val(Nvl(rsTmp!序号))
                rsTmp.MoveNext
            Loop
            If lngCurrentNo = 0 Then lngCurrentNo = lngLastNo + 1
            If lngCurrentNo > lngRegLimit Then
                ReserveRegNo = True
                Exit Function
            End If
            lng序号 = lngCurrentNo
        Else
            strSql = "Select A.限号数" & vbNewLine & _
                     "From 挂号安排限制 A, 挂号安排 B" & vbNewLine & _
                     "Where A.限制项目 =" & vbNewLine & _
                     "      Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And" & vbNewLine & _
                     "      A.安排id = B.ID And B.号码 = [1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "挂号限号限约", str号码)
            If Not rsTmp.EOF Then
                lngRegLimit = Val(Nvl(rsTmp!限号数))
            End If
            strSql = "Select 序号,状态" & vbNewLine & _
                     "From 挂号序号状态" & vbNewLine & _
                     "Where 号码 = [1] And 日期 Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60" & vbNewLine & _
                     "Order By 序号 Asc"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "挂号序号查询", str号码)
            Do While Not rsTmp.EOF
                If lngLastNo = 0 Then
                    lngLastNo = Val(Nvl(rsTmp!序号))
                Else
                    If Val(Nvl(rsTmp!序号)) - lngLastNo > 1 Then
                        lngCurrentNo = lngLastNo + 1
                    Else
                        lngLastNo = Val(Nvl(rsTmp!序号))
                    End If
                End If
                If Val(Nvl(rsTmp!状态)) = 4 Then lngRegLimit = lngRegLimit + 1
                rsTmp.MoveNext
            Loop
            If lngCurrentNo = 0 Then lngCurrentNo = lngLastNo + 1
            If lngCurrentNo > lngRegLimit Then '说明是加号，这种情况不用锁号
                ReserveRegNo = True
                Exit Function
            End If
            lng序号 = lngCurrentNo
        End If
    End If
    On Error GoTo errTry
    If lng序号 <> 0 Then
        If strTime <> "" Then
            strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & _
                      "'," & strTime & "," & lng序号 & "," & ZVal(lng记录ID) & ",'" & str备注 & "')"
        Else
            strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & _
                      "',To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-mm-dd") & "','YYYY-MM-DD')," & lng序号 & _
                      "," & ZVal(lng记录ID) & ",'" & str备注 & "')"
        End If
        Call gobjDatabase.ExecuteProcedure(strSql, "ReserveRegNo")
    End If
    ReserveRegNo = True
    Exit Function
errTry:
    intTimes = intTimes + 1
    If bln严格控制 And Not bln分时段 And intTimes < 4 Then GoTo Retry
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Function

Public Sub CancelRegNo(Optional ByVal lng记录ID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:取消挂号时删除锁定挂号序号
    '编制:李南春
    '日期:2019/2/18 15:34
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    On Error GoTo Errhand
    
    strSql = "Zl_挂号序号状态_Lock(2,'" & UserInfo.姓名 & "',Null,Null,Null," & ZVal(lng记录ID) & ")"
    Call gobjDatabase.ExecuteProcedure(strSql, "CancelRegNo")
    Exit Sub
Errhand:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrlog
End Sub

Public Sub InitAddressLength()
    Dim strSql As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select 家庭地址, 户口地址, 出生地点, 联系人地址 From 病人信息 Where Rownum < 2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "获取地址长度")
    If Not rsTmp.EOF Then
        glngMax家庭地址 = rsTmp.Fields("家庭地址").DefinedSize
        glngMax户口地址 = rsTmp.Fields("户口地址").DefinedSize
        glngMax出生地点 = rsTmp.Fields("出生地点").DefinedSize
        glngMax联系人地址 = rsTmp.Fields("联系人地址").DefinedSize
    End If
    If glngMax家庭地址 = 0 Then glngMax家庭地址 = 100: If glngMax户口地址 = 0 Then glngMax户口地址 = 100
    If glngMax出生地点 = 0 Then glngMax出生地点 = 100: If glngMax联系人地址 = 0 Then glngMax联系人地址 = 100
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrlog
End Sub

'========================================================================================================
'zlPlugIn外挂接口
'========================================================================================================
Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:外挂部件出错处理
    '入参:objErr 错误对象， strFunName 接口方法名称
    '出参:
    '编制:刘兴洪
    '日期:2014-04-09 13:27:19
    '说明:当方法不存在（错误号438）时不提示，其它错误弹出提示框
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & _
            objErr.Number & vbCrLf & _
            objErr.Description, vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

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

Public Function zlSaveRgstAfterByPlugIn(ByVal lngModule As Long, ByVal strNO As String, ByVal bln预约 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号完成后调用
    ' 入参 : lngModual=当前调用接口的主程序模块号
    '        strNo-挂号单号
    '        bln预约-预约挂号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/22 10:03
    '---------------------------------------------------------------------------------------
    If CreatePlugIn(lngModule, -1) = False Then Exit Function
    
    On Error Resume Next
    If gobjPlugIn.SaveRegisterAfter(glngSys, lngModule, strNO, bln预约) = False Then Exit Function
    If Err <> 0 Then
        Call zlPlugInErrH(Err, "SaveRegisterAfter")
        Set gobjPlugIn = Nothing
        Exit Function
    End If
    
    zlSaveRgstAfterByPlugIn = True
End Function

