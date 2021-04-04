Attribute VB_Name = "mdl泸州"
Option Explicit
'修改记录:
    '2004-05-27 ZYB 增加门诊医保和住院医保
    '1、费用分割()
    '2、基本统筹()
    '算法：如果原字段是否医保=0，当做非医保项目
    '如果是否医保=1，如果是门诊医保且当前是门诊交易，当做医保项目，住院同理

'一、IC卡函数所需结构定义
'1、基本结构:
'      1）病人信息结构       TIC泸州
'      2）IC卡就医信息结构   TBlockPayInfo    （或叫支付信息）
Public Type TIC泸州
    CenterCode       As String * 4      ' 中心代码
    Cardno           As String * 8      ' 卡号
    IDCardno         As String * 18     ' 身份证号 长度不足后补#0
    MediAccountNo    As String * 8      ' 医保号
    Name             As String * 10     ' 姓名
    Sex              As String * 1      ' 性别 1-男  0-女
    Birthday         As String * 8      ' 出生日期 YYYYMMDD
    UnitCode         As String * 5      ' 用人单位编码
    ClassCode        As String * 2      ' 职工身份：0x：在职1x：退休, 05和11为一次性缴费
    DomainCode       As String * 1      ' 职工属地 0-正常 1-常驻外地 2-异地安置
    MediYear         As String * 4      ' 医保年度
    InNo             As Long            ' 装钱期次
    OutSerialNo      As Long            ' 支付顺序号
    InPerAcc         As Double          ' 个人帐户累计注入金额
    OutPerAcc        As Double          ' 个人帐户累计支出金额
    PlanPaidFee      As Double          ' 统筹基金支付费用累计（基本+补充）
    PlanPaidAmt      As Double          ' 统筹基金支付金额累计（基本+补充）
    ChronicPaidFee   As Double          ' 慢性病支付费用累计
    ChronicPaidAmt   As Double          ' 慢性病支付金额累计
    InHosPaidAmt     As Double          ' 住院个人帐户支付金额
    ClinicPaidAmt    As Double          ' 门诊个人帐户支付金额
    Password         As String * 4      ' 个人密码
    InHosTimes       As Long            ' 本年有效住院次数
    IsOffical        As String * 1      ' 公务员 0-否；其他-是
    IsAttend         As String * 1      ' 医疗照顾对象 0-否；1-是
    InpatientFlag    As String * 1      ' 住院标志 0-不住院 1-住院
    Reserved         As String * 2      ' 保留不使用。主要为了能与DLL正常传递数据
    QuotaPaidAmt     As Double          ' 慢性病额度已支付金额
    ChronicSillPaidAmt    As Double     ' 慢性病起付金已支付金额
End Type

Private Type TPayInfo
    OccurDate        As String * 8 '  就医日期
    HospitalCode     As String * 4 '  医疗机构代码
    Tail             As String * 4
    Amount           As Double     '  本次费用合计
    AccPay           As Double     '  个人帐户支付
    CdFlag           As Long
End Type
Private Type TBlockPayInfo
    First            As TPayInfo   ' 第一次就医信息
    Second           As TPayInfo   ' 第二次就医信息
    Third            As TPayInfo   ' 第三次就医信息
End Type
Private Type TInMoneyParameter
    CenterCode       As String * 4 ' 中心代码
    Cardno           As String * 8 ' 卡号
    MediYear         As String * 4 ' 医保年度
    InNo             As Long       ' 装钱期次
    InPerAcc         As Double     ' 个人帐户累计注入金额
    InExAcc          As Double     ' 补充帐累计注入金额
    InSubAcc         As Double     ' 补助帐户累计注入金额
End Type
'二、IC卡读写函数定义说明

'2、基本读写
'      1）读IC卡病人信息
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC泸州 As TIC泸州) As Long
'      2）写IC卡病人信息
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC泸州 As TIC泸州) As Long

'记录住院情况
Private Declare Function ReadICCardPayInfo Lib "ICREAD.DLL" (BlockPayInfo As TBlockPayInfo) As Long
Private Declare Function WriteICCardPayInfo Lib "ICWRITE.DLL" (ByVal strCardNO As String, iIC泸州 As TPayInfo) As Long

'完成在线装钱
'Modified By 朱玉宝 2003-12-10 地区：泸州 参数增加
Private Declare Function OnLineInMoney Lib "InMoneyOnLine.dll" (ByVal IC_CenterCode As String, ByVal IC_CardNo As String, ByVal IC_MediYear As String, ByVal HosCode As String, ByVal serverIP As String) As Long

Private Enum card医保灰度
    deg停止支付 = 1
    deg上传明细 = 2 '也停止支持
    deg个人支付 = 3 '可用个人帐户支付，统筹停----改为：允许统筹支付，但不允许办理入院
    deg医保支付 = 4 '
    deg正常支付 = 5 '不下发
End Enum

Private Type 泸州结算数据  '本结构中的变量都是与本次结算有关，至于那些累计值，基本上都要求从卡上取
    灰度         As card医保灰度
    病人ID       As Long
    主页ID         As Long
    中心序号     As Long
    年度         As Long
    跨年住院     As Boolean
    跨年结算     As Boolean
    住院次数       As Long
    住院次数增加   As Long
    中途结帐       As Long
    起付线         As Currency
    封顶线         As Currency
    实际起付线     As Currency  '本次实际支付的起付线金额
    本次起付线     As Currency  '本次预计会支付的起线金额
    发生费用       As Currency
    全自费         As Currency
    首先自付       As Currency
    进入统筹       As Currency
    医保项目金额   As Currency
    乙类项目金额   As Currency
    个人帐户支付   As Currency
    住院床日       As Long
    统筹已支付金额 As Currency   '这两个变量不能省。有时不能从卡上去取这两个值，比如不使用累计的中途结算，就要从数据库中以前的结算记录汇总
    统筹已支付费用 As Currency
    进入统筹支付   As Currency
    进入统筹费用   As Currency
    进入慢病支付   As Currency
    进入慢病费用   As Currency
    统筹基金支付   As Currency
    统筹基金自付   As Currency
    参加补充保险   As Long
    补充基金支付   As Currency
    补充基金自付   As Currency
    补助基金支付   As Currency
    补助基金自付   As Currency
    超基本封顶线   As Currency
    超补充封顶线   As Currency
    进入额度支付   As Currency
    进入门诊个人帐户支付  As Currency
    进入住院个人帐户支付  As Currency
    进入慢性病起付金  As Currency
    病种代码 As String
    病种名称 As String
    病种类型 As String
End Type

Private Type 政策
    个人帐户支付全自费 As Boolean
    个人帐户支付首先自付 As Boolean
    个人帐户支付超限 As Boolean
    全额统筹 As Boolean
    费用封顶 As Boolean
    费用段值 As Boolean
    使用累计 As Boolean
    补充报销减起付金 As Boolean
    起付线在段中 As Boolean
    跨年起付金类型 As Long          '0-补原起付金；1补今年差价；2交起付金
    跨年增加住院次数 As Long        '1增加一次，0不增加
End Type
'-------------变量定义
Public gIC泸州 As TIC泸州                 '全局定义的存储IC卡信息的结构
Public gIC泸州Temp As TIC泸州             '主要用于与远程主机交换数据
Public gcn泸州 As New ADODB.Connection        '连接到医保前置服务器
Private m泸州 As 泸州结算数据
Private m政策 As 政策

'把这几个涉及报销的的变量作为全局

'-------------函数定义

Public Function 医保初始化_泸州() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    '因为经常要访问医保服务器，所以强制加上这个限制
    医保初始化_泸州 = 检查医保服务器_泸州
End Function

Public Function 身份标识_泸州(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；
'      bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回： 空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim strIdentify As String, strAddition As String
    Dim strBirthday As String, datToday As Date
    Dim str病种ID As String, lng病种ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If frmIdentify泸州.GetPatient(bytType, False, lng病种ID, str病种ID) = True Then
        '身份识别完成，返回病人信息
        With gIC泸州
            Call 医保灰度(.CenterCode, .Cardno)
            If m泸州.灰度 = deg停止支付 Then
                MsgBox "该病人暂时停止医保支付，请到医保中心处理。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bytType = 1 Then
                '对有限制的病人进行提醒
                If m泸州.灰度 = deg上传明细 Then
                    MsgBox "该病人不能使用统筹基金支付住院费用。", vbExclamation, gstrSysName
                End If
            End If
            
            '建立病人档案信息，传入格式：
            '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
            '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
            '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
            strIdentify = TrimStr(.Cardno)                              '0卡号
            strIdentify = strIdentify & ";" & TrimStr(.MediAccountNo)   '1医保号
            strIdentify = strIdentify & ";" & TrimStr(.Password)        '2密码
            strIdentify = strIdentify & ";" & TrimStr(.Name) '3姓名
            strIdentify = strIdentify & ";" & IIf(.Sex = "1", "男", "女")   '4性别
            
            strBirthday = TrimStr(.Birthday)
            datToday = zlDatabase.Currentdate
            If strBirthday = "" Then
                strBirthday = Format(datToday, "yyyy-MM-dd")
            Else
                strBirthday = Mid(strBirthday, 1, 4) & "-" & Mid(strBirthday, 5, 2) & "-" & Mid(strBirthday, 7, 2)
            End If
            strIdentify = strIdentify & ";" & strBirthday              '5出生日期
            strIdentify = strIdentify & ";" & TrimStr(.IDCardno)   '6身份证
            strIdentify = strIdentify & ";" & TrimStr(.UnitCode) & "(" & TrimStr(.UnitCode) & ")"  '7.单位名称(编码)
            
            '得到中心序号
            gstrSQL = "select 序号 from 保险中心目录 where 险类=[1] and 编码=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市, .CenterCode)
            
            If rsTemp.RecordCount = 0 Then
                身份标识_泸州 = ""
                MsgBox "该病人所属中心尚未建立，不能使用。", vbInformation, gstrSysName
                Exit Function
            Else
                m泸州.中心序号 = rsTemp("序号")
            End If
            
            '得到原住院病种
            If bytType <> 1 Then
                gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 险类=[1] And 医保号=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "得到原住院病种", TYPE_泸州市, CStr(TrimStr(.MediAccountNo)))
                If Not rsTemp.EOF Then
                    lng病种ID = rsTemp!病种ID
                End If
            End If
            
            strAddition = ";" & m泸州.中心序号                          '8.中心代码
            strAddition = strAddition & ";"                             '9.顺序号
            strAddition = strAddition & ";" & TrimStr(.ClassCode)       '10人员身份
            strAddition = strAddition & ";" & (.InPerAcc - .OutPerAcc)  '11帐户余额
            strAddition = strAddition & ";" & .InpatientFlag            '12当前状态
            strAddition = strAddition & ";" & IIf(lng病种ID > 0, lng病种ID, "") '13病种ID

'            strAddition = strAddition & ";" & IIf(Left(TrimStr(.ClassCode), 1) = "0", 1, 0)    '14在职
            Select Case Left(TrimStr(.ClassCode), 1)                    '14在职(1,2,3)
            Case "0"
                strAddition = strAddition & ";1"
            Case "1"
                strAddition = strAddition & ";2"
            Case "5"
                strAddition = strAddition & ";3"
            End Select
            strAddition = strAddition & ";"                             '15退休证号
            strAddition = strAddition & ";" & DateDiff("yyyy", CDate(strBirthday), datToday) '16年龄段
            strAddition = strAddition & ";" & m泸州.灰度                   '17灰度级
            strAddition = strAddition & ";" & .InPerAcc                 '18帐户增加累计
            strAddition = strAddition & ";" & .OutPerAcc                '19帐户支出累计
            strAddition = strAddition & ";" & .PlanPaidFee              '20进入统筹累计
            strAddition = strAddition & ";" & .PlanPaidAmt              '21统筹报销累计
            strAddition = strAddition & ";" & .InHosTimes               '22住院次数累计
            strAddition = strAddition & ";"                             '23就诊类型 (1、急诊门诊)
            
            lng病人ID = BuildPatiInfo(bytType, strIdentify & strAddition, lng病人ID, TYPE_泸州市)
            '返回格式:中间插入病人ID
            身份标识_泸州 = strIdentify & ";" & lng病人ID & strAddition
            
            '更新病种信息
            If bytType = 1 Then
                gstrSQL = "zlyb.zl_病种信息_INSERT(" & TYPE_泸州市 & "," & lng病人ID & ",0,0,'" & str病种ID & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
            End If
        End With
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_泸州(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'参数:
'返回: 返回个人帐户余额的金额
    
    On Error GoTo errHandle
    
    '执行装钱操作，顺便就读取了最新的个人数据
    If 装钱操作(lng病人ID) = True Then
        '检查黑名单
        Call 医保灰度(gIC泸州.CenterCode, gIC泸州.Cardno)
        If m泸州.灰度 > deg上传明细 Then
            '返回余额
            个人余额_泸州 = gIC泸州.InPerAcc - gIC泸州.OutPerAcc
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_泸州(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim cls医保 As New clsInsure, tmp泸州 As 泸州结算数据, tmp政策 As 政策
    Dim dbl全自费 As Currency, dbl首先自付 As Currency, dbl进入统筹 As Currency
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset
    
    m泸州 = tmp泸州         '初始化变量
    m政策 = tmp政策
    
    If rs明细.RecordCount = 0 Then
        MsgBox "没有费用，不能进行预结算。", vbInformation, gstrSysName
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    On Error GoTo errHandle
    
    'Modified By 朱玉宝 2003-12-10 地区：泸州
    If Calc费用分割(rs明细, False, dbl全自费, dbl首先自付, dbl进入统筹, False, True) = False Then
        Exit Function
    End If
    With m泸州
        .发生费用 = dbl全自费 + dbl进入统筹 + dbl首先自付
        .全自费 = dbl全自费
        .首先自付 = dbl首先自付
        .进入统筹 = dbl进入统筹
    End With
    
    gstrSQL = "Select B.编码 From 保险帐户 A,保险病种 B where A.险类=[1] And A.病人ID=[2] And A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊虚拟结算", TYPE_泸州市, lng病人ID)
    If rsTemp.EOF = False Then
        gstrSQL = "Select B.编码,名称,类别 From 保险病种 B where B.险类=" & TYPE_泸州市 & " And B.编码='" & rsTemp("编码") & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
        If rsTemp.EOF = False Then
            m泸州.病种代码 = rsTemp("编码")
            m泸州.病种名称 = Nvl(rsTemp("名称"))
            m泸州.病种类型 = Nvl(rsTemp("类别"))
        Else
            m泸州.病种代码 = ""
            m泸州.病种名称 = ""
            m泸州.病种类型 = ""
        End If
    End If
    
    '报销规定
    m政策.个人帐户支付全自费 = cls医保.GetCapability(support收费帐户全自费, 0, TYPE_泸州市)
    m政策.个人帐户支付首先自付 = cls医保.GetCapability(support收费帐户首先自付, 0, TYPE_泸州市)
    
    gstrSQL = "SELECT B.医保年,A.起付线在段中,A.段值类型,A.封顶类型,A.使用累计报销,A.个人账户可支付首先自付 " & _
               " FROM 保险中心目录 A,保险主机 B " & _
               " WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & gIC泸州.CenterCode & "' AND A.主机编码=B.编码 AND A.险类=B.险类 "
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        m政策.个人帐户支付首先自付 = Nvl(rsTemp("个人账户可支付首先自付")) = 1
    End If
    
    With m泸州
        .个人帐户支付 = 0
        If m政策.个人帐户支付全自费 = True Then
            .个人帐户支付 = dbl全自费
        End If
        
        If Is离休病人(lng病人ID) = True Then
            '部分费用可以用医保基金报销
            .进入统筹费用 = .进入统筹 + .首先自付
            If Is全额统筹(lng病人ID, TYPE_泸州市) = True Then
                '首先自付也是由医保基金支付
                .进入统筹支付 = .进入统筹 + .首先自付
            Else
                .进入统筹支付 = .进入统筹
                If m政策.个人帐户支付首先自付 = True Then
                    .个人帐户支付 = .个人帐户支付 + .首先自付
                End If
            End If
            .统筹基金支付 = .进入统筹支付
        Else
            '只能用个人帐户支付
            .个人帐户支付 = .个人帐户支付 + .进入统筹
            If m政策.个人帐户支付首先自付 = True Then
                .个人帐户支付 = .个人帐户支付 + dbl首先自付
            End If
        End If
        
        '检查帐户余额是否足够支付
        If .个人帐户支付 > gIC泸州.InPerAcc - gIC泸州.OutPerAcc Then
            .个人帐户支付 = gIC泸州.InPerAcc - gIC泸州.OutPerAcc
            If .个人帐户支付 < 0 Then .个人帐户支付 = 0
        End If
    End With
    
    '设置医保灰度
    Call 医保灰度(gIC泸州.CenterCode, gIC泸州.Cardno)
    If m泸州.灰度 < deg个人支付 Then m泸州.个人帐户支付 = 0
    
    str结算方式 = "个人帐户;" & m泸州.个人帐户支付 & ";1"
    If m泸州.灰度 >= deg医保支付 Then
        If m泸州.统筹基金支付 > 0 Then str结算方式 = str结算方式 & "|医保基金;" & m泸州.统筹基金支付 & ";0"
        If m泸州.补充基金支付 > 0 Then str结算方式 = str结算方式 & "|补充基金;" & m泸州.补充基金支付 & ";0"
    End If
    门诊虚拟结算_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_泸州(lng结帐ID As Long, cur个人帐户 As Currency, ByVal cur全自费 As Currency, ByVal cur首先自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rsTemp As New ADODB.Recordset, rs信息 As New ADODB.Recordset
    Dim ic门诊 As TIC泸州            '用上行结构，好象返回值有问题（主要是涉及金额的几个成员）
    Dim str医院编码 As String
    Dim lng年龄 As Long, lng病人ID As Long
    Dim cur统筹金额 As Currency
    Dim dat当前日期 As Date
    Dim bln离休 As Boolean
    
    On Error GoTo errHandle
        
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.收费细目ID,C.项目编码,B.编码,B.名称,A.实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价 " & _
              "  From 门诊费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where A.结帐ID=[1] And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0" & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= " & TYPE_泸州市 & _
              "  Order by A.病人ID,A.发生时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊结算", lng结帐ID)
    'Modified By 朱玉宝 2003-12-10 地区：泸州
    If Calc费用分割(rsTemp, True, cur全自费, cur首先自付, cur统筹金额, False, True) = False Then
        Exit Function
    End If
    With m泸州
        .全自费 = cur全自费
        .首先自付 = cur首先自付
        .进入统筹 = cur统筹金额
        .发生费用 = cur全自费 + cur首先自付 + cur统筹金额
    End With
    
    '读出可以填写到保险结算记录中的信息
    gstrSQL = "SELECT A.病人ID,A.NO,A.实际票号,A.记录性质,substr(B.姓名,1,8) as 姓名,substr(B.性别,1,2) as 性别,floor(MONTHS_BETWEEN(A.登记时间,B.出生日期)/12) AS 年龄" & _
              "         ,B.身份证号,C.卡号,C.医保号,a.登记时间,substr(A.操作员姓名,1,8) as 医生,D.名称 AS 部门" & _
              "  FROM 门诊费用记录 A,病人信息 B,保险帐户 C,部门表 D" & _
              "  Where A.结帐ID =[1] And A.病人ID = B.病人ID And B.病人ID = C.病人ID And C.险类 =[2] And A.开单部门ID = D.ID(+) and rownum<2"
    Set rs信息 = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID, TYPE_泸州市)
    lng病人ID = rs信息("病人ID")
    
    If ReadIC(lng病人ID, 0, False, "收费时读卡失败。", ic门诊, bln离休) = False Then
        Exit Function
    End If
    
    Call 医保灰度(ic门诊.CenterCode, ic门诊.Cardno)
    
    If m泸州.灰度 = deg停止支付 Then
        '不用再处理后续过程
        门诊结算_泸州 = True
        Exit Function
    End If
    
    dat当前日期 = zlDatabase.Currentdate
    
    '判断该病人的卡是否插入正确
    If 检查IC卡(lng病人ID, TrimStr(ic门诊.Cardno), TrimStr(ic门诊.CenterCode)) = False Then Exit Function
    
    With ic门诊
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_泸州市 & "," & Format(dat当前日期, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc + cur个人帐户 & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
        
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_泸州市 & "," & lng病人ID & "," & _
            Format(dat当前日期, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ",0,0,0," & _
            m泸州.发生费用 & "," & cur全自费 & "," & cur首先自付 & "," & cur统筹金额 & ",0,0,0," & _
            cur个人帐户 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    End With
    
    '医保服务器的事务虽然不能与主程序在一个事务中，但起码要与写卡在一起
    With ic门诊
        gstrSQL = "INSERT INTO 保险结算记录 " & _
                  "   (性质,记录id,险类,年度,中心代码,序号,病人id,主页id,姓名,性别,年龄,医保号,卡号,身份证号,身份代码,单位医保号 " & _
                  ",是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出,统筹已支付金额,统筹已支付费用 " & _
                  ",慢病已支付金额,慢病已支付费用,慢病起付金已支付,备案日期,增加住院次数 " & _
                  ",门诊个人帐户支付金额,住院个人帐户支付金额,额度已支付金额,部门名称,医生名称,病种代码,病种名称,病种类型 " & _
                  ",发生费用金额,个人帐户支付,全自付金额,首先自付金额,转外首先自付,起付线,封顶线,实际起付线 " & _
                  ",进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
                  ",补充基金支付,补充基金自付,补助基金支付,补助基金自付 " & _
                  ",超基本封顶线,超补充封顶线,进入额度支付,进入门诊个人帐户支付,进入住院个人帐户支付,进入慢性病起付金 " & _
                  ",卡灰度级,发票号,票据日期,冲票标志,被冲票据号,支付顺序号,是否上传) " & _
                  " Values "
         gstrSQL = gstrSQL & " (1," & lng结帐ID & "," & TYPE_泸州市 & "," & .MediYear & ",'" & .CenterCode & "','" & rs信息("NO") & "1" & rs信息!记录性质 & "'," & lng病人ID & ",0,'" & rs信息("姓名") & _
                  "','" & rs信息("性别") & "'," & rs信息("年龄") & ",'" & rs信息("医保号") & "','" & rs信息("卡号") & "','" & rs信息("身份证号") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & m泸州.参加补充保险 & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null," & m泸州.住院次数增加 & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs信息("部门") & "','" & rs信息("医生") & "','" & m泸州.病种代码 & "','" & m泸州.病种名称 & "','" & m泸州.病种类型 & "' " & _
                  "," & m泸州.发生费用 & "," & cur个人帐户 & "," & cur全自费 & "," & cur首先自付 & ",0,0,0,0 " & _
                  "," & m泸州.进入统筹支付 & "," & m泸州.进入统筹费用 & "," & m泸州.进入慢病支付 & "," & m泸州.进入慢病费用 & "," & _
                  (m泸州.统筹基金支付 + m泸州.补充基金支付 + m泸州.补助基金支付) & "," & (m泸州.统筹基金自付 + m泸州.补充基金自付 + m泸州.补助基金自付) & "," & m泸州.统筹基金支付 & "," & m泸州.统筹基金自付 & " " & _
                  "," & m泸州.补充基金支付 & "," & m泸州.补充基金自付 & ",0,0 " & _
                  "," & m泸州.超基本封顶线 & "," & m泸州.超补充封顶线 & "," & m泸州.进入额度支付 & "," & cur个人帐户 & ",0," & m泸州.进入慢性病起付金 & " " & _
                  "," & m泸州.灰度 & ",'" & Nvl(rs信息("实际票号"), " ") & "'," & GetOracleFormat(rs信息("登记时间")) & ",1,'','" & .OutSerialNo + 1 & "',0)"
        '准备写入的卡信息
        .OutPerAcc = .OutPerAcc + cur个人帐户                   '个人帐户累计支出金额
        .ClinicPaidAmt = .ClinicPaidAmt + cur个人帐户           '门诊个人帐户支出金额
        .InHosTimes = .InHosTimes + m泸州.住院次数增加          '有些慢特病会增加住院次数
        .PlanPaidFee = .PlanPaidFee + m泸州.进入统筹费用        '统筹基金支付费用累计（基本+补充）
        .PlanPaidAmt = .PlanPaidAmt + m泸州.进入统筹支付        ' 统筹基金支付金额累计（基本+补充）
        .ChronicPaidFee = .ChronicPaidFee + m泸州.进入慢病费用                 '慢性病支付费用累计
        .ChronicPaidAmt = .ChronicPaidAmt + m泸州.进入慢病支付                 '慢性病支付金额累计
        .QuotaPaidAmt = .QuotaPaidAmt + m泸州.进入额度支付                     '慢性病额度已支付金额
        .ChronicSillPaidAmt = .ChronicSillPaidAmt + m泸州.进入慢性病起付金     '慢性病起付金已支付金额
        .OutSerialNo = .OutSerialNo + 1           ' 支付顺序号
    End With
        
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
        .AccPay = m泸州.个人帐户支付
        .Amount = m泸州.发生费用
        .CdFlag = 1
    End With
    
    '完成卡写入
    Dim str数据体 As String
    With m泸州
        str数据体 = ic门诊.CenterCode & "|" & gstr医院编码 & "|0|" & rs信息("NO") & "1|" & _
                    TrimStr(ic门诊.MediAccountNo) & "|" & cur个人帐户 & "|" & .统筹基金支付 & "|" & .补充基金支付 & "|" & _
                    .进入统筹费用 & "|" & .进入统筹支付 & "|" & .住院次数增加 & "|" & .超基本封顶线 & "|1"
    End With
    
    If WriteIC(bln离休, True, 0, gstrSQL, ic门诊, payLog, str数据体) = False Then
        Exit Function
    End If
    
    门诊结算_泸州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_泸州(lng结帐ID As Long, cur个人帐户 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rsTemp As New ADODB.Recordset, rs结算 As New ADODB.Recordset
    Dim ic门诊 As TIC泸州
    Dim lng序号 As Long, lng病人ID As Long
    Dim dat当前日期 As Date
    Dim bln离休 As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "Select *  From 保险结算记录 Where 记录ID=" & lng结帐ID
    rs结算.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    
    lng病人ID = rs结算("病人ID")
        
    If ReadIC(lng病人ID, 0, True, "退费时读卡失败。", ic门诊, bln离休) = False Then
        Exit Function
    End If
    
    '取中心序号
    gstrSQL = "Select 序号 From 保险中心目录 Where 编码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取中心序号", ic门诊.CenterCode)
    m泸州.中心序号 = rsTemp!序号
    
    If Val(ic门诊.MediYear) <> rs结算("年度") Then
        Err.Raise 9000, gstrSysName, "跨年不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get当前医保年) <> rs结算("年度") Then
        Err.Raise 9000, gstrSysName, "跨年不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call 医保灰度(ic门诊.CenterCode, ic门诊.Cardno)
    
    If m泸州.灰度 = deg停止支付 Then
        '不用再处理后续过程
        '门诊结算冲销_泸州 = True
        Exit Function
    End If
    
    dat当前日期 = zlDatabase.Currentdate
        
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID)
    
    lng序号 = rsTemp("结帐ID")
    
    With ic门诊
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_泸州市 & "," & Format(dat当前日期, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc - cur个人帐户 & "," & .PlanPaidFee - rs结算("进入统筹费用") & "," & _
            .PlanPaidAmt - rs结算("进入统筹支付") & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
        
        gstrSQL = "zl_保险结算记录_insert(1," & lng序号 & "," & TYPE_泸州市 & "," & lng病人ID & "," & _
            Format(dat当前日期, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ",0,0,0," & _
            rs结算("发生费用金额") * -1 & "," & rs结算("全自付金额") * -1 & "," & rs结算("首先自付金额") * -1 & "," & rs结算("进入统筹费用") * -1 & ",0,0,0," & cur个人帐户 * -1 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    End With
    
    '医保服务器的事务虽然不能与主程序在一个事务中，但起码要与写卡在一起
    With ic门诊
        gstrSQL = "INSERT INTO 保险结算记录 " & _
                  "   (性质,记录id,险类,年度,中心代码,序号,病人id,主页id,姓名,性别,年龄,医保号,卡号,身份证号,身份代码,单位医保号 " & _
                  ",是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出,统筹已支付金额,统筹已支付费用 " & _
                  ",慢病已支付金额,慢病已支付费用,慢病起付金已支付,备案日期,增加住院次数 " & _
                  ",门诊个人帐户支付金额,住院个人帐户支付金额,额度已支付金额,部门名称,医生名称,病种代码,病种名称,病种类型 " & _
                  ",发生费用金额,个人帐户支付,全自付金额,首先自付金额,转外首先自付,起付线,封顶线,实际起付线 " & _
                  ",进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
                  ",补充基金支付,补充基金自付,补助基金支付,补助基金自付 " & _
                  ",超基本封顶线,超补充封顶线,进入额度支付,进入门诊个人帐户支付,进入住院个人帐户支付,进入慢性病起付金 " & _
                  ",卡灰度级,发票号,票据日期,冲票标志,被冲票据号,支付顺序号,是否上传) " & _
                  " Values "
         gstrSQL = gstrSQL & " (1," & lng序号 & "," & TYPE_泸州市 & "," & .MediYear & ",'" & .CenterCode & "','" & Mid(rs结算("序号"), 1, Len(rs结算("序号")) - 2) & "2" & Right(rs结算!序号, 1) & "'," & lng病人ID & ",0,'" & rs结算("姓名") & _
                  "','" & rs结算("性别") & "'," & rs结算("年龄") & ",'" & rs结算("医保号") & "','" & rs结算("卡号") & "','" & rs结算("身份证号") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & rs结算("参加补充保险") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null," & rs结算("增加住院次数") & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs结算("部门名称") & "','" & rs结算("医生名称") & "','" & rs结算("病种代码") & "','" & rs结算("病种名称") & "','" & rs结算("病种类型") & "' " & _
                  "," & rs结算("发生费用金额") & "," & cur个人帐户 & "," & rs结算("全自付金额") & "," & rs结算("首先自付金额") & ",0,0,0,0 " & _
                  "," & rs结算("进入统筹支付") & "," & rs结算("进入统筹费用") & "," & rs结算("进入慢病支付") & "," & rs结算("进入慢病费用") & "," & rs结算("统筹总支付") & "," & rs结算("统筹总自付") & "," & rs结算("统筹基金支付") & "," & rs结算("统筹基金自付") & " " & _
                  "," & rs结算("补充基金支付") & "," & rs结算("补充基金自付") & ",0,0 " & _
                  "," & rs结算("超基本封顶线") & "," & rs结算("超补充封顶线") & "," & rs结算("进入额度支付") & "," & rs结算("进入门诊个人帐户支付") & "," & rs结算("进入住院个人帐户支付") & "," & rs结算("进入慢性病起付金") & " " & _
                  "," & m泸州.灰度 & ",'" & Nvl(rs结算("发票号")) & "',sysdate,-1,'" & rs结算("序号") & "','" & .OutSerialNo + 1 & "',0)"
        '准备写卡
        .OutPerAcc = .OutPerAcc - cur个人帐户                  '个人帐户累计支出金额
        .ClinicPaidAmt = .ClinicPaidAmt - cur个人帐户           '门诊个人帐户支出金额
        .InHosTimes = .InHosTimes - rs结算("增加住院次数")      '有些慢特病会增加住院次数
        .PlanPaidFee = .PlanPaidFee - rs结算("进入统筹费用")      '统筹基金支付费用累计（基本+补充）
        .PlanPaidAmt = .PlanPaidAmt - rs结算("进入统筹支付")        ' 统筹基金支付金额累计（基本+补充）
        .ChronicPaidFee = .ChronicPaidFee - rs结算("进入慢病费用")                '慢性病支付费用累计
        .ChronicPaidAmt = .ChronicPaidAmt - rs结算("进入慢病支付")                '慢性病支付金额累计
        .QuotaPaidAmt = .QuotaPaidAmt - rs结算("进入额度支付")                     '慢性病额度已支付金额
        .ChronicSillPaidAmt = .ChronicSillPaidAmt - rs结算("进入慢性病起付金")      '慢性病起付金已支付金额
        .OutSerialNo = .OutSerialNo + 1           ' 支付顺序号
    End With
        
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
        .AccPay = cur个人帐户
        .Amount = rs结算("发生费用金额")
        .CdFlag = 0
    End With
    
    '完成卡写入
    Dim str数据体 As String
        
    str数据体 = ic门诊.CenterCode & "|" & gstr医院编码 & "|0|" & Mid(rs结算("序号"), 1, Len(rs结算("序号")) - 1) & "2|" & _
                TrimStr(ic门诊.MediAccountNo) & "|" & cur个人帐户 & "|" & rs结算("统筹基金支付") & "|" & rs结算("补充基金支付") & "|" & _
                rs结算("进入统筹费用") & "|" & rs结算("进入统筹支付") & "|" & rs结算("增加住院次数") & "|" & rs结算("超基本封顶线") & "|-1"
    
    
    If WriteIC(bln离休, True, 0, gstrSQL, ic门诊, payLog, str数据体) = False Then
        Exit Function
    End If
    
    门诊结算冲销_泸州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 个人帐户转预交_泸州(lng预交ID As Long, curMoney As Currency) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
           
    '由于泸州医保不支持该业务，所以强行返回失败
    个人帐户转预交_泸州 = False
End Function

Public Function 入院登记_泸州(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim ic入院 As TIC泸州       '入院登记读出结构
    Dim dat当前日期 As Date
    Dim bln离休 As Boolean
    
    On Error GoTo errHandle
    
    If ReadIC(lng病人ID, 1, True, "入院登记时读卡失败。", ic入院, bln离休) = False Then
        Exit Function
    End If
        
    dat当前日期 = zlDatabase.Currentdate
    
    Call 医保灰度(ic入院.CenterCode, ic入院.Cardno)
    
    If m泸州.灰度 = deg停止支付 Then
        '不用再处理后续过程
        入院登记_泸州 = False
        MsgBox "该病人已经停止医保支付，不能作为医保病人入院。", vbInformation, gstrSysName
        Exit Function
    End If
    If m泸州.灰度 = deg个人支付 Then
        '不用再处理后续过程
        入院登记_泸州 = False
        MsgBox "该病人已经停止医保支付（灰度为3），不能作为医保病人入院。", vbInformation, gstrSysName
        Exit Function
    End If

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_泸州市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    
    'Modified by 朱玉宝 2004-01-07 将当前医保年写入保险帐户
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_泸州市 & ",'医保年','''" & Get当前医保年 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    
    gstrSQL = "zlyb.zl_病种信息_入院(" & TYPE_泸州市 & "," & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将身份验证处录入的病种信息更新为本次的入院病种")
    
    Dim payLog As TPayInfo
    '完成卡写入
    With ic入院
        .InpatientFlag = 1
    End With
    If WriteIC(bln离休, False, 1, "", ic入院, payLog, "") = False Then
        Exit Function
    End If
    
    入院登记_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_泸州(ByVal lng病人ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'返回：交易成功返回true；否则，返回false
    Dim ic出院 As TIC泸州
    Dim bln离休 As Boolean
    
    On Error GoTo errHandle
    
    If ReadIC(lng病人ID, 1, True, "出院办理时读卡失败。", ic出院, bln离休) = False Then
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_泸州市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    
    Dim payLog As TPayInfo
    '完成卡写入
    With ic出院
        .InpatientFlag = 0
    End With
    If WriteIC(bln离休, False, 1, "", ic出院, payLog, "") = False Then
        Exit Function
    End If
        
    出院登记_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_泸州(rsExse As Recordset) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合
'      NO、序号、医保项目编码、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,收费类别
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim ic病人 As TIC泸州, tmp泸州 As 泸州结算数据, tmp政策 As 政策
    Dim cur全自费 As Currency, cur首先自付 As Currency, cur统筹 As Currency
    Dim bln离休 As Boolean
    Dim str入院年度 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If 检查医保服务器_泸州 = False Then
        '不能连接到前置服务器，就认为不可使用
        Exit Function
    End If
    
    gIC泸州 = ic病人 '如此可以进行所有内部变量的初始化
    m泸州 = tmp泸州
    m政策 = tmp政策
    
    If ReadIC(rsExse("病人ID"), 1, True, "读卡信息失败。", gIC泸州, bln离休) = False Then
        Exit Function
    End If
        
    '完成一些数据的初始化，黑名单人员也要使用的数据
    With m泸州
        .病人ID = rsExse("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", CLng(rsExse("病人ID")))
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
        g结算数据.主页ID = rsTemp("主页ID")
    
        '避免在出院结帐后再次进行结帐
        gstrSQL = "SELECT 病人ID FROM 保险结算记录 WHERE 中途结帐=0 AND 病人ID=[1] AND 主页ID=[2] AND 险类=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", .病人ID, .主页ID, TYPE_泸州市)
        
        If rsTemp.RecordCount > 0 Then
            MsgBox "病人已经进行过住院结算，不能再进行结帐操作。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    '检查病人的费用是否都已经重新计算过报销金额
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码,B.编码,B.名称,A.实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND A.实收金额 IS NOT NULL " & _
              "  And A.结帐ID Is NULL And nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "  Order by A.发生时间,A.NO,A.记录性质,Decode(A.记录状态,3,1,1,1,2),A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", m泸州.病人ID, m泸州.主页ID, TYPE_泸州市)
    If rsTemp.EOF = False Then
        '还存在没有分割费用的明细
        If Calc费用分割(rsTemp, True, cur全自费, cur首先自付, cur统筹) = False Then
            Exit Function
        End If
    End If
    
    '目前只是泸州医保使用该参数
    'Modified by 朱玉宝 2004-01-07
    gstrSQL = "select A.病种ID,A.医保年,B.编码,B.序号 " & _
            " from 保险帐户 A,保险中心目录 B " & _
            " where A.病人ID=[1] and A.险类=[2]" & _
            "  and A.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", m泸州.病人ID, TYPE_泸州市)
    If rsTemp.EOF = True Then
        MsgBox "请系统管理员完成医保中心的设置。", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTemp!病种ID, 0) = 0 Then
        MsgBox "没有选择病种，不允许结帐！", vbInformation, gstrSysName
        Exit Function
    End If
    
    m泸州.中心序号 = rsTemp("序号")
    'Modified by 朱玉宝 2004-01-07
    str入院年度 = Nvl(rsTemp!医保年)
    
    gstrSQL = "Select B.编码 From 保险帐户 A,保险病种 B where A.险类=[1] And A.病人ID=[2] And A.病种ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "住院虚拟结算", TYPE_泸州市, m泸州.病人ID)
    If rsTemp.EOF = False Then
        gstrSQL = "Select B.编码,名称,类别 From 保险病种 B where B.险类=" & TYPE_泸州市 & " And B.编码='" & rsTemp("编码") & "'"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
        m泸州.病种代码 = rsTemp("编码")
        m泸州.病种名称 = Nvl(rsTemp("名称"))
        m泸州.病种类型 = Nvl(rsTemp("类别"))
    End If
    
    '1.2 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期,sysdate 当前日期 " & _
              "from 病案主页 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", m泸州.病人ID, m泸州.主页ID)
    
    With m泸州
        If rsTemp("出院日期") = CDate("3000-01-01") Then
            .中途结帐 = 1
        Else
            '表示该病人已经出院
            .中途结帐 = 0
        End If
        'Modified By 朱玉宝 2003-12-10 地区：泸州
        .年度 = Get当前医保年
        'Modified by 朱玉宝 2004-01-07
        If str入院年度 = "" Then str入院年度 = Format(rsTemp!入院日期, "yyyy")
        If str入院年度 <> .年度 Then
            .跨年住院 = True '会影响起付线的值，以及是否增加住院次数
            
            '可能是跨年的第一次结算
            gstrSQL = "Select 年度 From 保险结算记录 Where 性质=2 and 险类=" & TYPE_泸州市 & _
                " And 病人ID=" & m泸州.病人ID & " And 主页ID=" & m泸州.主页ID & " And 年度=" & m泸州.年度
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn泸州
            
            If rsTemp.EOF = True Then
                .跨年结算 = True  '绝对要把累计金额全部清掉
            End If
        End If
    End With
        
    '此处使用装钱操作，主要目的是初始化病人的卡上的余额，以及累计进入统筹和统筹累计报销
    If 装钱操作(m泸州.病人ID) = False Then
        MsgBox "病人装钱操作失败，无法准确得到病人的余额与累计报销金额。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gIC泸州
        gstrSQL = "zl_帐户年度信息_insert(" & m泸州.病人ID & "," & TYPE_泸州市 & "," & .MediYear & "," & _
            .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidFee & "," & _
            .PlanPaidAmt & "," & .InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    End With
    
    Call 医保灰度(gIC泸州.CenterCode, gIC泸州.Cardno)
    
    If Calc基本统筹() = False Then
        Exit Function
    End If
    
    If m泸州.灰度 >= deg个人支付 Then
        With m泸州
            '现在这两个需要累计的数值出来了
            .进入统筹支付 = .统筹基金支付
            '进入统筹费用不能大于封顶线
            'takecare
            If m政策.费用封顶 = True Then
                .进入统筹费用 = .发生费用 - .全自费 - .超基本封顶线
                If .进入统筹费用 > (.封顶线 - .统筹已支付费用) Then .进入统筹费用 = (.封顶线 - .统筹已支付费用)
                If .实际起付线 > .进入统筹费用 Then .实际起付线 = .进入统筹费用
            Else
                .进入统筹费用 = .统筹基金支付 + .统筹基金自付
            End If
        
            If Calc补充报销() = False Then
                Exit Function
            End If
            
            If gIC泸州.IsOffical = 1 Then '公务员才进行补助报销
                If Calc补助报销() = False Then
                    Exit Function
                End If
            End If
            
            If m政策.全额统筹 = True Then
                住院虚拟结算_泸州 = "医保基金;" & .进入统筹 + .首先自付 & ";0"
            Else
                住院虚拟结算_泸州 = "医保基金;" & .统筹基金支付 & ";0"
                If .补充基金支付 > 0 Then
                    住院虚拟结算_泸州 = 住院虚拟结算_泸州 & "|补充基金;" & .补充基金支付 & ";0"
                End If
                If .补助基金支付 > 0 Then
                    住院虚拟结算_泸州 = 住院虚拟结算_泸州 & "|补助基金;" & .补助基金支付 & ";0"
                End If
            End If
        End With
    End If
'
    '还需要考虑个人帐户的支付范围
    '超过封顶线,则在Calc基本统筹()中计算的首先自付比实际的首先自付还多了,需要重新计算
    With m泸州
        If .灰度 >= deg个人支付 Then
            Dim dbl个人自付 As Double, dbl总自付 As Double, dbl首先自付 As Double '总自付=个人自付+首先自付+全自费
            dbl总自付 = .发生费用 - .统筹基金支付 - .补充基金支付
            dbl个人自付 = .进入统筹费用 - .统筹基金支付 - .补充基金支付
            '因个人自付中含首先自付，需要重新计算
            'takecare
            If m政策.费用封顶 Then
                If .进入统筹费用 > .医保项目金额 Then
                    dbl首先自付 = (.进入统筹费用 - .医保项目金额) * 0.2
                Else
                    dbl首先自付 = 0
                End If
                .首先自付 = dbl首先自付
            End If
            
            If m政策.费用封顶 = True Then
                dbl个人自付 = dbl个人自付 - .首先自付
            Else
                '支付封顶模式下，由于.进入统筹费用未包含起付线，因此，本处需要加上实际起付线
                dbl个人自付 = dbl个人自付 + .实际起付线
            End If
            If dbl个人自付 <= 0 Then dbl个人自付 = 0
            
            .个人帐户支付 = dbl个人自付
            If m政策.个人帐户支付首先自付 = True Then
                .个人帐户支付 = .个人帐户支付 + .首先自付
            End If
    
            If m政策.个人帐户支付全自费 = True Then
                .个人帐户支付 = .个人帐户支付 + .全自费
            End If
     
            '检查帐户余额是否足够支付
            If m泸州.个人帐户支付 > gIC泸州.InPerAcc - gIC泸州.OutPerAcc Then
                m泸州.个人帐户支付 = gIC泸州.InPerAcc - gIC泸州.OutPerAcc
            End If
            If m泸州.个人帐户支付 < 0 Then m泸州.个人帐户支付 = 0
   
            住院虚拟结算_泸州 = 住院虚拟结算_泸州 & IIf(住院虚拟结算_泸州 = "", "", "|") & "个人帐户;" & .个人帐户支付 & ";1"
        End If
    End With
    
    If 住院虚拟结算_泸州 = "" Then 住院虚拟结算_泸州 = "个人帐户;0;1"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_泸州(lng结帐ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID     病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim ic泸州 As TIC泸州               '住院结算读出结构
    Dim bln离休 As Boolean
    Dim rs信息 As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim var结算计算 As Variant, lng段数 As Long, str分段 As String
    
    On Error GoTo errHandle
    
    '读出可以填写到保险结算记录中的信息
    gstrSQL = "SELECT A.病人ID,A.NO,A.实际票号,substr(B.姓名,1,8) as 姓名,substr(B.性别,1,2) as 性别,floor(MONTHS_BETWEEN(A.收费时间,B.出生日期)/12) AS 年龄" & _
              "         ,B.身份证号,C.卡号,C.医保号,A.收费时间,substr(A.操作员姓名,1,8) as 医生" & _
              "," & IIf(m泸州.中途结帐 = 1, "A.开始日期", "D.入院日期") & " AS 入院日期," & IIf(m泸州.中途结帐 = 1, "A.结束日期", "D.出院日期") & " AS 出院日期 " & _
              "  FROM 病人结帐记录 A,病人信息 B,保险帐户 C,病案主页 D" & _
              "  Where A.ID =[1] And A.病人ID = B.病人ID And B.病人ID = C.病人ID And C.险类 =[2]" & _
             "         And B.病人ID=D.病人ID And D.主页ID=[3]"
    Set rs信息 = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID, TYPE_泸州市, m泸州.主页ID)
    '只要是中途结算，就要多加1天
    m泸州.住院床日 = Fix(CDate(Format(rs信息("出院日期"), "yyyy-MM-dd")) - _
                         CDate(Format(rs信息("入院日期"), "yyyy-MM-dd"))) + m泸州.中途结帐
    If m泸州.住院床日 <= 0 Then m泸州.住院床日 = 1
    
    If ReadIC(rs信息("病人ID"), 1, True, "结算时读卡失败。", ic泸州, bln离休) = False Then
        Exit Function
    End If
    
    Call 医保灰度(ic泸州.CenterCode, ic泸州.Cardno)
    
'    If m泸州.灰度 = deg停止支付 Then
'        '不用再处理后续过程
'        住院结算_泸州 = True
'        Exit Function
'    End If
        
    '求个人帐户支付金额
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID)
    
    If Not rsTemp.EOF Then
        m泸州.个人帐户支付 = rsTemp!金额
    Else
        m泸州.个人帐户支付 = 0
    End If
    
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    With m泸州
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & .病人ID & "," & TYPE_泸州市 & "," & .年度 & "," & _
            ic泸州.InPerAcc & "," & ic泸州.OutPerAcc + .个人帐户支付 & "," & ic泸州.PlanPaidFee + .进入统筹费用 & "," & _
            ic泸州.PlanPaidAmt + .进入统筹支付 & "," & ic泸州.InHosTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_泸州市 & "," & .病人ID & "," & _
            .年度 & "," & ic泸州.InPerAcc & "," & ic泸州.OutPerAcc & "," & ic泸州.PlanPaidFee & "," & _
            ic泸州.PlanPaidAmt & "," & ic泸州.InHosTimes & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用 & "," & .全自费 & "," & .首先自付 & "," & .进入统筹费用 & "," & .进入统筹支付 & ",0," & _
            .超基本封顶线 & "," & .个人帐户支付 & ",'" & ic泸州.OutSerialNo + 1 & "'," & .主页ID & "," & .中途结帐 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
        
        For Each var结算计算 In gcol结算计算
            '依次为档次、进入统筹金额、统筹报销金额、比例
            gstrSQL = "zl_保险结算计算_Insert(" & lng结帐ID & "," & _
                var结算计算(0) & "," & var结算计算(1) & "," & var结算计算(2) & "," & var结算计算(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
            
            lng段数 = lng段数 + 1
            If lng段数 <= 5 Then
                str分段 = str分段 & "," & var结算计算(2) & "," & IIf(m泸州.灰度 < deg个人支付, 0, (var结算计算(1) - var结算计算(2)))
            End If
        Next
        '补足五段
        For lng段数 = lng段数 + 1 To 5
            str分段 = str分段 & ",0,0"
        Next
    End With
    
    '医保服务器的事务虽然不能与主程序在一个事务中，但起码要与写卡在一起
    With ic泸州
        gstrSQL = "INSERT INTO 保险结算记录 " & _
                "   (性质,记录id,险类,年度,中心代码,序号,病人id,主页id,姓名,性别,年龄,医保号,卡号,身份证号,身份代码,单位医保号 " & _
                ",是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出,统筹已支付金额,统筹已支付费用 " & _
                ",慢病已支付金额,慢病已支付费用,慢病起付金已支付,备案日期 " & _
                ",门诊个人帐户支付金额,住院个人帐户支付金额,额度已支付金额,部门名称,医生名称,病种代码,病种名称,病种类型 " & _
                ",住院次数,增加住院次数,治愈情况,入院日期,出院日期,住院天数 " & _
                ",发生费用金额,个人帐户支付,全自付金额,首先自付金额,转外首先自付,起付线,封顶线,实际起付线 " & _
                ",进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
                ",补充基金支付,补充基金自付,补助基金支付,补助基金自付 " & _
                ",第一段支付,第一段自付,第二段支付,第二段自付,第三段支付,第三段自付,第四段支付,第四段自付,第五段支付,第五段自付 " & _
                ",超基本封顶线,超补充封顶线,进入额度支付,进入门诊个人帐户支付,进入住院个人帐户支付,进入慢性病起付金 " & _
                ",卡灰度级,发票号,票据日期,冲票标志,被冲票据号,支付顺序号,中途结帐,是否上传) " & _
                  " Values "
         gstrSQL = gstrSQL & " (2," & lng结帐ID & "," & TYPE_泸州市 & "," & .MediYear & ",'" & .CenterCode & "','" & rs信息("NO") & "1'," & m泸州.病人ID & "," & m泸州.主页ID & ",'" & rs信息("姓名") & _
                  "','" & rs信息("性别") & "'," & rs信息("年龄") & ",'" & rs信息("医保号") & "','" & rs信息("卡号") & "','" & rs信息("身份证号") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & m泸州.参加补充保险 & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null" & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & ToVarchar(UserInfo.部门, 20) & "','" & rs信息("医生") & "','" & m泸州.病种代码 & "','" & m泸州.病种名称 & "','" & m泸州.病种类型 & "' " & _
                  "," & ic泸州.InHosTimes & "," & m泸州.住院次数增加 & ",'0'," & GetOracleFormat(rs信息("入院日期")) & "," & GetOracleFormat(rs信息("出院日期")) & "," & m泸州.住院床日 & _
                  "," & m泸州.发生费用 & "," & m泸州.个人帐户支付 & "," & m泸州.全自费 & "," & m泸州.首先自付 & ",0," & m泸州.起付线 & "," & m泸州.封顶线 & "," & m泸州.实际起付线 & " " & _
                  "," & m泸州.进入统筹支付 & "," & m泸州.进入统筹费用 & "," & m泸州.进入慢病支付 & "," & m泸州.进入慢病费用 & "," & _
                  (m泸州.统筹基金支付 + m泸州.补充基金支付 + m泸州.补助基金支付) & "," & (m泸州.统筹基金自付 + m泸州.补充基金自付 + m泸州.补助基金自付) & "," & m泸州.统筹基金支付 & "," & m泸州.统筹基金自付 & " " & _
                  "," & m泸州.补充基金支付 & "," & m泸州.补充基金自付 & "," & m泸州.补助基金支付 & "," & m泸州.补助基金自付 & str分段 & _
                  "," & m泸州.超基本封顶线 & "," & m泸州.超补充封顶线 & "," & m泸州.进入额度支付 & ",0," & m泸州.个人帐户支付 & "," & m泸州.进入慢性病起付金 & " " & _
                  "," & m泸州.灰度 & ",'" & Nvl(rs信息("实际票号"), " ") & "'," & GetOracleFormat(rs信息("收费时间")) & ",1,'','" & .OutSerialNo + 1 & "'," & m泸州.中途结帐 & ",0)"
        '准备写卡
        .OutPerAcc = .OutPerAcc + m泸州.个人帐户支付                   '个人帐户累计支出金额
        .InHosPaidAmt = .InHosPaidAmt + m泸州.个人帐户支付             '住院个人帐户支出金额
        .InHosTimes = .InHosTimes + m泸州.住院次数增加                 '只有出院结算会增加住院次数
        .PlanPaidFee = .PlanPaidFee + m泸州.进入统筹费用        '统筹基金支付费用累计（基本+补充）
        .PlanPaidAmt = .PlanPaidAmt + m泸州.进入统筹支付        ' 统筹基金支付金额累计（基本+补充）
        .ChronicPaidFee = .ChronicPaidFee + m泸州.进入慢病费用                 '慢性病支付费用累计
        .ChronicPaidAmt = .ChronicPaidAmt + m泸州.进入慢病支付                 '慢性病支付金额累计
        .QuotaPaidAmt = .QuotaPaidAmt + m泸州.进入额度支付                     '慢性病额度已支付金额
        .ChronicSillPaidAmt = .ChronicSillPaidAmt + m泸州.进入慢性病起付金     '慢性病起付金已支付金额
        .OutSerialNo = .OutSerialNo + 1           ' 支付顺序号
    End With
    '记录住院情况。这一部分信息不是太重要，即使出错，也可以忽略，而不能回滚前一次写卡
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
        .AccPay = m泸州.个人帐户支付
        .Amount = m泸州.发生费用
        .CdFlag = 1
    End With
        
    '完成卡写入
    Dim str数据体 As String
    With m泸州
        str数据体 = ic泸州.CenterCode & "|" & gstr医院编码 & "|1|" & rs信息("NO") & "1|" & _
                    TrimStr(ic泸州.MediAccountNo) & "|" & m泸州.个人帐户支付 & "|" & .统筹基金支付 & "|" & .补充基金支付 & "|" & _
                    .进入统筹费用 & "|" & .进入统筹支付 & "|" & .住院次数增加 & "|" & IIf(.参加补充保险 = 1, .超补充封顶线, .超基本封顶线) & "|1"
    End With
    If WriteIC(bln离休, True, 1, gstrSQL, ic泸州, payLog, str数据体) = False Then
        Exit Function
    End If
    
    住院结算_泸州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_泸州(lng结帐ID As Long) As Boolean
'----------------------------------------------------------------
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID-需要作废的结帐单ID号；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'----------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rs结算 As New ADODB.Recordset, rs结算计算 As New ADODB.Recordset
    Dim ic住院 As TIC泸州                '住院结算读出结构
    Dim lng冲销ID As Long
    Dim bln离休 As Boolean
    Dim cur个人帐户 As Currency
    
    On Error GoTo errHandle
    
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rs结算 = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID)
    lng冲销ID = rs结算("ID") '冲销单据的ID
    rs结算.Close
    
    gstrSQL = "Select *  From 保险结算记录 Where 记录ID=" & lng结帐ID
    rs结算.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    
    If rs结算.RecordCount = 0 Then
        MsgBox "该病人的医保结算数据丢失，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    If Can住院结算冲销(rs结算("病人ID"), rs结算("主页ID")) = False Then Exit Function
    
    If ReadIC(rs结算("病人ID"), 1, True, "作废结算时读卡失败。", ic住院, bln离休) = False Then
        Exit Function
    End If
    
    '取中心序号
    gstrSQL = "Select 序号 From 保险中心目录 Where 编码= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取中心序号", ic住院.CenterCode)
    m泸州.中心序号 = rsTemp!序号
    
    If Val(ic住院.MediYear) <> rs结算("年度") Then
        Err.Raise 9000, gstrSysName, "跨年不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(Get当前医保年) <> rs结算("年度") Then
        Err.Raise 9000, gstrSysName, "跨年不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call 医保灰度(ic住院.CenterCode, ic住院.Cardno)
    
    If m泸州.灰度 = deg停止支付 Then
        '不用再处理后续过程
        住院结算冲销_泸州 = False
        Err.Raise 9000, gstrSysName, "该病人已经停止医保支付，不能进行冲销操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    '判断该病人的卡是否插入正确
    If 检查IC卡(rs结算("病人ID"), TrimStr(ic住院.Cardno), TrimStr(ic住院.CenterCode)) = False Then Exit Function
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    gstrSQL = "zl_帐户年度信息_insert(" & rs结算("病人ID") & "," & TYPE_泸州市 & "," & rs结算("年度") & "," & _
        ic住院.InPerAcc & "," & ic住院.OutPerAcc - rs结算("个人帐户支付") & "," & ic住院.PlanPaidFee - rs结算("进入统筹费用") & "," & _
        ic住院.PlanPaidAmt - rs结算("进入统筹支付") & "," & ic住院.InHosTimes & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    
    '冲销单据基本上是复制原单据
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_泸州市 & "," & rs结算("病人ID") & "," & _
        rs结算("年度") & "," & ic住院.InPerAcc & "," & ic住院.OutPerAcc & "," & ic住院.PlanPaidFee & "," & _
        ic住院.PlanPaidAmt & "," & ic住院.InHosTimes & "," & rs结算("起付线") * -1 & "," & rs结算("封顶线") & "," & rs结算("实际起付线") * -1 & "," & _
        rs结算("发生费用金额") * -1 & "," & rs结算("全自付金额") * -1 & "," & rs结算("首先自付金额") * -1 & "," & rs结算("进入统筹费用") * -1 & "," & _
        rs结算("进入统筹支付") * -1 & ",0," & rs结算("超基本封顶线") * -1 & "," & rs结算("个人帐户支付") * -1 & ",'" & ic住院.OutSerialNo + 1 & "'," & _
        IIf(IsNull(rs结算("主页ID")), "null", rs结算("主页ID")) & "," & rs结算("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
    cur个人帐户 = rs结算("个人帐户支付")
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=[1]"
    Set rs结算计算 = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", lng结帐ID)
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "泸州医保")
        
        rs结算计算.MoveNext
    Loop
    
    '医保服务器的事务虽然不能与主程序在一个事务中，但起码要与写卡在一起
    With ic住院
        gstrSQL = "INSERT INTO 保险结算记录 " & _
                "   (性质,记录id,险类,年度,中心代码,序号,病人id,主页id,姓名,性别,年龄,医保号,卡号,身份证号,身份代码,单位医保号 " & _
                ",是否公务员,是否医疗照顾对象,参加补充保险,帐户累计增加,帐户累计支出,统筹已支付金额,统筹已支付费用 " & _
                ",慢病已支付金额,慢病已支付费用,慢病起付金已支付,备案日期 " & _
                ",门诊个人帐户支付金额,住院个人帐户支付金额,额度已支付金额,部门名称,医生名称,病种代码,病种名称,病种类型 " & _
                ",住院次数,增加住院次数,治愈情况,入院日期,出院日期,住院天数 " & _
                ",发生费用金额,个人帐户支付,全自付金额,首先自付金额,转外首先自付,起付线,封顶线,实际起付线 " & _
                ",进入统筹支付,进入统筹费用,进入慢病支付,进入慢病费用,统筹总支付,统筹总自付,统筹基金支付,统筹基金自付 " & _
                ",补充基金支付,补充基金自付,补助基金支付,补助基金自付 " & _
                ",第一段支付,第一段自付,第二段支付,第二段自付,第三段支付,第三段自付,第四段支付,第四段自付,第五段支付,第五段自付 " & _
                ",超基本封顶线,超补充封顶线,进入额度支付,进入门诊个人帐户支付,进入住院个人帐户支付,进入慢性病起付金 " & _
                ",卡灰度级,发票号,票据日期,冲票标志,被冲票据号,支付顺序号,中途结帐,是否上传) " & _
                  " Values "
         gstrSQL = gstrSQL & " (2," & lng冲销ID & "," & TYPE_泸州市 & "," & .MediYear & ",'" & .CenterCode & "','" & Mid(rs结算("序号"), 1, Len(rs结算("序号")) - 1) & "2'," & rs结算("病人ID") & "," & rs结算("主页ID") & ",'" & rs结算("姓名") & _
                  "','" & rs结算("性别") & "'," & rs结算("年龄") & ",'" & rs结算("医保号") & "','" & rs结算("卡号") & "','" & rs结算("身份证号") & "','" & .ClassCode & "','" & .UnitCode & "' " & _
                  "," & .IsOffical & "," & .IsAttend & "," & rs结算("参加补充保险") & "," & .InPerAcc & "," & .OutPerAcc & "," & .PlanPaidAmt & "," & .PlanPaidFee & _
                  "," & .ChronicPaidAmt & "," & .ChronicPaidFee & "," & .ChronicSillPaidAmt & ",null" & _
                  "," & .ClinicPaidAmt & "," & .InHosPaidAmt & "," & .QuotaPaidAmt & ",'" & rs结算("部门名称") & "','" & rs结算("医生名称") & "','" & rs结算("病种代码") & "','" & rs结算("病种名称") & "','" & rs结算("病种类型") & "' " & _
                  "," & .InHosTimes & "," & rs结算("增加住院次数") & ",'0'," & GetOracleFormat(rs结算("入院日期")) & "," & GetOracleFormat(rs结算("出院日期")) & "," & rs结算("住院天数") & _
                  "," & rs结算("发生费用金额") & "," & rs结算("个人帐户支付") & "," & rs结算("全自付金额") & "," & rs结算("首先自付金额") & ",0," & rs结算("起付线") & "," & rs结算("封顶线") & "," & rs结算("实际起付线") & " " & _
                  "," & rs结算("进入统筹支付") & "," & rs结算("进入统筹费用") & "," & rs结算("进入慢病支付") & "," & rs结算("进入慢病费用") & "," & rs结算("统筹总支付") & "," & rs结算("统筹总自付") & "," & rs结算("统筹基金支付") & "," & rs结算("统筹基金自付") & " " & _
                  "," & rs结算("补充基金支付") & "," & rs结算("补充基金自付") & "," & rs结算("补助基金支付") & "," & rs结算("补助基金自付") & _
                  "," & rs结算("第一段支付") & "," & rs结算("第一段自付") & "," & rs结算("第二段支付") & "," & rs结算("第二段自付") & "," & rs结算("第三段支付") & _
                  "," & rs结算("第三段自付") & "," & rs结算("第四段支付") & "," & rs结算("第四段自付") & "," & rs结算("第五段支付") & "," & rs结算("第五段自付") & " " & _
                  "," & rs结算("超基本封顶线") & "," & rs结算("超补充封顶线") & "," & rs结算("进入额度支付") & "," & rs结算("进入门诊个人帐户支付") & "," & rs结算("进入住院个人帐户支付") & "," & rs结算("进入慢性病起付金") & " " & _
                  "," & m泸州.灰度 & ",'" & Nvl(rs结算("发票号"), " ") & "',sysdate,-1,'" & rs结算("序号") & "','" & .OutSerialNo + 1 & "'," & rs结算("中途结帐") & ",0)"
        '准备写卡
        .OutPerAcc = .OutPerAcc - cur个人帐户                  '个人帐户累计支出金额
        .InHosPaidAmt = .InHosPaidAmt - cur个人帐户            '门诊个人帐户支出金额
        .InHosTimes = .InHosTimes - rs结算("增加住院次数")      '有些慢特病会增加住院次数
        .PlanPaidFee = .PlanPaidFee - rs结算("进入统筹费用")      '统筹基金支付费用累计（基本+补充）
        .PlanPaidAmt = .PlanPaidAmt - rs结算("进入统筹支付")        ' 统筹基金支付金额累计（基本+补充）
        .ChronicPaidFee = .ChronicPaidFee - rs结算("进入慢病费用")                '慢性病支付费用累计
        .ChronicPaidAmt = .ChronicPaidAmt - rs结算("进入慢病支付")                '慢性病支付金额累计
        .QuotaPaidAmt = .QuotaPaidAmt - rs结算("进入额度支付")                     '慢性病额度已支付金额
        .ChronicSillPaidAmt = .ChronicSillPaidAmt - rs结算("进入慢性病起付金")      '慢性病起付金已支付金额
        .OutSerialNo = .OutSerialNo + 1           ' 支付顺序号
    End With
        
    '记录住院情况
    Dim payLog As TPayInfo
    With payLog
        .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
        .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
        .AccPay = cur个人帐户
        .Amount = rs结算("发生费用金额")
        .CdFlag = 0
    End With
        
    '完成卡写入
    Dim str数据体 As String
        
    str数据体 = ic住院.CenterCode & "|" & gstr医院编码 & "|1|" & Mid(rs结算("序号"), 1, Len(rs结算("序号")) - 1) & "2|" & _
                TrimStr(ic住院.MediAccountNo) & "|" & cur个人帐户 & "|" & rs结算("统筹基金支付") & "|" & rs结算("补充基金支付") & "|" & _
                rs结算("进入统筹费用") & "|" & rs结算("进入统筹支付") & "|" & rs结算("增加住院次数") & "|" & IIf(rs结算("参加补充保险") = 1, rs结算("超补充封顶线"), rs结算("超基本封顶线")) & "|-1"
    
    If WriteIC(bln离休, True, 1, gstrSQL, ic住院, payLog, str数据体) = False Then
        Exit Function
    End If
            
    住院结算冲销_泸州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 记帐传输_泸州(ByVal str单据号 As String, ByVal int性质 As Integer, Optional ByVal lng病人ID As Long = 0) As Boolean
'功能:上传新产生的记帐明细到医保中心
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
'返回:
    Dim rsTemp As New ADODB.Recordset
    Dim cur全自费 As Currency, cur首先自付 As Currency, cur统筹金额 As Currency
    
    '请注意：泸州医保是在记帐单保存后再调用传输过程的。
    
    On Error GoTo errHandle
    
    '读出该张单据的费用明细
    
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录状态,A.收费细目ID,C.项目编码,B.编码,B.名称,A.实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,病人信息 E " & _
              "  where A.NO=[1] and A.记录性质=[2] and A.记录状态=1 And Nvl(A.是否上传,0)=0 " & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "        and A.病人ID=D.病人ID And A.病人ID=E.病人ID And D.主页ID=E.住院次数 and D.险类=[3]" & _
              "  Order by A.病人ID,A.发生时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "记帐传输", str单据号, int性质, TYPE_泸州市)
    
    If Calc费用分割(rsTemp, True, cur全自费, cur首先自付, cur统筹金额) = False Then
        Exit Function
    End If
        
    记帐传输_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐作废_泸州(ByVal str单据号 As String, ByVal int性质 As Integer, 病人ID As Long) As Boolean
'功能:作废已经上传到医保中心的记帐明细
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'返回:
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput As Variant
    Dim lng上传标志 As Long
    
    On Error GoTo errHandle
    
    '读出该张单据的费用明细中有未上传的记录（取原始单据）
'    gstrSQL = "Select distinct nvl(A.是否上传,0) 上传标志 " & _
'              "  From 病人费用记录 A" & _
'              "  where A.NO='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态<>2 and nvl(A.实收金额,0)<>0 "
'    Call OpenRecordset(rsTemp, "记帐作废")
'
'    If rsTemp.RecordCount > 1 Then
'        MsgBox "该单据里的费用明细还未全部完成费用分割。", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    If rsTemp("上传标志") <> 0 Then
'        '已经完成费用分割或者上传的数据，作废的数据要与原始数据的分割金额相同
'        lng上传标志 = rsTemp("上传标志")
'        gstrSQL = "Select ID " & _
'                  "  From 病人费用记录 A" & _
'                  "  where A.NO='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态=2 and nvl(A.实收金额,0)<>0 "
'        Call OpenRecordset(rsTemp, "记帐作废")
'
'        Do Until rsTemp.EOF
'            '将作废了的单据改为已经完成了费用分割的状态
'            gstrSQL = "ZL_病人费用记录_更新医保(" & rsTemp("ID") & ",null,null,null,null,2)"
'            gcnOracle.Execute gstrSQL, , adCmdStoredProc
'
'            rsTemp.MoveNext
'        Loop
'    End If
    
    记帐作废_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 错误信息_泸州(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息
    Select Case lngErrCode
        Case -2
            错误信息_泸州 = "参数个数错误。"
        Case -3
            错误信息_泸州 = "操作端口失败。"
        Case -4
            错误信息_泸州 = "打开读卡器失败,请检查读卡器连接和电源。"
        Case -5
            错误信息_泸州 = "无卡。"
        Case 0
            错误信息_泸州 = "正确。"
        Case 2
            错误信息_泸州 = "读错误。"
        Case 3
            错误信息_泸州 = "文件结束。"
        Case 4
            错误信息_泸州 = "错误PIN。"
'        Case 5
'            错误信息_泸州 = "。"
        Case 6
            错误信息_泸州 = "复位失败。"
        Case 7
            错误信息_泸州 = "检验错误。"
        Case 8
            错误信息_泸州 = "修改数据失败。"
        Case 9
            错误信息_泸州 = "命令长度错误。"
        Case 10
            错误信息_泸州 = "状态错误。"
        Case 11
            错误信息_泸州 = "文件类别错误。"
        Case 12
            错误信息_泸州 = "文件未选择。"
        Case 13
            错误信息_泸州 = "不可重用。"
        Case 14
            错误信息_泸州 = "文件已经存在。"
        Case 15
            错误信息_泸州 = "错误的P1/P2。"
        Case 16
            错误信息_泸州 = "参数错误。"
        Case 17
            错误信息_泸州 = "错误的P2。"
        Case 18
            错误信息_泸州 = "文件没有找到。"
        Case 19
            错误信息_泸州 = "文件无足够空间。"
        Case 20
            错误信息_泸州 = "参数错误。"
        Case 21
            错误信息_泸州 = "偏移量错误。"
        Case 22
            错误信息_泸州 = "指令代码无效。"
        Case 23
            错误信息_泸州 = "无效的CLA。"
        Case 24
            错误信息_泸州 = "参数错误。"
        Case 25
            错误信息_泸州 = "写卡数据转换错误。"
        Case 26
            错误信息_泸州 = "个人帐户出现负数,交医保中心处理。"
        Case 33
            错误信息_泸州 = "IC卡已经被非法更换,写卡失败。"
        Case 100
            错误信息_泸州 = "一期卡，需要格式转换。"
        Case 101
            错误信息_泸州 = "非本系统卡。"
        Case 210
            错误信息_泸州 = "写卡失败。"
        Case 211
            错误信息_泸州 = "写卡失败,扣卡交医保中心处理。"
        Case 300
            错误信息_泸州 = "CRC校验错误。"
        Case 301
            错误信息_泸州 = "IC卡已经被非法更换,写卡失败.。"
        Case 600
            错误信息_泸州 = "读卡值转换错误。"
        Case Else
            错误信息_泸州 = "不可识别的错误。"
    End Select
End Function

Private Function 装钱操作(ByVal lng病人ID As Long) As Boolean
'功能：首先断断是否要装钱，然后完成相应操作
    Dim rsTemp As New ADODB.Recordset
    
    Dim str装钱模式 As String, bln强制装钱 As Boolean, bln远程验证 As Boolean, str远程地址 As String
    Dim str医保年  As String, lng装钱期次 As Long
    Dim dbl累计注入 As Double
    Dim ic卡 As TIC泸州
    Dim str医保年_IC  As String, lng装钱期次_IC As Long
    Dim dbl累计注入_IC As Double
    Dim lngTemp As Long, bln离休 As Boolean
    
    Dim str参数值 As String
    
    On Error GoTo errHandle
    
    If Get保险参数_泸州(bln远程验证, str远程地址) = False Then
        Exit Function
    End If
    
    If bln远程验证 = True Then
        装钱操作 = True
        Exit Function
    End If
    
    '得到最新的IC卡信息
    '使用本地的，因为可能进行更改但又不成功
    If ReadIC(lng病人ID, 1, True, "装钱时读卡失败。", gIC泸州, bln离休) = False Then
        Exit Function
    End If
    If bln离休 = True Then
        '离休人员不装钱
        装钱操作 = True
        Exit Function
    End If
    
    ic卡 = gIC泸州
    
    With ic卡
        str医保年_IC = .MediYear
        lng装钱期次_IC = .InNo
        dbl累计注入_IC = .InPerAcc
    End With
    
    '获得装钱模式
    '进行合法性验证
    gstrSQL = "SELECT B.医保年,B.装钱序号,B.装钱模式 " & _
               " FROM 保险中心目录 A,保险主机 B " & _
               " WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & ic卡.CenterCode & "' AND A.主机编码=B.编码 AND A.险类=B.险类 "
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        str装钱模式 = Nvl(rsTemp("装钱模式"))
        str医保年 = Nvl(rsTemp("医保年"))
        lng装钱期次 = Nvl(rsTemp("装钱序号"), 0)
    End If
    If str装钱模式 = "" Or str医保年 = "" Then
        MsgBox "请先请管理员完成医保数据的下载。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If str装钱模式 = "1" Then
'        If str医保年 > str医保年_IC Then
'            Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, ic卡.InPerAcc - ic卡.OutPerAcc)
'
'            '把信息写回卡中
'            If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
'                '更新全局变量，可能有用
'                gIC泸州 = ic卡
'                装钱操作 = True
'                Exit Function
'            Else
'                '装钱失败
'                Exit Function
'            End If
'        Else
'            lngTemp = OnLineInMoney(ic卡.CenterCode, ic卡.Cardno, str医保年_IC, Trim(gstr医院编码), serverIP)
'            If lngTemp <> 0 Then
'                Exit Function
'            Else
'                '装钱成功，从卡中读出新的值
'                If ReadICCard(gIC泸州) <> 0 Then
'                    装钱操作 = True
'                    Exit Function
'                End If
'            End If
'        End If
        '在线装钱
        Dim serverIP As String
        serverIP = Get主机IP
        lngTemp = OnLineInMoney(ic卡.CenterCode, ic卡.Cardno, str医保年_IC, Trim(gstr医院编码), serverIP)
        If lngTemp <> 0 Then
            '装钱不成功
            '判断是否列更换医保年
            If str医保年 > str医保年_IC Then
                MsgBox "装钱清单中没有此卡号信息，请到中心处理！", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            '装钱成功，从卡中读出新的值
            If ReadIC(lng病人ID, 1, False, "装钱时读卡失败。", gIC泸州, bln离休) = True Then
                装钱操作 = True
                Exit Function
            End If
        End If
    End If
    
    If str装钱模式 = "0" Then
        '不装钱
        If ic卡.MediYear = "2001" And ic卡.InNo = 0 Then
            '强制离线装钱模式
            bln强制装钱 = True
        Else
            '判断是否列更换医保年
            If str医保年 > ic卡.MediYear Then
                Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, ic卡.InPerAcc - ic卡.OutPerAcc)
                If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
                    '更新全局变量，可能有用
                    gIC泸州 = ic卡
                Else
                    '装钱失败
                    Exit Function
                End If
            End If
        End If
        
    End If
    
    If (str装钱模式 = "2" Or bln强制装钱 = True) And lng装钱期次 > ic卡.InNo Then
        '离线装钱
        If 检查医保服务器_泸州 = False Then
            '不能连接到前置服务器，就认为不可使用
            Exit Function
        End If
        
        '得到装钱清单
        With ic卡
            gstrSQL = "select 帐户注入 from 装钱清单 " & _
                     "where 中心代码='" & .CenterCode & "' and 卡号='" & .Cardno & "' and 装钱期次=" & lng装钱期次
                     '"where 中心代码='" & .CenterCode & "' and 卡号='" & .Cardno & "' and 医保年='" & str医保年 & "' and 装钱期次=" & lng装钱期次
        End With
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn泸州, adOpenStatic
        If rsTemp.RecordCount = 0 Then
            '判断是否列更换医保年
            If str医保年 > ic卡.MediYear Then
                MsgBox "装钱清单中没有此卡号信息，请到中心处理！", vbInformation, gstrSysName
                Exit Function
'                Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, ic卡.InPerAcc - ic卡.OutPerAcc)
'                If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
'                    '更新全局变量，可能有用
'                    gIC泸州 = ic卡
'                    装钱操作 = True
'                End If
            Else
                MsgBox "装钱清单中没有此卡号信息，请到中心处理！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        
        '注意：此处应该改为解密后得到金额
        dbl累计注入 = Val(EncryptStr(IIf(IsNull(rsTemp("帐户注入")), "", rsTemp("帐户注入")), "256", False))
        If str医保年 > ic卡.MediYear Then
            '更换医保年装钱
            Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, dbl累计注入)
        Else
            '不换医保年装钱
            With ic卡
                .InNo = lng装钱期次
                .InPerAcc = dbl累计注入
                .OutSerialNo = .OutSerialNo + 1
            End With
        End If
        If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
            '更新全局变量，可能有用
            gIC泸州 = ic卡
        Else
            '装钱失败
            Exit Function
        End If
    End If
    
    装钱操作 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub 更换医保年装钱(ic泸州 As TIC泸州, ByVal str医保年 As String, ByVal lng装钱期次 As Long, ByVal dbl累计注入 As Double)
    With ic泸州
        .MediYear = str医保年
        .InNo = lng装钱期次
        .InPerAcc = dbl累计注入
        .OutPerAcc = 0
        .PlanPaidAmt = 0
        .PlanPaidFee = 0
        .ChronicPaidAmt = 0
        .ChronicPaidFee = 0
        .InHosTimes = 0
        .QuotaPaidAmt = 0
        .InHosPaidAmt = 0
        .ClinicPaidAmt = 0
        .ChronicSillPaidAmt = 0
        .OutSerialNo = .OutSerialNo + 1
    End With
End Sub

Private Function 记录装钱日志(ic泸州 As TIC泸州, ByVal IC_MediYear As String, ByVal IC_InNo As Long, ByVal IC_InPerAcc As Double) As Boolean
    
    If 检查医保服务器_泸州 = False Then
        '不能连接到前置服务器，就认为不可使用
        Exit Function
    End If
    
    gcn泸州.BeginTrans
    On Error Resume Next
    
    '首先保存装钱日志
    With ic泸州
        gstrSQL = "insert into 装钱日志 (中心代码,卡号,卡中医保年,卡中装钱期次,卡中账户注入" & _
            ",库中医保年,库中装钱期次,库中账户注入,操作日期) values ('" & _
            .CenterCode & "','" & .Cardno & "','" & IC_MediYear & "'," & IC_InNo & "," & Format(IC_InPerAcc, "#####0.00") & ",'" & _
            .MediYear & "'," & .InNo & "," & Format(.InPerAcc, "#####0.00") & ",sysdate)"
        
    End With
    gcn泸州.Execute gstrSQL
    If Err <> 0 Then
        gcn泸州.RollbackTrans
        Err.Clear
        Exit Function
    End If
    
    '完成写卡操作
    If WriteICCard(ic泸州) <> 0 Then
        gcn泸州.RollbackTrans
        MsgBox "IC卡装钱操作失败。", vbInformation, gstrSysName
        Exit Function
    End If
    If Err <> 0 Then '有可能写卡时出现实时错误
        gcn泸州.RollbackTrans
        Err.Clear
        Exit Function
    End If
    
    gcn泸州.CommitTrans
    记录装钱日志 = True
End Function

Private Sub 医保灰度(ByVal str中心 As String, ByVal str卡号 As String)
'返回指定用户的医保灰度级
    Dim rsTemp As New ADODB.Recordset
    
    If 检查医保服务器_泸州 = False Then
        '不能连接到前置服务器，就认为不可使用
        m泸州.灰度 = deg停止支付
        Exit Sub
    End If
    
    gstrSQL = "select 灰度 from 黑名单 where 中心代码='" & str中心 & "' and 卡号='" & str卡号 & "'"
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount > 0 Then
        '设置灰度值
        m泸州.灰度 = Val(rsTemp("灰度"))
    Else
        '正常的不下发
        m泸州.灰度 = deg正常支付
    End If
    
End Sub

Private Function 检查IC卡(ByVal lng病人ID As Long, ByVal str卡号 As String, ByVal str中心 As String) As Boolean
'功能：判断该病人的卡是否插入正确
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.卡号,A.医保号,B.编码 from 保险帐户 A,保险中心目录 B " & _
              " where A.险类=[1] and A.病人ID=[2] and a.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市, lng病人ID)
    
    If rsTemp("卡号") <> str卡号 Or rsTemp("编码") <> str中心 Then
        MsgBox "刷卡器中的卡不是当前病人的，请插入正确的IC卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    检查IC卡 = True
End Function

Public Function 检查医保服务器_泸州() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn泸州.State = adStateOpen Then
        检查医保服务器_泸州 = True
        Exit Function
    End If
    
    '读出连接医保服务器的配置
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                '解密
                If strPass <> "" Then strPass = EncryptStr(strPass, 256, False)
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(gcn泸州, strServer, strUser, strPass, False) = True Then
        检查医保服务器_泸州 = True
        Exit Function
    End If
        
    MsgBox "医保前置服务器连接失败。", vbInformation, gstrSysName
End Function

Public Function Get离休病人_泸州(ByVal strIdentify As String, ic泸州 As TIC泸州, Optional ByVal bln医保号 As Boolean = True) As Boolean
'功能：从离休清单中读取病人情况，填入IC卡结构中
'参数：strIdentify     病人身份验证（bln医保号=False 为身份证 ，bln医保号=True 是医保号）
'      IC泸州        根据读出的信息填写IC卡结构
    Dim rsTemp As New ADODB.Recordset

    If 检查医保服务器_泸州 = False Then
        Exit Function
    End If
    
    gstrSQL = "select * from 离休人员 where " & IIf(bln医保号 = True, "医保号", "身份证号") & _
                "='" & strIdentify & "'"
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '没找到该离休病人的记录
        Exit Function
    End If
    
    With ic泸州
        .CenterCode = rsTemp("中心代码")     'As String * 4      ' 中心代码
        .Cardno = rsTemp("医保号")           'As String * 8      ' 卡号
        .IDCardno = rsTemp("身份证号")       'As String * 18     ' 身份证号 长度不足后补#0
        .MediAccountNo = rsTemp("医保号")    'As String * 8      ' 医保号
        .Name = rsTemp("姓名")               'As String * 10     ' 姓名
        .Sex = IIf(IsNull(rsTemp("性别")), "1", rsTemp("性别"))       'As String * 1      ' 性别 1-男  0-女
        .Birthday = rsTemp("生日")           'As String * 8      ' 出生日期 YYYYMMDD
        .UnitCode = rsTemp("单位医保号")     'As String * 5      ' 用人单位编码
        .ClassCode = rsTemp("身份代码")      'As String * 2      ' 职工身份：0x：在职1x：退休, 05和11为一次性缴费
        .DomainCode = 0     'As String * 1      ' 职工属地 0-正常 1-常驻外地 2-异地安置
        .MediYear = Year(zlDatabase.Currentdate)          'As String * 4      ' 医保年度
        .InNo = 0           'As Long            ' 装钱期次
        .OutSerialNo = 0    'As Long            ' 支付顺序号
        .InPerAcc = 0       'As Double          ' 个人帐户累计注入金额
        .OutPerAcc = 0      'As Double          ' 个人帐户累计支出金额
        .PlanPaidAmt = 0     'As Double          ' 本年统筹支付金额累计
        .PlanPaidFee = 0 'As Double          ' 本年进入统筹金额累计
        .ChronicPaidFee = 0 '   As Double          ' 慢性病支付费用累计
        .ChronicPaidAmt = 0 '   As Double          ' 慢性病支付金额累计
        .InHosPaidAmt = 0 '     As Double          ' 住院个人帐户支付金额
        .ClinicPaidAmt = 0 '    As Double          ' 门诊个人帐户支付金额
        .QuotaPaidAmt = 0 '     As Double          ' 慢性病额度已支付金额
        .ChronicSillPaidAmt = 0 '    As Double     ' 慢性病起付金已支付金额
        .IsOffical = "0" '        As String * 1      ' 公务员 0-否；其他-是
        .IsAttend = "0" '       As String * 1      ' 医疗照顾对象 0-否；1-是
        .Password = "9000"       'As String * 4      ' 个人密码
        .InHosTimes = 0 'As Long           ' 本年有效住院次数
        .InpatientFlag = 0  'As String * 1      ' 住院标志 0-不住院 1-住院
    End With
    
    Get离休病人_泸州 = True
End Function


Private Function Is离休病人(ByVal lng病人ID As Long) As Boolean
'功能：根据帐户信息判断病人是否离休病人
'参数：返回病人的医保号
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.在职 from 保险帐户 A where A.险类=[1] and A.病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市, lng病人ID)
    
    If rsTemp.EOF = True Then
        '该病人没发现
        Is离休病人 = False
    Else
        Is离休病人 = IIf(rsTemp("在职") = 3, True, False)
    End If
End Function

Private Function Get帐户信息(ByVal lng病人ID As Long, str医保号 As String, str身份证号 As String, str密码 As String) As Boolean
'功能：根据帐户信息判断病人是否离休病人
'参数：返回病人的医保号
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.医保号,A.密码,B.身份证号 from 保险帐户 A,病人信息 B where A.险类=[1]" & _
        " and A.病人ID=[2] And A.病人ID=B.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市, lng病人ID)
    
    If rsTemp.EOF = False Then
        '该病人发现
        str医保号 = Nvl(rsTemp("医保号"))
        str身份证号 = Nvl(rsTemp("身份证号"))
        str密码 = Nvl(rsTemp("密码"))
        Get帐户信息 = True
    Else
        MsgBox "无法读取帐户信息。", vbInformation, gstrSysName
    End If
End Function

'Modified By 朱玉宝 2003-12-10 地区：泸州 增加参数
Private Function Calc费用分割(rs费用明细 As ADODB.Recordset, ByVal 是否更新 As Boolean _
                , cur全自费 As Currency, cur首先自付 As Currency, cur统筹 As Currency, _
                Optional ByVal 费用分割 As Boolean = False, Optional ByVal bln门诊 As Boolean = False) As Boolean
'功能：根据费用明细，重新计算明细中费用的报销金额。计算好的金额可以直接上传
'参数：rs费用明细  费用明细，包含费用的细目ID、单价、数量、金额
'      是否更新     是否需要对数据库中病人费用记录的医保数据进行更新。门诊预算时不能做
'      cur全自费    输出参数，费用中全自费部分的金额
'      cur首先自付  输出参数，费用中首先自付部分的金额
'      cur统筹      输出参数，费用中统筹部分的金额
'      费用分割     输入参数，为否表示限价从病人费用记录中读取，仅计算当前那笔记录
'返回：本函数成功完成所有功能，为True
'调用位置：门诊预算、门诊结算、住院记帐、住院预算、住院结算、费用明细上传

    Dim str中心编码 As String, str病种编码 As String, lng病人ID As Long
    Dim rs保险大类 As New ADODB.Recordset
    Dim rs病种特准 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset, str项目编码 As String, str细目名称 As String
    Dim cur金额 As Currency, cur实际单价 As Currency, cur最大价格 As Currency, cur单价 As Currency, cur自付比例 As Currency, cur床位费 As Currency, cur乙类项目 As Currency
    Dim cur统筹金额 As Currency, cur自付 As Currency, lng保险大类ID As Long, lng保险项目否 As Long
    Dim bln医保病人 As Boolean, bln医保项目 As Boolean, bln更新 As Boolean, bln计算 As Boolean
    
    If 检查医保服务器_泸州 = False Then
        Exit Function
    End If
    cur全自费 = 0
    cur首先自付 = 0
    cur统筹 = 0
    
    On Error GoTo errHandle
    '得到所有医保大类
    gstrSQL = "SELECT A.ID,A.编码 FROM 保险支付大类 A Where A.险类 =" & TYPE_泸州市
    rs保险大类.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    'Modified by zyb ##2003-08-31
    If Not 费用分割 Then If rs费用明细.RecordCount > 0 Then rs费用明细.MoveFirst
    Do Until rs费用明细.EOF
        bln计算 = True
        If Nvl(rs费用明细!数量, 0) = 0 Then
            cur实际单价 = 0
        Else
            cur实际单价 = rs费用明细!实收金额 / Nvl(rs费用明细!数量, 0)
        End If
        
        If lng病人ID <> rs费用明细("病人ID") Then
            '先判断是不是医保病人
            bln医保病人 = False
            If Not bln门诊 Then
                gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(rs费用明细!病人ID), TYPE_泸州市)
                bln医保病人 = (rsTemp!Records = 1)
            Else
                bln医保病人 = True
            End If
            
            If bln医保病人 Then
                lng病人ID = rs费用明细("病人ID")
                '不同的病人，可能属于不同的中心，其床位限价也可能不同，所以要单独处理
                gstrSQL = "SELECT B.编码 中心,C.编码 AS 病种编码 " & _
                    "FROM 保险帐户 A,保险中心目录 B,保险病种 C " & _
                    "WHERE A.病人ID=" & lng病人ID & " AND A.险类=" & TYPE_泸州市 & " AND A.险类=B.险类 AND nvl(A.中心,0)=nvl(B.序号,0) AND A.病种ID=C.ID(+)"
                If rsTemp.State = adStateOpen Then rsTemp.Close
                rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
                
                '得到该医保病人的病种特准项目
                gstrSQL = "SELECT A.项目序号,A.首先自付比例 FROM 保险病种项目 A Where A.病种序号 ='" & rsTemp("病种编码") & "'"
                If rs病种特准.State = adStateOpen Then rs病种特准.Close
                rs病种特准.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
                
                '得到该中心规定的床位费限价
                str中心编码 = rsTemp("中心")
                gstrSQL = "Select 每天床位费限价,乙类项目价格 From 保险中心目录 Where 险类=" & TYPE_泸州市 & " And 编码='" & rsTemp("中心") & "'"
                If rsTemp.State = adStateOpen Then rsTemp.Close
                rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
                cur床位费 = rsTemp("每天床位费限价")
                cur乙类项目 = Nvl(rsTemp("乙类项目价格"), 0)
            End If
        End If
        
        If bln医保病人 Then
            If 是否更新 = False Then
                If Get医保编码(rs费用明细("收费细目ID"), str项目编码, str细目名称) = False Then
                    MsgBox str细目名称 & "还没有完成保险编码的对应，不能完成结算。", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                If IsNull(rs费用明细("项目编码")) = True Then
                    MsgBox "请为" & rs费用明细("名称") & "设置医保编码。", vbInformation, gstrSysName
                    Exit Function
                End If
                str项目编码 = rs费用明细("项目编码")
                str细目名称 = rs费用明细("名称")
            End If
            
            '获得保险项目的详细信息，方便计算
            bln医保项目 = False
            gstrSQL = "Select a.门诊医保,a.住院医保,a.最大价格限制,a.价格,a.是否医保,a.大类编码,b.首先自付比例 from 保险项目 a,保险支付大类比例 b Where a.险类=" & TYPE_泸州市 & " And a.编码='" & str项目编码 & "' and a.险类=b.险类 and a.大类编码=b.大类编码 and b.中心代码='" & str中心编码 & "'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
                MsgBox str细目名称 & "的保险编码有误，不能完成结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bln门诊 Then
                bln医保项目 = (Nvl(rsTemp!门诊医保, 0) = 1)
            Else
                bln医保项目 = (Nvl(rsTemp!住院医保, 0) = 1)
            End If
            If rs费用明细("收费类别") = "J" Then
                '床位费
                lng保险项目否 = 1
                If cur实际单价 <= cur床位费 Then
                    cur统筹金额 = rs费用明细("实收金额")
                Else
                    cur统筹金额 = cur床位费 * rs费用明细("数量")
                End If
                cur统筹 = cur统筹 + cur统筹金额
                cur全自费 = cur全自费 + (rs费用明细("实收金额") - cur统筹金额)
                cur最大价格 = cur床位费
            Else
                'Modified by zyb 20050429
                '求出该项目的最大可以报销的价格
                '如果当前记录是冲销记录，则取原始记录的限价
                If bln门诊 = False Then
                    If rs费用明细!记录状态 = 2 Then
                        Dim rsGet As New ADODB.Recordset
                        Set rsGet = New ADODB.Recordset
                        If rsGet.State = 1 Then rsGet.Close
                        rsGet.Open "Select 保险大类ID,保险项目否,保险编码,统筹金额,nvl(限价,0) AS 限价 From 住院费用记录 " & _
                            " Where (NO,记录性质,记录状态,序号) IN " & _
                            "     (Select NO,记录性质,3,序号 From 住院费用记录 Where ID=" & rs费用明细!ID & ") And (Nvl(是否上传,0)=1 or nvl(结帐id,0)>0)", gcnOracle
                        If rsGet.RecordCount <> 0 Then
                            cur最大价格 = rsGet!限价
                            cur统筹金额 = -1 * Nvl(rsGet!统筹金额, 0)
                            lng保险大类ID = Nvl(rsGet!保险大类id, 0)
                            lng保险项目否 = Nvl(rsGet!保险项目否, 0)
                            str项目编码 = Nvl(rsGet!保险编码)
                            bln计算 = False
                        Else
                            cur最大价格 = IIf(Nvl(rsTemp("最大价格限制"), 0) = 0, Nvl(rsTemp("价格"), 0), rsTemp("最大价格限制"))
                        End If
                    Else
                        cur最大价格 = IIf(Nvl(rsTemp("最大价格限制"), 0) = 0, Nvl(rsTemp("价格"), 0), rsTemp("最大价格限制"))
                    End If
                Else
                    cur最大价格 = IIf(Nvl(rsTemp("最大价格限制"), 0) = 0, Nvl(rsTemp("价格"), 0), rsTemp("最大价格限制"))
                End If
                
                If bln计算 Then
                    'Modified by zyb ##2003-08-31
                    If 费用分割 Then
                        If Nvl(rs费用明细("限价"), 0) = 0 And Nvl(rs费用明细("统筹金额"), 0) = 0 Then
                            '如果费用记录中保存的限价为零且统筹金额也为零，则说明以前是非医保病人，以当前的限价为准
                            '医保病人正常记帐，未启用限价或启用限价前记的帐，都可能产生病人费用记录中的限价为零的情况，但统筹金额不可能为零
                            '非医保项目不可能存在限价的情况
                        Else
                            cur最大价格 = Nvl(rs费用明细("限价"), 0)
                        End If
                    End If
                    'Modified end
                    If cur最大价格 > 0 And cur最大价格 < cur实际单价 Then
                        '该项目存在最大限价，并且比医院价格要低
                        cur单价 = cur最大价格
                    Else
                        cur单价 = cur实际单价
                    End If
                    
                    rs病种特准.Filter = "项目序号='" & str项目编码 & "'"
                    If rs病种特准.EOF = False Then
                        '是否医保项目，按此处作准
                        lng保险项目否 = IIf(rs病种特准("首先自付比例") = 1, 0, 1)
                        cur自付比例 = rs病种特准("首先自付比例")
                    Else
                        '以保险项目中的值为准
                        lng保险项目否 = rsTemp("是否医保")
                        cur自付比例 = rsTemp("首先自付比例")
                        
                        If lng保险项目否 = 1 And cur乙类项目 > 0 And _
                            (rs费用明细("收费类别") <> "5" And rs费用明细("收费类别") <> "6" And rs费用明细("收费类别") <> "7") Then
                            
                            '对于按价格开区分甲类或乙类项目的中心
                            If cur实际单价 >= cur乙类项目 Then
                                cur自付比例 = 0.2
                            Else
                                cur自付比例 = 0
                            End If
                        End If
                        
                        '虽然定义为保险项目，但由于自付比例，仍改为全自费
                        If lng保险项目否 = 1 And rsTemp("首先自付比例") = 1 Then lng保险项目否 = 0
                    End If
                    
                    If lng保险项目否 = 0 Or Not bln医保项目 Then
                        '全自费项目
                        '2005-09-12 by gzy lng保险项目否=rstemp("是否医保")*iff(rstemp("首先自负比例"),1,0,1)*bln(医保项目)
                        lng保险项目否 = 0
                        cur统筹金额 = 0
                        cur全自费 = cur全自费 + rs费用明细("实收金额")
                    Else
                        If cur最大价格 = 0 Or cur实际单价 <= cur最大价格 Then
                            '没有价格限制，或者限制的价格还没有超过
                            cur统筹金额 = rs费用明细("实收金额") * (1 - cur自付比例)
                        Else
                            '有价格限制，就只能取最大价格
                            cur统筹金额 = cur最大价格 * rs费用明细("数量") * (1 - cur自付比例)
                        End If
                        cur统筹 = cur统筹 + cur统筹金额
                        
                        'Modified by zyb ##2003-08-31
                        '当存在最大价格限制时,其首先自付的计算规则应该是(全自付=超限部分+非医保项目的费用;实收金额=统筹金额+首先自付+全自付)
                        If cur最大价格 > 0 And cur最大价格 < cur实际单价 Then
                            cur自付 = (cur最大价格 * rs费用明细("数量") - cur统筹金额)
                        Else
                            cur自付 = (rs费用明细("实收金额") - cur统筹金额)
                        End If
                        cur首先自付 = cur首先自付 + cur自付
                        cur全自费 = cur全自费 + (rs费用明细("实收金额") - cur统筹金额 - cur自付)
                        'Modified end
                    End If
                End If
            End If
            
            If bln计算 Then
                rs保险大类.Filter = "编码='" & rsTemp("大类编码") & "'"
                If rs保险大类.EOF = False Then
                    lng保险大类ID = rs保险大类("ID")
                Else
                    lng保险大类ID = 0
                End If
            End If
            
            '只有门诊预结算不更新
            If 是否更新 = True Then
                '不做事务控制，这样可以与门诊收费放在一个事务中。然后住院数据都是已经保存好了的，随便怎么计算都无所谓
                'Modified by zyb ##2003-09-01(因为统一改为预结算时全部重算,所以不更新是否上传标志)
                gstrSQL = "ZL_病人费用记录_更新医保(" & rs费用明细("ID") & "," & cur统筹金额 & "," & _
                    lng保险大类ID & "," & lng保险项目否 & ",'" & str项目编码 & "',NULL," & cur最大价格 & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            
            'Modified by zyb ##2003-08-31
            If 费用分割 Then Exit Do
        End If
        rs费用明细.MoveNext
    Loop
    
    Calc费用分割 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get医保编码(ByVal 明细ID As Long, 医保编码 As String, 细目名称 As String) As Boolean
'功能：根据费用明细ID，得到其医保编码
'参数：明细ID     收费细目的ID
'      医保编码   输出值，收费细目对应的医保编码
'返回：本函数成功完成所有功能，为True
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select A.项目编码,B.名称 From 保险支付项目 A,收费细目 B Where B.ID=" & 明细ID & " And B.ID=A.收费细目ID(+) And A.险类(+)=" & TYPE_泸州市
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    If rsTemp.EOF = False Then
        医保编码 = Nvl(rsTemp("项目编码"))
        细目名称 = Nvl(rsTemp("名称"))
    Else
        医保编码 = ""
        细目名称 = "ID为" & 明细ID & "的项目"
    End If
    
    Get医保编码 = (医保编码 <> "")
End Function

Private Function Calc基本统筹() As Boolean
'功能：计算出住院病人的普通基本统筹金额
'输入参数：
'输出参数：
'返回：成功计算，则返回True
    
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    
    Dim cls医保 As New clsInsure
    Dim dbl多次起付线和 As Currency, dbl原起付线 As Currency, dbl新起付线 As Currency
    Dim dbl多次进入统筹和 As Currency, dbl多次首先自付和 As Currency     '多次是指该病人本次住院以前结帐的累计
    Dim cur全自费 As Currency, cur首先自付 As Currency, cur统筹 As Currency
    Dim str项目编码 As String, str项目名称 As String
    '计算参数
    Dim bln无起付线 As Boolean, bln无封顶线 As Boolean, bln医保项目 As Boolean
    
    On Error GoTo errHandle
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量。以及计算参数
    Set gcol结算计算 = New Collection
    
    m政策.个人帐户支付全自费 = cls医保.GetCapability(support结算帐户全自费, 0, TYPE_泸州市)
    m政策.个人帐户支付首先自付 = cls医保.GetCapability(support结算帐户首先自付, 0, TYPE_泸州市)
    m政策.个人帐户支付超限 = cls医保.GetCapability(support结算帐户超限, 0, TYPE_泸州市)
    
    gstrSQL = "SELECT B.医保年,A.起付线在段中,A.段值类型,A.封顶类型,A.补充报销减起付金,A.使用累计报销,A.个人账户可支付首先自付 " & _
               " ,A.跨年起付金类型,A.跨年增加住院次数 " & _
               " FROM 保险中心目录 A,保险主机 B " & _
               " WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & gIC泸州.CenterCode & "' AND A.主机编码=B.编码 AND A.险类=B.险类 "
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        m泸州.年度 = Val(Nvl(rsTemp("医保年")))
        m政策.费用段值 = Nvl(rsTemp("段值类型")) = 1
        m政策.费用封顶 = Nvl(rsTemp("段值类型")) = 1
        m政策.起付线在段中 = Nvl(rsTemp("起付线在段中")) = 1
        m政策.使用累计 = Nvl(rsTemp("使用累计报销")) = 1
        m政策.补充报销减起付金 = Nvl(rsTemp("使用累计报销")) = 1
        m政策.个人帐户支付首先自付 = Nvl(rsTemp("个人账户可支付首先自付")) = 1
        m政策.跨年起付金类型 = Nvl(rsTemp("跨年起付金类型"), 0)
        m政策.跨年增加住院次数 = Nvl(rsTemp("跨年增加住院次数"), 0)
    End If
    If m泸州.年度 = 0 Then
        MsgBox "请系统管理员完成医保数据的下载。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.1、算出病人本次住院的各种费用
    'Modified by zyb ##2003-08-31(标准单价改为单价,金额改为实收金额)
    '计算公式:全自付=超限部分+非医保项目的费用;实收金额=统筹金额+首先自付+全自付
    gstrSQL = _
        "Select Mod(A.记录性质,10) as 记录性质,A.记录状态,A.NO,Nvl(A.价格父号,序号) as 序号,A.病人ID,A.主页ID," & _
        "   A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0) as 保险大类ID,Avg(Nvl(A.付数,1)*A.数次) as 数量,NVL(A.统筹金额,0) as 统筹金额," & _
        "   Sum(A.标准单价) as 单价,Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 实收金额,A.发生时间,Nvl(A.保险项目否,0) as 保险项目否,Nvl(Sum(限价),0) 限价" & _
        "   From 住院费用记录 A" & _
        "   Where A.记帐费用=1 And Nvl(A.记录状态,0)<>0 And A.病人ID=[1] and A.主页ID=[2] And A.操作员姓名 is not null" & _
        "   Group by Mod(A.记录性质,10),A.记录状态,A.NO,Nvl(A.价格父号,序号),A.病人ID,A.主页ID," & _
        "       A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0),A.发生时间,Nvl(A.保险项目否,0),NVL(A.统筹金额,0)" & _
        "   Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", m泸州.病人ID, m泸州.主页ID)
    
    With m泸州
        Do Until rsTemp.EOF
            '2004-05-27
            '判断是否属于医保项目（由于床位费未调用费用分割函数，只有单独判断）
            If Get医保编码(rsTemp("收费细目ID"), str项目编码, str项目名称) = False Then
                MsgBox str项目名称 & "还没有完成保险编码的对应，不能完成结算。", vbInformation, gstrSysName
                Exit Function
            End If
            '2005-09-12 by gzy 判断住院医保项目在cale费用分割中完成，以"是否医保"标识
            'gstrSQL = "Select * from 保险项目 Where 险类=" & TYPE_泸州市 & " And 编码='" & str项目编码 & "'"
            'If rsCheck.State = adStateOpen Then rsCheck.Close
            'rsCheck.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
            'bln医保项目 = (Nvl(rsCheck!住院医保, 0) = 1)
            
            If rsTemp("保险项目否") = 0 Then  'Or Not bln医保项目
                .全自费 = .全自费 + rsTemp("实收金额")
            Else
                If rsTemp("收费类别") = "J" Then
                    .进入统筹 = .进入统筹 + rsTemp("统筹金额")
                    .医保项目金额 = .医保项目金额 + rsTemp("统筹金额")
                    If rsTemp("实收金额") <> rsTemp("统筹金额") Then
                        .全自费 = .全自费 + rsTemp("实收金额") - rsTemp("统筹金额")
                    End If
                Else
                    .进入统筹 = .进入统筹 + rsTemp("统筹金额")
                    If rsTemp("实收金额") <> rsTemp("统筹金额") Then
                        '所有乙类项目的金额
                        'Modified by zyb ##2004-11-13   不能按结算时医保项目的状态再次重算，因为上面计算进入统筹是直接取的原来的值
'                        Call Calc费用分割(rsTemp, False, cur全自费, cur首先自付, cur统筹, True)
'                        If cur首先自付 = 0 Then '只有乙类存在首先自付，这里是对限价的处理
'                            .医保项目金额 = .医保项目金额 + cur统筹
'                        Else
'                            .乙类项目金额 = .乙类项目金额 + cur统筹 + cur首先自付
'                        End If
                        If rsTemp!限价 <> 0 Then
                            '限价是单个数量的限价
                            If rsTemp!限价 * rsTemp!数量 = rsTemp!统筹金额 Then
                                '甲类
                                .医保项目金额 = .医保项目金额 + rsTemp!统筹金额
                                cur全自费 = rsTemp!实收金额 - rsTemp!统筹金额
                                cur首先自付 = 0
                            Else
                                'Modified by zyb 20050429
                                '乙类项目
                                If rsTemp!限价 >= (rsTemp!实收金额 / Nvl(rsTemp!数量, 1)) Then
                                    '此种情况说明限价大于单价，以单价为准计算的统筹金额，此时全自费为零
                                    cur全自费 = 0
                                    cur首先自付 = rsTemp!实收金额 - rsTemp!统筹金额
                                    .乙类项目金额 = .乙类项目金额 + (rsTemp!统筹金额 + cur首先自付)
                                Else
                                    cur全自费 = rsTemp!实收金额 - (rsTemp!限价 * rsTemp!数量)
                                    cur首先自付 = (rsTemp!限价 * rsTemp!数量) - rsTemp!统筹金额
                                    .乙类项目金额 = .乙类项目金额 + (rsTemp!统筹金额 + cur首先自付)
                                End If
                            End If
                        Else
                            cur全自费 = 0
                            cur首先自付 = rsTemp!实收金额 - rsTemp!统筹金额
                            .乙类项目金额 = .乙类项目金额 + (rsTemp!统筹金额 + cur首先自付)
                        End If

                        .全自费 = .全自费 + cur全自费
                        .首先自付 = .首先自付 + cur首先自付
                        'Modified end
                    Else
                        .医保项目金额 = .医保项目金额 + rsTemp("实收金额")
                    End If
                End If
            End If
            
            .发生费用 = .发生费用 + rsTemp("实收金额")
            rsTemp.MoveNext
        Loop
    End With
    
    '1.2、得到帐户的相关信息
    With m泸州
        gstrSQL = "select A.人员身份,A.在职,A.年龄段," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) " & _
                  "     and B.年度(+)=[1] and A.病人ID=[2] and A.险类=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", .年度, .病人ID, TYPE_泸州市)
        
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        .住院次数 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        
        gstrSQL = "select 年龄段,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_泸州市 & " and nvl(中心,0)=" & .中心序号 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        Call OpenRecordset_OtherBase(rsTemp, "", gstrSQL, gcn泸州)
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln无起付线 = (rsTemp("无起付线") = 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
        
        m政策.全额统筹 = (rsTemp("全额统筹") = 1)
    End With
    
    '1.3 读出本次住院期间累计结帐情况
    gstrSQL = "select nvl(max(A.起付线),0) as 原起付线,nvl(sum(A.实际起付线*冲票标志),0) as 起付线,nvl(sum((A.发生费用金额-A.全自付金额-A.首先自付金额)*冲票标志),0) as 进入统筹金额,nvl(sum(A.首先自付金额*冲票标志),0) as 首先自付金额 " & _
              "  from 保险结算记录 A " & _
              "  Where A.病人ID = " & m泸州.病人ID & " And A.主页ID = " & m泸州.主页ID & _
              " And A.险类 = " & TYPE_泸州市
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    dbl原起付线 = rsTemp("原起付线")
    dbl多次起付线和 = rsTemp("起付线")
    dbl多次进入统筹和 = rsTemp("进入统筹金额")
    dbl多次首先自付和 = rsTemp("首先自付金额")
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、获得起付线、封顶线、支付比例等数据
    '3.1、获得起付线、封顶线
    'Modified By 朱玉宝 2004-05-08 原因：建立分中心
    With m泸州
        gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 ,max(decode(A.性质,'1',A.金额,0)) as 起付线 " & _
                  "         ,max(decode(A.性质,'" & (.住院次数 + 1) & "',A.金额,0)) as 实际起付线,min(A.金额) as 最低起付线 " & _
                  "  from 保险支付限额 A " & _
                  "  where A.险类=" & TYPE_泸州市 & " and A.中心=" & .中心序号 & " and A.年度=" & .年度 & " And A.在职=" & lng在职 & " And A.人员类别=" & IIf(gIC泸州.DomainCode = 0, 1, 2)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
                
        If bln无起付线 Then
            dbl新起付线 = 0
        Else
            dbl新起付线 = IIf(IsNull(rsTemp("实际起付线")), 0, rsTemp("实际起付线"))
            If dbl新起付线 = 0 Then
                '一般都会有，如果实在超过了住院次数，就取最后一次（也就是金额最小的一次）
                dbl新起付线 = IIf(IsNull(rsTemp("最低起付线")), 0, rsTemp("最低起付线"))
            End If
            If dbl新起付线 = 0 Then
                MsgBox "请在“年度结算规则”中设置本年度的起付线。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If bln无封顶线 Then
            .封顶线 = 0
        Else
            .封顶线 = IIf(IsNull(rsTemp("封项线")), 0, rsTemp("封项线"))
            If .封顶线 = 0 Then
                MsgBox "请在“年度结算规则”中设置本年度的封顶线。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    Dim bln补起付线 As Boolean
    If m泸州.跨年住院 = False Then
        m泸州.起付线 = dbl新起付线
        bln补起付线 = True
    Else
        Select Case m政策.跨年起付金类型
            Case 0
                m泸州.起付线 = dbl原起付线
                bln补起付线 = True
            Case 1
                m泸州.起付线 = dbl新起付线
                bln补起付线 = True
            Case Else
                m泸州.起付线 = dbl新起付线
                bln补起付线 = False
        End Select
    End If
    
    '算出本次需要扣除的起付线
    If bln补起付线 = True Then
        If m泸州.起付线 > dbl多次起付线和 Then
            '得到预计支付的起付线，还不是最终的
            m泸州.本次起付线 = m泸州.起付线 - dbl多次起付线和
        Else
            '没有要支付的起付线
            m泸州.本次起付线 = 0
        End If
    End If
    
    '是否增加住院次数
    If m泸州.中途结帐 = 0 Then
        '出院
        If m泸州.跨年住院 = True Then
            '本年度住院
            m泸州.住院次数增加 = m政策.跨年增加住院次数
        Else
            m泸州.住院次数增加 = IIf(m泸州.灰度 = deg停止支付, 0, 1)
        End If
    End If
    
    If m泸州.灰度 < deg个人支付 Then
        '不需要再计算与报销相关的值
        Calc基本统筹 = True
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额。为了比较明显地了解变量的使用，故把变量定义写在这里
    With m泸州
        If m政策.使用累计 = True Then
            '累计金额就从卡上取
            .统筹已支付费用 = gIC泸州.PlanPaidFee
            .统筹已支付金额 = gIC泸州.PlanPaidAmt
        Else
            '但本次住院的要累计
            gstrSQL = "SELECT nvl(sum(进入统筹支付*冲票标志),0) 累计支付,nvl(sum(进入统筹费用*冲票标志),0) 累计费用 " & _
                      "FROM 保险结算记录 WHERE 病人ID=" & .病人ID & " AND 主页ID=" & .主页ID & " AND 性质=2 AND 险类=" & TYPE_泸州市
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn泸州
            .统筹已支付费用 = rsTemp("累计费用")
            .统筹已支付金额 = rsTemp("累计支付")
        End If
    
        '跨年结算，这些金额都应该是0
        If .跨年结算 = True Then
            '跨年结算就不用考虑以前的结算金额
            dbl多次起付线和 = 0
            dbl多次进入统筹和 = 0
            .统筹已支付费用 = 0
            .统筹已支付金额 = 0
        End If
        
    
        '如果已经超过封顶，直接退出，不需要再扣起付线了
        If m政策.费用封顶 = True Then
            '费用封顶的超封顶线可能含有首先自付部分
            If .统筹已支付费用 >= .封顶线 And .封顶线 > 0 Then
                .超基本封顶线 = .发生费用 - .全自费
                Calc基本统筹 = True
                Exit Function
            End If
        Else
            '支付封顶的超封顶线只能含有进入统筹部分
            If .统筹已支付金额 >= .封顶线 And .封顶线 > 0 Then
                .超基本封顶线 = .进入统筹
                Calc基本统筹 = True
                Exit Function
            End If
        End If
    
        '3.3、取得费用档次
        If rsTemp.State = adStateOpen Then rsTemp.Close
        'Modified By 朱玉宝 2004-05-08 原因：建立分中心
        gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
                  "  from 保险支付比例 A,保险费用档 B " & _
                  "  Where A.险类 =" & TYPE_泸州市 & " And A.中心 =" & m泸州.中心序号 & " And A.年度 =" & m泸州.年度 & " And A.在职 =" & lng在职 & " And A.年龄段 =" & lng年龄段 & _
                  "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 And A.在职=B.在职 and A.人员类别=" & IIf(gIC泸州.DomainCode = 0, 1, 2) & _
                  "  order by B.档次"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '然后进入分段计算
        '能求出实际起付线、分段报销金额、分段进入费用
        If m政策.费用段值 = True Then
            '费用段值
            If m政策.费用封顶 = False Then
                '支付封顶，类似于自贡那种模式
                If Calc基本分段1(rsTemp, m政策.起付线在段中 = False, dbl多次起付线和, dbl多次进入统筹和) = False Then Exit Function
            Else
                '费用封顶，类似于泸州模式
                If Calc基本分段2(rsTemp, m政策.起付线在段中 = False, dbl多次起付线和, dbl多次进入统筹和, dbl多次首先自付和) = False Then Exit Function
            End If
        Else
            '支付段值
            If m政策.费用封顶 = False Then
                '支付封顶
                If Calc基本分段3(rsTemp) = False Then Exit Function
            Else
                '费用封顶
                If Calc基本分段4(rsTemp) = False Then Exit Function
            End If
        End If
        
        'takecare
        '计算超限自付部分
        If .封顶线 > 0 Then
            '有封顶线
            If m政策.费用封顶 = True Then
                '.进入统筹 和 .首先自付 都属于费用（但如果必扣起付金的话，超封顶线中也不能包含那部分金额）
                .超基本封顶线 = (.发生费用 + .统筹已支付费用) - .全自费 - .封顶线 '- IIf(m政策.补充报销减起付金 = True, .实际起付线, 0)
            Else
                '支付封顶，只有统筹部分
                .超基本封顶线 = .进入统筹 - .统筹基金支付 - .统筹基金自付 - .实际起付线
            End If
            If .超基本封顶线 < 0 Then .超基本封顶线 = 0                   '如果进入统筹金额还不到起付线，为负数
        End If
        
        '如果在黑名单中（灰度<医保支付），则实际起付线=0，相关统筹自付值=0
        If m泸州.灰度 = deg停止支付 Then
            .实际起付线 = 0
        End If
        If m泸州.灰度 < deg个人支付 Then
            .补充基金自付 = 0
            .补助基金自付 = 0
            .统筹基金自付 = 0
        End If
    End With
        
    Calc基本统筹 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc基本分段1(rs费用段 As ADODB.Recordset, bln先扣起付线 As Boolean, dbl多次起付线 As Currency, dbl多次进入统筹 As Currency) As Boolean
'功能：计算按费用分段，支付封顶的情况
    Dim dbl已支付金额 As Currency  '根据参数得到已经使用的金额或费用
    Dim dbl下限 As Currency       '每一段的最低值，可能是费用，也可能是支付金额
    Dim dbl上限 As Currency       '每一段的最高值，可能是费用，也可能是支付金额
    Dim dbl分段进入 As Currency   '进入某一段的统筹金额
    Dim dbl分段报销 As Currency   '进入某一段的统筹报销金额
    Dim dbl本次进入 As Currency   '本次总的进入统筹金额
    Dim dbl本次报销 As Currency   '本次总的进入报销金额
    
    Dim dbl起点 As Currency  '用于计算的起点值
    Dim dbl剩余 As Currency  '还可以利用的统筹金额
    
    Dim dblTemp As Currency, lng档次 As Long
    Dim dbl起付线 As Currency
    
    dbl起付线 = m泸州.本次起付线
    dbl已支付金额 = m泸州.统筹已支付金额
    
    If bln先扣起付线 = True Then
        '首先把起付线金额扣除
        If m泸州.进入统筹 > dbl起付线 Then
            '足额扣除
            m泸州.实际起付线 = dbl起付线
            dbl起付线 = 0
            '因为起付线已经完成扣除，所以起点就在历次的进入统筹金额减去起付线
            If dbl多次进入统筹 > dbl多次起付线 Then
                dbl起点 = dbl多次进入统筹 - dbl多次起付线
            Else
                dbl起点 = 0
            End If
            dbl剩余 = m泸州.进入统筹 - m泸州.实际起付线
        Else
            '连起付线都不足以支付，直接退出
            m泸州.实际起付线 = m泸州.进入统筹
            
            Do Until rs费用段.EOF
                lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
                dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
                    
                gcol结算计算.Add Array(lng档次, 0, 0, dblTemp)
                rs费用段.MoveNext
            Loop
            Calc基本分段1 = True
            Exit Function
        End If
    Else
        dbl起点 = dbl多次进入统筹
        dbl剩余 = m泸州.进入统筹
    End If
    
    Do Until rs费用段.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        dbl下限 = IIf(IsNull(rs费用段("下限")), 0, rs费用段("下限"))
        dbl上限 = IIf(IsNull(rs费用段("上限")), 0, rs费用段("上限"))
        
        '支付封顶，类似于自贡那种模式
        If dbl已支付金额 < m泸州.封顶线 Or m泸州.封顶线 = 0 Then    '未超过封顶线或无封顶线
            '还可以继续报销
            If dbl下限 = 0 Then
                '这一段主要是与进行数据正确性检查，与计算无关
                If m泸州.起付线 > dbl上限 And dbl上限 > 0 Then
                    MsgBox "该病人的实际起付线比第一档费用的上限还多，请检查保险费用档。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If dbl起点 >= dbl下限 And (dbl起点 < dbl上限 Or dbl上限 = 0) And dbl剩余 > 0 Then
                '该段以前还未计算完全，求出本段需要另外扣除的金额（已经计算过的段，或者金额不到的段不会进入）
                If dbl上限 = 0 Then
                    dbl分段进入 = dbl剩余 '可全部进入
                Else
                    '在剩余值与本段空间之间选最小值
                    dbl分段进入 = dbl上限 - dbl起点
                    If dbl分段进入 > dbl剩余 Then dbl分段进入 = dbl剩余
                End If
                '起点后移，可报销金额变化
                dbl起点 = dbl起点 + dbl分段进入
                dbl剩余 = dbl剩余 - dbl分段进入
                If dbl起付线 > 0 Then
                    '不需要扣起付线就进来，这样简化
                    If dbl分段进入 > dbl起付线 Then
                        '能完成满足起付线，'除扣除起付线外还有一部分用于报销
                        m泸州.实际起付线 = m泸州.实际起付线 + dbl起付线
                        dbl分段进入 = dbl分段进入 - dbl起付线
                        dbl起付线 = 0
                    Else
                        '全部用于满足扣起付线，剩余的起付线还要用于下一段
                        m泸州.实际起付线 = m泸州.实际起付线 + dbl分段进入
                        dbl起付线 = dbl起付线 - dbl分段进入
                        dbl分段进入 = 0
                    End If
                End If
                
                '按比例求出该段的报销金额
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                dbl分段报销 = Val(Format(dbl分段进入 * rs费用段("比例") / 100, "0.00")) '这是该段最多可以报销的金额
                
                If dbl已支付金额 + dbl分段报销 > m泸州.封顶线 And m泸州.封顶线 <> 0 Then
                    '报销金额超过了封顶线，并且存在封顶线限制
                    dbl分段报销 = m泸州.封顶线 - dbl已支付金额
                    
                    '倒推进入统筹金额
                    If rs费用段("比例") <> 0 Then
                        dbl分段进入 = dbl分段报销 * 100 / rs费用段("比例")
                    Else
                        dbl分段进入 = 0
                    End If
                End If
                
            End If
        End If
        
        dbl已支付金额 = dbl已支付金额 + dbl分段报销
        
        '档次、进入统筹金额、统筹报销金额、比例
        '进行格式化
        dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
        dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
        lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
        dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        
        dbl本次进入 = dbl分段进入 + dbl本次进入
        dbl本次报销 = dbl本次报销 + dbl分段报销
        rs费用段.MoveNext
    Loop
    
    m泸州.统筹基金支付 = dbl本次报销
    m泸州.统筹基金自付 = dbl本次进入 - dbl本次报销
    
    Calc基本分段1 = True
End Function

Private Function Calc基本分段2(rs费用段 As ADODB.Recordset, bln先扣起付线 As Boolean, dbl多次起付线 As Currency, dbl多次进入统筹 As Currency, dbl多次首先自付 As Currency) As Boolean
'功能：计算按费用分段，支付封顶的情况
    Dim dbl下限 As Currency       '每一段的最低值，可能是费用，也可能是支付金额
    Dim dbl上限 As Currency       '每一段的最高值，可能是费用，也可能是支付金额
    Dim dbl分段进入 As Currency   '进入某一段的统筹金额
    Dim dbl分段报销 As Currency   '进入某一段的统筹报销金额
    Dim dbl本次进入 As Currency   '本次总的进入统筹金额
    Dim dbl本次报销 As Currency   '本次总的进入报销金额
    
    Dim dbl起点 As Currency  '用于计算的起点值
    Dim dbl剩余费用 As Currency  '还可以利用的费用
    Dim dbl剩余统筹 As Currency  '还可以利用的统筹金额
    
    Dim dblTemp As Currency, lng档次 As Long
    Dim dbl起付线 As Currency
    
    dbl起付线 = m泸州.本次起付线
    If m泸州.封顶线 > 0 Then
        '首先求出还可以使用的费用
        dbl剩余费用 = m泸州.封顶线 - m泸州.统筹已支付费用
        If dbl剩余费用 < 0 Then dbl剩余费用 = 0
        
        '再求出这部分费用中的统筹金额
        If dbl剩余费用 > m泸州.医保项目金额 Then
            dbl剩余统筹 = m泸州.医保项目金额
            dbl剩余费用 = dbl剩余费用 - m泸州.医保项目金额
            
            If dbl剩余费用 > m泸州.乙类项目金额 Then
                dbl剩余统筹 = dbl剩余统筹 + m泸州.乙类项目金额 * 0.8
            Else
                'Modified by zyb ##2003-08-31
                '超过封顶线,则在Calc基本统筹()中计算的首先自付比实际的首先自付还多了,需要重新计算
                dbl剩余统筹 = dbl剩余统筹 + dbl剩余费用 * 0.8 '这里使用一个常值
                m泸州.首先自付 = dbl剩余费用 * 0.2
            End If
        Else
            dbl剩余统筹 = dbl剩余费用
            m泸州.首先自付 = 0
        End If
    Else
        dbl剩余统筹 = m泸州.进入统筹
    End If
    
    If bln先扣起付线 = True Then
        '首先把起付线金额扣除
        If dbl剩余统筹 > dbl起付线 Then
            '足额扣除
            m泸州.实际起付线 = dbl起付线
            dbl起付线 = 0
            '因为起付线已经完成扣除，所以起点就在历次的进入统筹金额减去起付线
            If dbl多次进入统筹 > dbl多次起付线 Then
                dbl起点 = m泸州.统筹已支付费用 - dbl多次起付线 '已支付费用中包含有dbl多次进入统筹，绝对比dbl多次起付线大
            Else
                dbl起点 = m泸州.统筹已支付费用 - dbl多次进入统筹 '起码还有以前的首先自付要算
            End If
            
            'Modified By 朱玉宝 2003-12-10 地区：泸州
            If dbl起点 < 0 Then dbl起点 = 0
            dbl剩余统筹 = dbl剩余统筹 - m泸州.实际起付线
        Else
            '连起付线都不足以支付，直接退出
            m泸州.实际起付线 = dbl剩余统筹
            
            Do Until rs费用段.EOF
                lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
                dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
                    
                gcol结算计算.Add Array(lng档次, 0, 0, dblTemp)
                rs费用段.MoveNext
            Loop
            Calc基本分段2 = True
            Exit Function
        End If
    Else
        dbl起点 = m泸州.统筹已支付费用
    End If
    
    Do Until rs费用段.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        dbl下限 = IIf(IsNull(rs费用段("下限")), 0, rs费用段("下限"))
        dbl上限 = IIf(IsNull(rs费用段("上限")), 0, rs费用段("上限"))
        If dbl上限 = 0 Then dbl上限 = m泸州.封顶线 '正好是费用封顶，也就可以作为段值
        
        '还可以继续报销
        If dbl下限 = 0 Then
            '这一段主要是与进行数据正确性检查，与计算无关
            If m泸州.起付线 > dbl上限 And dbl上限 > 0 Then
                MsgBox "该病人的实际起付线比第一档费用的上限还多，请检查保险费用档。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If dbl起点 >= dbl下限 And (dbl起点 < dbl上限 Or dbl上限 = 0) And dbl剩余统筹 > 0 Then
            '该段以前还未计算完全，求出本段需要另外扣除的金额（已经计算过的段，或者金额不到的段不会进入）
            If dbl上限 = 0 Then
                dbl分段进入 = dbl剩余统筹 '可全部进入
            Else
                '在剩余值与本段空间之间选最小值
                dbl分段进入 = dbl上限 - dbl起点
                If dbl分段进入 > dbl剩余统筹 Then dbl分段进入 = dbl剩余统筹
            End If
            
            '起点后移，可使用费用变化
            dbl起点 = dbl起点 + dbl分段进入
            dbl剩余统筹 = dbl剩余统筹 - dbl分段进入
            
            If dbl起付线 > 0 Then
                '不需要扣起付线就进来，这样简化
                If dbl分段进入 > dbl起付线 Then
                    '能完成满足起付线，'除扣除起付线外还有一部分用于报销
                    m泸州.实际起付线 = m泸州.实际起付线 + dbl起付线
                    dbl分段进入 = dbl分段进入 - dbl起付线
                    dbl起付线 = 0
                Else
                    '全部用于满足扣起付线，剩余的起付线还要用于下一段
                    m泸州.实际起付线 = m泸州.实际起付线 + dbl分段进入
                    dbl起付线 = dbl起付线 - dbl分段进入
                    dbl分段进入 = 0
                End If
            End If
            
            '按比例求出该段的报销金额（进入黑名单的参保人员，其分档报销应该为零）
            dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
            If m泸州.灰度 < deg个人支付 Then
                dbl分段报销 = 0
            Else
                dbl分段报销 = Val(Format(dbl分段进入 * rs费用段("比例") / 100, "0.00")) '这是该段最多可以报销的金额
            End If
        End If
        
        '档次、进入统筹金额、统筹报销金额、比例
        '进行格式化
        dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
        dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
        lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
        dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        
        dbl本次进入 = dbl分段进入 + dbl本次进入
        dbl本次报销 = dbl本次报销 + dbl分段报销
        rs费用段.MoveNext
    Loop
    
    m泸州.统筹基金支付 = dbl本次报销
    m泸州.统筹基金自付 = dbl本次进入 - dbl本次报销
    
    Calc基本分段2 = True
End Function

Private Function Calc基本分段3(rs费用段 As ADODB.Recordset) As Boolean
'功能：计算按费用分段，支付封顶的情况
    Dim dbl已支付金额 As Currency  '根据参数得到已经使用的金额或费用
    Dim dbl下限 As Currency       '每一段的最低值，可能是费用，也可能是支付金额
    Dim dbl上限 As Currency       '每一段的最高值，可能是费用，也可能是支付金额
    Dim dbl分段进入 As Currency   '进入某一段的统筹金额
    Dim dbl分段报销 As Currency   '进入某一段的统筹报销金额
    Dim dbl本次进入 As Currency   '本次总的进入统筹金额
    Dim dbl本次报销 As Currency   '本次总的进入报销金额
    
    Dim dbl起点 As Currency  '用于计算的起点值
    Dim dbl剩余 As Currency  '还可以利用的统筹金额
    
    Dim dblTemp As Currency, lng档次 As Long
    Dim dbl起付线 As Currency
    
    dbl起付线 = m泸州.本次起付线
    dbl已支付金额 = m泸州.统筹已支付金额
    
    '首先把起付线金额扣除（因为起付线永远不能报销的，所以不能放在哪一段去判断。行就行，不行就不行。
    If m泸州.进入统筹 > dbl起付线 Then
        '足额扣除
        m泸州.实际起付线 = dbl起付线
        dbl起点 = m泸州.统筹已支付金额   '说明已经支付过了
        dbl剩余 = m泸州.进入统筹 - m泸州.实际起付线
    Else
        '连起付线都不足以支付，直接退出
        m泸州.实际起付线 = m泸州.进入统筹
        
        Do Until rs费用段.EOF
            lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
            dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
                
            gcol结算计算.Add Array(lng档次, 0, 0, dblTemp)
            rs费用段.MoveNext
        Loop
        Calc基本分段3 = True
        Exit Function
    End If
    
    Do Until rs费用段.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        dbl下限 = IIf(IsNull(rs费用段("下限")), 0, rs费用段("下限"))
        dbl上限 = IIf(IsNull(rs费用段("上限")), 0, rs费用段("上限"))
        
        '支付封顶
        If dbl已支付金额 < m泸州.封顶线 Or m泸州.封顶线 = 0 Then    '未超过封顶线或无封顶线
            '还可以继续报销
            If dbl起点 >= dbl下限 And (dbl起点 < dbl上限 Or dbl上限 = 0) And dbl剩余 > 0 Then
                '该段以前还未计算完全，求出本段需要另外扣除的金额（已经计算过的段，或者金额不到的段不会进入）
                If dbl上限 = 0 Then
                    'dbl分段报销 = dbl剩余 * rs费用段("费用段") '可全部进入
                    dbl分段报销 = dbl剩余 * rs费用段("比例") / 100
                Else
                    '在剩余值与本段空间之间选最小值
                    dbl分段报销 = dbl上限 - dbl起点
                    'If dbl分段报销 > dbl剩余 * rs费用段("费用段") Then dbl分段报销 = dbl剩余 * rs费用段("费用段")
                    If dbl分段报销 > dbl剩余 * rs费用段("比例") / 100 Then
                        dbl分段报销 = dbl剩余 * rs费用段("比例") / 100
                    End If
                End If
                '倒推求出该段可以进入的最大统筹费用
                'dbl分段进入 = dbl分段报销 / rs费用段("费用段")
                dbl分段进入 = dbl分段报销 / (rs费用段("比例") / 100)
                
                dbl起点 = dbl起点 + dbl分段报销
                dbl剩余 = dbl剩余 - dbl分段进入
                
                '按比例求出该段的报销金额
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                If m泸州.灰度 < deg个人支付 Then
                    dbl分段报销 = 0
                Else
                    dbl分段报销 = Val(Format(dbl分段进入 * rs费用段("比例") / 100, "0.00")) '这是该段最多可以报销的金额
                End If
                
                If dbl已支付金额 + dbl分段报销 > m泸州.封顶线 And m泸州.封顶线 <> 0 Then
                    '报销金额超过了封顶线，并且存在封顶线限制
                    dbl分段报销 = m泸州.封顶线 - dbl已支付金额
                    
                    '倒推进入统筹金额
                    If rs费用段("比例") <> 0 Then
                        dbl分段进入 = dbl分段报销 * 100 / rs费用段("比例")
                    Else
                        dbl分段进入 = 0
                    End If
                End If
                
            End If
        End If
        
        dbl已支付金额 = dbl已支付金额 + dbl分段报销
        
        '档次、进入统筹金额、统筹报销金额、比例
        '进行格式化
        dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
        dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
        lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
        dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        
        dbl本次进入 = dbl分段进入 + dbl本次进入
        dbl本次报销 = dbl本次报销 + dbl分段报销
        rs费用段.MoveNext
    Loop
    
    m泸州.统筹基金支付 = dbl本次报销
    m泸州.统筹基金自付 = dbl本次进入 - dbl本次报销
    
    Calc基本分段3 = True
End Function

Private Function Calc基本分段4(rs费用段 As ADODB.Recordset) As Boolean
'功能：计算按费用分段，支付封顶的情况
    Dim dbl下限 As Currency       '每一段的最低值，可能是费用，也可能是支付金额
    Dim dbl上限 As Currency       '每一段的最高值，可能是费用，也可能是支付金额
    Dim dbl分段进入 As Currency   '进入某一段的统筹金额
    Dim dbl分段报销 As Currency   '进入某一段的统筹报销金额
    Dim dbl本次进入 As Currency   '本次总的进入统筹金额
    Dim dbl本次报销 As Currency   '本次总的进入报销金额
    
    Dim dbl起点 As Currency  '用于计算的起点值
    Dim dbl剩余费用 As Currency  '还可以利用的费用
    Dim dbl剩余统筹 As Currency  '还可以利用的统筹金额
    
    Dim dblTemp As Currency, lng档次 As Long
    Dim dbl起付线 As Currency
    
    dbl起付线 = m泸州.本次起付线
    If m泸州.封顶线 > 0 Then
        '首先求出还可以使用的费用
        dbl剩余费用 = m泸州.封顶线 - m泸州.统筹已支付费用
        If dbl剩余费用 < 0 Then dbl剩余费用 = 0
        
        '再求出这部分费用中的统筹金额
        If dbl剩余费用 > m泸州.医保项目金额 Then
            dbl剩余统筹 = m泸州.医保项目金额
            dbl剩余费用 = dbl剩余费用 - m泸州.医保项目金额
            
            If dbl剩余费用 > m泸州.乙类项目金额 Then
                dbl剩余统筹 = dbl剩余统筹 + m泸州.乙类项目金额 * 0.8
            Else
                'Modified by zyb ##2003-08-31
                '超过封顶线,则在Calc基本统筹()中计算的首先自付比实际的首先自付还多了,需要重新计算
                dbl剩余统筹 = dbl剩余统筹 + dbl剩余费用 * 0.8 '这里使用一个常值
                m泸州.首先自付 = dbl剩余费用 * 0.2
            End If
        Else
            dbl剩余统筹 = dbl剩余费用
        End If
    Else
        dbl剩余统筹 = m泸州.进入统筹
    End If
    
    '首先把起付线金额扣除（因为起付线永远不能报销的，所以不能放在哪一段去判断。行就行，不行就不行。
    If dbl剩余统筹 > dbl起付线 Then
        '足额扣除
        m泸州.实际起付线 = dbl起付线
        dbl起点 = m泸州.统筹已支付金额   '说明已经支付过了
        dbl剩余统筹 = dbl剩余统筹 - m泸州.实际起付线
    Else
        '连起付线都不足以支付，直接退出
        m泸州.实际起付线 = dbl剩余统筹
        
        Do Until rs费用段.EOF
            lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
            dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
                
            gcol结算计算.Add Array(lng档次, 0, 0, dblTemp)
            rs费用段.MoveNext
        Loop
        Calc基本分段4 = True
        Exit Function
    End If
    
    Do Until rs费用段.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        dbl下限 = IIf(IsNull(rs费用段("下限")), 0, rs费用段("下限"))
        dbl上限 = IIf(IsNull(rs费用段("上限")), 0, rs费用段("上限"))
        
        '还可以继续报销
        If dbl起点 >= dbl下限 And (dbl起点 < dbl上限 Or dbl上限 = 0) And dbl剩余统筹 > 0 Then
            '该段以前还未计算完全，求出本段需要另外扣除的金额（已经计算过的段，或者金额不到的段不会进入）
            If dbl上限 = 0 Then
                'dbl分段报销 = dbl剩余统筹 * rs费用段("费用段") '可全部进入
                dbl分段报销 = dbl剩余统筹 * rs费用段("比例") / 100 '可全部进入
            Else
                '在剩余值与本段空间之间选最小值
                dbl分段报销 = dbl上限 - dbl起点
                'If dbl分段报销 > dbl剩余统筹 * rs费用段("费用段") Then dbl分段报销 = dbl剩余统筹 * rs费用段("费用段")
                If dbl分段报销 > dbl剩余统筹 * rs费用段("比例") / 100 Then
                    dbl分段报销 = dbl剩余统筹 * rs费用段("比例") / 100
                End If
            End If
            '倒推求出该段可以进入的最大统筹费用
            'dbl分段进入 = dbl分段报销 / rs费用段("费用段")
            dbl分段进入 = dbl分段报销 / (rs费用段("比例") / 100)
            
            dbl起点 = dbl起点 + dbl分段报销
            dbl剩余统筹 = dbl剩余统筹 - dbl分段进入
        End If
        
        '档次、进入统筹金额、统筹报销金额、比例
        '进行格式化
        dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
        dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
        lng档次 = IIf(IsNull(rs费用段("档次")), 0, rs费用段("档次"))
        dblTemp = IIf(IsNull(rs费用段("比例")), 0, rs费用段("比例"))
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        
        dbl本次进入 = dbl分段进入 + dbl本次进入
        dbl本次报销 = dbl本次报销 + dbl分段报销
        rs费用段.MoveNext
    Loop
    
    m泸州.统筹基金支付 = dbl本次报销
    m泸州.统筹基金自付 = dbl本次进入 - dbl本次报销
    
    Calc基本分段4 = True
End Function

Private Function Calc慢特病() As Boolean
'功能：计算出门诊慢病、大病病人的普通基本统筹金额
'输入参数：
'输出参数：
'返回：成功计算，则返回True
    
    On Error GoTo errHandle
    
    Calc慢特病 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc补充报销() As Boolean
'功能：计算出住院病人、门诊慢病或大病病人的补充报销金额
'输入参数：
'输出参数：
'返回：成功计算，则返回True
    Dim rsTemp As New ADODB.Recordset
    Dim bln费用封顶 As Boolean, dbl比例 As Currency, dbl限额 As Currency
    Dim dbl剩余医保 As Currency, dbl剩余乙类 As Currency
    Dim dbl补充费用 As Currency, dbl补充支付 As Currency, dbl剩余统筹 As Currency, dbl补充起付 As Currency
    
    m泸州.参加补充保险 = 0
    On Error GoTo errHandle
    gstrSQL = "SELECT A.开展补充保险报销,A.补充报销比例,A.补充报销限额,A.补充报销限额类型 " & _
               " FROM 保险中心目录 A " & _
               " WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & gIC泸州.CenterCode & "'"
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp("开展补充保险报销") = 0 Then
        '不开展补充保险报销业务
        Calc补充报销 = True
        Exit Function
    End If
    
    bln费用封顶 = Nvl(rsTemp("补充报销限额类型")) = 1
    dbl比例 = rsTemp("补充报销比例")
    dbl限额 = rsTemp("补充报销限额")
    
    gstrSQL = "Select * From 补充人员 Where 中心代码='" & gIC泸州.CenterCode & "' and to_number(职工编码)=" & Val(TrimStr(gIC泸州.MediAccountNo))
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '该人没有参数补充保险
        Calc补充报销 = True
        Exit Function
    End If
    
    m泸州.参加补充保险 = 1
    
    
    '得到进入被充段的费用
    With m泸州
        If bln费用封顶 = True Then
            '统筹已支付费用应该从卡上得到。因为补充保险从来都是要累计的
            '费用封顶，所以能首先得到超补充封顶线的费用
            If .统筹已支付费用 + .进入统筹 + .首先自付 > dbl限额 Then
                '超封顶线
                .超补充封顶线 = .统筹已支付费用 + .进入统筹 + .首先自付 - dbl限额
            End If
            '进入补充的费用与最多可以使用的统筹金额
            dbl补充费用 = .发生费用 - .全自费 - .进入统筹费用 - .超补充封顶线
            
            If m泸州.医保项目金额 > .进入统筹费用 Then
                '医保项目的金额已经超过了基本统筹所需要的金额
                dbl剩余医保 = m泸州.医保项目金额 - .进入统筹费用
                dbl剩余乙类 = .乙类项目金额
            Else
                dbl剩余医保 = 0
                dbl剩余乙类 = .乙类项目金额 - (.进入统筹费用 - .医保项目金额)
            End If
            
            If dbl补充费用 > dbl剩余医保 Then
                dbl剩余统筹 = dbl剩余医保
                dbl补充费用 = dbl补充费用 - dbl剩余医保
                
                If dbl补充费用 > dbl剩余乙类 Then
                    dbl剩余统筹 = dbl剩余统筹 + dbl剩余乙类 * 0.8
                Else
                    dbl剩余统筹 = dbl剩余统筹 + dbl补充费用 * 0.8 '这里使用一个常值
                End If
            Else
                dbl剩余统筹 = dbl补充费用
            End If
            
            If m政策.补充报销减起付金 = True Then
                If dbl剩余统筹 > .本次起付线 - .实际起付线 Then
                    '扣取所有起付线
                    dbl剩余统筹 = dbl剩余统筹 - (.本次起付线 - .实际起付线)
                    .实际起付线 = .本次起付线
                Else
                    '只够支付部分起付线
                    .实际起付线 = .实际起付线 + dbl剩余统筹
                    dbl剩余统筹 = 0
                End If
            End If
            
            .补充基金支付 = dbl剩余统筹 * dbl比例
            .补充基金自付 = dbl剩余统筹 - dbl剩余统筹 * dbl比例
            
        Else
            '支付封顶
            dbl剩余统筹 = .超基本封顶线 '都是统筹金额
            If m政策.补充报销减起付金 = True Then
                If dbl剩余统筹 > .本次起付线 - .实际起付线 Then
                    '扣取所有起付线
                    dbl补充起付 = (.本次起付线 - .实际起付线)
                    dbl剩余统筹 = dbl剩余统筹 - dbl补充起付
                    .实际起付线 = .本次起付线
                Else
                    '只够支付部分起付线
                    dbl补充起付 = dbl剩余统筹
                    dbl剩余统筹 = 0
                    .实际起付线 = .实际起付线 + dbl补充起付
                End If
            End If
            
            dbl补充支付 = dbl剩余统筹 * dbl比例     '支付封顶中超封顶线全是进入统筹的金额
            If dbl补充支付 > dbl限额 - .统筹已支付金额 - .统筹基金支付 Then
                '已经超过能报销的了
                dbl补充支付 = dbl限额 - .统筹已支付金额 - .统筹基金支付
                If dbl补充支付 < 0 Then dbl补充支付 = 0 '可以早已经超过限额了
                dbl补充费用 = dbl补充支付 / dbl比例
            Else
                dbl补充费用 = dbl剩余统筹
            End If
            
            .补充基金支付 = dbl补充支付
            .补充基金自付 = dbl补充费用 - dbl补充支付
            .超补充封顶线 = .超基本封顶线 - dbl补充费用 - dbl补充起付
        End If
        
        '但要记住，这要在支持补充登记的情况下才改变
        .进入统筹支付 = .进入统筹支付 + .补充基金支付
        .进入统筹费用 = .进入统筹费用 + dbl补充费用 '这一部分也要累加
    End With
    Calc补充报销 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Calc补助报销() As Boolean
'功能：计算出住院病人的补助报销金额（基本上是公务员）
'输入参数：
'输出参数：
'返回：成功计算，则返回True
    Dim dbl总费用 As Currency, dbl总自付 As Currency
    Dim dbl起始值 As Currency, dbl终止值 As Currency, dbl比例 As Currency
    Dim dbl补助自付 As Currency, dbl补助支付 As Currency
    Dim dbl分段报销 As Currency, dbl分段费用 As Currency
    Dim rs比例 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "SELECT A.开展补助报销 " & _
               " FROM 保险中心目录 A " & _
               " WHERE A.险类=" & TYPE_泸州市 & " AND A.编码='" & gIC泸州.CenterCode & "'"
    rsTemp.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    If rsTemp("开展补助报销") = 0 Then
        '不开展补充保险报销业务
        Calc补助报销 = True
        Exit Function
    End If
    
    gstrSQL = "Select 段值,比例 From 保险补助比例 Where 险类=" & TYPE_泸州市 & _
            " And 中心=" & m泸州.中心序号 & " and 年度=" & m泸州.年度 & " Order by 段值"
    rs比例.Open gstrSQL, gcn泸州, adOpenStatic, adLockReadOnly
    
    With m泸州
        If .参加补充保险 = 1 Then
            dbl总费用 = .首先自付 + .进入统筹 - .超补充封顶线
        Else
            dbl总费用 = .首先自付 + .进入统筹 - .超基本封顶线
        End If
        dbl总自付 = dbl总费用 - .统筹基金支付 - .补充基金支付
        
        '分段计算。段值是一个比例，等于 dbl总自付/dbl总费用
        Do Until rs比例.EOF
            If rs比例.AbsolutePosition = 1 Then
                '第一段只作为起始值
                dbl起始值 = dbl总费用 * rs比例("段值")
                dbl比例 = rs比例("比例")
            Else
                dbl终止值 = dbl总费用 * rs比例("段值")
                If dbl总自付 > dbl起始值 Then
                    If dbl总自付 <= dbl终止值 Then
                        dbl分段费用 = dbl总自付 - dbl起始值
                    Else
                        dbl分段费用 = dbl终止值 - dbl起始值
                    End If
                    
                    dbl分段报销 = dbl分段费用 * dbl比例
                    m泸州.补助基金支付 = m泸州.补助基金支付 + dbl分段报销
                    m泸州.补助基金自付 = m泸州.补助基金自付 + dbl分段费用 - dbl分段报销
                End If
                '作为下一段的起始值
                dbl起始值 = dbl终止值
                dbl比例 = rs比例("比例")
            End If
            
            If rs比例.AbsolutePosition = rs比例.RecordCount Then
                '最后一段
                If dbl总自付 > dbl起始值 Then
                    dbl分段费用 = dbl总自付 - dbl起始值
                    dbl分段报销 = dbl分段费用 * dbl比例
                    m泸州.补助基金支付 = m泸州.补助基金支付 + dbl分段报销
                    m泸州.补助基金自付 = m泸州.补助基金自付 + dbl分段费用 - dbl分段报销
                End If
            
            End If
            rs比例.MoveNext
        Loop
    End With
    
    Calc补助报销 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ReadIC(ByVal 病人ID As Long, ByVal 场合 As Integer, ByVal 检查卡正确 As Boolean, ByVal 错误信息 As String _
                       , ic卡 As TIC泸州, 离休病人 As Boolean) As Boolean
'功能：从读卡器、数据库、远程得到病人的信息
'输入参数：病人ID           用于判断病人是否是离休病人
'          场合             1-门诊收费、2-住院
'          检查卡正确       对于要完成写卡操作的业务，如门诊收费，就需要判断是否是该病人的卡
'          错误信息         为了更准确的显示错误信息
'输出参数：ic卡             读出后的IC卡信息
'          离休病人         当前病人是否离休人员
'返回：成功读取，返回True
    Dim str医保号 As String, str身份证号 As String, str密码 As String
    Dim lngReturn As Long
    Dim bln远程验证 As Boolean, str远程地址 As String
    
    On Error GoTo errHandle
    
    If Get保险参数_泸州(bln远程验证, str远程地址) = False Then
        Exit Function
    End If
    
    If Get帐户信息(病人ID, str医保号, str身份证号, str密码) = False Then Exit Function
    离休病人 = Is离休病人(病人ID)
    
    If 离休病人 = False Then
        If bln远程验证 = False Then
            If ReadICCard(ic卡) <> 0 Then
                MsgBox 错误信息, vbInformation, gstrSysName
                Exit Function
            End If
        Else
            gIC泸州Temp.IDCardno = str身份证号
            If frmSock沪州.CommIC(str远程地址, True, 场合, str身份证号 & "|" & str密码) = False Then
                Exit Function
            End If
            ic卡 = gIC泸州Temp
        End If
        If ic卡.InpatientFlag = "1" And 场合 = 0 Then
            MsgBox "该病人仍然在院，不能继续。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If 检查卡正确 = True Then
            '判断该病人的卡是否插入正确
            If 检查IC卡(病人ID, TrimStr(ic卡.Cardno), TrimStr(ic卡.CenterCode)) = False Then Exit Function
        End If
    Else
        If Get离休病人_泸州(str医保号, ic卡) = False Then
            Exit Function
        End If
    End If
    
    ReadIC = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function WriteIC(ByVal 离休病人 As Boolean, ByVal 收费日志 As Boolean, ByVal 场合 As Integer, ByVal Insert结算 As String, ic卡 As TIC泸州 _
    , payLog As TPayInfo, ByVal 数据体 As String) As Boolean
'功能：从读卡器、数据库、远程得到病人的信息
'输入参数：离休病人         如果是离休病人，则不进行写卡
'          收费日志         对于入院出院的写卡，就不需要写日志
'          场合             0-门诊;1-住院
'输出参数：ic卡             准备写入的IC卡信息
'          payLog           准备写入的日志信息
'返回：成功读取，返回True
    Dim lngReturn As Long
    Dim bln远程验证 As Boolean, str远程地址 As String
    
    If Get保险参数_泸州(bln远程验证, str远程地址) = False Then
        Exit Function
    End If
    
    gcn泸州.BeginTrans
    On Error GoTo errHandle
    '首先完成数据库的操作
    If Insert结算 <> "" Then gcn泸州.Execute Insert结算
    
    If 离休病人 = False Then
        '进行写卡
        If bln远程验证 = False Then
            lngReturn = WriteICCard(ic卡)
            If lngReturn <> 0 Then
                gcn泸州.RollbackTrans
                MsgBox "写入卡失败。" & 错误信息_泸州(lngReturn), vbInformation, gstrSysName
                Exit Function
            End If
            If 收费日志 = True Then
                '记录费用日志情况。这一部分信息不是太重要，即使出错，也可以忽略，而不能回滚前一次写卡
                On Error Resume Next
                lngReturn = WriteICCardPayInfo(ic卡.Cardno, payLog)
            End If
        ElseIf 数据体 <> "" Then
            '除了是远程控制外，还要是收费才行
            If frmSock沪州.CommIC(str远程地址, False, 场合, 数据体) = False Then
                gcn泸州.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
    gcn泸州.CommitTrans
    
    WriteIC = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcn泸州.RollbackTrans
End Function

Public Function Get保险参数_泸州(是否远程 As Boolean, 远程地址 As String) As Boolean
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.险类=[1] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "泸州医保", TYPE_泸州市)

    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "中心身份验证"
                是否远程 = Nvl(rsTemp("参数值")) = "是"
            Case "医保中心地址"
                远程地址 = Nvl(rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    Get保险参数_泸州 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'Modified by 朱玉宝 2004-01-07
Public Function 挂号结算_泸州(ByVal lng结帐ID As Long) As Boolean
    Dim lng病人ID As Long
    Dim cur总额 As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取病人ID
    gstrSQL = "Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    '取费用总额
    gstrSQL = "Select Sum(实收金额 ) as 金额 From 门诊费用记录 Where 结帐ID= [1] And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取费用总额", lng结帐ID)
    cur总额 = rsTemp!金额
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_泸州市 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        cur总额 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & 0 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存挂号数据")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
    
    挂号结算_泸州 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'Modified by 朱玉宝 2004-01-07
Public Function 挂号结算冲销_泸州(ByVal lng结帐ID As Long) As Boolean
    Dim lng原结帐ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    lng原结帐ID = lng结帐ID
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID", lng结帐ID)
    lng结帐ID = rsTemp("结帐ID")
    
    '读取原结算记录信息
    gstrSQL = "Select 病人ID,发生费用金额,全自付金额 from 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取原结算记录信息", lng原结帐ID)
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_泸州市 & "," & rsTemp!病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & 0 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存挂号数据")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
    
    挂号结算冲销_泸州 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Function Get当前医保年() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.医保年 From 保险主机 A,保险中心目录 B" & _
              " Where A.险类=" & TYPE_泸州市 & " And A.编码=B.主机编码 And B.序号=" & m泸州.中心序号
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn泸州
    End With
    
    Get当前医保年 = Nvl(rsTemp!医保年)
End Function

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Function Get主机IP() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select 装钱IP地址1 IP From 保险主机参数 A,保险中心目录 B " & _
             " Where A.险类=" & TYPE_泸州市 & " And A.主机=B.主机编码 And B.序号=" & m泸州.中心序号
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcn泸州
    End With
    
    Get主机IP = Nvl(rsTemp!IP)
End Function

Public Function 身份标识_泸州2(ByVal strCard As String, ByVal strPass As String, Optional lng病人ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    Dim mIC泸州 As TIC泸州
    Dim bln远程验证 As Boolean, str远程地址 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Get保险参数_泸州(bln远程验证, str远程地址) = False Then
        Exit Function
    End If
    
    If strCard <> "1" Then
        If bln远程验证 = False Then
            lngReturn = ReadICCard(mIC泸州)
        Else
            '远程连接
            If Trim(strPass) = "" Then
                Exit Function
            End If
            If frmSock沪州.CommIC(str远程地址, True, 0, strPass & "|" & strNewPass) = False Then
                Exit Function
            End If
            mIC泸州 = gIC泸州Temp
        End If
    Else
        '从离休清单中读取病人情况，填入IC卡结构中
        If Get离休病人_泸州(strPass, mIC泸州, False) = False Then
            MsgBox "没有找到该医保病人的基本信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If lngReturn <> 0 Then Exit Function
    
    'H） 密码校验是否正确：验证IC卡中Password。（如果Password为9000不对密码进行验证）。
    If TruncZero(mIC泸州.Password) <> "9000" Then
        If bln远程验证 = False Then
            If TruncZero(mIC泸州.Password) <> strPass Then
                MsgBox "密码输入错误！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '读取医保病人的ID
    gstrSQL = "Select 病人ID From 保险帐户 Where 险类=[1] And 卡号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医保病人的ID", TYPE_泸州市, CStr(mIC泸州.Cardno))
    If rsTemp.EOF Then
        MsgBox "该病人无任何费用！", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rsTemp!病人ID
    身份标识_泸州2 = lng病人ID
End Function


