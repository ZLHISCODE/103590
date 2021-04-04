Attribute VB_Name = "mdl中软"
Option Explicit
'一、IC卡函数所需结构定义
'1、基本结构:
'      1）病人信息结构       TIC中软
'      2）IC卡就医信息结构   TBlockPayInfo    （或叫支付信息）
'2、业务结构
'      1）门诊挂号读出结构   TRegisterResult
'      2）门诊收费读出结构   TChargeResult
'      3）门诊收费写入结构   TChargeParameter
'      4）住院登记读出卡结构 TInpatientRegResult
'      5）住院登记写卡结构   TInpatientRegParameter
'      6）出院结算写卡结构   TInpatientPayParameter
'      7）装钱写卡结构       TInMoneyParameter
Public Type TIC中软
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
    InExAcc          As Double          ' 补充帐累计注入金额
    OutExAcc         As Double          ' 补充帐户累计支出金额
    InSubAcc         As Double          ' 补助帐户累计注入金额
    OutSubAcc        As Double          ' 补助帐户累计支出金额
    OutAnnPlan       As Double          ' 本年统筹支付金额累计
    OutAnnOverLine   As Double          ' 本年进入统筹金额累计
    Password         As String * 4      ' 个人密码
    AnnInpatientTimes As Long           ' 本年有效住院次数
    InpatientFlag    As String * 1      ' 住院标志 0-不住院 1-住院
    HasSubInsurance  As String * 1      ' 是否参加公务员补助保险  0-否  其他-是
    HasExInsurance   As String * 1      ' 是否参加补充保险0：否1是
    HasBigIllness    As String * 1      ' 是否参加大病医保
End Type
Private Type TPayInfo
    OccurDate        As String * 8 '  就医日期
    HospitalCode     As String * 4 '  医疗机构代码
    Amount           As Double     '  本次费用合计
    AccPay           As Double     '  个人帐户支付
End Type
Private Type TBlockPayInfo
    First            As TPayInfo   ' 第一次就医信息
    Second           As TPayInfo   ' 第二次就医信息
    Third            As TPayInfo   ' 第三次就医信息
End Type
Private Type TRegisterResult
    CenterCode       As String * 4 ' 中心代码
    Cardno           As String * 8 ' 卡号
    Name             As String * 10 ' 姓名
    Sex              As String * 1 ' 性别 1-男  0-女
    Birthday         As String * 8 ' 出生日期 YYYYMMDD
    MediAccountNo    As String * 8 ' 医保号
    UnitCode         As String * 5 ' 用人单位编码
    ClassCode        As String * 2 ' 职工身份 0X-在职 1X-退休
    DomainCode       As String * 1 ' 职工属地 0-正常 1-常驻外地 2-异地安置
    Password         As String * 4 ' 个人密码
    MediYear         As String * 4 ' 医保年度
    InNo             As Long       ' 装钱期次
    InPerAcc         As Double     ' 个人帐户累计注入金额
    InExAcc          As Double     ' 补充帐累计注入金额
    InSubAcc         As Double     ' 补助帐户累计注入金额
    OutPerAcc        As Double     ' 个人帐户累计支出金额
    OutExAcc         As Double     ' 补充帐户累计支出金额
    OutSubAcc        As Double     ' 补助帐户累计支出金额
    InpatientFlag    As String * 1 ' 住院标志 0-不住院 1-住院
End Type
Private Type TChargeResult
    CenterCode       As String * 4 ' 中心代码
    Cardno           As String * 8 ' 卡号
    Name             As String * 10 ' 姓名
    Sex              As String * 1 ' 性别 1-男  0-女
    Birthday         As String * 8 ' 出生日期 YYYYMMDD
    MediAccountNo    As String * 8 ' 医保号
    UnitCode         As String * 5 ' 用人单位编码
    ClassCode        As String * 2 ' 职工类别 0X-在职 1X-退休
    DomainCode       As String * 1 ' 职工状态 0-正常 1-常驻外地
    Password         As String * 4 ' 个人密码
    MediYear         As String * 4 ' 医保年度
    InNo             As Long       ' 装钱期次
    InPerAcc         As Double     ' 个人帐户累计注入金额
    InExAcc          As Double     ' 补充帐累计注入金额
    InSubAcc         As Double     ' 补助帐户累计注入金额
    OutPerAcc        As Double     ' 个人帐户累计支出金额
    OutExAcc         As Double     ' 补充帐户累计支出金额
    OutSubAcc        As Double     ' 补助帐户累计注入金额
    OutSerialNo      As Long       ' 支付顺序号
    InpatientFlag    As String * 1 ' 住院标志
End Type
Private Type TChargeParameter
    Cardno           As String * 8 ' 卡号
    OutPerAcc        As Double     ' 个人帐户累计支出金额
    OutExAcc         As Double     ' 补充帐户累计支出金额
    OutSubAcc        As Double     ' 补助帐户累计支出金额
    OutSerialNo      As Long       ' 支付顺序号
    PayOccurDate     As String * 8 ' 日期
    PayHospitalCode  As String * 4 ' 医院代码
    PayAccPay        As Double     ' 个人帐户支付
    PayAmount        As Double     ' 总额
End Type
Private Type TInpatientRegResult
    CenterCode       As String * 4 ' 中心代码
    Cardno           As String * 8 ' 卡号
    IDCardno         As String * 18 ' 身份证号 长度不足后补#0
    MediAccountNo    As String * 8 ' 医保号
    Name             As String * 10 ' 姓名
    Sex              As String * 1 ' 性别 1-男  0-女
    Birthday         As String * 8 ' 出生日期 YYYYMMDD
    UnitCode         As String * 5 ' 用人单位编码
    ClassCode        As String * 2 ' 职工类别 0X-在职 1X-退休
    DomainCode       As String * 1 ' 职工状态 0-正常 1-常驻外地
    MediYear         As String * 4 ' 医保年度
    InNo             As Long       ' 装钱期次
    OutSerialNo      As Long       ' 支付顺序号
    InPerAcc         As Double     ' 个人帐户累计注入金额
    OutPerAcc        As Double     ' 个人帐户累计支出金额
    InExAcc          As Double     ' 补充帐累计注入金额
    OutExAcc         As Double     ' 补充帐户累计支出金额
    InSubAcc         As Double     ' 补助帐户累计注入金额
    OutSubAcc        As Double     ' 补助帐户累计支出金额
    OutAnnPlan       As Double     ' 本年统筹支付金额累计
    OutAnnOverLine   As Double     ' 本年进入统筹金额累计
    Password         As String * 4 ' 个人密码
    AnnInpatientTimes As Long       ' 本年有效住院次数
    InpatientFlag    As String * 1 ' 住院标志 0-不住院 1-住院
    HasSubInsurance  As String * 1 ' 公务员标志  0-否  其他-是
    HasExInsurance   As String * 1 ' 是否参加补充保险
    HasBigIllness    As String * 1 ' 是否参加大病医保
End Type
Private Type TInpatientRegParameter
    Cardno           As String * 8 ' 卡号
    InpatientFlag    As String * 1 ' 住院标志 0-不住院 1-住院
End Type
Private Type TInpatientPayParameter
    Cardno           As String * 8 ' 卡号
    OutPerAcc        As Double     ' 个人帐户累计支出金额
    OutExAcc         As Double     ' 补充帐户累计支出金额
    OutSubAcc        As Double     ' 补助帐户累计支出金额
    OutSerialNo      As Long       ' 支付顺序号
    OutAnnOverLine   As Double     ' 本年起付段以上基本医疗费
    OutAnnPlan       As Double     ' 本年统筹支付金额累计
    InpatientFlag    As String * 1 ' 住院标志 0-不住院 1-住院
    AnnInpatientTimes As Long       ' 本年有效住院次数
    PayOccurDate     As String * 8 ' 日期
    PayHospitalCode  As String * 4 ' 医院代码
    PayAccPay        As Double     ' 个人帐户支付
    PayAmount        As Double     ' 总额
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
Private Type TPayLog
    OccurDate        As String * 8   '  就医日期
    HospitalCode     As String * 4   '  医疗机构代码
    Amount           As String * 8   '  本次费用合计
    AccPay           As String * 8   '  个人帐户支付
End Type
'记录住院情况
Private Declare Function ChargeLog Lib "ICAPI.DLL" (payLog As TPayLog) As Long
'
''二、IC卡读写函数定义说明
''1、初始化
''      1）卡还原(将IC卡的PIN还原成初始值)
'Private Declare Function ReturnICCard Lib "ICWRITE.DLL" () As Long
''      2）制IC卡
'Private Declare Function MakeICCard Lib "ICWRITE.DLL" (iIC中软 As TIC中软) As Long
'
''2、基本读写
''      1）读IC卡病人信息
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC中软 As TIC中软) As Long
''      2）写IC卡病人信息
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC中软 As TIC中软) As Long
''      3）读IC卡就医信息
'Private Declare Function ReadICCardPayInfo Lib "ICREAD.DLL" (BlockPayInfo As TBlockPayInfo) As Long
'
''3、业务读写
''      1）挂号读卡
'Private Declare Function RegisterRead Lib "ICAPI.DLL" (RegisterResult As TRegisterResult) As Long
''      2）划价收费读卡
'Private Declare Function ChargeRead Lib "ICAPI.DLL" (ChargeResult As TChargeResult) As Long
''      3）划价收费写卡
Private Declare Function ChargeWrite Lib "ICAPI.DLL" (ChargeParameter As TChargeParameter) As Long
''      4）住院登记读卡
Private Declare Function InpatientRegRead Lib "ICAPI.DLL" (InpatientRegResult As TInpatientRegResult) As Long
''      5）住院登记写卡
Private Declare Function InpatientRegWrite Lib "ICAPI.DLL" (InpatientRegParameter As TInpatientRegParameter) As Long
''      6）住院中结、结算写卡
'Private Declare Function InpatientPayWrite Lib "ICAPI.DLL" (InpatientPayParameter As TInpatientPayParameter) As Long
'
''4、修改密码和装钱
''      1）修改密码
'Private Declare Function ChangePassword Lib "ICAPI.DLL" (Cardno As Variant, Password As Variant) As Long
''      2）新年不装钱初始化:将支出清零,累计清零,注入为卡余额,支付号加1
''         只用CardNo , MediYear字段
'Private Declare Function YearInitICCard Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
''      3）新年装钱初始化:将支出清零,累计清零,注入为传入注入金额, 支付号加1
''         用CardNo, MediYear, InNo, InPerAcc, InSubAcc, InExAcc字段.
'Private Declare Function YearInitICCardWithInMoney Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
''      4）不换医保年装钱
''         用CardNo, InNo, InPerAcc, InSubAcc, InExAcc字段.
'Private Declare Function InMoney Lib "ICAPI.DLL" (InMoneyParameter As TInMoneyParameter) As Long
'
''完成在线装钱
Private Declare Function OnLineInMoney Lib "InMoneyOnLine.dll" (ByVal IC_CenterCode As String, ByVal IC_CardNo As String, ByVal IC_MediYear As String, ByVal HosCode As String) As Long


Private Enum card医保灰度
    deg停止支付 = 1
    deg上传明细 = 2 '也停止支持
    deg个人支付 = 3 '可用个人帐户支付，统筹停
    deg医保支付 = 4 '
    deg正常支付 = 5 '不下发
End Enum

'-------------变量定义

Public gIC中软 As TIC中软                 '全局定义的存储IC卡信息的结构
Public gcn中软 As New ADODB.Connection        '连接到医保前置服务器

'-------------函数定义

Public Function 医保初始化_中软() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false

    医保初始化_中软 = True
End Function

Public Function 身份标识_中软(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回： 空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim strIdentify As String, strAddition As String
    Dim strBirthday As String, datToday As Date
    Dim lng中心 As Long, lng病种 As Long, str病种 As String
    Dim rsTemp As New ADODB.Recordset, rs病种 As ADODB.Recordset
    Dim lng灰度 As card医保灰度
    
    On Error GoTo errHandle
    
    If frmIdentify中软.GetPatient(bytType <> 2) = True Then
        '身份识别完成，返回病人信息
        With gIC中软
            lng灰度 = 医保灰度(.CenterCode, .Cardno)
            If lng灰度 = deg停止支付 Then
                MsgBox "该病人暂时停止医保支付，请到医保中心处理。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If bytType = 1 Then
                '对有限制的病人进行提醒
                If lng灰度 = deg个人支付 Or lng灰度 = deg上传明细 Then
                    MsgBox "该病人不能使用统筹基金支付住院费用。", vbExclamation, gstrSysName
                End If
            End If
            
            If bytType = 1 Then
                Dim rsSelected As New ADODB.Recordset
                gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                        " From 保险病种 A where 1=2 And A.险类=[1]"
                Set rsSelected = zlDatabase.OpenSQLRecord(gstrSQL, "获取已选择的病种", TYPE_自贡市)
                
                '住院要选择病种，以确认一些特殊收费项目
                gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                        " From 保险病种 A where A.险类=[1]"
                Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", TYPE_自贡市)
                If rs病种.RecordCount > 0 Then
VirusSelect:
                    If frm多病种选择.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：", rsSelected, False) = True Then
                        lng病种 = 0
                        str病种 = ""
                        With rs病种
                            If .RecordCount <> 0 Then .MoveFirst
                            lng病种 = rs病种("ID")
                            Do While Not .EOF
                                str病种 = str病种 & "|" & rs病种!ID
                                .MoveNext
                            Loop
                            If str病种 <> "" Then str病种 = Mid(str病种, 2)
                        End With
                    Else
                        MsgBox "必须要选择病种！", vbInformation, gstrSysName
                        GoTo VirusSelect
                    End If
                End If
            End If
            
            '建立病人档案信息，传入格式：
            '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
            '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
            '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
            strIdentify = TrimStr(.Cardno)                              '0卡号
            strIdentify = strIdentify & ";" & TrimStr(.MediAccountNo)   '1医保号
            strIdentify = strIdentify & ";" & .Password        '2密码
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
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_自贡市, .CenterCode)
            
            If rsTemp.RecordCount = 0 Then
                身份标识_中软 = ""
                MsgBox "该病人所属中心尚未建立，不能使用。", vbInformation, gstrSysName
                Exit Function
            Else
                lng中心 = rsTemp("序号")
            End If
            
            '得到原住院病种
            If bytType <> 1 Then
                gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 险类=[1] And 医保号=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "得到原住院病种", TYPE_自贡市, TrimStr(.MediAccountNo))
                If Not rsTemp.EOF Then
                    lng病种 = rsTemp!病种ID
                End If
            End If

            strAddition = ";" & lng中心                                 '8.中心代码
            strAddition = strAddition & ";"                             '9.顺序号
            strAddition = strAddition & ";" & TrimStr(.ClassCode)       '10人员身份
            strAddition = strAddition & ";" & (.InPerAcc - .OutPerAcc)  '11帐户余额
            strAddition = strAddition & ";" & .InpatientFlag            '12当前状态
            strAddition = strAddition & ";" & IIf(lng病种 > 0, lng病种, "") '13病种ID

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
            strAddition = strAddition & ";" & lng灰度                   '17灰度级
            strAddition = strAddition & ";" & .InPerAcc                 '18帐户增加累计
            strAddition = strAddition & ";" & .OutPerAcc                '19帐户支出累计
            strAddition = strAddition & ";" & .OutAnnOverLine           '20进入统筹累计
            strAddition = strAddition & ";" & .OutAnnPlan               '21统筹报销累计
            strAddition = strAddition & ";" & .AnnInpatientTimes        '22住院次数累计
            strAddition = strAddition & ";"                             '23就诊类型 (1、急诊门诊)
            
            lng病人ID = BuildPatiInfo(bytType, strIdentify & strAddition, lng病人ID, TYPE_自贡市)
            '返回格式:中间插入病人ID
            身份标识_中软 = strIdentify & ";" & lng病人ID & strAddition
            
            If bytType = 1 Then
                gstrSQL = "zl_病种信息_INSERT(" & TYPE_自贡市 & "," & lng病人ID & ",'" & str病种 & "')"
                gcn中软.Execute gstrSQL, , adCmdStoredProc
            End If
        End With
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_中软(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'参数:
'返回: 返回个人帐户余额的金额
    Dim lngReturn As Long
    
    On Error GoTo errHandle
    
    '执行装钱操作，顺便就读取了最新的个人数据
    If 装钱操作(lng病人ID) = True Then
        '检查黑名单
        If 医保灰度(gIC中软.CenterCode, gIC中软.Cardno) > deg上传明细 Then
            '返回余额
            个人余额_中软 = gIC中软.InPerAcc - gIC中软.OutPerAcc
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_中软(lng结帐ID As Long, cur个人帐户 As Currency, ByVal cur全自费 As Currency, ByVal cur首先自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rsTemp As New ADODB.Recordset
    'Dim ic门诊读卡 As TChargeResult       '门诊收费读出结构
    Dim ic门诊读卡 As TIC中软            '用上行结构，好象返回值有问题（主要是涉及金额的几个成员）
    Dim ic门诊写卡 As TChargeParameter    '门诊收费写入结构
    Dim card灰度 As card医保灰度
    Dim str医院编码 As String
    Dim lng年龄 As Long, lngReturn As Long, lng病人ID As Long
    Dim cur票据总金额 As Currency
    Dim dat当前日期 As Date
    Dim bln离休 As Boolean, str医保号 As String
    
    On Error GoTo errHandle
        
    '此时所有收费细目必然有对应的医保编码
    gstrSQL = "Select 病人ID,结帐金额  From 门诊费用记录 Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    lng病人ID = rsTemp("病人ID")
    Do Until rsTemp.EOF
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    bln离休 = Is离休病人(lng病人ID, str医保号)
    
    If bln离休 = False Then
        If ReadICCard(ic门诊读卡) <> 0 Then
            Err.Raise 9000, gstrSysName, "收费时读卡失败。"
            Exit Function
        End If
        If ic门诊读卡.InpatientFlag = "1" Then
            Err.Raise 9000, gstrSysName, "该病人仍然在院，不能继续。 "
            Exit Function
        End If
    Else
        If Get离休病人_中软(str医保号, ic门诊读卡) = False Then
            Exit Function
        End If
    End If
    
    card灰度 = 医保灰度(ic门诊读卡.CenterCode, ic门诊读卡.Cardno)
    
    If card灰度 = deg停止支付 Then
        '不用再处理后续过程
        门诊结算_中软 = True
        Exit Function
    End If
    
    dat当前日期 = zlDatabase.Currentdate
    
    '判断该病人的卡是否插入正确
    If 检查IC卡(lng病人ID, TrimStr(ic门诊读卡.Cardno), TrimStr(ic门诊读卡.CenterCode)) = False Then Exit Function
    
    With ic门诊读卡
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_自贡市 & "," & Format(dat当前日期, "yyyy") & "," & _
            .InPerAcc & "," & .OutPerAcc + cur个人帐户 & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_自贡市 & "," & lng病人ID & "," & _
            Format(dat当前日期, "yyyy") & "," & .InPerAcc & "," & .OutPerAcc & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ",0,0,0," & _
            cur票据总金额 & "," & cur全自费 & "," & cur首先自付 & "," & cur票据总金额 - cur全自费 - cur首先自付 & ",0,0,0," & _
            cur个人帐户 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    End With
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    If bln离休 = False Then
        With ic门诊写卡
            .Cardno = ic门诊读卡.Cardno         ' 卡号
            .OutPerAcc = ic门诊读卡.OutPerAcc + cur个人帐户  ' 个人帐户累计支出金额
            .OutExAcc = ic门诊读卡.OutExAcc                  ' 补充帐户累计支出金额
            .OutSubAcc = ic门诊读卡.OutSubAcc                ' 补助帐户累计支出金额
            .OutSerialNo = ic门诊读卡.OutSerialNo + 1  ' 支付顺序号
            .PayOccurDate = Format(dat当前日期, "yyyyMMdd")  ' 日期
            .PayHospitalCode = Trim(Mid(gstr医院编码, 1, 4)) ' 医院代码
            .PayAccPay = cur个人帐户      ' 个人帐户支付
            .PayAmount = cur票据总金额    ' 总额
        End With
        
        lngReturn = ChargeWrite(ic门诊写卡)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "费用写入卡失败。" & 错误信息_中软(lngReturn)
            Exit Function
        End If
    End If
        
    门诊结算_中软 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_中软(lng结帐ID As Long, cur个人帐户 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rsTemp As New ADODB.Recordset
    'Dim ic门诊读卡 As TChargeResult       '门诊收费读出结构
    Dim ic门诊读卡 As TIC中软            '用上行结构，好象返回值有问题（主要是涉及金额的几个成员）
    Dim ic门诊写卡 As TChargeParameter    '门诊收费写入结构
    Dim card灰度 As card医保灰度
    Dim lngReturn As Long, lng序号 As Long, lng病人ID As Long
    Dim cur票据总金额 As Currency, cur全自费 As Currency, cur首先自付 As Currency, cur进入统筹 As Currency
    Dim dat当前日期 As Date
    Dim bln离休 As Boolean, str医保号 As String
    Dim lng原医保年 As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 病人ID,发生费用金额,全自付金额,首先自付金额,进入统筹金额,年度  From 保险结算记录 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
        
    lng病人ID = rsTemp("病人ID")
    lng原医保年 = rsTemp("年度")
    cur票据总金额 = IIf(IsNull(rsTemp("发生费用金额")), 0, rsTemp("发生费用金额"))
    cur全自费 = IIf(IsNull(rsTemp("全自付金额")), 0, rsTemp("全自付金额")) * -1
    cur首先自付 = IIf(IsNull(rsTemp("首先自付金额")), 0, rsTemp("首先自付金额")) * -1
    cur进入统筹 = IIf(IsNull(rsTemp("进入统筹金额")), 0, rsTemp("进入统筹金额")) * -1
    
    bln离休 = Is离休病人(lng病人ID, str医保号)
    
    If bln离休 = False Then
        If ReadICCard(ic门诊读卡) <> 0 Then
            Err.Raise 9000, gstrSysName, "退费时读卡失败。"
            Exit Function
        End If
        If ic门诊读卡.InpatientFlag = "1" Then
            Err.Raise 9000, gstrSysName, "该病人仍然在院，不能继续。"
            Exit Function
        End If
    Else
        If Get离休病人_中软(str医保号, ic门诊读卡) = False Then
            Exit Function
        End If
    End If
    
    If Not Check有效期(ic门诊读卡.CenterCode) Then Exit Function
    
    card灰度 = 医保灰度(ic门诊读卡.CenterCode, ic门诊读卡.Cardno)
    If card灰度 = deg停止支付 Then
        '不用再处理后续过程
        '门诊结算冲销_中软 = True
        Exit Function
    End If
    
    gstrSQL = "select B.编码,B.序号 " & _
            " from 保险帐户 A,保险中心目录 B " & _
            " where A.病人ID=[1] and A.险类=[2]" & _
            "  and A.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng病人ID, TYPE_自贡市)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "请系统管理员完成医保中心的设置。"
        Exit Function
    End If
    '取当前医保年
    g结算数据.年度 = Val(Get保险参数_中软(rsTemp("编码"), "医保年", True))
    If g结算数据.年度 = 0 Then
        Err.Raise 9000, gstrSysName, "请系统管理员完成医保数据的下载。"
        Exit Function
    End If
    '只能冲销本医保年度的收费记录
    If lng原医保年 < g结算数据.年度 Then
       Err.Raise 9000, gstrSysName, "不能冲销非本医保年度的门诊收费记录。"
       Exit Function
    End If
    dat当前日期 = zlDatabase.Currentdate
        
    '判断该病人的卡是否插入正确
    If 检查IC卡(lng病人ID, TrimStr(ic门诊读卡.Cardno), TrimStr(ic门诊读卡.CenterCode)) = False Then Exit Function
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    lng序号 = rsTemp("结帐ID")
    
    With ic门诊读卡
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_自贡市 & "," & g结算数据.年度 & "," & _
            .InPerAcc & "," & .OutPerAcc - cur个人帐户 & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        
        gstrSQL = "zl_保险结算记录_insert(1," & lng序号 & "," & TYPE_自贡市 & "," & lng病人ID & "," & _
            g结算数据.年度 & "," & .InPerAcc & "," & .OutPerAcc - cur个人帐户 & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ",0,0,0," & _
            cur票据总金额 * -1 & "," & cur全自费 & "," & cur首先自付 & "," & cur进入统筹 & ",0,0,0," & cur个人帐户 * -1 & ",'" & .OutSerialNo + 1 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    End With
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    If bln离休 = False Then
        With ic门诊写卡
            .Cardno = ic门诊读卡.Cardno         ' 卡号
            .OutPerAcc = ic门诊读卡.OutPerAcc - cur个人帐户 ' 个人帐户累计支出金额
            .OutExAcc = ic门诊读卡.OutExAcc                  ' 补充帐户累计支出金额
            .OutSubAcc = ic门诊读卡.OutSubAcc                ' 补助帐户累计支出金额
            .OutSerialNo = ic门诊读卡.OutSerialNo + 1  ' 支付顺序号
            .PayOccurDate = Format(dat当前日期, "yyyyMMdd")  ' 日期
            .PayHospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
            .PayAccPay = cur个人帐户      ' 个人帐户支付
            .PayAmount = cur票据总金额    ' 总额
        End With
        
        lngReturn = ChargeWrite(ic门诊写卡)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "退费时写入卡失败。" & 错误信息_中软(lngReturn)
            Exit Function
        End If
    End If
    
    门诊结算冲销_中软 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人帐户转预交_中软(lng预交ID As Long, curMoney As Currency) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
           
    '由于中软医保不支持该业务，所以强行返回失败
    个人帐户转预交_中软 = False
End Function

Public Function 入院登记_中软(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim ic中软 As TIC中软
    Dim ic入院读卡 As TInpatientRegResult       '入院登记读出结构
    Dim ic入院写卡 As TInpatientRegParameter    '入院登记写入结构
    Dim lngReturn As Long
    Dim dat当前日期 As Date, card灰度 As card医保灰度
    Dim bln离休 As Boolean
    
    On Error GoTo errHandle
    
    bln离休 = Is离休病人(lng病人ID, str医保号)
    
    If bln离休 = False Then
        If InpatientRegRead(ic入院读卡) <> 0 Then
            MsgBox "入院登记时读卡失败。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Get离休病人_中软(str医保号, ic中软) = False Then
            Exit Function
        End If
        '将数据传递过来，只处理要用到的几个字段
        With ic入院读卡
            .Cardno = ic中软.Cardno
            .CenterCode = ic中软.CenterCode
        End With
    End If
        
        
    dat当前日期 = zlDatabase.Currentdate
    
    '检查刷卡器的的卡是否当前病人的
    If 检查IC卡(lng病人ID, TrimStr(ic入院读卡.Cardno), TrimStr(ic入院读卡.CenterCode)) = False Then Exit Function

    card灰度 = 医保灰度(ic入院读卡.CenterCode, ic入院读卡.Cardno)
    
    If card灰度 = deg停止支付 Then
        '不用再处理后续过程
        入院登记_中软 = False
        MsgBox "该病人已经停止医保支付，不能作为医保病人入院。", vbInformation, gstrSysName
        Exit Function
    End If

    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_自贡市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    
    If bln离休 = False Then
        With ic入院写卡
            .Cardno = ic入院读卡.Cardno         ' 卡号
            .InpatientFlag = 1
        End With
        
        lngReturn = InpatientRegWrite(ic入院写卡)
        If lngReturn Then
            MsgBox "入院登记写入卡失败。" & 错误信息_中软(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    入院登记_中软 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_中软(ByVal lng病人ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'返回：交易成功返回true；否则，返回false
    Dim ic中软 As TIC中软
    Dim ic入院读卡 As TInpatientRegResult       '入院登记读出结构
    Dim ic入院写卡 As TInpatientRegParameter    '入院登记写入结构
    Dim lngReturn As Long
    Dim bln离休 As Boolean, str医保号 As String
    
    On Error GoTo errHandle
    
    bln离休 = Is离休病人(lng病人ID, str医保号)
    
    If bln离休 = False Then
        If InpatientRegRead(ic入院读卡) <> 0 Then
            MsgBox "出院办理时读卡失败。", vbInformation, gstrSysName
            Exit Function
        End If
        '检查刷卡器的卡是否当前病人的
        If 检查IC卡(lng病人ID, TrimStr(ic入院读卡.Cardno), TrimStr(ic入院读卡.CenterCode)) = False Then Exit Function
    Else
        If Get离休病人_中软(str医保号, ic中软) = False Then
            Exit Function
        End If
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_自贡市 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    
    If bln离休 = False Then
        With ic入院写卡
            .Cardno = ic入院读卡.Cardno         '卡号
            .InpatientFlag = 0                  '表示出院
        End With
        
        lngReturn = InpatientRegWrite(ic入院写卡)
        If lngReturn <> 0 Then
            MsgBox "出院办理写入卡失败。" & 错误信息_中软(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    出院登记_中软 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_中软(rsExse As Recordset) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合
'      NO、序号、医保项目编码、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,收费类别
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim lngReturn As Long, ic病人 As TIC中软
    Dim bln离休 As Boolean, str医保号 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If 检查医保服务器_中软 = False Then
        '不能连接到前置服务器，就认为不可使用
        Exit Function
    End If
    
    bln离休 = Is离休病人(rsExse("病人ID"), str医保号)
    
    If bln离休 = False Then
        lngReturn = ReadICCard(ic病人)
        If lngReturn <> 0 Then
            MsgBox "读卡信息失败。" & 错误信息_中软(lngReturn), vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Get离休病人_中软(str医保号, ic病人) = False Then Exit Function
    End If
    
    '完成一些数据的初始化，黑名单人员也要使用的数据
    With g结算数据
        .病人ID = rsExse("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", CLng(rsExse("病人ID")))
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
    
        '避免在出院结帐后再次进行结帐
        gstrSQL = "SELECT 病人ID FROM 保险结算记录 WHERE 中途结帐=0 AND 病人ID=[1] AND 主页ID=[2] AND 险类=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", .病人ID, .主页ID, TYPE_自贡市)
        
        If rsTemp.RecordCount > 0 Then
            MsgBox "病人已经进行过住院结算，不能再进行结帐操作。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    
    '目前只是自贡医保使用该参数
    '年度使用保险参数中定义的（因此只要没有下载，医院就还在以前的年度上处理）
    gstrSQL = "select A.病种ID,B.编码,B.序号 " & _
            " from 保险帐户 A,保险中心目录 B " & _
            " where A.病人ID=[1] and A.险类=[2]" & _
            "  and A.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, TYPE_自贡市)
    If rsTemp.EOF = True Then
        MsgBox "请系统管理员完成医保中心的设置。", vbInformation, gstrSysName
        Exit Function
    End If
    If Nvl(rsTemp!病种ID) = 0 Then
        MsgBox "没有选择病种，不允许结帐！", vbInformation, gstrSysName
        Exit Function
    End If
    
    g结算数据.年度 = Val(Get保险参数_中软(rsTemp("编码"), "医保年", True))
    If g结算数据.年度 = 0 Then
        MsgBox "请系统管理员完成医保数据的下载。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '1.2 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID)
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
    End If

    '此处使用装钱操作，主要目的是初始化病人的卡上的余额，以及累计进入统筹和统筹累计报销
    If 装钱操作(rsExse("病人ID")) = False Then
        MsgBox "病人装钱操作失败，无法准确得到病人的余额与累计报销金额。", vbInformation, gstrSysName
        Exit Function
    End If
    
    With gIC中软
        gstrSQL = "zl_帐户年度信息_insert(" & rsExse("病人ID") & "," & TYPE_自贡市 & "," & .MediYear & "," & _
            .InPerAcc & "," & .OutPerAcc & "," & .OutAnnOverLine & "," & _
            .OutAnnPlan & "," & .AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    End With
    '其余的算法与中联医保相同
    住院虚拟结算_中软 = 住院虚拟结算(rsExse, 医保灰度(ic病人.CenterCode, ic病人.Cardno))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function 住院虚拟结算(rs费用明细 As Recordset, ByVal deg灰度 As card医保灰度) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '结算要求：NO、序号、病人ID、医保项目编码、收费类别、收费名称、开单部门、规格、产地、数量、价格、金额、医生,登记时间(发生时间),婴儿费,保险大类ID
    Dim rs大类汇总 As Recordset     '按医保支付大类汇总得到
    Dim rs特准项目大类 As New ADODB.Recordset
    Dim rs算法 As New ADODB.Recordset, rs大类 As New ADODB.Recordset         '保存
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng大类ID As Long
    Dim lng中心 As Long, str中心 As String
    Dim lng在职 As Long, lng年龄段 As Long, lng年龄 As Long
    Dim dblTemp As Double, lng档次 As Long
    
    Dim dbl最大金额  As Double ''对一个按住院日计算的项目，最多能得到的金额
    Dim dbl已报销金额 As Double, dbl累计进入 As Double
    Dim dbl下限 As Double, dbl上限 As Double, dbl分段进入 As Double, dbl分段报销 As Double
    
    Dim cls医保 As New clsInsure
    Dim bln个人帐户支付全自费 As Boolean, bln个人帐户支付首先自付 As Boolean, bln个人帐户支付超限 As Boolean
    Dim cur全自费 As Currency, cur首先自付 As Currency
    Dim bln全额统筹 As Boolean, bln无起付线 As Boolean, bln无封顶线 As Boolean
    
    Dim bln跨年结算 As Boolean   '对于自贡医保，如果是跨年结算，即使该病人是第二次结帐。各分段计算也是从头开始
    Dim dbl多次起付线和 As Double, dbl多次进入统筹和 As Double   '多次是指该病人以前结帐的累计
    Dim dbl计算起付线 As Double, dbl本次起付线 As Double
    Dim lng原医保年 As Long
    
    On Error GoTo errHandle
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '1、初始化一些变量
    Set gcol结算计算 = New Collection
    
'    gstrSQL = "select D.ID 大类ID,A.收费细目ID " & _
'             " from  " & _
'             " (select C.收费细目ID " & _
'             " from 保险病种 A,ZLYB.病种信息 B,保险特准项目 C " & _
'             " Where A.险类=" & TYPE_自贡市 & " And A.险类=B.险类 And B.病人ID=" & g结算数据.病人ID & " And Nvl(C.性质,0)=0 And Nvl(C.性质,0)<>2 And C.病种ID=B.病种ID And B.病种ID=A.ID) A, " & _
'             " 保险项目 B,保险支付项目  C,保险支付大类 D " & _
'             " Where A.收费细目ID=C.收费细目ID And C.险类=B.险类 And B.编码=C.项目编码 And B.大类编码=D.编码 And B.险类=" & TYPE_自贡市 & _
'             " And D.险类=B.险类"
'    Call OpenRecordset(rs特准项目大类, "获取该病人所有病种的批准项目大类")
    
    bln个人帐户支付全自费 = cls医保.GetCapability(support结算帐户全自费, 0, TYPE_自贡市)
    bln个人帐户支付首先自付 = cls医保.GetCapability(support结算帐户首先自付, 0, TYPE_自贡市)
    bln个人帐户支付超限 = cls医保.GetCapability(support结算帐户超限, 0, TYPE_自贡市)
    
    gstrSQL = "select B.编码,B.序号 " & _
            " from 保险帐户 A,保险中心目录 B " & _
            " where A.病人ID=[1] and A.险类=[2]" & _
            "  and A.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, TYPE_自贡市)
    If rsTemp.EOF = True Then
        MsgBox "请系统管理员完成医保中心的设置。", vbInformation, gstrSysName
        Exit Function
    End If
    lng中心 = rsTemp("序号")
    str中心 = rsTemp("编码")
    
    gstrSQL = "select max(年度) as 年度 from 保险结算记录 where 病人id=[1] and 主页id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID)
    If rsTemp.EOF = False Then
       lng原医保年 = IIf(IsNull(rsTemp("年度")) = True, g结算数据.年度, rsTemp("年度"))
    Else
       lng原医保年 = g结算数据.年度
    End If
    
    'If g结算数据.年度 > Val(Format(rs费用明细("发生时间"), "yyyy")) Then
    If g结算数据.年度 > lng原医保年 Then
        bln跨年结算 = True
    End If
        
    '1.3 读出本次住院期间累计结帐情况
    gstrSQL = "select nvl(sum(A.起付线),0) as 起付线,nvl(sum(A.进入统筹金额),0) as 进入统筹金额 " & _
              "  from 保险结算记录 A,病人结帐记录 B " & _
              "  Where A.病人ID = [1] And A.主页ID = [2]" & _
              " And A.险类 = [3] And A.记录ID = B.ID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", g结算数据.病人ID, g结算数据.主页ID, TYPE_自贡市)
    dbl多次起付线和 = rsTemp("起付线")
    dbl多次进入统筹和 = rsTemp("进入统筹金额")
    
    With g结算数据
        gstrSQL = "select A.中心,A.人员身份,A.在职,A.年龄段," & _
                  "      B.住院次数累计,B.帐户增加累计,B.帐户支出累计,B.进入统筹累计,B.统筹报销累计" & _
                  " from 保险帐户 A,帐户年度信息 B" & _
                  " where A.病人ID=B.病人ID(+) and A.险类=B.险类(+) " & _
                  "     and B.年度(+)=[1] and A.病人ID=[2] and A.险类=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "虚拟结算", .年度, .病人ID, TYPE_自贡市)
        
        lng中心 = IIf(IsNull(rsTemp("中心")), 0, rsTemp("中心"))
        lng在职 = IIf(IsNull(rsTemp("在职")), 1, rsTemp("在职"))
        lng年龄 = IIf(IsNull(rsTemp("年龄段")), 0, rsTemp("年龄段"))
        .住院次数 = IIf(IsNull(rsTemp("住院次数累计")), 0, rsTemp("住院次数累计"))
        .帐户累计增加 = IIf(IsNull(rsTemp("帐户增加累计")), 0, rsTemp("帐户增加累计"))
        .帐户累计支出 = IIf(IsNull(rsTemp("帐户支出累计")), 0, rsTemp("帐户支出累计"))
        .累计进入统筹 = IIf(IsNull(rsTemp("进入统筹累计")), 0, rsTemp("进入统筹累计"))
        .累计统筹报销 = IIf(IsNull(rsTemp("统筹报销累计")), 0, rsTemp("统筹报销累计"))
    
        
        gstrSQL = "select 年龄段,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
                " from 保险年龄段" & _
                " where 险类=" & TYPE_自贡市 & " and nvl(中心,0)=" & lng中心 & _
                "       and 在职=" & lng在职 & " and 下限<=" & lng年龄 & " and (" & lng年龄 & "<=上限 or 上限=0)"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "请在“保险类别管理”中设置年龄段与费用档。", vbInformation, gstrSysName
            Exit Function
        End If
        lng年龄段 = rsTemp("年龄段")
        bln全额统筹 = (rsTemp("全额统筹") = 1)
        bln无起付线 = (rsTemp("无起付线") = 1)
        bln无封顶线 = (rsTemp("无封顶线") = 1)
    End With
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '2、按统筹支付项目合计发生金额和数量
    '2.1、初始化记录集
    Set rs大类汇总 = New ADODB.Recordset
    With rs大类汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "保险大类ID", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 8, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "统筹金额", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Do Until rs费用明细.EOF
    '装数据写入记录集，供其它窗体使用
        If rs费用明细("保险项目否") = 1 Then
'            rs特准项目大类.Filter = "收费细目ID=" & rs费用明细!收费细目ID
'            If rs特准项目大类.EOF Then
'                lng大类ID = rs费用明细!保险大类ID
'            Else
'                lng大类ID = rs特准项目大类!大类ID
'            End If
'            '更新费用明细
'            gstrSQL = ""
'            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保大类ID")
            
            lng大类ID = rs费用明细!保险大类id
            If rs大类汇总.RecordCount = 0 Then
                rs大类汇总.AddNew
                rs大类汇总("保险大类ID") = lng大类ID
                rs大类汇总("数量") = rs费用明细("数量")
                rs大类汇总("金额") = rs费用明细("金额")
            Else
                rs大类汇总.MoveFirst
                rs大类汇总.Find "保险大类ID=" & lng大类ID
                If rs大类汇总.EOF Then
                    rs大类汇总.AddNew
                    rs大类汇总("保险大类ID") = lng大类ID
                    rs大类汇总("数量") = rs费用明细("数量")
                    rs大类汇总("金额") = rs费用明细("金额")
                Else
                    rs大类汇总("数量") = rs大类汇总("数量") + rs费用明细("数量")
                    rs大类汇总("金额") = rs大类汇总("金额") + rs费用明细("金额")
                End If
            End If
            rs大类汇总.Update
        Else
            cur全自费 = cur全自费 + rs费用明细("金额")
        End If
            
        dblTemp = dblTemp + rs费用明细("金额")
        rs费用明细.MoveNext
    Loop
    g结算数据.发生费用金额 = dblTemp
    
    '2.2、计算进入统筹金额
    gstrSQL = "select ID,编码,算法,统筹比额,特准定额,特准天数,是否医保 FROM 保险支付大类  where 险类=[1]"
    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_自贡市)
    rs算法.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
    
    dblTemp = 0
    If rs大类汇总.RecordCount > 0 Then rs大类汇总.MoveFirst
    Do Until rs大类汇总.EOF
        
        rs大类.Filter = "ID=" & rs大类汇总("保险大类ID")
        If rs大类.EOF = False Then
            rs算法.Filter = "编码='" & rs大类("编码") & "'"
        Else
            rs算法.Filter = "编码='90009'"
        End If
        If rs算法.RecordCount > 0 Then
            If rs算法("是否医保") = 1 Then
                '算法:1-总额计算项目；2-住院日核定项目
                If rs算法("算法") = 1 Then
                    If rs算法("统筹比额") = 0 Then
                        cur全自费 = cur全自费 + rs大类汇总("金额")
                    Else
                        dblTemp = dblTemp + rs大类汇总("金额") * rs算法("统筹比额") / 100
                    End If
                Else
                    If Val(rs大类汇总("数量")) > Val(rs算法("特准天数")) Then
                        '如果住院日超过特准天数，那么最大金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                        dbl最大金额 = rs算法("特准定额") * rs算法("特准天数") + _
                            (rs大类汇总("数量") - IIf(rs算法("特准定额") = 0 Or rs算法("特准天数") = 0, 0, rs算法("特准天数"))) * rs算法("统筹比额")
                    Else
                        '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                        If rs算法("特准定额") = 0 Or rs算法("特准天数") = 0 Then
                            dbl最大金额 = rs大类汇总("数量") * rs算法("统筹比额")
                        Else
                            dbl最大金额 = rs大类汇总("数量") * rs算法("特准定额")
                        End If
                    End If
                    
                    '总金额比最大金额小，就取全部金额；否则只最大金额
                    dblTemp = dblTemp + IIf(rs大类汇总("金额") < dbl最大金额, rs大类汇总("金额"), dbl最大金额)
                    
                    If rs大类汇总("金额") > dbl最大金额 Then
                        '全部算作全自费
                        cur全自费 = cur全自费 + rs大类汇总("金额") - dbl最大金额
                    End If
                End If
            Else
                cur全自费 = cur全自费 + rs大类汇总("金额")
            End If
        Else
            cur全自费 = cur全自费 + rs大类汇总("金额")
        End If
        rs大类汇总.MoveNext
    Loop
    g结算数据.进入统筹金额 = dblTemp
    g结算数据.全自费金额 = cur全自费
    g结算数据.首先自付金额 = g结算数据.发生费用金额 - cur全自费 - dblTemp
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '3、获得起付线、封顶线、支付比例等数据
    '3.1、获得起付线、封顶线
    With g结算数据
        
        gstrSQL = "select max(decode(A.性质,'A',A.金额,0)) as 封项线 ,max(decode(A.性质,'1',A.金额,0)) as 起付线 " & _
                  "         ,max(decode(A.性质,'" & (.住院次数 + 1) & "',A.金额,0)) as 实际起付线,min(A.金额) as 最低起付线 " & _
                  "  from 保险支付限额 A " & _
                  "  where A.险类=" & TYPE_自贡市 & " and A.中心=" & lng中心 & " and A.年度=" & .年度
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
                
        If bln无起付线 Then
            .实际起付线 = 0
            .起付线 = 0
        Else
            .起付线 = IIf(IsNull(rsTemp("实际起付线")), 0, rsTemp("实际起付线"))
            If .起付线 = 0 Then
                '一般都会有，如果实在超过了住院次数，就取最后一次（也就是金额最小的一次）
                .起付线 = IIf(IsNull(rsTemp("最低起付线")), 0, rsTemp("最低起付线"))
            End If
            If .起付线 = 0 Then
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
    
        '3.2、根据以前扣除的起付线金额，得出本次的实际起付线
        If dbl多次起付线和 > 0 Then
            '表明该病人肯定有多次结帐
            
            If dbl多次起付线和 > dbl多次进入统筹和 Then
                '该病人的本次结算还要扣除一部分起付线金额
                dbl计算起付线 = dbl多次起付线和 - dbl多次进入统筹和
            Else
                '起付线已经扣完
                dbl计算起付线 = 0
            End If
            
            If .起付线 > dbl多次起付线和 Then
                '调高了起付线，要补这段差值
                .起付线 = .起付线 - dbl多次起付线和
            Else
                '以前的起付线金额已经全额保存，本次不用再保存了
                .起付线 = 0
            End If
            
            dbl计算起付线 = dbl计算起付线 + .起付线
        Else
            dbl计算起付线 = .起付线
        End If
        dbl本次起付线 = dbl计算起付线
    End With
    
    '3.3、取得费用档次
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "select B.档次,B.下限,B.上限,A.比例 " & _
              "  from 保险支付比例 A,保险费用档 B " & _
              "  Where A.险类 =" & TYPE_自贡市 & " And A.中心 =" & lng中心 & " And A.年度 =" & g结算数据.年度 & " And A.在职 =" & lng在职 & " And A.年龄段 =" & lng年龄段 & _
              "       and A.险类=B.险类 and A.中心=b.中心 and A.档次=B.档次 " & _
              "  order by B.档次"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        MsgBox "请在“年度结算规则”中设置本年度的统筹支付比例。。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    '4、计算该次结算可报销的金额
    dbl累计进入 = 0   '保存分段累计进入统筹
    dbl已报销金额 = g结算数据.累计统筹报销
    g结算数据.统筹报销金额 = 0
    
    If bln跨年结算 = True Then
        '跨年结算就不用考虑以前的结算金额
        dbl多次进入统筹和 = 0
    End If
    Do Until rsTemp.EOF
        dbl分段进入 = 0
        dbl分段报销 = 0
        
        If dbl已报销金额 < g结算数据.封顶线 Or g结算数据.封顶线 = 0 Then    '未超过封顶线或无封顶线
            '还可以继续报销
            dbl下限 = IIf(IsNull(rsTemp("下限")), 0, rsTemp("下限"))
            dbl上限 = IIf(IsNull(rsTemp("上限")), 0, rsTemp("上限"))
            If dbl下限 = 0 Then
                If g结算数据.起付线 > dbl上限 Then
                    MsgBox "该病人的实际起付线比第一档费用的上限还多，请检查保险费用档。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If g结算数据.进入统筹金额 + dbl多次进入统筹和 > dbl下限 And (dbl多次进入统筹和 < dbl上限 Or dbl上限 = 0) Then
                '该段以前还未计算完全，求出本段需要另外扣除的金额
                dblTemp = 0
                If dbl多次进入统筹和 > dbl下限 Then
                    '以前已经计算过的
                    dblTemp = dbl多次进入统筹和 - dbl下限
                End If
                
                '由于要扣除一部分起付线和已结金额，所以下限金额会有变化
                If dbl下限 + dblTemp + dbl计算起付线 > dbl上限 And dbl上限 > 0 Then
                    dbl下限 = dbl上限
                    dbl计算起付线 = dbl计算起付线 - (dbl上限 - dbl下限 - dblTemp) '本段已经扣完，留着下段扣
                Else
                    dbl下限 = dbl下限 + dbl计算起付线 + dblTemp
                    dbl计算起付线 = 0
                End If
                
                If g结算数据.进入统筹金额 + dbl多次进入统筹和 <= dbl上限 Or dbl上限 = 0 Then
                    '按实际值进入
                    dbl分段进入 = g结算数据.进入统筹金额 + dbl多次进入统筹和 - dbl下限
                    
                    '如果由于加上起付线、或以前的结帐金额，导致进入统筹的金额还不能达到下限，那只能取0
                    If dbl分段进入 < 0 Then dbl分段进入 = 0
                Else
                    '全额进入
                    dbl分段进入 = dbl上限 - dbl下限
                End If
                '按比例求出该段的报销金额
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                dbl分段报销 = Val(Format(dbl分段进入 * rsTemp("比例") / 100, "0.00"))
                
                If dbl已报销金额 + dbl分段报销 > g结算数据.封顶线 And g结算数据.封顶线 <> 0 Then
                    '报销金额超过了封顶线，并且存在封顶线限制
                    dbl分段报销 = g结算数据.封顶线 - dbl已报销金额
                    
                    '倒推进入统筹金额
                    If rsTemp("比例") <> 0 Then
                        dbl分段进入 = dbl分段报销 * 100 / rsTemp("比例")
                    Else
                        dbl分段进入 = 0
                    End If
                End If
                
                '进行格式化
                dbl分段进入 = Val(Format(dbl分段进入, "0.00"))
                dbl分段报销 = Val(Format(dbl分段报销, "0.00"))
                
                dbl已报销金额 = dbl已报销金额 + dbl分段报销
                g结算数据.统筹报销金额 = g结算数据.统筹报销金额 + dbl分段报销
            End If
        End If
        
        '档次、进入统筹金额、统筹报销金额、比例
        lng档次 = IIf(IsNull(rsTemp("档次")), 0, rsTemp("档次"))
        dblTemp = IIf(IsNull(rsTemp("比例")), 0, rsTemp("比例"))
        dbl累计进入 = dbl分段进入 + dbl累计进入
            
        gcol结算计算.Add Array(lng档次, dbl分段进入, dbl分段报销, dblTemp)
        rsTemp.MoveNext
    Loop
    
    g结算数据.实际起付线 = dbl本次起付线 - dbl计算起付线
    
    With g结算数据
        '计算超限自付部分
        .超限自付金额 = .进入统筹金额 - dbl本次起付线 - dbl累计进入
        If .超限自付金额 < 0 Then .超限自付金额 = 0                   '如果进入统筹金额还不到起付线，为负数
    End With
    
    If deg灰度 < deg医保支付 Then
        '不能用医保基金支付
        g结算数据.统筹报销金额 = 0
        g结算数据.超限自付金额 = 0
        
        住院虚拟结算 = "医保基金;" & g结算数据.统筹报销金额 & ";0"
    Else
        If bln全额统筹 = True Then
            住院虚拟结算 = "医保基金;" & g结算数据.统筹报销金额 + g结算数据.首先自付金额 & ";0"
        Else
            住院虚拟结算 = "医保基金;" & g结算数据.统筹报销金额 & ";0"
        End If
    End If
    
    '还需要考虑个人帐户的支付范围
    With g结算数据
        dblTemp = 0   '暂时保存可使用的个人帐户余额
        
        If bln个人帐户支付全自费 = True Then
            dblTemp = dblTemp + .全自费金额
        End If
        
        If bln个人帐户支付首先自付 = True And bln全额统筹 = False Then
            dblTemp = dblTemp + .首先自付金额
        End If
        
        If bln个人帐户支付超限 = True Then
            '只能支付进入统筹，但未报销的部分
            dblTemp = dblTemp + .进入统筹金额 - .统筹报销金额
        Else
            dblTemp = dblTemp + .进入统筹金额 - .统筹报销金额 - .超限自付金额
        End If
        
        If deg灰度 >= deg个人支付 Then
            If .帐户累计增加 - .帐户累计支出 - dblTemp > 0 Then
               住院虚拟结算 = 住院虚拟结算 & "|个人帐户;" & dblTemp & ";1"
            Else
               住院虚拟结算 = 住院虚拟结算 & "|个人帐户;" & IIf(.帐户累计增加 - .帐户累计支出 > 0, .帐户累计增加 - .帐户累计支出, 0) & ";1"
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    住院虚拟结算 = ""
End Function

Public Function 住院结算_中软(lng结帐ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID     病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim ic住院读卡 As TIC中软               '住院结算读出结构
    Dim ic住院写卡 As TIC中软               '住院结算读出结构
    Dim card灰度 As card医保灰度
    Dim lngReturn As Long
    Dim bln离休 As Boolean, str医保号 As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim cur个人帐户 As Currency, var结算计算 As Variant
    
    On Error GoTo errHandle
    
    bln离休 = Is离休病人(g结算数据.病人ID, str医保号)
    If bln离休 = False Then
        If ReadICCard(ic住院读卡) <> 0 Then
            Err.Raise 9000, gstrSysName, "结算时读卡失败。"
            Exit Function
        End If
        
        '判断该病人的卡是否插入正确
        If 检查IC卡(g结算数据.病人ID, TrimStr(ic住院读卡.Cardno), TrimStr(ic住院读卡.CenterCode)) = False Then Exit Function
    Else
        If Get离休病人_中软(str医保号, ic住院读卡) = False Then Exit Function
    End If
    
    If Not Check有效期(ic住院读卡.CenterCode) Then Exit Function
    card灰度 = 医保灰度(ic住院读卡.CenterCode, ic住院读卡.Cardno)
    
'    If card灰度 = deg停止支付 Then
'        '不用再处理后续过程
'        住院结算_中软 = True
'        Exit Function
'    End If
        
    '求个人帐户支付金额
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    If Not rsTemp.EOF Then cur个人帐户 = rsTemp!金额
    
    
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    If g结算数据.中途结帐 = 0 Then
        '表示该病人已经出院
        ic住院读卡.AnnInpatientTimes = ic住院读卡.AnnInpatientTimes + 1
    End If
        
    With g结算数据
        '为了保证安全，累计数据还是读卡中数据

        gstrSQL = "zl_帐户年度信息_insert(" & .病人ID & "," & TYPE_自贡市 & "," & .年度 & "," & _
            ic住院读卡.InPerAcc & "," & ic住院读卡.OutPerAcc + cur个人帐户 & "," & ic住院读卡.OutAnnOverLine + .进入统筹金额 & "," & _
            ic住院读卡.OutAnnPlan + .统筹报销金额 & "," & ic住院读卡.AnnInpatientTimes & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_自贡市 & "," & .病人ID & "," & _
            .年度 & "," & ic住院读卡.InPerAcc & "," & ic住院读卡.OutPerAcc & "," & ic住院读卡.OutAnnOverLine & "," & _
            ic住院读卡.OutAnnPlan & "," & ic住院读卡.AnnInpatientTimes & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
            .发生费用金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
            .超限自付金额 & "," & cur个人帐户 & ",'" & ic住院读卡.OutSerialNo + 1 & "'," & .主页ID & "," & .中途结帐 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        
        For Each var结算计算 In gcol结算计算
            '依次为档次、进入统筹金额、统筹报销金额、比例
            gstrSQL = "zl_保险结算计算_Insert(" & lng结帐ID & "," & _
                var结算计算(0) & "," & var结算计算(1) & "," & var结算计算(2) & "," & var结算计算(3) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        Next
    End With
    
    If bln离休 = False Then
        With ic住院读卡
            .Cardno = ic住院读卡.Cardno         ' 卡号
            .OutPerAcc = ic住院读卡.OutPerAcc + cur个人帐户         ' 个人帐户累计支出金额
            .OutExAcc = ic住院读卡.OutExAcc                         ' 补充帐户累计支出金额
            .OutSubAcc = ic住院读卡.OutSubAcc                       ' 补助帐户累计支出金额
            .OutSerialNo = ic住院读卡.OutSerialNo + 1               ' 支付顺序号
            .OutAnnOverLine = ic住院读卡.OutAnnOverLine + g结算数据.进入统筹金额  ' 本年进入统筹金额累计
            .OutAnnPlan = ic住院读卡.OutAnnPlan + g结算数据.统筹报销金额          ' 本年统筹支付金额累计
            .InpatientFlag = ic住院读卡.InpatientFlag                             ' 住院标志 0-不住院 1-住院
            .AnnInpatientTimes = ic住院读卡.AnnInpatientTimes                     ' 本年有效住院次数
'            .PayOccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
'            .PayHospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
'            .PayAccPay = cur个人帐户      ' 个人帐户支付
'            .PayAmount = g结算数据.发生费用金额    ' 总额
        End With
        
        lngReturn = WriteICCard(ic住院读卡)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "费用写入卡失败。" & 错误信息_中软(lngReturn)
            Exit Function
        End If
        
        '记录住院情况
        Dim payLog As TPayLog
        With payLog
            .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
            .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
            .AccPay = Space(8 - Len(CStr(cur个人帐户 * 100))) & CStr(cur个人帐户 * 100)
            .Amount = Space(8 - Len(CStr(g结算数据.发生费用金额 * 100))) & CStr(g结算数据.发生费用金额 * 100)
        End With
        ChargeLog payLog
    End If
        
    住院结算_中软 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_中软(lng结帐ID As Long) As Boolean
'----------------------------------------------------------------
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID-需要作废的结帐单ID号；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'----------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rs结算计算 As New ADODB.Recordset
    Dim ic住院读卡 As TIC中软                '住院结算读出结构
    Dim ic住院写卡 As TIC中软                '住院结算写入结构
    Dim card灰度 As card医保灰度
    Dim lng冲销ID As Long, lngReturn As Long
    Dim bln离休 As Boolean, str医保号 As String
    Dim cur个人帐户 As Currency
    Dim lng病人ID As Long, lng原医保年 As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "select distinct A.ID,A.病人id from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    lng病人ID = rsTemp("病人id")
    gstrSQL = "select B.编码,B.序号 " & _
            " from 保险帐户 A,保险中心目录 B " & _
            " where A.病人ID=[1] and A.险类=[2]" & _
            "  and A.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng病人ID, TYPE_自贡市)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "请系统管理员完成医保中心的设置。"
        Exit Function
    End If
    '取当前医保年
    g结算数据.年度 = Val(Get保险参数_中软(rsTemp("编码"), "医保年", True))
    If g结算数据.年度 = 0 Then
        Err.Raise 9000, gstrSysName, "请系统管理员完成医保数据的下载。"
        Exit Function
    End If
    
    '只允许对中途结帐进行作废
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 中途结帐=1 and 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算不是中途结帐，不能作废。"
        Exit Function
    End If
    lng原医保年 = rsTemp("年度")
    '只能冲销本医保年度的医保结算记录
    If lng原医保年 < g结算数据.年度 Then
       Err.Raise 9000, gstrSysName, "不能冲销非本医保年度的医保结算记录。"
       Exit Function
    End If
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "该病人的医保结算数据丢失，不能作废。"
        Exit Function
    End If
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    bln离休 = Is离休病人(rsTemp("病人ID"), str医保号)
    If bln离休 = False Then
        If ReadICCard(ic住院读卡) <> 0 Then
            Err.Raise 9000, gstrSysName, "结算时读卡失败。"
            Exit Function
        End If
    Else
        If Get离休病人_中软(str医保号, ic住院读卡) = False Then Exit Function
    End If
    
    If Not Check有效期(ic住院读卡.CenterCode) Then Exit Function
    
    card灰度 = 医保灰度(ic住院读卡.CenterCode, ic住院读卡.Cardno)
    If card灰度 = deg停止支付 Then
        '不用再处理后续过程
        住院结算冲销_中软 = False
        Err.Raise 9000, gstrSysName, "该病人已经停止医保支付，不能进行冲销操作。"
        Exit Function
    End If
    
    
    '判断该病人的卡是否插入正确
    If 检查IC卡(rsTemp("病人ID"), TrimStr(ic住院读卡.Cardno), TrimStr(ic住院读卡.CenterCode)) = False Then Exit Function
    
    '将此处的数据保存与主程序的数据保存想成一个事务
    '因此就不需要单独的事务控制
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & TYPE_自贡市 & "," & rsTemp("年度") & "," & _
        ic住院读卡.InPerAcc & "," & ic住院读卡.OutPerAcc - rsTemp("个人帐户支付") & "," & ic住院读卡.OutAnnOverLine - rsTemp("进入统筹金额") & "," & _
        ic住院读卡.OutAnnPlan - rsTemp("统筹报销金额") & "," & ic住院读卡.AnnInpatientTimes & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    
    '冲销单据基本上是复制原单据
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_自贡市 & "," & rsTemp("病人ID") & "," & _
        rsTemp("年度") & "," & rsTemp("帐户累计增加") & "," & rsTemp("帐户累计支出") & "," & rsTemp("累计进入统筹") & "," & _
        rsTemp("累计统筹报销") & "," & rsTemp("住院次数") & "," & rsTemp("起付线") * -1 & "," & rsTemp("封顶线") & "," & rsTemp("实际起付线") * -1 & "," & _
        rsTemp("发生费用金额") * -1 & "," & rsTemp("全自付金额") * -1 & "," & rsTemp("首先自付金额") * -1 & "," & rsTemp("进入统筹金额") * -1 & "," & _
        rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") * -1 & "," & rsTemp("个人帐户支付") * -1 & ",'" & ic住院读卡.OutSerialNo + 1 & "'," & _
        IIf(IsNull(rsTemp("主页ID")), "null", rsTemp("主页ID")) & "," & rsTemp("中途结帐") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
    cur个人帐户 = rsTemp("个人帐户支付")
    
    gstrSQL = "select 档次,进入统筹金额,统筹报销金额,比例 from 保险结算计算 where 结帐ID=[1]"
    Set rs结算计算 = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", lng结帐ID)
    
    Do Until rs结算计算.EOF
        '依次为档次、进入统筹金额、统筹报销金额、比例
        gstrSQL = "zl_保险结算计算_Insert(" & lng冲销ID & "," & _
            rs结算计算("档次") & "," & rs结算计算("进入统筹金额") * -1 & "," & rs结算计算("统筹报销金额") * -1 & "," & rs结算计算("比例") & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "中软医保")
        
        rs结算计算.MoveNext
    Loop
    
    If bln离休 = False Then
        With ic住院读卡
            .Cardno = ic住院读卡.Cardno         ' 卡号
            .OutPerAcc = ic住院读卡.OutPerAcc - rsTemp("个人帐户支付") ' 个人帐户累计支出金额
            .OutExAcc = ic住院读卡.OutExAcc                            ' 补充帐户累计支出金额
            .OutSubAcc = ic住院读卡.OutSubAcc                          ' 补助帐户累计支出金额
            .OutSerialNo = ic住院读卡.OutSerialNo + 1                  ' 支付顺序号
            .OutAnnOverLine = ic住院读卡.OutAnnOverLine - rsTemp("进入统筹金额")  ' 本年进入统筹金额累计
            .OutAnnPlan = ic住院读卡.OutAnnPlan - rsTemp("统筹报销金额")          ' 本年统筹支付金额累计
            .InpatientFlag = ic住院读卡.InpatientFlag                  ' 住院标志 0-不住院 1-住院
            .AnnInpatientTimes = ic住院读卡.AnnInpatientTimes          ' 本年有效住院次数
'            .PayOccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")           ' 日期
'            .PayHospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
'            .PayAccPay = cur个人帐户 * -1    ' 个人帐户支付
'            .PayAmount = g结算数据.发生费用金额    ' 总额
        End With
        
        lngReturn = WriteICCard(ic住院读卡)
        If lngReturn Then
            Err.Raise 9000, gstrSysName, "费用写入卡失败。" & 错误信息_中软(lngReturn)
            Exit Function
        End If
        '记录住院情况
        Dim payLog As TPayLog
        With payLog
            .HospitalCode = Mid(gstr医院编码, 1, 4) ' 医院代码
            .OccurDate = Format(zlDatabase.Currentdate, "yyyyMMdd")                       ' 日期
            .AccPay = Space(8 - Len(CStr(cur个人帐户 * -100))) & CStr(cur个人帐户 * -100)
            .Amount = Space(8 - Len(CStr(g结算数据.发生费用金额 * 100))) & CStr(g结算数据.发生费用金额 * 100)
        End With
        ChargeLog payLog
    End If
        
    住院结算冲销_中软 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 错误信息_中软(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息
    Select Case lngErrCode
        Case -2
            错误信息_中软 = "参数个数错误。"
        Case -3
            错误信息_中软 = "操作端口失败。"
        Case -4
            错误信息_中软 = "打开读卡器失败,请检查读卡器连接和电源。"
        Case -5
            错误信息_中软 = "无卡。"
        Case 0
            错误信息_中软 = "正确。"
        Case 2
            错误信息_中软 = "读错误。"
        Case 3
            错误信息_中软 = "文件结束。"
        Case 4
            错误信息_中软 = "错误PIN。"
'        Case 5
'            错误信息_中软 = "。"
        Case 6
            错误信息_中软 = "复位失败。"
        Case 7
            错误信息_中软 = "检验错误。"
        Case 8
            错误信息_中软 = "修改数据失败。"
        Case 9
            错误信息_中软 = "命令长度错误。"
        Case 10
            错误信息_中软 = "状态错误。"
        Case 11
            错误信息_中软 = "文件类别错误。"
        Case 12
            错误信息_中软 = "文件未选择。"
        Case 13
            错误信息_中软 = "不可重用。"
        Case 14
            错误信息_中软 = "文件已经存在。"
        Case 15
            错误信息_中软 = "错误的P1/P2。"
        Case 16
            错误信息_中软 = "参数错误。"
        Case 17
            错误信息_中软 = "错误的P2。"
        Case 18
            错误信息_中软 = "文件没有找到。"
        Case 19
            错误信息_中软 = "文件无足够空间。"
        Case 20
            错误信息_中软 = "参数错误。"
        Case 21
            错误信息_中软 = "偏移量错误。"
        Case 22
            错误信息_中软 = "指令代码无效。"
        Case 23
            错误信息_中软 = "无效的CLA。"
        Case 24
            错误信息_中软 = "参数错误。"
        Case 25
            错误信息_中软 = "写卡数据转换错误。"
        Case 26
            错误信息_中软 = "个人帐户出现负数,交医保中心处理。"
        Case 33
            错误信息_中软 = "IC卡已经被非法更换,写卡失败。"
        Case 100
            错误信息_中软 = "一期卡，需要格式转换。"
        Case 101
            错误信息_中软 = "非本系统卡。"
        Case 210
            错误信息_中软 = "写卡失败。"
        Case 211
            错误信息_中软 = "写卡失败,扣卡交医保中心处理。"
        Case 300
            错误信息_中软 = "CRC校验错误。"
        Case 301
            错误信息_中软 = "IC卡已经被非法更换,写卡失败.。"
        Case 600
            错误信息_中软 = "读卡值转换错误。"
        Case Else
            错误信息_中软 = "不可识别的错误。"
    End Select
End Function

Private Function 装钱操作(ByVal lng病人ID As Long) As Boolean
'功能：首先断断是否要装钱，然后完成相应操作
    Dim rsTemp As New ADODB.Recordset
    Dim str医保号 As String
    
    Dim str装钱模式 As String, bln强制装钱 As Boolean
    Dim str医保年  As String, lng装钱期次 As Long
    Dim dbl累计注入 As Double
    Dim ic卡 As TIC中软
    Dim str医保年_IC  As String, lng装钱期次_IC As Long
    Dim dbl累计注入_IC As Double
    Dim lngTemp As Long
    
    Dim str参数值 As String
    
    On Error GoTo errHandle
    
    '得到最新的IC卡信息
    '使用本地的，因为可能进行更改但又不成功
    If Is离休病人(lng病人ID, str医保号) = False Then
        If ReadICCard(gIC中软) <> 0 Then
            Exit Function
        End If
    Else
        '医保病人不需要装钱
        If Get离休病人_中软(str医保号, gIC中软) = False Then Exit Function
        装钱操作 = True
        Exit Function
    End If
    '判断卡是否当前病人的
    If lng病人ID > 0 Then
        If 检查IC卡(lng病人ID, TrimStr(gIC中软.Cardno), TrimStr(gIC中软.CenterCode)) = False Then
            Exit Function
        End If
    End If
    ic卡 = gIC中软
    
    With ic卡
        str医保年_IC = .MediYear
        lng装钱期次_IC = .InNo
        dbl累计注入_IC = .InPerAcc
    End With
    
    '获得装钱模式
    '进行合法性验证
    str装钱模式 = Left(Get保险参数_中软(ic卡.CenterCode, "装钱模式", False), 1)
    str医保年 = Get保险参数_中软(ic卡.CenterCode, "医保年", True)
    lng装钱期次 = Val(Get保险参数_中软(ic卡.CenterCode, "装钱序号", True))
    
    If str装钱模式 = "" Or str医保年 = "" Then
        MsgBox "请先请管理员完成医保数据的下载。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If str装钱模式 = "1" Then
        '在线装钱
        'Modified By 朱玉宝 2003-12-10 地区：泸州 （以前的模式没有错，所以改回来）
        lngTemp = OnLineInMoney(ic卡.CenterCode, ic卡.Cardno, str医保年_IC, Trim(gstr医院编码))
        If lngTemp <> 0 Then
            '装钱不成功
            '判断是否列更换医保年
            If str医保年 > str医保年_IC Then
                MsgBox "装钱清单中没有此卡号信息，请到中心处理！", vbInformation, gstrSysName
                Exit Function
'                Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, ic卡.InPerAcc - ic卡.OutPerAcc)
'                '把信息写回卡中
'                If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
'                    '更新全局变量，可能有用
'                    gIC中软 = ic卡
'                Else
'                    '装钱失败
'                    Exit Function
'                End If
            End If
        Else
            '装钱成功，从卡中读出新的值
            If ReadICCard(gIC中软) <> 0 Then
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
                    gIC中软 = ic卡
                Else
                    '装钱失败
                    Exit Function
                End If
            End If
        End If
        
    End If
    
    If (str装钱模式 = "2" Or bln强制装钱 = True) And lng装钱期次 > ic卡.InNo Then
        '离线装钱
        If 检查医保服务器_中软 = False Then
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
        rsTemp.Open gstrSQL, gcn中软, adOpenStatic
        If rsTemp.RecordCount = 0 Then
            '判断是否列更换医保年
            If str医保年 > ic卡.MediYear Then
                MsgBox "装钱清单中没有此卡号信息，请到中心处理！", vbInformation, gstrSysName
                Exit Function
'                Call 更换医保年装钱(ic卡, str医保年, lng装钱期次, ic卡.InPerAcc - ic卡.OutPerAcc)
'                If 记录装钱日志(ic卡, str医保年_IC, lng装钱期次_IC, dbl累计注入_IC) = True Then
'                    '更新全局变量，可能有用
'                    gIC中软 = ic卡
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
            gIC中软 = ic卡
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

Private Sub 更换医保年装钱(ic中软 As TIC中软, ByVal str医保年 As String, ByVal lng装钱期次 As Long, ByVal dbl累计注入 As Double)
    With ic中软
        .MediYear = str医保年
        .InNo = lng装钱期次
        .InPerAcc = dbl累计注入
        .OutPerAcc = 0
        .OutAnnPlan = 0
        .OutAnnOverLine = 0
        .AnnInpatientTimes = 0
        .OutSerialNo = .OutSerialNo + 1
    End With
End Sub

Private Function 记录装钱日志(ic中软 As TIC中软, ByVal IC_MediYear As String, ByVal IC_InNo As Long, ByVal IC_InPerAcc As Double) As Boolean
    
    If 检查医保服务器_中软 = False Then
        '不能连接到前置服务器，就认为不可使用
        Exit Function
    End If
    
    gcn中软.BeginTrans
    On Error Resume Next
    
    '首先保存装钱日志
    With ic中软
        gstrSQL = "insert into 装钱日志 (中心代码,卡号,卡中医保年,卡中装钱期次,卡中账户注入" & _
            ",库中医保年,库中装钱期次,库中账户注入,操作日期) values ('" & _
            .CenterCode & "','" & .Cardno & "','" & IC_MediYear & "'," & IC_InNo & "," & Format(IC_InPerAcc, "#####0.00") & ",'" & _
            .MediYear & "'," & .InNo & "," & Format(.InPerAcc, "#####0.00") & ",sysdate)"
        
    End With
    gcn中软.Execute gstrSQL
    If Err <> 0 Then
        Err.Clear
        gcn中软.RollbackTrans
        Exit Function
    End If
    
    '完成写卡操作
    If WriteICCard(ic中软) <> 0 Then
        gcn中软.RollbackTrans
        MsgBox "IC卡装钱操作失败。", vbInformation, gstrSysName
        Exit Function
    End If
    
    gcn中软.CommitTrans
    记录装钱日志 = True
End Function

Private Function 医保灰度(ByVal str中心 As String, ByVal str卡号 As String) As card医保灰度
'返回指定用户的医保灰度级
    Dim rsTemp As New ADODB.Recordset
    
    If 检查医保服务器_中软 = False Then
        '不能连接到前置服务器，就认为不可使用
        医保灰度 = deg停止支付
        Exit Function
    End If
    
    gstrSQL = "select 灰度 from 黑名单 where 中心代码='" & str中心 & "' and 卡号='" & str卡号 & "'"
    rsTemp.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount > 0 Then
        '设置灰度值
        医保灰度 = Val(rsTemp("灰度"))
    Else
        '正常的不下发
        医保灰度 = deg正常支付
    End If
    
End Function

Private Function 检查IC卡(ByVal lng病人ID As Long, ByVal str卡号 As String, ByVal str中心 As String) As Boolean
'功能：判断该病人的卡是否插入正确
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.卡号,A.医保号,B.编码 from 保险帐户 A,保险中心目录 B " & _
              " where A.险类=[1] and A.病人ID=[2] and a.险类=B.险类 and A.中心=B.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_自贡市, lng病人ID)
    
    If rsTemp("卡号") <> str卡号 Or rsTemp("编码") <> str中心 Then
        MsgBox "刷卡器中的卡不是当前病人的，请插入正确的IC卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    检查IC卡 = True
End Function

Private Function 检查医保服务器_中软() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If gcn中软.State = adStateOpen Then
        检查医保服务器_中软 = True
        Exit Function
    End If
    
    '读出连接医保服务器的配置
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_自贡市)
    
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
    
    If OraDataOpen(gcn中软, strServer, strUser, strPass, False) = True Then
        检查医保服务器_中软 = True
        Exit Function
    End If
        
    MsgBox "医保前置服务器连接失败。", vbInformation, gstrSysName
End Function

Public Function Get离休病人_中软(ByVal strIdentify As String, ic中软 As TIC中软, Optional ByVal bln医保号 As Boolean = True) As Boolean
'功能：从离休清单中读取病人情况，填入IC卡结构中
'参数：strIdentify     病人身份验证（bln医保号=False 为身份证 ，bln医保号=True 是医保号）
'      IC中软        根据读出的信息填写IC卡结构
    Dim rsTemp As New ADODB.Recordset

    If 检查医保服务器_中软 = False Then
        Exit Function
    End If
    
    gstrSQL = "select * from 离休人员 where " & IIf(bln医保号 = True, "医保号", "身份证号") & _
                "='" & strIdentify & "'"
    rsTemp.Open gstrSQL, gcn中软, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '没找到该离休病人的记录
        Exit Function
    End If
    
    With ic中软
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
        .InExAcc = 0        'As Double          ' 补充帐累计注入金额
        .OutExAcc = 0       'As Double          ' 补充帐户累计支出金额
        .InSubAcc = 0       'As Double          ' 补助帐户累计注入金额
        .OutSubAcc = 0      'As Double          ' 补助帐户累计支出金额
        .OutAnnPlan = 0     'As Double          ' 本年统筹支付金额累计
        .OutAnnOverLine = 0 'As Double          ' 本年进入统筹金额累计
        .Password = "9000"       'As String * 4      ' 个人密码
        .AnnInpatientTimes = 0 'As Long           ' 本年有效住院次数
        .InpatientFlag = 0  'As String * 1      ' 住院标志 0-不住院 1-住院
        .HasSubInsurance = 0 'As String * 1      ' 是否参加公务员补助保险  0-否  其他-是
        .HasExInsurance = 0 'As String * 1      ' 是否参加补充保险0：否1是
        .HasBigIllness = 0  'As String * 1      ' 是否参加大病医保
    End With
    
    Get离休病人_中软 = True
End Function


Private Function Is离休病人(ByVal lng病人ID As Long, str医保号 As String) As Boolean
'功能：根据帐户信息判断病人是否离休病人
'参数：返回病人的医保号
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 在职,医保号 from 保险帐户 where 险类=[1] and 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保", TYPE_自贡市, lng病人ID)
    
    If rsTemp.EOF = True Then
        '该病人没发现
        Is离休病人 = False
    Else
        Is离休病人 = IIf(rsTemp("在职") = 3, True, False)
        str医保号 = rsTemp("医保号")
    End If
End Function

Public Function Get保险参数_中软(ByVal str中心代码 As String, ByVal str参数名 As String, bln医保服务器 As Boolean) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    If 检查医保服务器_中软 = False Then
        Exit Function
    End If
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名='" & str参数名 & "' and A.险类=" & TYPE_自贡市 & " and (A.中心 is null or A.中心 in (select B.序号 from 保险中心目录 B where B.险类=" & TYPE_自贡市 & " and B.编码='" & str中心代码 & "'))"
    If bln医保服务器 = True Then
        Call OpenRecordset_OtherBase(rsTemp, "", gstrSQL, gcn中软)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "中软医保")
    End If
    
    If rsTemp.EOF = False Then
        Get保险参数_中软 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function Check有效期(ByVal strCenterCode As String) As Boolean
    '检查中心的有效期
    Dim str有效期  As String
    
    str有效期 = Get保险参数_中软(strCenterCode, "有效期", True)
    
    If IsDate(str有效期) = False Then
        MsgBox "请先从医保中心下载数据后再使用本功能。", vbInformation, gstrSysName
        Exit Function
    End If
    If CDate(str有效期) < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
        MsgBox "病人所属医保中心已经过了有效期。", vbInformation, gstrSysName
        Exit Function
    End If
    Check有效期 = True
End Function
