Attribute VB_Name = "mdl凯里"
Option Explicit

'Public glngReturn As Long                       'API函数返回值
Public gstrReturn As String                    'API函数出口参数
Public gstr医保号 As String

'=================================================================================================================
'功能：获取病人信息
'入口参数：密码
'出口参数：0.IC卡号,1.社会保障号,2.姓名,3.性别,4.出生日期,5.人员类别,6.身份证号码,7.参保日期,9.民族编码,
'   9.特殊疾病标志,10.特殊病病种,11.特殊病有效截止日期,12.特殊病特殊治疗起付线,
'   13.特殊病特殊治疗医保范围内金额,14.特殊病特殊治疗统筹基金支付金额,15.本年普通门诊医保范围内费用累计,
'   16.本年个人帐户收入累计,17.本年个人帐户支付累计,18.本年住院次数,19.本年基本医疗保险符合政策医疗费用累计,
'   20.本年基本医疗保险基金支付累计,21.本年大额保险符合政策医疗费用累计,22.本年大额保险基金支付累计,
'   23.本年公务员补助保险符合政策医疗费用累计,24.本年公务员补助基金支付累计,25.本年工伤保险符合政策医疗费用,
'   26.本年工伤保险基金制服累计,27.本年生育保险符合政策医疗费用累计,28.本年生育保险基金支付累计,29.当前余额,
'   30.名单类别
'参数说明：性别-----1.男,2.女
'          人员类别-----01.在职,02.退休,03.离休,04.老红军,05.在职二等乙,06.退休二等乙,07.离休二等乙
'          特殊病标志-----0.普通,1.特殊疾病,2.特殊治疗
'          特殊病种-----疾病名称,如有多个,以"，"分隔
'=================================================================================================================
Public Declare Function GetPersonInfo Lib "API接口" Alias "f_comm_getpersoninfo" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：修改IC卡密码
'入口参数：旧密码,新密码
'出口参数：无
'=================================================================================================================
Public Declare Function ChangePass Lib "API接口" Alias "f_comm_alterpasswd" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：门诊挂号
'入口参数：门诊挂号流水号,社会保障号,就诊日期,就诊科室名称
'出口参数：无
'=================================================================================================================
Public Declare Function OupReg Lib "API接口" Alias "f_oup_register" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：门诊退号
'入口参数：门诊挂号流水号
'出口参数：无
'=================================================================================================================
Public Declare Function unOupReg Lib "API接口" Alias "f_oup_unregister" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：从政策机获得明细项目的分解信息
'入口参数：项目代码,项目名称,目录类别,单价,计价单位,数量,金额
'出口参数：0.项目代码,1.项目名称,2.目录类别,3.单价,4.计价单位,5.数量,6.金额,7.项目编码,8.项目类别,9.支付级别,
'          10.对照表中医院药品单价,11.自付比例,12.医保范围内金额,13.医保范围外金额
'=================================================================================================================
Public Declare Function ItemDivide Lib "API接口" Alias "f_item_divide" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：将患者门诊处方信息传入政策机，并从政策机获得费用分解信息
'入口参数：outp_divide.in
'   首行：门诊收费流水号,挂号流水号,社会保障号,费用发生日期,疾病标志,诊断,ICD9码,签字医师代码,文件记录条数,
'         本次收费费用总金额
'   记录行：项目代码 , 项目名称, 目录类别, 单价, 计价单位, 数量, 金额
'出口参数：outp_divide.out
'   首行：0.门诊收费流水号,1.挂号流水号,2.社会保障号,3.文件记录条数,4.本次收费费用总金额,5.医保范围内金额,
'         6.医保范围外金额,7.甲类药品金额,8.乙类药品医保支付金额,9.乙类药品政策自付金额,10.丙类药品全额自付金额,
'         11.全额支付诊疗项目金额,12.部分支付诊疗项目医保支付金额,13.部分支付诊疗项目政策自付金额,
'         14.不予支付诊疗项目全额自付金额,15.全额支付服务设施金额,16.部分支付服务设施医保支付金额,
'         17.部分支付服务设施政策自付金额,18.不予支付服务设施全额自付金额,19.个人帐户余额,20.现金支付金额,
'         21.个人帐户支付金额,22.公务员补助支付金额,23.统筹基金支付金额
'   记录行：0.项目代码,1.项目名称,2.目录类别,3.单价,4.计价单位,5.数量,6.金额,7.项目类别,8.支付级别,
'           9.对照表中医院药品单价,10.自付比例,11.医保范围内金额,12.医保范围外金额
'=================================================================================================================
Public Declare Function OutpDivide Lib "API接口" Alias "f_outp_divide" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：取消政策机分解的结果。政策机删除HIS在调用门诊费用分解函数时传来的记录。此函数只能在费用尚未支付确认时使用
'入口参数：门诊收费流水号,挂号流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function OutpCancel Lib "API接口" Alias "f_outp_cancel" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：按政策机分解的结果完成门诊收费。政策机记录确认信息，并冲减IC卡。须插卡
'入口参数：门诊收费流水号,挂号流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function OutpAffirm Lib "API接口" Alias "f_outp_affirm" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：将患者退费有关的信息传入政策机，并从政策机获得费用分解信息。须插卡
'入口参数：门诊收费流水号,挂号流水号,社会保障号,新门诊收费流水号
'出口参数：0.门诊收费流水号,1.挂号流水号,2.社会保障号,3.项目条数,4.本次收费费用总金额,5.医保范围内金额,
'          6.医保范围外金额,7.甲类药品金额,8.乙类药品医保支付金额,9.乙类药品政策自付金额,10.丙类药品全额自付金额,
'          11.全额支付诊疗项目金额,12.部分支付诊疗项目医保支付金额,13.部分支付诊疗项目政策自付金额,
'          14.不予支付诊疗项目全额自付金额,15.全额支付服务设施金额,16.部分支付服务设施医保支付金额,
'          17.部分支付服务设施政策自付金额,18.不予支付服务设施全额自付金额,19.个人帐户支付金额,
'          20.公务员补助支付金额,21.调剂金支付金额,22.现金支付金额
'=================================================================================================================
Public Declare Function OutpWithdrawal Lib "API接口" Alias "f_outp_withdrawal" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：将入院登记信息写入政策机
'入口参数：入院登记流水号,社会保障号,家庭病床标志,住院号,入院方式,入院日期,入院诊断,科室名称,床位号,收治医师姓名
'出口参数：无
'=================================================================================================================
Public Declare Function InpRegister Lib "API接口" Alias "f_inp_register" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：将转科转床信息写入政策机
'入口参数：入院登记流水号,社会保障号,住院号,科室名称,床位号,收治医师姓名
'出口参数：无
'=================================================================================================================
Public Declare Function InpTransfer Lib "API接口" Alias "f_inp_transfer" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：退院。政策机删除该入院登记信息。如果患者已发生费用，禁止使用此函数
'入口参数：入院登记流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function unInpRegister Lib "API接口" Alias "f_inp_unregister" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：政策机删除该入院登记信息。如果患者已结算，禁止使用此函数
'入口参数：入院登记流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function InpClean Lib "API接口" Alias "f_inp_clean" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：将患者住院的费用明细信息传入政策机，并从政策机获得费用分解信息
'入口参数：inp_append.in
'　　首行：入院登记流水号,社会保障号,文件记录条数,本次收费费用总金额
'　记录行：费用发生日期,项目代码,项目名称,目录类别,单价,计价单位,数量,金额,序列号
'出口参数：inp_append.out
'　　首行：0.入院登记流水号,1.社会保障号,2.文件记录条数,3.本次收费费用总金额,4.医保范围内金额,5.医保范围外金额,
'　　　　　6.甲类药品金额,7.乙类药品医保支付金额,8.乙类药品政策自付金额,9.丙类药品全额自付金额,
'          10.全额支付诊疗项目金额,11.部分支付诊疗项目医保支付金额,12.部分支付诊疗项目政策自付金额,
'          13.不予支付诊疗项目全额自付金额,14.全额支付服务设施金额,15.部分支付服务设施医保支付金额,
'          16.部分支付服务设施政策自付金额,17.不予支付服务设施全额自付金额
'　记录行：0.费用发生日期,1.项目代码,2.项目名称,3.目录类别,4.单价,5.计价单位,6.数量,7.金额,8.项目代码,9.项目类别,
'          10.支付级别,11.对照表中医院药品单价,12.自付比例,13.医保范围内金额,14.医保范围外金额,15.序列号
'=================================================================================================================
Public Declare Function InpAppend Lib "API接口" Alias "f_inp_append" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：将患者住院的退费明细信息传入政策机，并从政策机获得费用分解信息
'入口参数：inp_withdrawal.in
'　　首行：入院登记流水号,社会保障号,文件记录条数,本次收费费用总金额
'　记录行：费用发生日期,项目代码,项目名称,目录类别,单价,计价单位,数量,金额,序列号
'出口参数：inp_withdrawal.out
'　　首行：0.入院登记流水号,1.社会保障号,2.文件记录条数,3.本次收费费用总金额,4.医保范围内金额,5.医保范围外金额,
'　　　　　6.甲类药品金额,7.乙类药品医保支付金额,8.乙类药品政策自付金额,9.丙类药品全额自付金额,
'          10.全额支付诊疗项目金额,11.部分支付诊疗项目医保支付金额,12.部分支付诊疗项目政策自付金额,
'          13.不予支付诊疗项目全额自付金额,14.全额支付服务设施金额,15.部分支付服务设施医保支付金额,
'          16.部分支付服务设施政策自付金额,17.不予支付服务设施全额自付金额
'　记录行：0.费用发生日期,1.项目代码,2.项目名称,3.项目分类,4.单价,5.计价单位,6.数量,7.金额,8.项目代码,9.项目类别,
'          10.支付级别,11.对照表中医院药品单价,12.自付比例,13.医保范围内金额,14.医保范围外金额,15.序列号
'=================================================================================================================
Public Declare Function InpWithDrawal Lib "API接口" Alias "f_inp_withdrawal" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：从政策机获得费用分解信息
'入口参数：入院登记流水号,社会保障号
'出口参数：0 入院登记流水号,1 社会保障号,2 记录条数,3 费用总金额,4 医保范围内金额,5 医保范围外金额,6 甲类药品金额,
'          7 乙类药品医保支付金额,8 乙类药品政策自付金额,9 丙类药品全额自付金额,10 全额支付诊疗项目金额,
'          11 部分支付诊疗项目医保支付金额,12 部分支付诊疗项目政策自付金额,13 不予支付诊疗项目全额自付金额,
'          14 全额支付服务设施金额,15 部分支付服务设施医保支付金额,16 部分支付服务设施政策自付金额,
'          17 不予支付服务设施全额自付金额,18 个人帐户余额,19 现金支付金额,20 个人帐户支付金额,21 起付额,
'          22 统筹支付,23 统筹比例自付,24 大额支付,25 大额比例自付,26 公务员医疗补助支付金额
'=================================================================================================================
Public Declare Function InpDivide Lib "API接口" Alias "f_inp_divide" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：住院费用结算结算。政策机记录结算信息，并冲减IC卡。须插卡
'入口参数：住院收费流水号,入院登记流水号,社会保障号,中途结算/出院结算标志
'出口参数：无
'=================================================================================================================
Public Declare Function InpBalance Lib "API接口" Alias "f_inp_balance" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：将患者药店购药信息传入政策机，并从政策机获得费用分解信息
'入口参数：drugs_divide.in
'    首行：药店收费流水号,费用发生日期,社会保障号,外配处方号,疾病标志,诊断,签字医师代码,文件记录条数,
'          本次收费费用总金额
'　记录行：项目代码,项目名称,目录类别,单价,计价单位,数量,金额
'出口参数：drugs_divide.out
'　　首行：0.药店收费流水号,1.社会保障号,2.文件记录条数,3.本次收费费用总金额,4.医保范围内金额,5.医保范围外金额,
'          6.甲类药品金额,7.乙类药品医保支付金额,8.乙类药品政策自付金额,9.丙类药品全额自付金额,10.个人帐户余额,
'          11.现金支付金额,12.个人帐户支付金额,13.公务员补助支付金额,14.统筹支付金额
'　记录行：0.项目代码,1.项目名称,2.目录类别,3.单价,4.计价单位,5.数量,6.金额,7.项目类别,8.支付级别,
'          9.对照表中医院药品单价,10.自付比例,11.医保范围内金额,12.医保范围外金额
'=================================================================================================================
Public Declare Function DrugsDivide Lib "API接口" Alias "f_drugs_divide" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：取消政策机分解的结果。政策机删除HIS在调用 药店费用分解函数时传来的记录。此函数只能在费用尚未支付确认时使用。
'入口参数：药店收费流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function DrugsCancel Lib "API接口" Alias "f_drugs_cancel" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：按政策机分解的结果完成药店收费。政策机记录确认信息，并冲减IC卡。须插卡。
'入口参数：药店收费流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function DrugsAffirm Lib "API接口" Alias "f_drugs_affirm" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：将患者退费有关的信息传入政策机，并从政策机获得费用分解信息。须插卡。
'入口参数：药店收费流水号,社会保障号,新药店收费流水号(也可理解为退费流水号)
'出口参数：0.药店收费流水号,1.社会保障号,2.项目条数,3.本次收费费用总金额,4.医保范围内金额,5.医保范围外金额,
'          6.甲类药品金额,7.乙类药品医保支付金额,8.乙类药品政策自付金额,9.丙类药品全额自付金额,
'          10.个人帐户支付金额,11.公务员补助支付金额,12.统筹支付金额,13.现金支付金额
'=================================================================================================================
Public Declare Function DrugsWithDrawal Lib "API接口" Alias "f_drugs_withdrawal" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：从政策机获得退结算的预算结果，如退还现金、个人帐户等的金额。
'入口参数：入院登记流水号,原入院登记流水号,社会保障号,新收费流水号
'出口参数：0.入院登记流水号,1.社会保障号,2.记录条数,3.费用总金额,4.医保范围内金额,5.医保范围外金额,6.甲类药品金额,
'          7.乙类药品医保支付金额,8.乙类药品政策自付金额,9.丙类药品全额自付金额,10.全额支付诊疗项目金额,
'          11.部分支付诊疗项目医保支付金额,12.部分支付诊疗项目政策自付金额,13.不予支付诊疗项目全额自付金额,
'          14.全额支付服务设施金额,15.部分支付服务设施医保支付金额,16.部分支付服务设施政策自付金额,
'          17.不予支付服务设施全额自付金额,18.个人帐户余额,19.现金支付金额,20.个人帐户支付金额,21.起付额,
'          22.统筹支付,23.统筹比例自付,24.大额支付,25.大额比例自付,26.公务员医疗补助支付金额
'=================================================================================================================
Public Declare Function inpCutoffDivide Lib "API接口" Alias "f_inp_cutoff_divide" (ByVal strInPara As String, _
    strOutPara As String) As Long

'=================================================================================================================
'功能：取消退结算的预算结果。
'入口参数：住院收费流水号,入院登记流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function inpCutoffCancel Lib "API接口" Alias "f_inp_cutoff_cancel" (ByVal strInPara As String) As Long

'=================================================================================================================
'功能：确认退结算的预算结果。
'入口参数：住院收费流水号,入院登记流水号,社会保障号
'出口参数：无
'=================================================================================================================
Public Declare Function inpCutoffAffirm Lib "API接口" Alias "f_inp_cutoff_affirm" (ByVal strInPara As String) As Long

'读取writeOut函数写入文件中的数据，并存放在数组中，每行占用一个元素
Public Function readIn(strFileName As String, strInArray() As String) As Boolean
    Dim lngL As Long, fs, f, strTemp As String, lngFileLen As Long
    On Error Resume Next
    '检查文件是否存在，并获取文件大小
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(strFileName)
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbInformation, "读取文件"
        Exit Function
    End If
    lngFileLen = f.Size
    Err.Clear
    '读取数据
    Open strFileName For Input As #1
    lngL = 0
    Do While True
        Input #1, strTemp
        ReDim Preserve strInArray(lngL)
        strInArray(lngL) = strTemp
        lngL = lngL + 1
        lngFileLen = lngFileLen - LenC(strTemp) - 2         '计算剩余字节数
        If lngFileLen <= 0 Then Exit Do
    Loop
    Close #1
    If Err.Number = 0 Then readIn = True
End Function

'将数组写入文本文件中，每个元素占一行
Public Function writeOut(strFileName As String, strOutArray() As String) As Boolean
    Dim lngL As Long, fs, f
    On Error Resume Next
    '如果文件已经存在则先删除
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(strFileName)
    f.Delete True
    Err.Clear
    '写入数据
    Open strFileName For Binary As #1
    For lngL = 0 To UBound(strOutArray)
        Put #1, , strOutArray(lngL) & vbCrLf                '写入元素，并以回车换行符分隔
    Next
    Close #1
    If Err.Number = 0 Then writeOut = True
End Function

'获取字符串的字节数（中文字符按2字节计算）
Public Function LenC(sString As String) As Integer
    LenC = LenB(StrConv(sString, vbFromUnicode))
End Function

Public Function cNumber(strInPara As String) As Currency
    If IsNumeric(strInPara) Then
        cNumber = CCur(strInPara)
    Else
        cNumber = 0
    End If
End Function

Public Function 医保初始化_凯里() As Boolean
    医保初始化_凯里 = True
End Function

Public Function 身份标识_凯里(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify凯里
    Dim strPatiInfo As String

    strPatiInfo = frmIDentified.GetPatient(bytType)

    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
        If lng病人ID = 0 Then
            lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_凯里)
        End If
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
        If bytType = 0 Then
            If 门诊挂号_凯里(lng病人ID) = False Then
                身份标识_凯里 = ""
                MsgBox "注册门诊挂号信息失败", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        身份标识_凯里 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_凯里 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    身份标识_凯里 = ""
End Function

Public Function 个人余额_凯里(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset

    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_凯里)

    If rsTemp.EOF Then
        个人余额_凯里 = 0
    Else
        个人余额_凯里 = Nvl(rsTemp("帐户余额"), 0)
    End If
End Function

Public Function 门诊挂号_凯里(lng病人ID As Long) As Boolean
    
End Function

Public Function 费用明细传递_凯里(lng结帐ID As Long, Optional rs明细IN As ADODB.Recordset = Nothing) As Boolean
'    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
'    Dim str操作员 As String, cur发生费用, str就诊编号 As String, strBillNo As String
'    Dim lng病种ID As Long, str病种名称 As String, str病种编码 As String, int特病标志 As Integer
'    Dim str科室编号 As String, str科室名称 As String, lng科室ID As Long
'    Dim str明细编码 As String, str明细名称 As String
'    Dim strTemp As String
'
'    On Error GoTo errHandle
'
'    If rs明细IN Is Nothing Then
'        gstrSQL = "Select * From 病人费用记录 Where 结帐ID=" & lng结帐ID
'        Call OpenRecordset(rs明细, "凯里医保")
'    Else
'        Set rs明细 = rs明细IN.Clone
'    End If
'    If rs明细.EOF = True Then
'        费用明细传递_凯里 = False
'        Exit Function
'    End If
'
'    lng病人ID = rs明细("病人ID")
'    str操作员 = NVL(rs明细("操作员姓名"), UserInfo.姓名)
'
'
'    '写处方信息
'
'    initType
'    mblnReturn = wrecipe(gstr医保机构编码, gstr医院编码, str就诊编号, NVL(rs明细!主页ID, 0) & Right(rs明细!NO, 2), str病种编码, str病种名称, _
'                         int特病标志, NVL(rs明细!开单人, rs明细!划价人), NVL(rs明细!操作员姓名, UserInfo.姓名), str科室编号, _
'                         str科室名称, Format(rs明细!登记时间, "yyyy-MM-dd"), gstrReturn)
'    TrimType
'    If mblnReturn = False Then
'        If gstrReturn.errtext = "插入处方信息失败！ORA-00001: ???????? (YBYY.PRI_QTYL42_T)" Then
'            费用明细传递_凯里 = True
'        Else
'            MsgBox gstrReturn.errtext, vbInformation, "凯里医保"
'            费用明细传递_凯里 = False
'            Exit Function
'        End If
'    End If
'    '写处方明细
'    Do Until rs明细.EOF
'        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, "凯里医保")
'        str明细编码 = rsTemp!编码
'        str明细名称 = rsTemp!名称
'        initType
'        If InStr(NVL(rsTemp!规格, " "), "┆") > 0 Then
'            strTemp = Left(rsTemp!规格, InStr(rsTemp!规格, "┆") - 1)
'        Else
'            strTemp = NVL(rsTemp!规格, " ")
'        End If
''入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号,医院明细编码,医院明细名称,产地,规格,类别,
''         单位,单价,数量,时间,录入人,标志
'        If IsNull(rs明细!是否上传) Or rs明细!是否上传 = 0 Then
'            mblnReturn = wdetails(gstr医保机构编码, gstr医院编码, str就诊编号, NVL(rs明细!主页ID, 0) & Right(rs明细!NO, 2), rs明细!序号, _
'                rsTemp!类别 & rsTemp!编码, rsTemp!名称, " ", strTemp, NVL(rsTemp!费用类型, " "), NVL(rsTemp!计算单位, " "), rs明细!标准单价, _
'                rs明细!付数 * rs明细!数次, Format(rs明细!登记时间, "yyyy-MM-dd"), NVL(rs明细!操作员姓名, UserInfo.姓名), _
'                IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), gstrReturn)
'        Else
'            mblnReturn = udetails(gstr医保机构编码, gstr医院编码, str就诊编号, NVL(rs明细!主页ID, 0) & Right(rs明细!NO, 2), rs明细!序号, _
'                rsTemp!类别 & rsTemp!编码, rsTemp!名称, " ", strTemp, NVL(rsTemp!费用类型, " "), NVL(rsTemp!计算单位, " "), rs明细!标准单价, _
'                rs明细!付数 * rs明细!数次, Format(rs明细!登记时间, "yyyy-MM-dd"), NVL(rs明细!操作员姓名, UserInfo.姓名), _
'                IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), gstrReturn)
'        End If
'        TrimType
'        rs明细.MoveNext
'    Loop
'    rs明细.MoveFirst
'    If lng结帐ID = 0 Then
'        gstrSQL = "Update 病人费用记录 Set 是否上传=1 Where ID='" & rs明细!ID & "'"
'    Else
'        gstrSQL = "Update 病人费用记录 Set 是否上传=1 Where 结帐ID=" & lng结帐ID
'    End If
'    gcnOracle.Execute gstrSQL
'    费用明细传递_凯里 = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
End Function
