Attribute VB_Name = "mdl四川眉山"
Option Explicit
'门诊支付的统筹基金与住院支付的统筹金额不累加
'--------------------工业公司结算规则--------------------
'门诊：
'   一般-       个人帐户使用完后，不能再报销
'   慢特病-     先减个人帐户，余下部分统筹基金报50%，年最高限额为2000元（实际报销最高限额为1000元）
'   工伤-       统筹基金支付
'住院：
'   一般-       按报销比例结算
'--------------------  工厂结算规则  --------------------
'门诊（除一般外，都不减个人帐户）：
'   一般-       个人帐户使用完后，不能再报销
'   慢特病-     统筹基金报50%，年最高限额为4000元（实际报销最高限额为2000元）
'   伤残军人-   统筹基金报80%
'   特殊门诊-   统筹基金支付
'   离休人员-   统筹基金支付
'生育报销：无限额，100%
'住院：
'   一般-       按报销比例结算

Private Type ComInfo_眉山
    病人ID As Long
    中心 As Long
    卡号 As String
    医保号 As String
    人群 As String
    病种ID As Long
    病种名称 As String
    住院次数 As Integer
    住院天数 As Integer
    起付线 As Currency
    本次起付线 As Currency
    费用总额 As Currency
    帐户余额 As Currency
    进入统筹 As Currency
    实际报销 As Currency            '实际报销比例不等于100%的统筹支付的汇总金额
    进入实际报销 As Currency        '进入实际报销部分的统筹金额，这部分金额不进入分档计算
    全自付 As Currency
    首先自付 As Currency
    统筹支付 As Currency
    统筹自付 As Currency
    帐户支付 As Currency
    最高限额 As Currency            '限额，使用于当前对应的病种或人群
    已报销金额 As Currency          '本年度已经报销金额
    报销比例 As Single
End Type
Public gstr实际报销比例_米易 As String                           '保存用户输入的报销比例，用于计算
Public gComInfo_眉山 As ComInfo_眉山
Public rs大类_米易 As New ADODB.Recordset                   '大类汇总，用于计算统筹
Public rs支付大类_米易 As New ADODB.Recordset               '保险支付大类
Public rs分档支付_米易 As New ADODB.Recordset               '分档支付明细，用于保存结算计算表
Public Const gstrFormat_眉山 As String = "#####0.0;-#####0.0; ;"

Public Function 身份标识_眉山(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    身份标识_眉山 = frmIdentify眉山.ShowCard(bytType, lng病人ID)
End Function

Public Function 医保初始化_眉山() As Boolean
    医保初始化_眉山 = True
End Function

Public Function 个人余额_眉山(ByVal strSelfNo As String) As Currency
    '功能: 直接读出卡内金额
    '参数: 是否读卡
    '返回: 返回个人帐户余额
    Dim rsAccount As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select Nvl(帐户余额,0) 帐户余额 From 保险帐户 " & _
              " Where 险类=[1] And 医保号=[2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "返回个人帐户余额", TYPE_四川眉山, strSelfNo)
    
    个人余额_眉山 = rsAccount!帐户余额
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 门诊虚拟结算_眉山(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim curTotal As Currency, cur个人帐户 As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    
    '先计算出进入统筹（按大类汇总）
    Call Calc_基本统筹(rs明细, True)
    
    '根据中心、人群、病种计算统筹报销（注意限额的处理）
    Call Calc_门诊报销计算_米易(True, True)
    
    '门诊仅允许帐户支付，统筹报销需到医保中心处理
'    If gComInfo_眉山.统筹支付 <> 0 Then str结算方式 = "医保基金;" & gComInfo_眉山.统筹支付 & ";0"
    If gComInfo_眉山.帐户支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 = "", "", "|") & "个人帐户;" & gComInfo_眉山.帐户支付 & ";0"
    If str结算方式 = "" Then str结算方式 = "个人帐户;0;0"
    门诊虚拟结算_眉山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_眉山(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, _
Optional ByVal bln本院治疗 As Boolean = True) As Boolean
    Dim int性质 As Integer
    Dim int本院 As Integer, int外院 As Integer
    Dim cur帐户余额 As Currency, cur统筹累计 As Currency, int住院次数 As Integer
    Dim rsTemp As New ADODB.Recordset
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    On Error GoTo errHand
    int性质 = IIf(bln本院治疗, 1, 2)
    
    '先下个人帐户
    If cur个人帐户 <> 0 Then
        If Not 下个人帐户(gComInfo_眉山.病人ID, cur个人帐户 * -1) Then Exit Function
    End If
    
    '将结算信息保存到保险结算记录中
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_四川眉山 & "," & gComInfo_眉山.病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & ",0,0,0,0,NULL,0,0,0," & _
        gComInfo_眉山.费用总额 & "," & gComInfo_眉山.全自付 & "," & gComInfo_眉山.首先自付 & "," & gComInfo_眉山.进入统筹 & "," & gComInfo_眉山.统筹支付 & ",0," & _
        0 & "," & cur个人帐户 & ",null,null,null,null," & gComInfo_眉山.病种ID & ",'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    '保险各大类的报销明细
    With rs大类_米易
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!费用总额, 0) <> 0 Then
                gstrSQL = "ZL_保险报销记录_INSERT(" & int性质 & "," & lng结帐ID & "," & _
                "'" & !大类编码 & "','" & !大类名称 & "'," & !统筹比额 & "," & _
                "" & !特准定额 & "," & !特准天数 & "," & !费用总额 & "," & !报销总额 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保险门诊收费数据")
            End If
            .MoveNext
        Loop
    End With
    
    If cur个人帐户 <> 0 Then
        cur帐户余额 = 0: cur统筹累计 = 0: int住院次数 = 0
        gstrSQL = " Select Nvl(帐户增加累计,0) 帐户余额,Nvl(进入统筹累计,0) 统筹累计,Nvl(住院次数累计,0) 本院,Nvl(外院住院次数,0) 外院  From 帐户年度信息" & _
                  " Where 病人ID=[1] ANd 年度=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", gComInfo_眉山.病人ID, Format(zlDatabase.Currentdate, "yyyy"))
        '由于下个人帐户时，已经冲减了帐户，所以本次不作加减运算
        If Not rsTemp.EOF Then
            cur帐户余额 = Nvl(rsTemp!帐户余额, 0)
            cur统筹累计 = Nvl(rsTemp!统筹累计, 0)
            int本院 = rsTemp!本院
            int外院 = rsTemp!外院
        End If
        gstrSQL = "zl_帐户年度信息_Insert(" & gComInfo_眉山.病人ID & ",25," & Format(zlDatabase.Currentdate, "yyyy") & _
                  "," & cur帐户余额 & ",0," & cur统筹累计 & ",0," & int本院 & "," & int外院 & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新住院次数")
    End If
    
    门诊结算_眉山 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_眉山(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim cur帐户余额 As Currency, cur统筹累计 As Currency, int住院次数 As Integer
    Dim lng年度 As Long, lng记录ID As Long
    Dim int本院 As Integer, int外院 As Integer
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrH
    '产生结算冲销记录
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng记录ID = rsTemp!结帐ID
    
    gstrSQL = "Select * From 保险结算记录 Where 险类=25 And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取结算记录", lng结帐ID)
    lng年度 = Format(rsTemp!结算时间, "yyyy")
    If lng年度 <> Format(zlDatabase.Currentdate, "yyyy") Then
        Err.Raise 9000, gstrSysName, "不能冲销以往年度的单据！"
        Exit Function
    End If
    With rsTemp
        gstrSQL = "zl_保险结算记录_insert(" & !性质 & "," & lng记录ID & ",25," & !病人ID & "," & _
            lng年度 & ",0,0,0,0," & Nvl(!住院次数, 0) & "," & -1 * Nvl(!起付线, 0) & ",0," & -1 * Nvl(!实际起付线, 0) & "," & _
            -1 * Nvl(!发生费用金额, 0) & "," & -1 * Nvl(!全自付金额, 0) & "," & -1 * Nvl(!首先自付金额, 0) & "," & -1 * Nvl(!进入统筹金额, 0) & "," & -1 * Nvl(!统筹报销金额, 0) & ",0," & _
            0 & "," & -1 * cur个人帐户 & ",'" & lng结帐ID & "',null,null,null,null,'" & gstrUserName & "')" '支付顺序号用来保存被冲销的记录ID
        Call zlDatabase.ExecuteProcedure(gstrSQL, "产生冲销结算记录")
        
        gstrSQL = "Select * From 保险报销记录 Where 记录ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取大类报销记录", lng结帐ID)
        '保险各大类的报销明细
        Do While Not .EOF
            If Nvl(!费用总额, 0) <> 0 Then
                gstrSQL = "ZL_保险报销记录_INSERT(" & !性质 & "," & lng记录ID & "," & _
                "'" & !大类编码 & "','" & !大类名称 & "'," & !统筹比额 & "," & _
                "" & !特准定额 & "," & !特准天数 & "," & -1 * !费用总额 & "," & -1 * !报销总额 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保险门诊收费数据")
            End If
            .MoveNext
        Loop
    End With
    
    '还原个人帐户
    If cur个人帐户 <> 0 Then
        If Not 下个人帐户(lng病人ID, cur个人帐户) Then Exit Function
        
        cur帐户余额 = 0: cur统筹累计 = 0: int住院次数 = 0
        gstrSQL = " Select Nvl(帐户增加累计,0) 帐户余额,Nvl(进入统筹累计,0) 统筹累计,Nvl(住院次数累计,0) 本院,Nvl(外院住院次数,0) 外院  From 帐户年度信息" & _
                  " Where 病人ID=[1] ANd 年度=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户余额", lng病人ID, lng年度)
        '由于下个人帐户时，已经冲减了帐户，所以本次不作加减运算
        If Not rsTemp.EOF Then
            cur帐户余额 = Nvl(rsTemp!帐户余额, 0)
            cur统筹累计 = Nvl(rsTemp!统筹累计, 0)
            int本院 = rsTemp!本院
            int外院 = rsTemp!外院
        End If
        gstrSQL = "zl_帐户年度信息_Insert(" & lng病人ID & ",25," & Format(zlDatabase.Currentdate, "yyyy") & _
                  "," & cur帐户余额 & ",0," & cur统筹累计 & ",0," & int本院 & "," & int外院 & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新住院次数")
    End If
    
    门诊结算冲销_眉山 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 入院登记_眉山(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    On Error GoTo errHand
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_四川眉山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    入院登记_眉山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_眉山(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    On Error GoTo errHand
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_四川眉山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    出院登记_眉山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 入院登记撤销_眉山(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    gstrSQL = " Select Count(*) Records From 住院费用记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "撤销入院检查", lng病人ID, lng主页ID)
    If rsTemp!Records <> 0 Then
        MsgBox "已经存在费用记录，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_四川眉山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_眉山 = True
End Function

Public Function 出院登记撤销_眉山(lng病人ID As Long, lng主页ID As Long) As Boolean
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_四川眉山 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_眉山 = True
End Function

Public Function 住院虚拟结算_眉山(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim curTotal As Currency
    Dim lng主页ID As Long
    Dim cur个人自付 As Currency, cur个人帐户 As Long
    Dim str入院年份 As String, str结算年份 As String
    Dim str结算时间 As String, str经办时间 As String
    Dim blnUpload As Boolean
    Dim rsTemp As New ADODB.Recordset
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    On Error GoTo errHand
    
    住院虚拟结算_眉山 = "个人帐户;0;0"
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院结算_眉山(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim cur个人帐户 As Currency
    Dim lng主页ID As Long
    Dim str入院年份 As String, str结算年份 As String
    Dim str经办时间 As String, str结算时间 As String
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
        '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo errHand
    
    住院结算_眉山 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 住院结算冲销_眉山(lng结帐ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str退单编号 As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    MsgBox "本医保不支持冲销，请到医保办办理！", vbInformation, gstrSysName
    住院结算冲销_眉山 = False
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保终止_眉山() As Boolean
    医保终止_眉山 = True
End Function

Public Function 挂号结算_眉山(ByVal lng结帐ID As Long, ByVal cur金额 As Currency) As Boolean
    '仅用来提取病人身份
    
    挂号结算_眉山 = True
End Function

Public Function 挂号结算冲销_眉山(ByVal lng结帐ID As Long) As Boolean
    挂号结算冲销_眉山 = True
End Function

Private Function CalcPrepare(Optional ByVal bln本院治疗 As Boolean = True) As Boolean
    '用于结算前，获取该病人的相关信息
    Dim rsTemp As New ADODB.Recordset
    Dim str年度 As String
    
    '基本信息
    gstrSQL = " Select A.*,B.名称 病种名称,C.名称 人群 " & _
              " From 保险帐户 A,(Select * From 保险病种 Where 险类=" & TYPE_四川眉山 & ") B, " & _
              " (Select * From 保险人群 Where 险类=[1]) C" & _
              " Where A.病种ID=B.ID(+) And A.险类=[1] And A.病人ID=[2]" & _
              " And A.在职=C.序号 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人帐户信息", TYPE_四川眉山, gComInfo_眉山.病人ID)
    gComInfo_眉山.病种ID = Nvl(rsTemp!病种ID, 0)
    gComInfo_眉山.病种名称 = Nvl(rsTemp!病种名称)
    gComInfo_眉山.卡号 = rsTemp!卡号
    gComInfo_眉山.医保号 = rsTemp!医保号
    gComInfo_眉山.中心 = rsTemp!中心
    gComInfo_眉山.人群 = rsTemp!人群
    gComInfo_眉山.帐户余额 = Nvl(rsTemp!帐户余额, 0)
    
    '检查是否设置保险参数
    gstrSQL = "Select count(*) Records From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取系统参数", TYPE_四川眉山)
    If rsTemp!Records = 0 Then
        MsgBox "请设置保险参数！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '先检查是否设置好本年度的结算规则
    str年度 = Format(zlDatabase.Currentdate(), "yyyy")
    gstrSQL = " Select Count(*) Records From 保险报销政策 A,保险人群 B" & _
            " Where A.险类=[1] And A.中心=[2]" & _
            " And A.性质=1 And A.本院=[3] And A.年度=[4]" & _
            " And A.人群=B.序号 And A.险类=B.险类 And B.名称=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "支付比例", TYPE_四川眉山, gComInfo_眉山.中心, IIf(bln本院治疗, 1, 2), str年度, gComInfo_眉山.人群)
    If rsTemp!Records = 0 Then
        MsgBox "还未设置本年度的保险报销政策！[年度结算规则]", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = " Select Count(*) Records From 保险报销政策 A,保险人群 B" & _
            " Where A.险类=[1] And A.中心=[2]" & _
            " And A.性质=2 And A.本院=[3] And A.年度=[4]" & _
            " And A.人群=B.序号 And A.险类=B.险类 And B.名称=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "支付比例", TYPE_四川眉山, gComInfo_眉山.中心, IIf(bln本院治疗, 1, 2), str年度, gComInfo_眉山.人群)
    If rsTemp!Records = 0 Then
        MsgBox "还未设置本年度的保险报销政策！[年度结算规则]", vbInformation, gstrSysName
        Exit Function
    End If
    
    '非本院治疗时，报销模块会为下面的变量赋值
    CalcPrepare = True
    If bln本院治疗 = False Then Exit Function
    '取本年度住院次数
    gstrSQL = "Select Nvl(住院次数累计,0) 住院次数 From 帐户年度信息 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "本年度住院次数统计", TYPE_四川眉山, gComInfo_眉山.病人ID)
    If Not rsTemp.EOF Then
        gComInfo_眉山.住院次数 = rsTemp!住院次数 + 1
    Else
        gComInfo_眉山.住院次数 = 1
    End If
End Function

Public Function 医保设置_眉山() As Boolean
    医保设置_眉山 = frmSet眉山.ShowSet()
End Function

Public Sub Init_大类_米易()
    '初始化大类记录集
    Set rs大类_米易 = New ADODB.Recordset
    With rs大类_米易
        If .State = 1 Then .Close
        .Fields.Append "大类编码", adLongVarChar, 10  '0:表示新增
        .Fields.Append "大类名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "统筹比额", adDouble, 18, adFldIsNullable
        .Fields.Append "特准定额", adDouble, 18, adFldIsNullable
        .Fields.Append "特准天数", adDouble, 5, adFldIsNullable
        .Fields.Append "费用总额", adDouble, 18, adFldIsNullable
        .Fields.Append "报销总额", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    gstrSQL = "Select * From 保险支付大类 Where 险类=[1]"
    Set rs支付大类_米易 = zlDatabase.OpenSQLRecord(gstrSQL, "保险支付大类", TYPE_四川眉山)
    With rs支付大类_米易
        Do While Not .EOF
            rs大类_米易.AddNew
            rs大类_米易!大类编码 = !编码
            rs大类_米易!大类名称 = Nvl(!名称)
            rs大类_米易!统筹比额 = Nvl(!统筹比额, 0)
            rs大类_米易!特准天数 = Nvl(!特准天数, 0)
            rs大类_米易!特准定额 = Nvl(!特准定额, 0)
            rs大类_米易!费用总额 = 0
            rs大类_米易!报销总额 = 0
            rs大类_米易!数量 = 0
            rs大类_米易.Update
            .MoveNext
        Loop
    End With
    
    Set rs分档支付_米易 = New ADODB.Recordset
    With rs分档支付_米易
        If .State = 1 Then .Close
        .Fields.Append "档次", adDouble, 10  '0:表示新增
        .Fields.Append "比例", adDouble, 18, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "进入统筹", adDouble, 18, adFldIsNullable
        .Fields.Append "统筹报销", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub Init_结构体_米易()
    With gComInfo_眉山
        .住院次数 = 0
        .住院天数 = 0
        .费用总额 = 0
        .进入统筹 = 0
        .实际报销 = 0
        .进入实际报销 = 0
        .全自付 = 0
        .首先自付 = 0
        .统筹支付 = 0
        .统筹自付 = 0
        .帐户支付 = 0
        .最高限额 = 0
        .报销比例 = 0
        .病种ID = 0
        .人群 = 0
    End With
End Sub

Private Function Calc_基本统筹(ByVal rsExse As ADODB.Recordset, Optional ByVal bln门诊 As Boolean = True) As Boolean
    Dim cur金额 As Currency, cur统筹 As Currency, cur进入统筹 As Currency
    Dim int住院天数 As Integer
    Dim str大类编码 As String
    Dim rsTemp As New ADODB.Recordset
    '按保险大类计算：本次发生总额、本次起付线、进入统筹金额、全自付金额及首先自付金额，并更新结构体（不分中心）
    'rsExse的字段：病人ID,收费类别,收据费目,计算单位,开单人,收费细目ID,数量,单价,
    '              实收金额,统筹金额,保险支付大类ID,是否医保
    Call Init_大类_米易
    Call Init_结构体_米易
    
    cur金额 = 0: cur统筹 = 0:  cur进入统筹 = 0
    gComInfo_眉山.病人ID = rsExse!病人ID
    If Not CalcPrepare Then Exit Function        '门诊病人结算时，不需要获取相关信息，因为在门诊刷卡时就已经得到
    
    '汇总发生金额
    With rsExse
        Do While Not .EOF
            '先判断是否都设置了医保对应项目编码
            gstrSQL = " Select 是否医保 From 保险支付项目" & _
                      " Where 险类=[1] And 收费细目ID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否设置了对应的医保编码", TYPE_四川眉山, CLng(!收费细目ID))
            If rsTemp.EOF Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Nvl(!保险支付大类ID, 0) = 0 Then
                MsgBox "有项目未设置对应的医保大类，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            cur金额 = cur金额 + Nvl(!实收金额, 0)
            str大类编码 = ""
            If rs支付大类_米易.RecordCount <> 0 Then
                rs支付大类_米易.MoveFirst
                rs支付大类_米易.Find "ID=" & !保险支付大类ID
                If .EOF Then
                    MsgBox "保险大类发生改变，请重新预结算！", vbInformation, gstrSysName
                    Exit Function
                End If
                str大类编码 = rs支付大类_米易!编码
            End If
            
            If str大类编码 = "" Then
                MsgBox "保险大类发生改变，请重新预结算！", vbInformation, gstrSysName
                Exit Function
            End If
            
            '因为rs大类_米易是依据rs支付大类_米易产生的，走到这步，肯定不会为EOF
            rs大类_米易.MoveFirst
            rs大类_米易.Find "大类编码='" & str大类编码 & "'"
            rs大类_米易!费用总额 = rs大类_米易!费用总额 + Nvl(!实收金额, 0)
            If rs支付大类_米易!是否医保 = 1 And Val(rs支付大类_米易!统筹比额) <> 0 Then
                '如果大类是医保项目且统筹比额不等于零
                If rs支付大类_米易!服务对象 = 3 Or rs支付大类_米易!服务对象 = IIf(bln门诊, 1, 2) Then
                    '如果服务对象正确
                    If rsTemp!是否医保 = 1 Then
                        '如果明细项目也是医保项目
                        cur统筹 = cur统筹 + Nvl(!实收金额, 0)
                        rs大类_米易!报销总额 = rs大类_米易!报销总额 + Nvl(!实收金额, 0)
                        rs大类_米易!数量 = rs大类_米易!数量 + Nvl(!数量, 0)
                    End If
                End If
            End If
            rs大类_米易.Update
            .MoveNext
        Loop
    End With
    
    gComInfo_眉山.费用总额 = cur金额
    gComInfo_眉山.全自付 = cur金额 - cur统筹
    '计算进入统筹金额
    With rs大类_米易
        .MoveFirst
        Do While Not .EOF
            If !特准定额 = 0 And !特准天数 = 0 Then
                !报销总额 = !报销总额 * Nvl(!统筹比额, 0) / 100
            Else
                If !数量 > !特准天数 Then
                    '如果住院日超过特准天数，那么金额就是 特准天数*特准定额 +  (数量-特准天数)*统筹比额
                    '当特准定额或特准天数任一个为0时，就相当于不要特准天数
                    !报销总额 = !特准定额 * !特准天数 + _
                        (!数量 - IIf(!特准定额 = 0 Or !特准天数 = 0, 0, !特准天数)) * !统筹比额
                Else
                    If !特准定额 = 0 Or !特准天数 = 0 Then
                        '如果住院日低于特准天数，那么最大金额就是 数量*特准定额 或者 数量*统筹比额
                        '当特准定额或特准天数任一个为0时，就相当于不要特准定额
                        !报销总额 = !报销总额 * !统筹比额 / 100
                    Else
                        !报销总额 = !数量 * !特准定额
                    End If
                End If
            End If
            cur进入统筹 = cur进入统筹 + !报销总额
            .Update
            .MoveNext
        Loop
    End With
    gComInfo_眉山.首先自付 = cur统筹 - cur进入统筹
    gComInfo_眉山.进入统筹 = cur进入统筹
    
    Call Calc_实际报销_米易
    Calc_基本统筹 = True
End Function

Public Sub Calc_实际报销_米易()
    Dim intBound As Integer, lngRow As Long
    Dim sin比例 As Single
    Dim rsTemp As New ADODB.Recordset
    
    '根据实际报销比例计算，只要比例不等于100%的大类，要进入分档计算；否则不进入
    gstrSQL = " Select A.名称,B.参数值 比例 From 保险支付大类 A,(Select * From 保险参数 Where 险类=[1] And 中心=1 And 序号>10) B " & _
              " Where A.险类=[2] And A.名称=B.参数名(+) Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", 25)
    
    With rs大类_米易
        .MoveFirst
        Do While Not .EOF
            If !报销总额 <> 0 Then
                rsTemp.MoveFirst
                rsTemp.Find "名称='" & !大类名称 & "'"
                sin比例 = 100
                If Not rsTemp.EOF Then sin比例 = Nvl(rsTemp!比例, 0)
                '如果用户调整了比例，按用户输入的为准
                If InStr(1, gstr实际报销比例_米易, "|" & !大类名称 & ";") <> 0 Then
                    intBound = UBound(Split(Mid(gstr实际报销比例_米易, 2), "|"))
                    For lngRow = 0 To intBound
                        If Split(Split(Mid(gstr实际报销比例_米易, 2), "|")(lngRow), ";")(0) = !大类名称 Then
                            sin比例 = Val(Split(Split(Mid(gstr实际报销比例_米易, 2), "|")(lngRow), ";")(1))
                            Exit For
                        End If
                    Next
                End If
                
                '如果实际报销比例不为100%，则该部分金额直接按比例算出，不必再进入分档计算
                Debug.Print sin比例
                If sin比例 <> 100 And !大类名称 <> "浮动比例" Then
                    gComInfo_眉山.进入实际报销 = gComInfo_眉山.进入实际报销 + !报销总额
                    gComInfo_眉山.实际报销 = gComInfo_眉山.实际报销 + (!报销总额 * sin比例 / 100)
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Public Sub Calc_门诊报销计算_米易(Optional ByVal bln本院 As Boolean = True, Optional ByVal bln医院调用 As Boolean = False)
    Dim rsPara As New ADODB.Recordset
    Dim rsScale As New ADODB.Recordset
    Dim rsSum As New ADODB.Recordset
    Dim bln先扣个人帐户 As Boolean
    Dim sin比例 As Single
    Dim sin实际允许进入统筹比例 As Single
    '先取保险参数
    gstrSQL = "Select 中心,序号,参数值 From 保险参数 Where 险类=[1] Order by 中心,序号"
    Set rsPara = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险参数", TYPE_四川眉山)
    
    '计算实际进入统筹金额
    gstrSQL = "Select 比例 From 保险报销政策 A,保险人群 B" & _
             " Where A.性质=1 And A.中心=[1] And A.本院=[2]" & _
             " And B.险类=A.险类 And A.险类=[3]" & _
             " And A.人群=B.序号 And B.名称=[4] And A.档次=0 And A.年度=" & Format(zlDatabase.Currentdate, "yyyy")
    Set rsScale = zlDatabase.OpenSQLRecord(gstrSQL, "读取门诊报销比例", gComInfo_眉山.中心, IIf(bln本院, 1, 2), TYPE_四川眉山, gComInfo_眉山.人群)
    If rsScale.EOF Then
        sin实际允许进入统筹比例 = 0
    Else
        sin实际允许进入统筹比例 = Nvl(rsScale!比例, 0)
    End If
    
    '初始化
    gComInfo_眉山.帐户支付 = 0
    gComInfo_眉山.最高限额 = 0
    gComInfo_眉山.报销比例 = 0
    gComInfo_眉山.已报销金额 = 0
    rsPara.Filter = "中心=" & gComInfo_眉山.中心
    
    Dim cur扣除 As Currency
    If (gComInfo_眉山.病种名称 = "慢特病" Or gComInfo_眉山.人群 = "离休") And gComInfo_眉山.帐户余额 > 0 Then
        If sin实际允许进入统筹比例 <> 0 Then
            cur扣除 = gComInfo_眉山.帐户余额 / sin实际允许进入统筹比例 * 100
            If Calc_进入统筹 > cur扣除 Then
                gComInfo_眉山.进入统筹 = Calc_进入统筹 - cur扣除 + gComInfo_眉山.进入实际报销
                gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户余额
            Else
                cur扣除 = Calc_进入统筹 * sin实际允许进入统筹比例 / 100
                gComInfo_眉山.帐户支付 = cur扣除
                gComInfo_眉山.进入统筹 = gComInfo_眉山.实际报销
            End If
        End If
    End If
    
    If gComInfo_眉山.中心 = 1 Then  '工厂
        If gComInfo_眉山.病种名称 = "慢特病" Then
            bln先扣个人帐户 = (cur扣除 = 0)
            '取报销比例
            rsPara.MoveFirst
            rsPara.Find "序号=1"
            gComInfo_眉山.报销比例 = Nvl(rsPara!参数值, 0)
            rsPara.MoveFirst
            rsPara.Find "序号=2"
            gComInfo_眉山.最高限额 = Nvl(rsPara!参数值, 0)
            
            '取本年已经报销金额
            If gComInfo_眉山.最高限额 <> 0 Then
                gstrSQL = " Select Sum(进入统筹金额) 已报销金额 From 保险结算记录 " & _
                          " Where 险类=[2] And 性质=1 And 病人ID=" & gComInfo_眉山.病人ID & _
                          " And 年度='" & Format(zlDatabase.Currentdate, "yyyy") & "' And 病种ID=[1]"
                Set rsSum = zlDatabase.OpenSQLRecord(gstrSQL, "取本年度已报销金额", CLng(gComInfo_眉山.病种ID), TYPE_四川眉山)
                gComInfo_眉山.已报销金额 = Nvl(rsSum!已报销金额, 0)
            End If
        ElseIf gComInfo_眉山.人群 = "离休" Then
            bln先扣个人帐户 = (cur扣除 = 0)
            '取报销比例
            rsPara.MoveFirst
            rsPara.Find "序号=3"
            gComInfo_眉山.报销比例 = sin实际允许进入统筹比例
        ElseIf gComInfo_眉山.人群 = "伤残军人" Then
            '取报销比例
            rsPara.MoveFirst
            rsPara.Find "序号=4"
            gComInfo_眉山.报销比例 = sin实际允许进入统筹比例
        ElseIf gComInfo_眉山.病种名称 = "特殊人群" Then
            '取报销比例
            rsPara.MoveFirst
            rsPara.Find "序号=5"
            gComInfo_眉山.报销比例 = 100
        ElseIf gComInfo_眉山.病种名称 = "计划生育" Then
            gComInfo_眉山.报销比例 = 100
        Else
            If sin实际允许进入统筹比例 <> 0 Then
                cur扣除 = gComInfo_眉山.帐户余额 / sin实际允许进入统筹比例 * 100
                If Calc_进入统筹 > cur扣除 Then
                    gComInfo_眉山.进入统筹 = Calc_进入统筹 - cur扣除 + gComInfo_眉山.进入实际报销
                    gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户余额
                Else
                    cur扣除 = Calc_进入统筹 * sin实际允许进入统筹比例 / 100
                    gComInfo_眉山.帐户支付 = cur扣除
                    gComInfo_眉山.进入统筹 = 0
                End If
            End If
        End If
    Else                            '工业公司
        If gComInfo_眉山.病种名称 = "慢特病" Then
            bln先扣个人帐户 = (cur扣除 = 0)
            '取报销比例
            rsPara.MoveFirst
            rsPara.Find "序号=1"
            gComInfo_眉山.报销比例 = Nvl(rsPara!参数值, 0)
            rsPara.MoveFirst
            rsPara.Find "序号=2"
            gComInfo_眉山.最高限额 = Nvl(rsPara!参数值, 0)
            
            '取本年已经报销金额
            If gComInfo_眉山.最高限额 <> 0 Then
                gstrSQL = " Select Sum(进入统筹金额) 已报销金额 From 保险结算记录 " & _
                          " Where 险类=[1] And 性质=1 And 病人ID=[2]" & _
                          " And 年度='" & Format(zlDatabase.Currentdate, "yyyy") & "' And 病种ID=[3]"
                Set rsSum = zlDatabase.OpenSQLRecord(gstrSQL, "取本年度已报销金额", TYPE_四川眉山, gComInfo_眉山.病人ID, gComInfo_眉山.病种ID)
                gComInfo_眉山.已报销金额 = Nvl(rsSum!已报销金额, 0)
            End If
        ElseIf gComInfo_眉山.病种名称 = "工伤" Then
            rsPara.MoveFirst
            rsPara.Find "序号=3"
            gComInfo_眉山.报销比例 = 100
        Else
            If sin实际允许进入统筹比例 <> 0 Then
                cur扣除 = gComInfo_眉山.帐户余额 / sin实际允许进入统筹比例 * 100
                If Calc_进入统筹 > cur扣除 Then
                    gComInfo_眉山.进入统筹 = Calc_进入统筹 - cur扣除 + gComInfo_眉山.进入实际报销
                    gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户余额
                Else
                    cur扣除 = Calc_进入统筹 * sin实际允许进入统筹比例 / 100
                    gComInfo_眉山.帐户支付 = cur扣除
                    gComInfo_眉山.进入统筹 = gComInfo_眉山.进入实际报销
                End If
            End If
        End If
    End If
    
    '如果报销比例为零，报销限额也为零，则需要冲减个人帐户
    If Calc_进入统筹 > 0 Then
        If gComInfo_眉山.报销比例 <> 0 Or gComInfo_眉山.最高限额 <> 0 Then
            If gComInfo_眉山.最高限额 <> 0 Then
                gComInfo_眉山.最高限额 = gComInfo_眉山.最高限额 - gComInfo_眉山.已报销金额
            Else
                gComInfo_眉山.最高限额 = gComInfo_眉山.费用总额
            End If
            
            If bln先扣个人帐户 Then
                gComInfo_眉山.帐户支付 = IIf(gComInfo_眉山.帐户余额 > Calc_进入统筹, Calc_进入统筹, gComInfo_眉山.帐户余额)
                gComInfo_眉山.进入统筹 = Calc_进入统筹 - gComInfo_眉山.帐户支付 + gComInfo_眉山.进入实际报销
            End If
            If gComInfo_眉山.实际报销 <> 0 Then
                If gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付 > gComInfo_眉山.实际报销 Then
                    gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户支付 + gComInfo_眉山.实际报销
                    gComInfo_眉山.实际报销 = 0
                    gComInfo_眉山.进入实际报销 = 0
                Else
                    Dim sin报销比例 As Single
                    sin报销比例 = gComInfo_眉山.实际报销 / gComInfo_眉山.进入实际报销
                    gComInfo_眉山.实际报销 = gComInfo_眉山.实际报销 - (gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付)
                    gComInfo_眉山.进入实际报销 = gComInfo_眉山.实际报销 / sin报销比例
                    gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户余额
                End If
            End If
            gComInfo_眉山.统筹支付 = IIf(gComInfo_眉山.进入统筹 > gComInfo_眉山.最高限额, _
                gComInfo_眉山.最高限额, Calc_进入统筹) * Val(gComInfo_眉山.报销比例) / 100 + gComInfo_眉山.实际报销
            gComInfo_眉山.统筹自付 = gComInfo_眉山.进入统筹 - gComInfo_眉山.统筹支付
        Else
            If gComInfo_眉山.帐户支付 = 0 Then
                gComInfo_眉山.帐户支付 = IIf(gComInfo_眉山.帐户余额 > gComInfo_眉山.进入统筹, gComInfo_眉山.进入统筹, gComInfo_眉山.帐户余额)
                gComInfo_眉山.进入统筹 = 0
                gComInfo_眉山.统筹支付 = 0
                gComInfo_眉山.统筹自付 = 0
            End If
        End If
    Else
        gComInfo_眉山.进入统筹 = gComInfo_眉山.实际报销
        gComInfo_眉山.帐户支付 = gComInfo_眉山.帐户支付 + IIf(gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付 > gComInfo_眉山.进入统筹, gComInfo_眉山.进入统筹, gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付)
        gComInfo_眉山.进入统筹 = gComInfo_眉山.进入统筹 - IIf(gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付 > gComInfo_眉山.进入统筹, gComInfo_眉山.进入统筹, gComInfo_眉山.帐户余额 - gComInfo_眉山.帐户支付)
    End If
    Call FormatData(bln医院调用)
    
    rsPara.Filter = 0
    rsPara.Close
    Set rsPara = Nothing
End Sub

Private Sub FormatData(Optional ByVal bln医院调用 As Boolean = False)
    '将数据格式化为一位小数
    With gComInfo_眉山
        .统筹支付 = Val(Format(.统筹支付, gstrFormat_眉山))
        If Not bln医院调用 Then .帐户支付 = Val(Format(.帐户支付, gstrFormat_眉山))
    End With
End Sub

Private Function Calc_进入统筹() As Currency
    '由于计算方法特殊，需要去掉实际报销比例不等100%的项目的进入统筹金额进行计算，但进入统筹的金额又不发生变化
    '举例：甲类药：100%进入统筹；实际报销比例：100%；发生费用100
    '      半费  ：100%进入统筹；实际报销比例：50% ；发生费用100
    '计算规则：
    '      本次进入统筹总额 ：200
    '      本次报销总额     ：100(甲)+50(半)
    Calc_进入统筹 = gComInfo_眉山.进入统筹 - gComInfo_眉山.进入实际报销
End Function

Public Sub Calc_住院报销计算_米易(Optional ByVal bln本院 As Boolean = True)
    Dim cur进入统筹 As Currency, cur剩余统筹 As Currency, cur统筹支付 As Currency, cur档次报销 As Currency
    Dim cur进入统筹累计 As Currency, cur允许进入档次 As Currency '本年度进入统筹累计，允许进入档次用于计算起始档次的报销金额
    Dim sin比例 As Single
    Dim lng年度 As Long, int起始档 As Integer
    Dim blnExit As Boolean
    Dim rs档次 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    '计算统筹支付额、并更新结构体（需要区分中心、人群）
    gstrSQL = "Select A.比例,C.名称,C.档次,C.下限,C.上限 From 保险报销政策 A,保险人群 B,保险费用档 C " & _
             " Where A.性质=1 And A.中心=[1] And A.本院=" & IIf(bln本院, 1, 2) & _
             " And B.险类=A.险类 And A.险类= [2]  And A.档次<>0 " & _
             " And A.人群=B.序号 And B.名称=[3] And A.年度=" & Format(zlDatabase.Currentdate(), "yyyy") & _
             " And A.险类=C.险类 And C.档次=A.档次 And C.中心=A.中心" & _
             " Order by 档次"
    Set rs档次 = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险费用档", gComInfo_眉山.中心, TYPE_四川眉山, gComInfo_眉山.人群)
    
    cur进入统筹 = 0: cur统筹支付 = 0
    '应该先减去实际报销部分（因为这部分比例不一致，如果先减普通统筹部分，则需要重算实际报销部分金额）
    
    If gComInfo_眉山.进入统筹 >= gComInfo_眉山.起付线 Then
        If Calc_进入统筹 > gComInfo_眉山.起付线 Then
            cur剩余统筹 = Calc_进入统筹 - gComInfo_眉山.起付线
'            gComInfo_眉山.进入统筹 = cur剩余统筹 + gComInfo_眉山.进入实际报销
        Else
            sin比例 = gComInfo_眉山.实际报销 / gComInfo_眉山.进入实际报销
            cur剩余统筹 = 0
            gComInfo_眉山.进入实际报销 = gComInfo_眉山.进入实际报销 - (gComInfo_眉山.起付线 - Calc_进入统筹)
            gComInfo_眉山.实际报销 = gComInfo_眉山.进入实际报销 * sin比例
'            gComInfo_眉山.进入统筹 = gComInfo_眉山.进入实际报销
        End If
    Else
        cur剩余统筹 = 0
        gComInfo_眉山.进入实际报销 = 0
        gComInfo_眉山.实际报销 = 0
    End If
    gComInfo_眉山.本次起付线 = gComInfo_眉山.起付线
    If cur剩余统筹 <= 0 Then
        cur剩余统筹 = 0
        gComInfo_眉山.本次起付线 = gComInfo_眉山.进入统筹
    End If
    
    '获取本年度进入统筹累计
    cur允许进入档次 = 0
    cur进入统筹累计 = 0
    lng年度 = Format(zlDatabase.Currentdate, "yyyy")
    gstrSQL = " Select Nvl(进入统筹累计,0) 进入统筹累计 From 帐户年度信息 " & _
              " Where 年度=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取本年度进入统筹累计", lng年度, gComInfo_眉山.病人ID)
    If Not rsTemp.EOF Then cur进入统筹累计 = Nvl(rsTemp!进入统筹累计, 0)
    
    With rs档次
        '更新每档次的数据至全局记录集中
        .MoveFirst
        Do While Not .EOF
            If int起始档 = 0 Then
                If (cur进入统筹累计 >= Nvl(!下限, 0) And cur进入统筹累计 < Nvl(!上限, 0)) Or Nvl(!上限, 0) = 0 Then
                    int起始档 = !档次
                    If cur进入统筹累计 <> 0 Then
                        '只有进入统筹累计不为零，才进行计算，否则cur允许进入档次都为零，表示全进
                        cur允许进入档次 = IIf(Nvl(!上限, 0) = 0, 0, (Nvl(!上限, 0) - Nvl(!下限, 0)) - (cur进入统筹累计 - Nvl(!下限, 0)))
                    End If
                End If
            End If
            
            rs分档支付_米易.AddNew
            rs分档支付_米易!档次 = !档次
            rs分档支付_米易!比例 = !比例
            rs分档支付_米易!名称 = !名称
            rs分档支付_米易.Update
            .MoveNext
        Loop
        
        '计算各档次实际统筹金额
        blnExit = False
        .MoveFirst
        .Find "档次=" & int起始档
        Do While Not .EOF
            If (blnExit Or cur剩余统筹 <= 0) Then Exit Do
            rs分档支付_米易.MoveFirst
            rs分档支付_米易.Find "档次=" & !档次
            
            If !档次 <> int起始档 Then cur允许进入档次 = 0
            If !上限 <> 0 Then
                If cur剩余统筹 + cur进入统筹累计 > !上限 Then
                    '(1)
                    cur进入统筹 = !上限 - !下限
                Else
                    '(2)
                    cur进入统筹 = (cur剩余统筹 + cur进入统筹累计) - Nvl(!下限, 0)
                    blnExit = True
                End If
            Else
                '全部进入(3)
                cur进入统筹 = (cur剩余统筹 + cur进入统筹累计) - Nvl(!下限, 0)
                blnExit = True
            End If
            '如果进入部分大于本次总的进入统筹，则进入部分等于本次总的进入统筹部分-校对(2)
            If cur进入统筹 > cur剩余统筹 Then cur进入统筹 = cur剩余统筹
            '大于允许进入档次金额，表示进入金额过大-校对(2),(3)
            If cur进入统筹 > cur允许进入档次 And cur允许进入档次 <> 0 Then
                cur进入统筹 = cur允许进入档次
            End If
            '如果进入统筹小于零，则退出，再进行后面的计算没有意义
            If cur进入统筹 <= 0 Then
                cur进入统筹 = 0
                blnExit = True
            End If
            cur档次报销 = cur进入统筹 * !比例 / 100
            cur统筹支付 = cur统筹支付 + CCur(Format(cur档次报销, "#####0.00;-#####0.00;0;"))
            rs分档支付_米易!进入统筹 = cur进入统筹
            rs分档支付_米易!统筹报销 = cur档次报销
            rs分档支付_米易.Update

            .MoveNext
        Loop
    End With
    gComInfo_眉山.统筹自付 = (cur剩余统筹 - cur统筹支付) + (gComInfo_眉山.进入实际报销 - gComInfo_眉山.实际报销)
    gComInfo_眉山.统筹支付 = cur统筹支付 + gComInfo_眉山.实际报销
    If gComInfo_眉山.统筹支付 > 0 Then
        gComInfo_眉山.进入统筹 = cur剩余统筹 + gComInfo_眉山.进入实际报销
    Else
        gComInfo_眉山.进入统筹 = 0
    End If
    
    Call FormatData
End Sub

Private Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
    Dim RecTarget As New ADODB.Recordset
    Dim intFields As Integer, LngLocate As Long
    '编制人:朱玉宝
    '编制日期:2000-11-02
    '该记录集与凭证控件对应
    '也使用于保存
    
    LngLocate = -1
    Set RecTarget = New ADODB.Recordset
    With RecTarget
        If .State = 1 Then .Close
        If SourceRec.RecordCount <> 0 Then
            On Error Resume Next
            Err = 0
            LngLocate = SourceRec.AbsolutePosition
            If Err <> 0 Then LngLocate = -1
            SourceRec.MoveFirst
        End If
        For intFields = 0 To SourceRec.Fields.Count - 1
            .Fields.Append SourceRec.Fields(intFields).Name, SourceRec.Fields(intFields).Type, SourceRec.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        Do While Not SourceRec.EOF
            .AddNew
            For intFields = 0 To SourceRec.Fields.Count - 1
                .Fields(intFields) = SourceRec.Fields(intFields).Value
            Next
            .Update
            SourceRec.MoveNext
        Loop
    End With
    
    If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
    If LngLocate > 0 Then SourceRec.Move LngLocate - 1
    Set CopyNewRec = RecTarget
End Function

Public Function Encrypt(ByVal strMoney As String, ByVal strCardNO As String) As String
    Dim intLen As Integer, LngProcess As Long
    Dim strTmp As String, strTmp_Source As String, strTmp_Target As String, strTmp_CardNO As String
    Dim strEncrypt As String
    '加密解密
    
    Encrypt = ""
    If Val(strMoney) = 0 Then Exit Function
    
    strEncrypt = "thisisajokebyzybzl"
    For intLen = 1 To Len(strMoney)
        strTmp_Source = Mid(strMoney, intLen, 1)
        strTmp_Target = Mid(strEncrypt, intLen, 1)
        If intLen Mod Len(strCardNO) = 0 Then
            strTmp_CardNO = Mid(strCardNO, intLen, 1)
        Else
            strTmp_CardNO = Mid(strCardNO, intLen Mod Len(strCardNO), 1)
        End If
        LngProcess = asc(strTmp_Source) Xor asc(strTmp_Target) Xor asc(strTmp_CardNO)
        
        If LngProcess < 32 Then
            LngProcess = LngProcess + 32
        ElseIf LngProcess > 127 Then
            LngProcess = LngProcess - (LngProcess - 107)
        End If
        
        If LngProcess = 34 Then
            Encrypt = Encrypt & """"
        ElseIf LngProcess = 39 Then
            Encrypt = Encrypt & "''"
        Else
            Encrypt = Encrypt & Chr(LngProcess)
        End If
    Next
End Function

Public Function 检查帐户信息_米易(ByVal strCardNO As String, Optional ByVal blnUpdate As Boolean = False, Optional ByVal bln卡号 As Boolean = True) As Boolean
    Dim str比较串 As String, lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(帐户余额,0)*100 金额,卡号,加密串,病人ID From 保险帐户 Where 险类=[1]"
    If bln卡号 Then
        gstrSQL = gstrSQL & " And 医保号=[2]"
    Else
        gstrSQL = gstrSQL & " And 病人ID=[3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查帐户信息", TYPE_四川眉山, strCardNO, CLng(Val(strCardNO)))
    If Not rsTemp.EOF Then
        strCardNO = rsTemp!卡号
        lng病人ID = rsTemp!病人ID
        str比较串 = Encrypt(rsTemp!金额, Nvl(rsTemp!卡号))
        If str比较串 <> Nvl(rsTemp!加密串) Then
            If Not blnUpdate Then
                MsgBox "有人在非法修改医保病人的个人帐户，请检查！", vbInformation, gstrSysName
                Exit Function
            Else
                gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & ",25,'加密串','''" & str比较串 & "''')"
                gcnOracle.Execute gstrSQL
            End If
        End If
    End If
    检查帐户信息_米易 = True
End Function

Public Function 下个人帐户(ByVal lng病人ID As Long, ByVal cur金额 As Currency) As Boolean
    Dim lngNextID As Long
    Dim rsBalance As New ADODB.Recordset
    
    On Error GoTo errHand
    If Not 检查帐户信息_米易(lng病人ID, False, False) Then Exit Function
    
    lngNextID = zlDatabase.GetNextID("帐户变动记录")
    gstrSQL = "ZL_帐户变动记录_INSERT(" & lngNextID & "," & TYPE_四川眉山 & ",2," & lng病人ID & "," & _
                cur金额 & ",'" & gstrUserName & "','门诊" & IIf(cur金额 > 0, "冲帐", "支出") & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "下个人帐户")
    
    If Not 检查帐户信息_米易(lng病人ID, True, False) Then Exit Function
    下个人帐户 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


