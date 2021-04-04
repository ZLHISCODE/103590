Attribute VB_Name = "mdl浙江"
Option Explicit

'-----------------------------------------------------------
'没有对码的药品（除中草药外），医保中心按丙类结算
'没有对码的中草药、诊疗及材料，对码的为丙类，未对码的按甲类结算
'-----------------------------------------------------------

'查询操作
Public Declare Function QUERY_HANDLE Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'交易操作
Public Declare Function BUSINESS_HANDLE Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'人工应答
Public Declare Function TRADE_ANSWER Lib "SiInterface.DLL" (ByVal InputData As String, ByVal OutputData As String) As Long
'电子存折交易(intType:1存入,2消费)
Public Declare Function UF_DLPK Lib "CardOpe.dll" (ByVal intType As Integer, ByRef strPass As String, ByRef dbl金额 As Double) As Long
'读取卡内数据(intPathID:1--MF 11,2--MF 12,3--DF04 31,4--DF04 32,5--DF04 33,6--DF04 34,7--DF04 35,8--DF04 36)
Public Declare Function UF_Read_Info Lib "CardOpe.dll" (ByVal intPathID As Integer, ByRef strPass As String, _
    ByRef strInfo As Byte) As Long
'对卡内指定数据进行修改(intPathID:1--MF 11,2--MF 12,3--DF04 31,4--DF04 32,5--DF04 33,6--DF04 34,7--DF04 35,8--DF04 36)
Public Declare Function UF_Update_Info Lib "CardOpe.dll" (ByVal intPathID As Integer, ByRef strPass As String, _
    ByRef strInfo As Byte) As Long
'获取错误信息
Public Declare Function GetErrorDesc Lib "CardOpe.dll" (ByRef strDesc As Byte) As Long
Public Declare Function readCardID Lib "cardhandle.DLL" Alias "readCard" (ByRef strCardID As String) As Long

Public gcn浙江 As New ADODB.Connection, int浙江适用 As Integer, gstrInfo As String

Private str门诊号 As String, mstr卡号 As String

Public Function CheckReturn浙江(Optional int调用方式 As Integer = 0) As Boolean
    Dim strDesc As String, bytDesc(2048) As Byte
    If int调用方式 = 1 Then
        If glngReturn < 0 Then
            glngReturn = GetErrorDesc(bytDesc(0))
            strDesc = StrConv(bytDesc, vbUnicode)
            strDesc = Trim(Split(strDesc, Chr(0))(0))
            MsgBox "在进行医保调用时，医保返回以下错误：" & vbCrLf & "    " & strDesc, vbInformation, "接口错误"
            Exit Function
        End If
    Else
        If glngReturn < 0 Then
            If MsgBox("在进行医保调用时发生错误" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "接口错误") = vbRetry Then
                gstrInfo = ""
            Else
                gstrInfo = "-1"
            End If
            Exit Function
        ElseIf IsNumeric(Left(gstrInfo, InStr(gstrInfo, "|") - 1)) Then
            If Val(Left(gstrInfo, InStr(gstrInfo, "|") - 1)) < 0 Then
                If MsgBox("在进行医保调用时发生错误" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "接口错误") = vbRetry Then
                    gstrInfo = ""
                Else
                    gstrInfo = "-1"
                End If
                Exit Function
            End If
        Else
            If MsgBox("在进行医保调用时发生错误" & vbCrLf & gstrInfo, vbInformation + vbRetryCancel, "接口错误") = vbRetry Then
                gstrInfo = ""
            Else
                gstrInfo = "-1"
            End If
            Exit Function
        End If
    End If
    gstrInfo = Mid(gstrInfo, InStr(gstrInfo, "|") + 1)
    CheckReturn浙江 = True
End Function

Public Function Get保险参数_浙江(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名=[1] and A.险类=[2] and A.中心 is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, str参数名, TYPE_浙江)
    
    If rsTemp.EOF = False Then
        Get保险参数_浙江 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Private Function Get病人ID(str医保号 As String) As Long
'功能：通过医保中心号码和医保号求出病人ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 病人ID from 保险帐户 where 险类 = [1] and 医保号 = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江, str医保号)
    If Not rsTmp.BOF Then
        Get病人ID = CLng(rsTmp("病人ID"))
    Else
        Get病人ID = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get病人ID = 0
End Function

Private Function Get卡号(lng病人ID As Long, Optional flag As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 险类=[1] And 病人id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江, lng病人ID)
    If Not rsTemp.EOF Then
        If flag Then
            Get卡号 = Nvl(rsTemp!医保号)
        Else
            Get卡号 = Nvl(rsTemp!卡号)
        End If
    Else
        Get卡号 = ""
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Get卡号 = ""
End Function

Public Function openConn浙江() As Boolean
    If gcn浙江.State = 1 Then
        openConn浙江 = True
        Exit Function
    End If
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs浙江 As New ADODB.Recordset
    '如果连接已经打开，那就不用再测试
    If gcn浙江.State = adStateOpen Then
        openConn浙江 = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "浙江服务器"
                strServer = strTemp
            Case "浙江用户名"
                strUser = strTemp
            Case "浙江用户密码"
                strPass = strTemp
            Case "适用地区"
                int浙江适用 = Val(strTemp)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If int浙江适用 = 0 Then
        gcn浙江.CursorLocation = adUseClient
        gcn浙江.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    End If
    If Err.Number = 0 Then
        openConn浙江 = True
    Else
        openConn浙江 = False
    End If
    Exit Function
    
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保初始化_浙江() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String, rs浙江 As New ADODB.Recordset
    '如果连接已经打开，那就不用再测试
    If gcn浙江.State = adStateOpen Then
        gstrSQL = "Select * From 保险中心目录 Where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_浙江)
        gstr医保机构编码 = rsTemp!编码
        gstrSQL = "Select * From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_浙江)
        gstr医院编码 = Trim(rsTemp!医院编码)
        医保初始化_浙江 = True
        Exit Function
    End If
     
    On Error GoTo ErrH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "浙江服务器"
                strServer = strTemp
            Case "浙江用户名"
                strUser = strTemp
            Case "浙江用户密码"
                strPass = strTemp
            Case "适用地区"
                int浙江适用 = Val(strTemp)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    If int浙江适用 = 0 Then
        gcn浙江.CursorLocation = adUseClient
        gcn浙江.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    End If
    
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败。", vbInformation, gstrSysName
        医保初始化_浙江 = False
        Exit Function
    End If
    gstrSQL = "Select * From 保险中心目录 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_浙江)
    gstr医保机构编码 = rsTemp!编码
    gstrSQL = "Select * From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化", TYPE_浙江)
    gstr医院编码 = Trim(rsTemp!医院编码)
    
    医保初始化_浙江 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    医保初始化_浙江 = False
End Function

Public Function 个人余额_浙江(lng病人ID As Long) As Currency
'功能：通过病人的信息求出个人余额
    
    On Error GoTo errHandle
    glngReturn = QUERY_HANDLE("13|" & Get卡号(lng病人ID, True) & "|DF0432|", gstrInfo)
    If CheckReturn浙江() = False Then
        MsgBox "提取个人帐户余额失败", vbInformation, gstrSysName
        个人余额_浙江 = 0
    Else
        个人余额_浙江 = Trim(Split(gstrInfo, "|")(0))
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_浙江 & ",'帐户余额','" & 个人余额_浙江 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_浙江")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    个人余额_浙江 = 0
End Function

Public Function 身份标识_浙江(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify浙江
    Dim strPatiInfo As String, cur余额 As Currency
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strPatiInfo = frmIDentified.GetPatient(bytType, mstr卡号)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_浙江)
        
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        身份标识_浙江 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_浙江 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_浙江 = ""
End Function

Public Function 门诊虚拟结算_浙江(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
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
'    str结算方式  "报销方式;金额;是否允许修改|...."
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, _
        strTemp As String, cur个帐支付 As Currency, cur统筹支付 As Currency, cur救助金支付 As Currency, _
        cur公务员补助 As Currency, bytTemp(2048) As Byte
    Dim cur费用总额 As Currency
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有费用，不能进行预结算。", vbInformation, gstrSysName
        门诊虚拟结算_浙江 = False
        Exit Function
    End If
    cur费用总额 = 0
    While Not rs明细.EOF
        cur费用总额 = cur费用总额 + rs明细!实收金额
        rs明细.MoveNext
    Wend
    WriteInfo "开始预结算"
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    If 费用明细传递_浙江(0, rs明细) = False Then Exit Function
    
    str卡号 = Get卡号(lng病人ID)
    datCurr = zlDatabase.Currentdate
    
    strTemp = "09|1|" & UserInfo.姓名 & "|" & str门诊号 & "|" & str卡号 & "|" & str门诊号 & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||1|" & Trim(gstr医院编码) & "|"
    WriteInfo "调用：" & strTemp
    gstrInfo = Space(1024)
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    If UBound(Split(gstrInfo, "|")) < 18 Then
        MsgBox "医保预结算数据格式错误，请检查前置机与医保中心网络连接是否正常。", vbInformation, gstrSysName
        Exit Function
    End If
    If cur费用总额 <> Val(Split(gstrInfo, "|")(1)) Then
      If MsgBox("医保中心返回费用总额与发生额不符，请核对" & vbCrLf & "　　发生额:" & cur费用总额 & "　　　中心返回:" & Split(gstrInfo, "|")(1) & vbCrLf & "是否继续执行？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
           Exit Function
        End If
   End If
    cur个帐支付 = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    cur统筹支付 = Val(Split(gstrInfo, "|")(12))
    cur救助金支付 = Val(Split(gstrInfo, "|")(14))
    cur公务员补助 = Val(Split(gstrInfo, "|")(15))
    If cur个帐支付 <> 0 Then str结算方式 = "个人帐户;" & cur个帐支付 & ";0"
    If cur统筹支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "统筹基金;" & cur统筹支付 & ";0"
    If cur救助金支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "救助金支付;" & cur救助金支付 & ";0"
    If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "公务员补助;" & cur公务员补助 & ";0"
    
    门诊虚拟结算_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_浙江(lng结帐ID As Long, cur个帐支付 As Currency, str医保号 As String, cur全自付 As Currency, cur先自付 As Currency, cur医保基金 As Currency) As Boolean
'功能：对门诊费用进行明细传递并且进行结算
'如果门诊费用明细传递失败，就直接结束函数，返回函数失败
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim rs浙江 As New ADODB.Recordset, lng病人ID As Long, strTemp As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, datCurr As Date
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个人帐户 As Currency, cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String, str卡号 As String
    On Error GoTo errHandle
    '如果个人余额不足，无法进行结算
    lng病人ID = Get病人ID(str医保号)
    cur余额 = 个人余额_浙江(lng病人ID)
    
    str卡号 = Get卡号(lng病人ID)
    datCurr = zlDatabase.Currentdate
    WriteInfo "开始结算"
    strTemp = "10|1|" & UserInfo.姓名 & "|" & str门诊号 & "|" & str卡号 & "|" & str门诊号 & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||0|" & Trim(gstr医院编码) & "|"
    WriteInfo "调用：" & strTemp
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    str交易流水号 = Split(gstrInfo, "|")(0)
    cur发生费用 = Val(Split(gstrInfo, "|")(1))
    cur个帐支付 = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    cur统筹支付 = Val(Split(gstrInfo, "|")(12))
    cur救助金支付 = Val(Split(gstrInfo, "|")(14))
    cur公务员补助 = Val(Split(gstrInfo, "|")(15))
    
    WriteInfo "应答：" & str交易流水号
    glngReturn = TRADE_ANSWER(str交易流水号, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_浙江, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_浙江 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_浙江 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur统筹支付 + cur救助金支付 + cur公务员补助 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL," & cur统筹支付 + cur救助金支付 + cur公务员补助 & ",NULL,NULL," & _
            cur个帐支付 & ",NULL,NULL,NULL,'" & str交易流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    cur余额 = 个人余额_浙江(lng病人ID)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_浙江 & ",'帐户余额','" & cur余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    门诊结算_浙江 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_浙江(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, lng冲销ID As Long, strTemp As String
    Dim cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String
    Dim datCurr As Date

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    WriteInfo "准备冲销"
    '取原单据交易流水号
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    If IsNull(rsTemp!备注) Then
        Exit Function
    End If
    str就诊编号 = rsTemp!备注
    WriteInfo "调用：" & "99|" & str就诊编号 & "|" & Trim(gstr医院编码) & "|"
    '调用接口数冲销
    glngReturn = BUSINESS_HANDLE("99|" & str就诊编号 & "|" & Trim(gstr医院编码) & "|", gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    str就诊编号 = Split(gstrInfo, "|")(0)
    cur发生费用 = Val(Split(gstrInfo, "|")(1))
    cur个人帐户 = Val(Split(gstrInfo, "|")(10)) + Val(Split(gstrInfo, "|")(11))
    cur统筹支付 = Val(Split(gstrInfo, "|")(12))
    cur救助金支付 = Val(Split(gstrInfo, "|")(14))
    cur公务员补助 = Val(Split(gstrInfo, "|")(15))
    
'    WriteInfo "应答：" & str就诊编号
'    glngReturn = TRADE_ANSWER(str就诊编号, gstrInfo)
'    WriteInfo "返回：" & gstrInfo
'    If CheckReturn浙江() = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_浙江, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_浙江 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 - cur统筹支付 - cur救助金支付 - cur公务员补助 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            0 - cur发生费用 & ",0,0,NULL," & 0 - (cur统筹支付 + cur救助金支付 + cur公务员补助) & ",NULL,NULL," & _
            0 - cur个人帐户 & ",NULL,NULL,NULL,'" & str就诊编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    cur余额 = 个人余额_浙江(lng病人ID)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_浙江 & ",'帐户余额','" & cur余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    门诊结算冲销_浙江 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 费用明细传递_浙江(lng结帐ID As Long, Optional rs明细IN As ADODB.Recordset = Nothing, Optional str住院号 As String = "") As Boolean
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, _
        str操作员 As String, cur发生费用, str处方号 As String, datCurr As Date, rs大类 As New ADODB.Recordset, _
        strTemp As String, iLoop As Long, str明细编码 As String, str明细名称 As String, str明细内码 As String, _
        str收费类别 As String, str药品等级 As String, str自理金额 As String, cur自付比例 As Currency, _
        rs药品 As New ADODB.Recordset
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
'    str结算方式  "报销方式;金额;是否允许修改|...."
    Dim bln对码 As Boolean, bln药品 As Boolean, str医生 As String, str科室 As String
    Dim str终止日期 As String, str当前日期 As String, str场合 As String, str项目应用场合 As String
    Dim cur计算金额 As Currency, cur最高限额 As Currency
    Dim bln门诊 As Boolean
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    WriteInfo vbCrLf & "开始费用明细传递"
    
    bln门诊 = IIf(str住院号 = "", True, False) '2006-3-6
    If rs明细IN Is Nothing Then
        gstrSQL = "Select * From " & IIf(bln门诊, "门诊费用记录", "住院费用记录") & " Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 结帐ID=[1]"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    Else
        Set rs明细 = rs明细IN.Clone
    End If
    
    If rs明细.EOF = True Then
        费用明细传递_浙江 = False
        Exit Function
    End If
    
    lng病人ID = rs明细!病人ID
    str操作员 = ToVarchar(UserInfo.姓名, 20)
    Randomize
    If str住院号 = "" Then
        str门诊号 = Chr(Year(Date) - 1939) & hex(Month(datCurr)) & IIf(Day(datCurr) < 10, Day(datCurr), Chr(Day(datCurr) + 55)) & Format(datCurr, "HHMMSS") & Format(999 * Rnd + 1, "0##")
        str住院号 = str门诊号
    Else
        str门诊号 = Chr(Year(Date) - 1939) & hex(Month(datCurr)) & IIf(Day(datCurr) < 10, Day(datCurr), Chr(Day(datCurr) + 55)) & Format(datCurr, "HHMMSS") & Format(999 * Rnd + 1, "0##")
    End If
    str处方号 = Format(datCurr, "yyyymmddHHMMSS") & Format(999 * Rnd + 1, "0##") & Format(gstr医院编码, "0####") & "00000110"
    
    iLoop = 1
    '写处方明细
    
    Do Until rs明细.EOF
        'Beging 2006-3-1 陈东
        If bln门诊 = True Then
            str医生 = Nvl(rs明细!开单人, "无")
            gstrSQL = "Select 名称 as 科室 From 部门表 Where Id In (Select 部门ID From 部门人员 Where 缺省=1 And 人员ID In (Select A.Id From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And B.人员性质='医生' and 姓名=[1]))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊传明细取科室", CStr(rs明细!开单人))
            If rsTemp.RecordCount > 0 Then
                str科室 = Nvl(rsTemp!科室)
            Else
                str科室 = "无"
            End If
        Else
            str医生 = Nvl(rs明细!医生, "无")
            If rs明细IN Is Nothing Then
                str科室 = "无"
            Else
                str科室 = Nvl(rs明细!开单部门, "无")
            End If
        End If
        'End 2006-3-1 陈
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        str明细编码 = rsTemp!编码
        str明细名称 = rsTemp!名称
        Select Case rsTemp!类别
            Case "4"
                str收费类别 = "11"
            Case "5"
                str收费类别 = "11"
            Case "6"
                str收费类别 = "12"
            Case "7"
                str收费类别 = "13"
            Case "C"
                str收费类别 = "25"
            Case "D"
                str收费类别 = "21"
            Case "E"
                str收费类别 = "31"
            Case "F"
                str收费类别 = "24"
            Case "G"
                str收费类别 = "91"
            Case "H"
                str收费类别 = "33"
            Case "I"
                str收费类别 = "91"
            Case "J"
                str收费类别 = "34"
            Case "K"
                str收费类别 = "26"
            Case "L"
                str收费类别 = "23"
            Case "M"
                str收费类别 = "91"
            Case "Z"
                str收费类别 = "91"
            Case "1"
                str收费类别 = "91"
            Case Else
                str收费类别 = "91"
        End Select
        gstrSQL = "Select * From 保险支付项目 Where 险类=" & TYPE_浙江 & " And 是否医保=1 And 收费细目ID=[1]"
'        gstrSQL = "Select A.*,B.名称 As 大类名称,B.编码 As 大类编码 from 保险支付项目 A,保险支付大类 B Where A.险类=B.险类 And " & _
            "A.大类ID=B.ID And A.险类=" & TYPE_浙江 & " And A.是否医保=1 And A.收费细目ID=" & rs明细!收费细目ID
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then
            'Beging 2006-3-1
            bln对码 = False
            bln药品 = False
            'End 2006-3-1
            str明细内码 = str明细编码
            If InStr("11 12", str收费类别) > 0 Then
                gstrSQL = "Select * From 药品目录 Where 药品ID=[1]"
                Set rs药品 = zlDatabase.OpenSQLRecord(gstrSQL, "取药品信息", CLng(rs明细!收费细目ID))
                If rs药品.EOF Then
                    str药品等级 = "1"
                ElseIf Nvl(rs药品!标识码, "药品") = "药品" Then
                    'Beging 2006-3-1
                    bln药品 = True
                    'End 2006-3-1
                    str药品等级 = "3"
                Else
                    str药品等级 = "1"
                End If
            Else
                str药品等级 = "1"
            End If
        Else
            'Beging 2006-3-1
            bln对码 = True
            bln药品 = False
            'End 2006-3-1
            If InStr("11 12", str收费类别) > 0 Then
                gstrSQL = "Select * From 药品目录 Where 药品ID=[1]"
                Set rs药品 = zlDatabase.OpenSQLRecord(gstrSQL, "取药品信息", CLng(rs明细!收费细目ID))
                If rs药品.EOF Then
                    Select Case Nvl(rsTemp!附注, "")
                        Case "丙类", "丙类药"
                            str药品等级 = "3"
                        Case "乙类", "乙类药"
                            str药品等级 = "2"
                        Case Else
                            str药品等级 = "1"
                    End Select
                ElseIf Nvl(rs药品!标识码, "药品") = "药品" Then
                    'Beging 2006-3-1
                    bln药品 = True
                    'End 2006-3-1
                    Select Case Nvl(rsTemp!附注, "")
                        Case "甲类", "甲类药"
                            str药品等级 = "1"
                        Case "乙类", "乙类药"
                            str药品等级 = "2"
                        Case Else
                            str药品等级 = "3"
                    End Select
                Else
                    Select Case Nvl(rsTemp!附注, "")
                        Case "丙类", "丙类药"
                            str药品等级 = "3"
                        Case "乙类", "乙类药"
                            str药品等级 = "2"
                        Case Else
                            str药品等级 = "1"
                    End Select
                End If
            Else
                Select Case Nvl(rsTemp!附注, "")
                    Case "丙类", "丙类药"
                        str药品等级 = "3"
                    Case "乙类", "乙类药"
                        str药品等级 = "2"
                    Case Else
                        str药品等级 = "1"
                End Select
                If Nvl(rsTemp!大类id, "") <> "" Then
                    gstrSQL = "Select * From 保险支付大类 Where ID=[1] And 险类=[2]"
                    Set rs大类 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!大类id), TYPE_浙江)
                    If Not rs大类.EOF Then
                        If Nvl(rs大类!名称, "") = "特殊" Or Nvl(rs大类!名称, "") = "特殊检查" Or Nvl(rs大类!名称, "") = "特殊治疗" Then
                            str收费类别 = "22"
                        End If
                    End If
                End If
            End If
            str明细内码 = Nvl(rsTemp!项目编码, str明细编码)
            '医保中心要求上传医院的项目名称
            
'            str明细名称 = Nvl(rsTemp!项目名称, str明细名称)
        End If
        
        '>beging 2006-3-1
        '未对码,并且日期在指定日期之后,为自费
        str终止日期 = "2006-03-09 23:59:59"
        str当前日期 = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:MM:ss")
        
        If DateDiff("s", CDate(str当前日期), CDate(str终止日期)) < 0 Then
            If bln对码 = False Then
                str药品等级 = 3
            End If
        End If
        '>End 2006-3-1
        
        If str药品等级 = "2" Then
            If str收费类别 = "11" Or str收费类别 = "12" Or str收费类别 = "13" Then
                strTemp = "Select nvl(AKA069,0) AKA069 From KA02 Where AKA060='" & str明细内码 & "'"
            Else
                strTemp = "Select nvl(AKA069,0) AKA069 From ka03 Where AKA090='" & str明细内码 & "'"
            End If
            Set rsTemp = gcn浙江.Execute(strTemp)
            If rsTemp.EOF Then
                cur自付比例 = 0.05
            Else
                cur自付比例 = Nvl(rsTemp(0), 0)
                '>beging 2006-3-1
                '自付比例为0,药品等级为1 ,避免HIS中的附注有错造成上传错误
                If cur自付比例 = 0 Then
                    str药品等级 = "1"
                End If
                '>End 2006-3-1
            End If
        ElseIf str药品等级 = "3" Then
            cur自付比例 = 1
        Else
            cur自付比例 = 0
        End If
        
        '>Beging 2006-3-1
        If DateDiff("s", CDate(str当前日期), CDate(str终止日期)) < 0 Then
            strTemp = "Select nvl(AKA069,0) AKA069,AKA063 From KA02 Where AKA060='" & str明细内码 & "'"
            Set rsTemp = gcn浙江.Execute(strTemp)
            If rsTemp.RecordCount > 0 Then
                cur自付比例 = Nvl(rsTemp(0), 0)
                str药品等级 = 2
                str收费类别 = Nvl(rsTemp(1), "91")
                If cur自付比例 = 0 Then
                    str药品等级 = "1"
                End If
                If cur自付比例 = 1 Then
                    str药品等级 = 3
                End If
                '药品无最高限额
                cur最高限额 = 0
                str项目应用场合 = "无"
            Else
                strTemp = "Select nvl(AKA069,0) AKA069,AKA063,decode(CKC202,'1',aka068,0) as CKC202,decode(sign(instr(aae015,'限住院')),1,'住院') as 应用场合 From ka03 Where AKA090='" & str明细内码 & "'"
                Set rsTemp = gcn浙江.Execute(strTemp)
                If rsTemp.RecordCount > 0 Then
                    cur自付比例 = Nvl(rsTemp(0), 0)
                    str药品等级 = 2
                    str收费类别 = Nvl(rsTemp(1), "91")
                    cur最高限额 = Nvl(rsTemp("CKC202"), 0)
                    str项目应用场合 = Nvl(rsTemp!应用场合, "无")
                    
                    If cur自付比例 = 0 Then
                        str药品等级 = "1"
                    End If
                    If cur自付比例 = 1 Then
                        str药品等级 = 3
                    End If
                Else
                    str收费类别 = "91"
                    cur自付比例 = 1
                    str药品等级 = 3
                    cur最高限额 = 0
                    str项目应用场合 = "无"
                End If
            End If
        End If
        '>end  2006-3-1
        
        str自理金额 = 0
        If str药品等级 = "1" Then
            str自理金额 = 0
        ElseIf str药品等级 = "2" Then
            If rs明细.Fields.Count < 26 Then
                str自理金额 = rs明细!实收金额 * cur自付比例
            Else
                str自理金额 = rs明细!金额 * cur自付比例
            End If
        ElseIf str药品等级 = "3" Then
            If rs明细.Fields.Count < 26 Then
                str自理金额 = rs明细!实收金额
            Else
                str自理金额 = rs明细!金额
            End If
        End If
        
        '>Beging 2006-3-3
        If DateDiff("s", CDate(str当前日期), CDate(str终止日期)) < 0 Then
            str自理金额 = 0
            
            If bln门诊 = True Then
                str场合 = "门诊"
            Else
                str场合 = "住院"
            End If
            
            If cur最高限额 > 0 And InStr(str项目应用场合, str场合) > 0 Then
          
                If rs明细.Fields.Count < 26 Then
                    cur计算金额 = rs明细!单价
                Else
                    cur计算金额 = rs明细!价格
                End If
            
                If cur计算金额 - cur最高限额 > 0 Then
                    cur计算金额 = cur最高限额 * rs明细!数量
                Else
                    If rs明细.Fields.Count < 26 Then
                        cur计算金额 = rs明细!实收金额
                    Else
                        cur计算金额 = rs明细!金额
                    End If
                End If
                
                If str药品等级 = "1" Then
                    str自理金额 = cur计算金额
                ElseIf str药品等级 = "2" Then
                    str自理金额 = cur计算金额 * cur自付比例
                    
                ElseIf str药品等级 = "3" Then
                    If rs明细.Fields.Count < 26 Then
                        str自理金额 = rs明细!实收金额
                    Else
                        str自理金额 = rs明细!金额
                    End If
                End If
            Else
                If str药品等级 = "1" Then
                    str自理金额 = 0
                ElseIf str药品等级 = "2" Then
                    If rs明细.Fields.Count < 26 Then
                        str自理金额 = rs明细!实收金额 * cur自付比例
                    Else
                        str自理金额 = rs明细!金额 * cur自付比例
                    End If
                ElseIf str药品等级 = "3" Then
                    If rs明细.Fields.Count < 26 Then
                        str自理金额 = rs明细!实收金额
                    Else
                        str自理金额 = rs明细!金额
                    End If
                End If
            End If
        End If
        '>end  2006-3-3
        
        If rs明细.Fields.Count < 26 Then
        '2006-3-2 加入科室,医生
            strTemp = "Insert Into KC22 (AKB020,AKC190,CKC250,AAE072,CKC130,CKC131,AKC515,AKC221,AKA063,AKC220," & _
                "AKC222,AKC223,AKC224,AKA065,AKC225,AKC226,AKC227,AKC228,AKC253,ckc111,cdc100) Values ('" & Trim(gstr医院编码) & "','" & _
                str住院号 & "','" & iLoop & "','" & str门诊号 & "',1,'" & str处方号 & "','" & str明细编码 & "',to_date('" & _
                Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),'" & str收费类别 & "','" & _
                str门诊号 & "','" & str明细内码 & "','" & str明细名称 & "','" & IIf(InStr("11 12", str收费类别) > 0, "0", IIf(str收费类别 = "13", "1", "2")) & "','" & _
                str药品等级 & "'," & rs明细!单价 & "," & rs明细!数量 & "," & rs明细!实收金额 & "," & IIf(str药品等级 = "3", "0," & str自理金额, str自理金额 & ",0") & _
                ",'" & str科室 & "','" & str医生 & "')"
        Else
        '2006-3-2 加入科室,医生
            gstrSQL = "Select * From " & IIf(bln门诊, "门诊费用记录", "住院费用记录") & " Where NO=[1] And 序号=[2] And 门诊标志=2 And 记录性质=[3] And 记录状态=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rs明细!NO), CLng(rs明细!序号), CLng(rs明细!记录性质), CLng(rs明细!记录状态))
            strTemp = "Insert Into KC22 (AKB020,AKC190,CKC250,AAE072,CKC130,CKC131,AKC515,AKC221,AKA063,AKC220," & _
                "AKC222,AKC223,AKC224,AKA065,AKC225,AKC226,AKC227,AKC228,AKC253,ckc111,cdc100) Values ('" & Trim(gstr医院编码) & "','" & _
                str住院号 & "','" & iLoop & "','" & str门诊号 & "',1,'" & str处方号 & "','" & str明细编码 & "',to_date('" & _
                Format(rsTemp!发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),'" & str收费类别 & "','" & _
                str门诊号 & "','" & str明细内码 & "','" & str明细名称 & "','" & IIf(InStr("11 12", str收费类别) > 0, "0", IIf(str收费类别 = "13", "1", "2")) & "','" & _
                str药品等级 & "'," & rs明细!价格 & "," & rs明细!数量 & "," & rs明细!金额 & "," & IIf(str药品等级 = "3", "0," & str自理金额, str自理金额 & ",0") & _
                ",'" & str科室 & "','" & str医生 & "')"
        End If
        WriteInfo strTemp
        If rs明细.Fields.Count >= 26 Then
            gstrSQL = "Select * From " & IIf(bln门诊, "门诊费用记录", "住院费用记录") & " Where NO=[1] And 序号=[2] And 门诊标志=2 And 记录性质=[3] And 记录状态=[4]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rs明细!NO), CLng(rs明细!序号), CLng(rs明细!记录性质), CLng(rs明细!记录状态))
            If Nvl(rsTemp!是否上传, 0) = 0 Then gcn浙江.Execute strTemp
            gstrSQL = "zl_病人记帐记录_上传 ('" & rsTemp("ID") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        Else
            gcn浙江.Execute strTemp
        End If
        rs明细.MoveNext
        iLoop = iLoop + 1
    Loop
    费用明细传递_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_浙江(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, strInNote As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记_浙江 = False
        Exit Function
    End If

    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, True) '入院诊断
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    str卡号 = Get卡号(lng病人ID)
    WriteInfo "病人入院登记"
    WriteInfo "调用：" & "01|" & str卡号 & "|1|ZY" & lng病人ID & "_" & lng主页ID & "|" & lng病人ID & "_" & lng主页ID & "|" & _
        UserInfo.编号 & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstr医院编码) & "|"
    '调用接口数冲销
    glngReturn = BUSINESS_HANDLE("01|" & str卡号 & "|1|ZY" & lng病人ID & "_" & lng主页ID & "|" & lng病人ID & "_" & lng主页ID & "|" & _
        UserInfo.编号 & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstr医院编码) & "|", gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    WriteInfo "完成入院登记"
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_浙江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_浙江 = False
End Function

Public Function 入院登记撤消_浙江(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将入院登记撤消信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, strInNote As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记撤消_浙江 = False
        Exit Function
    End If

    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, True) '入院诊断
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    str卡号 = Get卡号(lng病人ID)
    WriteInfo "调用：" & "01|" & str卡号 & "|-1|ZY" & lng病人ID & "_" & lng主页ID & "|" & lng病人ID & "_" & lng主页ID & "|" & _
        UserInfo.编号 & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstr医院编码) & "|"
    '调用接口数冲销
    glngReturn = BUSINESS_HANDLE("01|" & str卡号 & "|-1|ZY" & lng病人ID & "_" & lng主页ID & "|" & lng病人ID & "_" & lng主页ID & "|" & _
        UserInfo.编号 & "|0|0||2|" & Format(datCurr, "yyyymmdd") & "|0|" & strInNote & "|" & Trim(gstr医院编码) & "|", gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function

     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_浙江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记撤消_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记撤消_浙江 = False
End Function

Public Function 住院虚拟结算_浙江(rs明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
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
'    str结算方式  "报销方式;金额;是否允许修改|...."
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, datCurr As Date, _
        strTemp As String, cur个帐支付 As Currency, cur统筹支付 As Currency, cur救助金支付 As Currency, _
        cur公务员补助 As Currency, bytTemp(2048) As Byte, lng主页ID As Long
    Dim cur费用总额 As Currency, str住院号 As String, str结算方式 As String
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有费用，不能进行预结算。", vbInformation, gstrSysName
        Exit Function
    End If
    cur费用总额 = 0
    While Not rs明细.EOF
        cur费用总额 = cur费用总额 + rs明细!金额
        rs明细.MoveNext
    Wend
    WriteInfo "开始预结算"
    rs明细.MoveFirst
    lng病人ID = rs明细!病人ID
    gstrSQL = "Select max(主页id) from 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    
    str住院号 = lng病人ID & "_" & lng主页ID
    If 费用明细传递_浙江(0, rs明细, str住院号) = False Then Exit Function
    
    str卡号 = Get卡号(lng病人ID)
    datCurr = zlDatabase.Currentdate
    
    strTemp = "09|2|" & UserInfo.姓名 & "|" & str住院号 & "|" & str卡号 & "|" & str住院号 & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||1|" & Trim(gstr医院编码) & "|"
    WriteInfo "调用：" & strTemp
    gstrInfo = Space(1024)
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    If cur费用总额 <> Val(Split(gstrInfo, "|")(1)) Then
        If MsgBox("医保中心返回费用总额与发生额不符，请核对" & vbCrLf & "　　发生额:" & cur费用总额 & "　　　中心返回:" & Split(gstrInfo, "|")(1) & vbCrLf & "是否继续执行？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    cur个帐支付 = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    cur统筹支付 = Val(Split(gstrInfo, "|")(15))
    cur救助金支付 = Val(Split(gstrInfo, "|")(17))
    cur公务员补助 = Val(Split(gstrInfo, "|")(18))
   ' If cur个帐支付 <> 0 Then
    str结算方式 = "个人帐户;" & cur个帐支付 & ";0"
    If cur统筹支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "统筹基金;" & cur统筹支付 & ";0"
    If cur救助金支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "救助金支付;" & cur救助金支付 & ";0"
    If cur公务员补助 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "公务员补助;" & cur公务员补助 & ";0"
    
    住院虚拟结算_浙江 = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_浙江(lng结帐ID As Long, lng病人ID As Long) As Boolean
'功能：对住院费用进行明细传递并且进行结算
'如果住院费用明细传递失败，就直接结束函数，返回函数失败
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim rs浙江 As New ADODB.Recordset, strTemp As String, lng主页ID As Long, str住院号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, datCurr As Date
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个帐支付 As Currency, cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String, str卡号 As String
    On Error GoTo errHandle
    '如果个人余额不足，无法进行结算
    gstrSQL = "Select max(主页id) from 病案主页 where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    cur余额 = 个人余额_浙江(lng病人ID)
    lng主页ID = rsTemp(0)
    
    str卡号 = Get卡号(lng病人ID)
    datCurr = zlDatabase.Currentdate
    WriteInfo "开始结算"
    str住院号 = lng病人ID & "_" & lng主页ID
    strTemp = "10|2|" & UserInfo.姓名 & "|" & str住院号 & "|" & str卡号 & "|" & str住院号 & "|||" & _
        Format(datCurr, "yyyymmdd|yyyymmdd") & "|0|||||0|" & Trim(gstr医院编码) & "|"
    WriteInfo "调用：" & strTemp
    glngReturn = BUSINESS_HANDLE(strTemp, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    str交易流水号 = Split(gstrInfo, "|")(0)
    cur发生费用 = Val(Split(gstrInfo, "|")(1))
    cur个帐支付 = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    cur统筹支付 = Val(Split(gstrInfo, "|")(15))
    cur救助金支付 = Val(Split(gstrInfo, "|")(17))
    cur公务员补助 = Val(Split(gstrInfo, "|")(18))
    
    WriteInfo "应答：" & str交易流水号
    glngReturn = TRADE_ANSWER(str交易流水号, gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_浙江, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & TYPE_浙江 & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_浙江 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 + cur个帐支付 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 + cur统筹支付 + cur救助金支付 + cur公务员补助 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & ",0,0,NULL," & cur统筹支付 + cur救助金支付 + cur公务员补助 & ",NULL,NULL," & _
            cur个帐支付 & ",NULL,NULL,NULL,'" & str交易流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    cur余额 = 个人余额_浙江(lng病人ID)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_浙江 & ",'帐户余额','" & cur余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    住院结算_浙江 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_浙江(lng结帐ID As Long) As Boolean
'功能：将住院收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, lng冲销ID As Long, strTemp As String
    Dim cur余额 As Currency, cur发生费用 As Currency, cur统筹支付 As Currency
    Dim cur救助金支付 As Currency, cur公务员补助 As Currency, str交易流水号 As String
    Dim datCurr As Date, cur个人帐户 As Currency, lng病人ID As Long

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    MsgBox "医保病人不能进行住院结算冲销", vbInformation, gstrSysName
    Exit Function
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("ID")
    WriteInfo "准备冲销"
    '取原单据交易流水号
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_浙江, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
    End If
    If IsNull(rsTemp!备注) Then
        Err.Raise 9000 + VbMsgBoxStyle.vbExclamation, gstrSysName, "该单据的交易流水号丢失，不能作废。"
        Exit Function
    End If
    str就诊编号 = rsTemp!备注
    WriteInfo "调用：" & "99|" & str就诊编号 & "|" & Trim(gstr医院编码) & "|"
    '调用接口数冲销
    glngReturn = BUSINESS_HANDLE("99|" & str就诊编号 & "|" & Trim(gstr医院编码) & "|", gstrInfo)
    WriteInfo "返回：" & gstrInfo
    If CheckReturn浙江() = False Then Exit Function
    
    str就诊编号 = Split(gstrInfo, "|")(0)
    cur发生费用 = Val(Split(gstrInfo, "|")(1))
    cur个人帐户 = Val(Split(gstrInfo, "|")(13)) + Val(Split(gstrInfo, "|")(14))
    cur统筹支付 = Val(Split(gstrInfo, "|")(15))
    cur救助金支付 = Val(Split(gstrInfo, "|")(17))
    cur公务员补助 = Val(Split(gstrInfo, "|")(18))
    
'    WriteInfo "应答：" & str就诊编号
'    glngReturn = TRADE_ANSWER(str就诊编号, gstrInfo)
'    WriteInfo "返回：" & gstrInfo
'    If CheckReturn浙江() = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(TYPE_浙江, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_浙江 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 - cur统筹支付 - cur救助金支付 - cur公务员补助 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            0 - cur发生费用 & ",0,0,NULL," & 0 - (cur统筹支付 + cur救助金支付 + cur公务员补助) & ",NULL,NULL," & _
            0 - cur个人帐户 & ",NULL,NULL,NULL,'" & str就诊编号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    cur余额 = 个人余额_浙江(lng病人ID)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_浙江 & ",'帐户余额','" & cur余额 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "浙江医保")
    
    住院结算冲销_浙江 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 出院登记撤消_浙江(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_浙江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记撤消_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记撤消_浙江 = False
End Function

Public Function 出院登记_浙江(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_浙江 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    出院登记_浙江 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记_浙江 = False
End Function

