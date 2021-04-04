Attribute VB_Name = "mdl徐州"
Option Explicit

Declare Function dysEncrypt Lib "ApiDll.DLL" (ByVal strPass As String) As String
Declare Function init_com% Lib "SURE32WC.DLL" (ByVal str As Long)
Declare Function close_com% Lib "SURE32WC.DLL" ()
Declare Function sele_card% Lib "SURE32WC.DLL" (ByVal crdno As Long)
Declare Function power_on% Lib "SURE32WC.DLL" ()
Declare Function power_off% Lib "SURE32WC.DLL" ()
Declare Function rd_str% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Address As Long, ByVal Length As Long, ByVal Buffer$)
Declare Function wr_str% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Address As Long, ByVal Length As Long, ByVal Buffer$)
Declare Function chk_sc% Lib "SURE32WC.DLL" (ByVal apz As Long, ByVal Length As Long, ByVal Buffer$)

Public gcn徐州 As New ADODB.Connection, intCOM徐州 As Integer
Private mcur个帐支付 As Currency, mcur统筹支付 As Currency, mcur总额 As Currency, mint本年住院次数 As Integer, _
        mcur公务员 As Currency, mcur全自理 As Currency, mcur部分自理 As Currency, mcur医保不支付 As Currency, _
        mstr结帐费用 As String

Public Function WriteCard(strState As String) As Boolean
    Dim lngReturn As Long, strReturn As String, strErrInfo As String, strInfo() As String
    lngReturn = init_com(intCOM徐州)
    If lngReturn <> 0 Then
        MsgBox "初始化端口错误", vbInformation, "读卡"
        Exit Function
    End If
    
    lngReturn = sele_card(43)
    If lngReturn <> 0 Then
        MsgBox "定义卡类型错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "卡上电错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
    If lngReturn <> 0 Then
        MsgBox "读取卡信息错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    strInfo = Split(Trim(strReturn), "@")
    strReturn = ""
    For lngReturn = 0 To 11
        strReturn = strReturn & IIf(strReturn <> "", "@", "") & strInfo(lngReturn)
    Next
    
    strReturn = "FFFF"
    lngReturn = chk_sc(0, 2, strReturn)
    If lngReturn <> 0 Then
        strErrInfo = "校验卡失败"
        Select Case lngReturn
            Case 2
                strErrInfo = strErrInfo & "-无卡"
            Case 3
                strErrInfo = strErrInfo & "-未上电"
            Case 4
                strErrInfo = strErrInfo & "-串口错误"
            Case 9
                strErrInfo = strErrInfo & "-数据长度错误"
            Case 11
                strErrInfo = strErrInfo & "-密码错误"
            Case 14
                strErrInfo = strErrInfo & "-卡已损坏"
        End Select
        MsgBox strErrInfo, vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    strInfo(3) = strState
    If InStr(strInfo(11), Chr(0)) > 0 Then strInfo(11) = Left(strInfo(11), InStr(strInfo(11), Chr(0)) - 1)
    strInfo(11) = strInfo(11) & "@"
    strReturn = ""
    For lngReturn = 0 To 11
        strReturn = strReturn & IIf(strReturn <> "", "@", "") & strInfo(lngReturn)
    Next
    lngReturn = wr_str(1, 0, 200, strReturn)
    If lngReturn <> 0 Then
        MsgBox "写卡数据失败", vbInformation, "写卡"
        GoTo powerOFF
    End If
    
    WriteCard = True
powerOFF:
    Call power_off
    Call close_com
End Function

Public Function MakeTransNO() As String
    Randomize
    MakeTransNO = Format(Date, "yymmdd") & Format(Time, "hhmmss") & Format(900099 * Rnd + 1, "0#####")
End Function

Public Function Get医保号(lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保号", lng病人ID)
    If rsTemp.EOF Then
        Get医保号 = ""
    Else
        Get医保号 = Nvl(rsTemp!医保号, "")
    End If
End Function

Public Function openConn徐州() As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    On Error GoTo errHandle
    If gcn徐州.State <> adStateOpen Then
        gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州)
        
        Do Until rsTemp.EOF
            str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Select Case rsTemp("参数名")
                Case "徐州用户名"
                    strUser = str参数值
                Case "徐州服务器"
                    strServer = str参数值
                Case "徐州用户密码"
                    strPass = str参数值
                Case "徐州数据库"
                    strDatabase = str参数值
            End Select
            rsTemp.MoveNext
        Loop
        
        intCOM徐州 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", 0)
        
        On Error Resume Next
        gcn徐州.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn徐州.CursorLocation = adUseClient
        gcn徐州.Open
        
        If Err <> 0 Then
            MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    openConn徐州 = True
    Exit Function

errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 医保初始化_徐州() As Boolean
    Dim rsTemp As New ADODB.Recordset, str参数值 As String, strUser As String, strServer As String, _
        strPass As String, strDatabase As String
    
    If openConn徐州() = False Then Exit Function
    
    gstrSQL = "Select * From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "医保初始化")
    gstr医院编码 = Trim(rsTemp!医院编码)
    
    医保初始化_徐州 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 身份标识_徐州(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能:识别指定人员是否为参保病人，返回病人的信息
'参数:bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回:空或信息串
'注意:1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify徐州, strPatiInfo As String
    
    WriteInfo vbCrLf & "开始身份验证"
    
    strPatiInfo = frmIDentified.GetPatient(bytType)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_徐州)
        
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        mint本年住院次数 = frmIDentified.mint住院次数
        Unload frmIDentified
    Else
        身份标识_徐州 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    
    WriteInfo "结束身份验证"
    
    身份标识_徐州 = strPatiInfo
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_徐州 = ""
End Function

Public Function 个人余额_徐州(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID=[1] and 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, TYPE_徐州)
    
    If rsTemp.EOF Then
        个人余额_徐州 = 0
    Else
        个人余额_徐州 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
End Function

Public Function 门诊虚拟结算_徐州(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数:rsDetail     费用明细(传入)
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
'字段:病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim cur全自理 As Currency, cur部分自理 As Currency, cur总额 As Currency, str医保号 As String, _
        strReturn As String, cur现金 As String, rsTemp As New ADODB.Recordset, strTransNO As String, _
        strSQL As String, bln单味草药 As Boolean, cur草药 As Currency, cur单味草药 As Currency, _
        blnIS药品 As Boolean, lng病人ID As Long, strPara As String, i As Long
        
    On Error GoTo errHandle
    WriteInfo vbCrLf & "开始门诊预结算"
    
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用记录，不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rs明细!病人ID
    
    While Not rs明细.EOF
        gstrSQL = "Select A.类别,B.项目编码,B.项目名称 From 收费细目 A,保险支付项目 B Where A.ID=B.收费细目ID " & _
            "And A.ID=[1] And B.是否医保=1 And B.险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID), TYPE_徐州)
        
        bln单味草药 = True
        If rsTemp.EOF Then
            cur全自理 = cur全自理 + rs明细!实收金额
        Else
            If rsTemp!类别 = "7" Then
                If cur草药 <> 0 Then bln单味草药 = False
                cur草药 = cur草药 + rs明细!实收金额
            End If
            If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
                blnIS药品 = True
                strSQL = "Select * From mi_drug_trade_list Where trade_code='" & rsTemp!项目编码 & "' And cancel_sign<>'1'"
            Else
                blnIS药品 = False
                strSQL = "Select * From mi_dt_item Where item_code='" & rsTemp!项目编码 & "' And cancal_sign<>'1'"
            End If
            Set rsTemp = gcn徐州.Execute(strSQL)
            If rsTemp.EOF Then
                cur全自理 = cur全自理 + rs明细!实收金额
            Else
                If Trim(rsTemp!mi_class = 2) Then
                    If blnIS药品 = True Then
                        cur部分自理 = cur部分自理 + rs明细!实收金额 * 0.2
                    Else
                        cur部分自理 = cur部分自理 + rs明细!实收金额 * Nvl(rsTemp!self_rate, 0) / 100
                    End If
                ElseIf Trim(rsTemp!mi_class) = 4 Then
                    cur单味草药 = cur单味草药 + rs明细!实收金额
                ElseIf Trim(rsTemp!mi_class) <> 1 Then
                    cur全自理 = cur全自理 + rs明细!实收金额
                End If
            End If
        End If
        cur总额 = cur总额 + rs明细!实收金额
        rs明细.MoveNext
    Wend
    If bln单味草药 = True Then cur全自理 = cur全自理 + cur单味草药
    
    If cur总额 = 0 Then
        MsgBox "没有产生病人费用，不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    str医保号 = Get医保号(lng病人ID)
    strTransNO = MakeTransNO()
    strPara = str医保号 & "," & cur总额 & "," & Format(cur部分自理, "0.00") & "," & Format(cur全自理, "0.00")
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','60','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    '0-成功标志，1-医保不支付费用，2-帐户支付费用，3-统筹支付费用，4-公务员基金支付费用
    If InStr(strReturn, ",") > 0 Then
        If Split(strReturn, ",")(0) = "01" Then
            MsgBox "医保交易失败", vbInformation, gstrSysName
            Exit Function
        Else
            mcur医保不支付 = CCur(Split(strReturn, ",")(1))
            mcur个帐支付 = CCur(Split(strReturn, ",")(2))
            mcur统筹支付 = CCur(Split(strReturn, ",")(3))
            mcur公务员 = CCur(Split(strReturn, ",")(4))
        End If
    Else
        MsgBox "医保交易失败", vbInformation, gstrSysName
        Exit Function
    End If
    
    mcur总额 = cur总额: mcur部分自理 = CCur(Format(cur部分自理, "0.00")): mcur全自理 = cur全自理
'    mstr结帐费用 = Right(Space(15) & mcur总额, 15)
'    For i = 1 To 9
'        mstr结帐费用 = mstr结帐费用 & Right(Space(15) & Split(strReturn, ",")(i), 15)
'    Next
    
    If mcur个帐支付 <> 0 Then str结算方式 = "个人帐户;" & mcur个帐支付 & ";0"
    If mcur统筹支付 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "统筹基金;" & mcur统筹支付 & ";0"
    If mcur公务员 <> 0 Then str结算方式 = str结算方式 & IIf(str结算方式 <> "", "|", "") & "公务员基金;" & mcur公务员 & ";0"
    
    WriteInfo "结束门诊预结算:" & str结算方式
    门诊虚拟结算_徐州 = True
    Exit Function
    
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_徐州(lng结帐ID As Long, cur个帐支付 As Currency, str医保号 As String, cur全自付 As Currency, cur先自付 As Currency, cur医保基金 As Currency) As Boolean
    Dim str密码 As String, lng病人ID As String, rsTemp As New ADODB.Recordset, STR姓名 As String, _
        strSQL As String, strReturn As String, strPara As String, strTransNO As String, _
        int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, _
        cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, datCurr As Date, cur余额 As Currency
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "开始门诊结算"
    '门诊号，发票号，医保号，门诊总费用，部分自理费用（乙类费用20%），完全自理费用，密码
    gstrSQL = "Select * From 保险帐户 Where 医保号=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保信息", str医保号, TYPE_徐州)
    
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到该病人的医保信息"
        Exit Function
    End If
    cur余额 = Nvl(rsTemp!帐户余额, 0)
    str密码 = Nvl(rsTemp!密码, "666666")
    lng病人ID = rsTemp!病人ID
    
    gstrSQL = "Select * From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", lng病人ID)
    STR姓名 = rsTemp!姓名
    
    strPara = lng结帐ID & "," & lng结帐ID & "," & str医保号 & "," & mcur总额 & "," & mcur部分自理 & "," & _
        mcur全自理 & "," & dysEncrypt(str密码)
    strTransNO = MakeTransNO()
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','61','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        Err.Raise 9000, gstrSysName, "请求被中止"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        Err.Raise 9000, gstrSysName, "医保交易失败"
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn

    '医院号,门诊号,发票号,医保号,姓名,门诊总费用,完全自理费用,部分自理费用,医保不支付费用,统筹支付,公务员支付,帐户支付,时间
    strSQL = "Insert Into hospital_clinic_payment (hospital_no,clinic_no,invoice_no,medical_card_no,name," & _
        "clinic_expense,full_self_expense,part_self_expense,mi_unpayment,social_payment,servant_payment," & _
        "account_payment,exectime) Values ('" & Trim(gstr医院编码) & "','" & lng结帐ID & "','" & lng结帐ID & "'," & _
        "'" & str医保号 & "','" & STR姓名 & "'," & mcur总额 & "," & mcur全自理 & "," & mcur部分自理 & "," & _
        mcur医保不支付 & "," & mcur统筹支付 & "," & mcur公务员 & "," & mcur个帐支付 & ",'" & Format(datCurr, "yyyy-mm-dd hh:mm:ss") & "')"
    WriteInfo "插入结算数据:" & strSQL
    gcn徐州.Execute strSQL
    
    '帐户年度信息
    Call Get帐户信息(TYPE_徐州, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_徐州 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 + mcur统筹支付 & _
        "," & cur统筹报销累计 + mcur统筹支付 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_徐州 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur余额 & "," & cur帐户支出累计 + mcur个帐支付 & "," & _
        cur进入统筹累计 + mcur统筹支付 + mcur公务员 & "," & cur统筹报销累计 + mcur统筹支付 + mcur公务员 & _
        "," & int住院次数累计 & ",NULL,NULL,NULL," & mcur总额 & "," & mcur全自理 & "," & mcur部分自理 & _
        ",NULL," & mcur统筹支付 + mcur公务员 & ",NULL,NULL," & mcur个帐支付 & ",NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    
    WriteInfo "门诊结算成功"
    
    门诊结算_徐州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    WriteInfo "发生错误:" & Err.Description
    Exit Function
End Function

Public Function 门诊结算冲销_徐州(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能:将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数:lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, str医保号 As String, _
        cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, _
        cur统筹报销累计 As Currency, int住院次数累计 As Integer, cur余额 As Currency, strSQL As String, _
        strPara As String, strReturn As String, strTransNO As String, str密码 As String, lng冲销ID As Long, _
        cur全自理 As Currency, cur部分自理 As Currency, STRNAME As String

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    WriteInfo vbCrLf & "准备门诊退费"
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取帐户信息", lng病人ID)
    cur余额 = Nvl(rsTemp!帐户余额, 0): str密码 = Nvl(rsTemp!密码, "666666")
    
    gstrSQL = "Select * From 病人信息 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", lng病人ID)
    STRNAME = rsTemp!姓名
    
    '取原单据交易数据
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    
    mcur总额 = rsTemp!发生费用金额
    mcur个帐支付 = rsTemp!个人帐户支付
    mcur统筹支付 = rsTemp!统筹报销金额
    cur全自理 = rsTemp!全自付金额
    cur部分自理 = rsTemp!首先自付金额
    
    '门诊号，发票号，医保号，'0'，密码
    str医保号 = Get医保号(lng病人ID)
    
'    strSql = "Select * From hospital_clinic_payment Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
'        "clinic_no='" & lng结帐ID & "' And invoice_no='" & lng结帐ID & "' and medical_card_no='" & str医保号 & "'"
'
'    WriteInfo "取原医保记录:" & strSql
'
'    Set rsTemp = gcn徐州.Execute(strSql)
'
'    If rsTemp.EOF Then
'        WriteInfo "取原医保记录失败"
'        Err.Raise 9000,gstrSysName, "前置机数据库中未找到原有交易记录，不能作废"
'        Exit Function
'    End If
    
    strPara = lng结帐ID & "," & lng结帐ID & "," & str医保号 & ",0," & dysEncrypt(str密码)
    strTransNO = MakeTransNO()

    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara

    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','62','" & UserInfo.编号 & "','" & strPara & "','9')"
    WriteInfo "写交易表:" & strSQL
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        Err.Raise 9000, gstrSysName, "请求被中止"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        Err.Raise 9000, gstrSysName, "医保交易失败"
        Exit Function
    End If

    WriteInfo "请求返回:" & strReturn
    
    strSQL = "Insert Into hospital_clinic_payment (hospital_no,clinic_no,invoice_no,medical_card_no,name," & _
        "clinic_expense,full_self_expense,part_self_expense,mi_unpayment,social_payment,servant_payment," & _
        "account_payment,exectime) Values ('" & Trim(gstr医院编码) & "','" & lng结帐ID & "','" & lng结帐ID & "'," & _
        "'" & str医保号 & "','" & STRNAME & "',-" & mcur总额 & ",-" & cur全自理 & _
        ",-" & cur部分自理 & ",-" & CStr(mcur总额 - mcur个帐支付) & ",0,0" & _
        ",-" & mcur个帐支付 & ",'" & Format(datCurr, "yyyy-mm-dd hh:mm:ss") & "')"
    WriteInfo "插入退费数据:" & strSQL
    gcn徐州.Execute strSQL
    
    '帐户年度信息
    Call Get帐户信息(TYPE_徐州, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_徐州 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - mcur个帐支付 & "," & cur进入统筹累计 - mcur统筹支付 & "," & _
        cur统筹报销累计 - mcur统筹支付 & "," & int住院次数累计 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_徐州 & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 - mcur个帐支付 & "," & cur进入统筹累计 - mcur统筹支付 & "," & _
            cur统筹报销累计 - mcur统筹支付 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            0 - mcur总额 & ",0,0,NULL," & 0 - mcur统筹支付 & ",NULL,NULL," & _
            0 - mcur个帐支付 & ",NULL,NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    
    门诊结算冲销_徐州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "发生错误:" & Err.Description
End Function

Public Function 入院登记_徐州(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能:将入院登记信息发送医保前置服务器确认；
'参数:lng病人ID-病人ID；lng主页ID-主页ID
'返回:交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, strInNote As String, str病种编码 As String, str病种名称 As String, _
        str住院科室 As String, str入院病床 As String, strSQL As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "办理医保入院"
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID = [1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记_徐州 = False
        Exit Function
    End If
    str医保号 = Nvl(rsTemp!医保号, "")
    strPara = lng病人ID & "_" & lng主页ID & "," & str医保号 & ",0"
    strTransNO = MakeTransNO()
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记_徐州 = False
        Exit Function
    End If
    
    str住院科室 = ToVarchar(Nvl(rsTemp!住院科室), 20)
    str入院病床 = ToVarchar(Nvl(rsTemp!入院病床), 3)

    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, True, True, True) '入院诊断
    If strInNote <> "" Then
        str病种名称 = Left(strInNote, InStr(strInNote, "|") - 1)
        str病种编码 = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    strSQL = "Insert Into hos_hospital_daybook (hospital_no,inhospital_no,medical_card_no," & _
        "last_examine_date,examine_date,disease_code,disease_name,inhospital_circs,inhospital_route," & _
        "inhospital_times,inhospital_type,sickarea_section_name,sickbed_no,outhospital_circs,checkout_date," & _
        "hospital_expense,part_self_payment,full_self_payment,start_payment,social_payment,social_unpayment," & _
        "supplement_payment,supplement_unpayment,servant_self_payment,servant_social_payment,cancel_sign) " & _
        "values ('" & Trim(gstr医院编码) & "','" & lng病人ID & "_" & lng主页ID & "','" & str医保号 & "',NULL,'" & _
        Format(datCurr, "yyyy-mm-dd") & "','" & str病种编码 & "','" & str病种名称 & "','1','1'," & _
        mint本年住院次数 & ",'1','" & str住院科室 & "','" & str入院病床 & "'," & _
        "NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'9')"
    gcn徐州.Execute strSQL
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    '进行医保入院交易
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        '如果失败，则删除插入到住院记录表中的记录
        gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
            "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='9'"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        '如果失败，则删除插入到住院记录表中的记录
        gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
            "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='9'"
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    
    '改写IC卡状态
    If WriteCard(1) = False Then
        MsgBox "写卡状态时失败，但病人入院操作不受影响", vbInformation, "写卡"
    End If
     
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_徐州 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_徐州 & ",'顺序号','''0''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_江苏")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_徐州 & ",'退休证号','''0''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_江苏")
    
    WriteInfo "完成入院登记"
    入院登记_徐州 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_徐州 = False
End Function

Public Function 入院登记撤消_徐州(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, str医保号 As String, strSQL As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "撤消入院登记"
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID = [1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        入院登记撤消_徐州 = False
        Exit Function
    End If
    str医保号 = Nvl(rsTemp!医保号, "")
    strPara = lng病人ID & "_" & lng主页ID & "," & str医保号 & ",9"
    strTransNO = MakeTransNO()
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    '进行医保入院交易
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    
    '撤消时删除入院时插入的入院记录
    gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
        "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='9'"
    
    '改写IC卡状态
    If WriteCard(0) = False Then
        MsgBox "写卡状态时失败，但撤消入院的操作不受影响", vbInformation, "写卡"
    End If
     
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_徐州 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    WriteInfo "完成撤消登记操作"
    入院登记撤消_徐州 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记撤消_徐州 = False
End Function

Public Function 出院登记_徐州(lng病人ID As Long, lng主页ID As Long) As Boolean
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_徐州 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    出院登记_徐州 = True
End Function

Public Function 医保出院(lng病人ID As Long, lng主页ID As Long, ByRef str请求 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, str医保号 As String, strInNote As String, str病种编码 As String, str病种名称 As String, _
        str住院科室 As String, str入院病床 As String, strSQL As String

    '出院必须结清
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    WriteInfo vbCrLf & "病人出院登记"
    
    gstrSQL = "Select * From 住院费用记录 Where 病人ID=[1] And 主页ID=[2] And 门诊标志=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    strPara = Nvl(rsTemp!结帐ID, "")
    If strPara = "" Then
        MsgBox "必须结清费用才能办理出院", vbInformation, gstrSysName
        Exit Function
    End If
'    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 险类=" & TYPE_徐州 & " And 记录ID=" & strPara
'    Call OpenRecordset(rsTemp, gstrSysName)
'    mstr结帐费用 = Nvl(rsTemp!备注, "")
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID = [1] And 险类= [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        医保出院 = False
        Exit Function
    End If
    str医保号 = Nvl(rsTemp!医保号, "")
    strPara = lng病人ID & "_" & lng主页ID & "," & str医保号 & ",1"
    strTransNO = MakeTransNO()
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,D.编码 as 科室编码,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        医保出院 = False
        Exit Function
    End If
    str住院科室 = ToVarchar(Nvl(rsTemp!住院科室), 20)
    str入院病床 = ToVarchar(Nvl(rsTemp!入院病床), 3)

    '获取入院诊断（病种编码）
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, False, True, True)  '入院诊断
    If strInNote <> "" Then
        str病种名称 = Left(strInNote, InStr(strInNote, "|") - 1)
        str病种编码 = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    strSQL = "Insert Into hos_hospital_daybook (hospital_no,inhospital_no,medical_card_no," & _
        "last_examine_date,examine_date,disease_code,disease_name,inhospital_circs,inhospital_route," & _
        "inhospital_times,inhospital_type,sickarea_section_name,sickbed_no,outhospital_circs,checkout_date," & _
        "hospital_expense,part_self_payment,full_self_payment,start_payment,social_payment,social_unpayment," & _
        "supplement_payment,supplement_unpayment,servant_self_payment,servant_social_payment,cancel_sign) " & _
        "values ('" & Trim(gstr医院编码) & "','" & lng病人ID & "_" & lng主页ID & "','" & str医保号 & "',NULL,'" & _
        Format(datCurr, "yyyy-mm-dd") & "','" & str病种编码 & "','" & str病种名称 & "','1','1'," & _
        mint本年住院次数 & ",'1','" & str住院科室 & "','" & str入院病床 & "'," & _
        "'1','" & Format(datCurr, "yyyy-mm-dd") & "'," & mstr结帐费用 & ",'0')"
    gcn徐州.Execute strSQL
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    '进行医保入院交易
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        '如果失败，则删除插入到住院记录表中的记录
        gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
            "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='0'"
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        '如果失败，则删除插入到住院记录表中的记录
        gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
            "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='0'"
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    str请求 = strSQL
    WriteInfo "完成出院登记"
    医保出院 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    医保出院 = False
End Function

Public Function 出院登记撤消_徐州(lng病人ID As Long, lng主页ID As Long) As Boolean
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_徐州 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    出院登记撤消_徐州 = True
End Function

Public Function 撤消医保出院_徐州(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, strPara As String, strReturn As String, _
        strTransNO As String, str医保号 As String, strSQL As String

    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    WriteInfo vbCrLf & "撤消出院登记"
    gstrSQL = "Select * From 保险帐户 Where 病人ID = [1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
    If rsTemp.EOF Then
        MsgBox "未能获取入院病人的相关信息。", vbInformation, gstrSysName
        撤消医保出院_徐州 = False
        Exit Function
    End If
    str医保号 = Nvl(rsTemp!医保号, "")
    strPara = lng病人ID & "_" & lng主页ID & "," & str医保号 & ",8"
    strTransNO = MakeTransNO()
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    '进行医保入院交易
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','09','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    
    '撤消后删除出院时插入的出院记录
    gcn徐州.Execute "Delete From hos_hospital_daybook Where hospital_no='" & Trim(gstr医院编码) & "' And " & _
        "inhospital_no='" & lng病人ID & "_" & lng主页ID & "' And cancel_sign='0'"
        
    WriteInfo "完成撤消出院操作"
    撤消医保出院_徐州 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    撤消医保出院_徐州 = False
End Function

Public Function 记帐传输_徐州(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng主页ID As Long, rs明细 As New ADODB.Recordset, str类别 As String, _
        str项目编码 As String, str项目名称 As String, str单位 As String, str通用名编码 As String, _
        cur自理金额 As Currency, lngTemp As Long, strReturn As String, strPara As String, cur部分自理 As Currency, _
        strTransNO As String, str医保号 As String, str医保状态 As String, datCurr As Date, strSQL As String, _
        cur部分自理Sum As Currency, cur全自理Sum As Currency
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    WriteInfo vbCrLf & "开始明细传递"
    If lng病人ID <> 0 Then
        gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
        lng主页ID = rsTemp(0)
    End If
    
    If str单据号 <> "" Then
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B" & _
                  " Where A.门诊标志=2 And A.记录状态<>0 And nvl(A.附加标志,0)<>9 and nvl(A.实收金额,0)<>0 " & _
                  " and A.记录性质=[1] and A.NO=[2]" & _
                  " And A.病人ID=B.病人ID And B.险类=[3]" & _
                  " order by A.主页ID,A.序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", int性质, str单据号, TYPE_徐州)
    Else
        gstrSQL = " Select A.* From 住院费用记录 A,保险帐户 B" & _
                  " Where A.门诊标志=2 And A.记录状态<>0 And nvl(A.附加标志,0)<>9 and nvl(A.实收金额,0)<>0 " & _
                  " and A.病人id=[1] And A.主页id=[2]" & _
                  " And A.病人ID=B.病人ID And B.险类=[3]" & _
                  " order by A.主页ID,A.序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, lng主页ID, TYPE_徐州)
    End If
    If rs明细.EOF Then
'        MsgBox "没有需要传的费用明细", vbInformation, gstrSysName
        WriteInfo "没有需要传递的明细，退出"
        记帐传输_徐州 = True
        Exit Function
    End If
    
    lng病人ID = rs明细!病人ID: lng主页ID = rs明细!主页ID
    
    str医保号 = Get医保号(lng病人ID)
    strPara = str医保号
    strTransNO = MakeTransNO()
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    '取医保状态
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','04','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        Exit Function
    End If
    str医保状态 = Trim(Split(strReturn, ",")(1))
    WriteInfo "请求返回:医保状态(" & str医保状态 & ")"
    
    gcnOracle.Execute "Delete From 医保病人状态 Where 病人ID=" & lng病人ID & " And to_char(日期,'yyyy-mm-dd')='" & Format(datCurr, "yyyy-mm-dd") & "'"
    gcnOracle.Execute "Insert into 医保病人状态 (病人ID,日期,医保状态) values (" & lng病人ID & ",to_date('" & Format(datCurr, "yyyy-mm-dd") & _
        "','yyyy-mm-dd')," & IIf(str医保状态 = "", "NULL", str医保状态) & ")"
    
    lngTemp = 0
    While Not rs明细.EOF
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        If rsTemp!类别 = "7" Then
            lngTemp = lngTemp + 1
        End If
        rs明细.MoveNext
    Wend
    cur部分自理Sum = 0: cur全自理Sum = 0
    rs明细.MoveFirst
'    gcnOracle.Execute "Delete From 县医保明细 Where 病人id=" & lng病人id & " And 主页ID=" & lng主页ID
    
    While Not rs明细.EOF
        gstrSQL = "Select * From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        str项目编码 = rsTemp!编码
        str项目名称 = rsTemp!名称
        str单位 = Nvl(rsTemp!计算单位, "")
        str类别 = rsTemp!类别
        
        strSQL = "Insert Into hos_advice_carryout (hospital_no,advice_serial_no,inhospital_no," & _
            "item_drug_code,conv_price,quantity,norm_unit,all_expense,self_payment,carryout_date," & _
            "item_drug_name,general_code,cease_reason) values ('" & Trim(gstr医院编码) & "','" & IIf(rs明细!实收金额 < 0, "-", "") & _
            rs明细!NO & "_" & rs明细!序号 & "','" & rs明细!病人ID & "_" & rs明细!主页ID & "',"
        
        gstrSQL = "Select * From 保险支付项目 Where 收费细目ID=[1] And 险类=[2] And 是否医保=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险支付项目", CLng(rs明细!收费细目ID), TYPE_徐州)
        cur自理金额 = 0: cur部分自理 = 0
        If rsTemp.EOF Then
            If str类别 = "5" Or str类别 = "6" Or str类别 = "7" Then
'                str项目编码 = "yz" & str项目编码
                str通用名编码 = "8888"
            Else
'                str项目编码 = "zz" & str项目编码
                str通用名编码 = "9000"
            End If
            cur自理金额 = rs明细!实收金额
        Else
            str项目名称 = rsTemp!项目名称
            str通用名编码 = rsTemp!项目编码
            If str类别 = "5" Or str类别 = "6" Or str类别 = "7" Then
'                str项目编码 = "@@" & rsTemp!项目编码
                gstrSQL = "Select * From mi_drug_trade_list Where trade_code='" & rsTemp!项目编码 & "'"
                Set rsTemp = gcn徐州.Execute(gstrSQL)
                If rsTemp.EOF Then
                    cur自理金额 = rs明细!实收金额
                ElseIf Trim(rsTemp!mi_class) = "1" Then
                    cur自理金额 = 0
                ElseIf Trim(rsTemp!mi_class) = "2" Then
                    cur部分自理 = rs明细!实收金额 * 0.2
                ElseIf Trim(rsTemp!mi_class) = "4" Then
                    If lngTemp < 2 Then
                        cur自理金额 = rs明细!实收金额
                    Else
                        cur自理金额 = 0
                    End If
                Else
                    cur自理金额 = rs明细!实收金额
                End If
            ElseIf str类别 = "J" Then
                If Val(rs明细!实收金额) > 0 Then
                    cur部分自理 = IIf(rs明细!实收金额 > 15, rs明细!实收金额 - 15, 0)
                Else
                    cur部分自理 = IIf(Abs(rs明细!实收金额) > 15, rs明细!实收金额 + 15, 0)
                End If
            Else
'                str项目编码 = "$$" & rsTemp!项目编码
                gstrSQL = "Select * From mi_dt_item Where item_code='" & rsTemp!项目编码 & "'"
                Set rsTemp = gcn徐州.Execute(gstrSQL)
                If rsTemp.EOF Then
                    cur自理金额 = rs明细!实收金额
                ElseIf Trim(rsTemp!mi_class) = "1" Then
                    cur自理金额 = 0
                ElseIf Trim(rsTemp!mi_class) <> "2" Then
                    cur自理金额 = rs明细!实收金额
                Else
                    cur部分自理 = rs明细!实收金额 * rsTemp!self_rate / 100
                End If
            End If
        End If
        
        gstrSQL = "Select * From 医保病人状态 Where to_char(日期,'yyyy-mm-dd')='" & Format(rs明细!发生时间, "yyyy-mm-dd") & "' And 病人ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人医保状态", lng病人ID)
        If rsTemp.EOF Then
            str医保状态 = "0"
        Else
            str医保状态 = Nvl(rsTemp!医保状态, "0")
        End If
        
        If str医保状态 <> "0" Then
            str医保状态 = "1"
            cur自理金额 = rs明细!实收金额
            cur部分自理 = 0
        End If
        
        strSQL = strSQL & "'" & str项目编码 & "'," & Format(rs明细!实收金额 / (rs明细!数次 * rs明细!付数), "0.0000") & _
            "," & rs明细!数次 * rs明细!付数 & ",'" & str单位 & "'," & rs明细!实收金额 & "," & _
            Format(cur自理金额 + cur部分自理, "0.0000") & ",'" & Format(datCurr, "yyyy-mm-dd") & "','" & _
            ToVarchar(str项目名称, 60) & "','" & str通用名编码 & "','" & str医保状态 & "')"
            
        If Nvl(rs明细!是否上传, 0) = 0 Then
            WriteInfo "写入前置机明细:" & strSQL
            gcn徐州.Execute strSQL
        End If
        
        cur部分自理Sum = cur部分自理Sum + cur部分自理
        cur全自理Sum = cur全自理Sum + cur自理金额
        
        gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细("ID") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
        strSQL = "Insert Into 县医保明细 Values ('" & ToVarchar(str项目名称, 50) & "','" & str单位 & "'," & _
            Format(rs明细!实收金额 / (rs明细!数次 * rs明细!付数), "0.0000") & "," & rs明细!数次 * rs明细!付数 & _
            "," & rs明细!实收金额 & "," & Format(cur自理金额 + cur部分自理, "0.0000") & "," & lng病人ID & "," & lng主页ID & _
            ",to_date('" & Format(rs明细!发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')," & rs明细("ID") & "," & str医保状态 & ")"
        If Nvl(rs明细!是否上传, 0) = 0 Then
            WriteInfo "写入县医保明细:" & strSQL
            gcnOracle.Execute strSQL
        End If
        rs明细.MoveNext
    Wend
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
'    cur自理金额 = CCur(rsTemp!退休证号): cur部分自理 = CCur(rsTemp!顺序号)
    cur自理金额 = 0: cur部分自理 = 0
    cur部分自理Sum = cur部分自理Sum + cur部分自理
    cur全自理Sum = cur全自理Sum + cur自理金额
    
    WriteInfo "保存病人已传输费用:部分自理费用(" & cur部分自理Sum & ")  全自理费用(" & cur全自理Sum & ")"
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_徐州 & ",'顺序号','''" & cur部分自理Sum & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_徐州")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_徐州 & ",'退休证号','''" & cur全自理Sum & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "身份标识_徐州")
    WriteInfo "完成明细传递"
    记帐传输_徐州 = True
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_徐州(rs明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
    Dim rsTemp As New ADODB.Recordset, lng主页ID As Long, cur全自理 As Currency, cur部分自理 As Currency, _
        cur总额 As Currency, int跨年标志 As Integer, strReturn As String, strPara As String, _
        strTransNO As String, lng草药 As Long, datCurr As Date, cur单味草药 As Currency, _
        rs徐州 As New ADODB.Recordset, strTemp As String, str类别 As String, strSQL As String, _
        rs费用明细 As New ADODB.Recordset, i As Long
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    If 记帐传输_徐州("", 0, "", lng病人ID) = False Then
        Exit Function
    End If
    
    WriteInfo vbCrLf & "开始住院预结算"
    
    gstrSQL = "Select max(主页ID) From 住院费用记录 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    
    gstrSQL = "Select NO From 住院费用记录 Where 门诊标志=2 And 病人id=[1] And 主页ID=[2] Group By NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    
    mcur全自理 = 0: mcur部分自理 = 0: mcur总额 = 0
    
    While Not rsTemp.EOF
        gstrSQL = "Select A.类别,B.实收金额,A.ID,B.NO From 收费细目 A,住院费用记录 B Where A.ID=B.收费细目ID And " & _
            "B.门诊标志=2 And B.NO='" & rsTemp!NO & "' And B.病人ID=[1] And B.主页ID=[2] And nvl(实收金额,0)<>0"
        Set rs费用明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
        
        lng草药 = 0
        '计算单味草药金额
        While Not rs费用明细.EOF
            If rs费用明细!类别 = "7" Then lng草药 = lng草药 + 1
            rs费用明细.MoveNext
        Wend
        rs费用明细.MoveFirst
        
        While Not rs费用明细.EOF
            gstrSQL = "Select * From 保险支付项目 Where 收费细目ID=[1]"
            Set rs徐州 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs费用明细!ID))
            If rs徐州.EOF Then
                mcur全自理 = mcur全自理 + rs费用明细!实收金额
            Else
                str类别 = rs费用明细!类别
                If str类别 = "5" Or str类别 = "6" Or str类别 = "7" Then
                    gstrSQL = "Select * From mi_drug_trade_list Where trade_code='" & rs徐州!项目编码 & "'"
                    Set rs徐州 = gcn徐州.Execute(gstrSQL)
                    If rs徐州.EOF Then
                        mcur全自理 = mcur全自理 + rs费用明细!实收金额
                    ElseIf rs徐州!mi_class = "1" Then
                        
                    ElseIf rs徐州!mi_class = "2" Then
                        mcur部分自理 = mcur部分自理 + rs费用明细!实收金额 * 0.2
                    ElseIf rs徐州!mi_class = "4" Then
                        If lng草药 < 2 Then
                            mcur全自理 = mcur全自理 + rs费用明细!实收金额
                        End If
                    Else
                        mcur全自理 = mcur全自理 + rs费用明细!实收金额
                    End If
                ElseIf str类别 = "J" Then
                    mcur部分自理 = mcur部分自理 + IIf(rs费用明细!实收金额 > 15, rs费用明细!实收金额 - 15, 0)
                Else
                    gstrSQL = "Select * From mi_dt_item Where item_code='" & rs徐州!项目编码 & "'"
                    Set rs徐州 = gcn徐州.Execute(gstrSQL)
                    If rs徐州.EOF Then
                        mcur全自理 = mcur全自理 + rs费用明细!实收金额
                    Else
                        mcur部分自理 = mcur部分自理 + rs费用明细!实收金额 * rs徐州!self_rate
                    End If
                End If
            End If
            mcur总额 = mcur总额 + rs费用明细!实收金额
            rs费用明细.MoveNext
        Wend
        
        rsTemp.MoveNext
    Wend
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_徐州)
    mcur全自理 = CCur(Nvl(rsTemp!退休证号, "0"))
    mcur部分自理 = CCur(Nvl(rsTemp!顺序号, "0"))
    
    WriteInfo "计算自理费用:总额(" & mcur总额 & ")  全自理(" & mcur全自理 & ")  部分自理(" & mcur部分自理 & ")"
    
    gstrSQL = "Select 入院日期 From 病案主页 Where 病人id=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If Format(rsTemp(0), "yyyy") <> Format(datCurr, "yyyy") Then
        int跨年标志 = 1
    Else
        int跨年标志 = 0
    End If
    
    If mcur总额 = 0 Then
        mcur部分自理 = 0: mcur全自理 = 0
    End If
    
    strTransNO = MakeTransNO()
    '医院号，医保号，住院总费用，部分自理部分费用，完全自理部分费用，入院跨年度
    strPara = Trim(gstr医院编码) & "," & str医保号 & "," & mcur总额 & "," & mcur部分自理 & "," & mcur全自理 & "," & int跨年标志
    
    WriteInfo "发送请求:流水号---" & strTransNO
    WriteInfo "　　　　　 参数---" & strPara
    
    strSQL = "Insert Into ins_tranask (transerial,trantype,hdcode,parm,tranflag) Values ('" & _
        strTransNO & "','22','" & UserInfo.编号 & "','" & strPara & "','9')"
    gcn徐州.Execute strSQL
    If frm等待响应徐州.Result(strTransNO, strReturn) = False Then
        WriteInfo "交易中止"
        MsgBox "请求被中止", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(Trim(strReturn), 2) = "01" Then
        MsgBox "医保交易失败", vbInformation, gstrSysName
        Exit Function
    End If
    
    WriteInfo "请求返回:" & strReturn
    
    '0-成功标志，1-部分自理部分费用，2-完全自理部分费用，3-起付线支付，4-统筹支付，5-统筹不支付，6-大病支付
    '7-大病不支付，8-公务员自理费用，9-公务员基金支付
    mcur个帐支付 = CCur(Split(strReturn, ",")(6))
    mcur统筹支付 = CCur(Split(strReturn, ",")(4))
    mcur公务员 = CCur(Split(strReturn, ",")(9))
'    mstr结帐费用 = Right(Space(15) & mcur总额, 15)
    'mstr结帐费用 = mcur总额 & Mid(strReturn, InStr(strReturn, ","))
    mstr结帐费用 = mcur总额
    strSQL = "Delete 县医保结帐 Where 病人id=" & lng病人ID & " And 主页id=" & lng主页ID
    gcnOracle.Execute strSQL
    
    strSQL = "Insert Into 县医保结帐 Values (" & lng病人ID & "," & lng主页ID & "," & Format(mcur总额, "0.0000")
    For i = 1 To 9
'        mstr结帐费用 = mstr结帐费用 & Right(Space(15) & Split(strReturn, ",")(i), 15)
        strSQL = strSQL & "," & Format(Split(strReturn, ",")(i), "0.0000")
        mstr结帐费用 = mstr结帐费用 & "," & Format(Split(strReturn, ",")(i), "0.0000")
    Next
    strSQL = strSQL & ")"
    WriteInfo "写入县医保结算数据:" & strSQL
    gcnOracle.Execute strSQL
    
    If mcur个帐支付 <> 0 Then 住院虚拟结算_徐州 = "大病支付;" & mcur个帐支付 & ";0"
    住院虚拟结算_徐州 = 住院虚拟结算_徐州 & IIf(住院虚拟结算_徐州 <> "", "|", "") & "统筹基金;" & mcur统筹支付 & ";0"
    If mcur公务员 <> 0 Then 住院虚拟结算_徐州 = 住院虚拟结算_徐州 & IIf(住院虚拟结算_徐州 <> "", "|", "") & "公务员基金;" & mcur公务员 & ";0"
    
    WriteInfo "完成预结算:" & 住院虚拟结算_徐州
    Exit Function
errHandle:
    WriteInfo "发生错误:" & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_徐州(lng结帐ID As Long, lng病人ID As Long) As Boolean
    Dim datCurr As Date, int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, _
        cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, cur余额 As String, rsTemp As New ADODB.Recordset, _
        str出院请求 As String
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    cur余额 = 个人余额_徐州(lng病人ID)
    
    gstrSQL = "Select max(主页id) from 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    If 医保出院(lng病人ID, rsTemp(0), str出院请求) = False Then Exit Function
    str出院请求 = Replace(str出院请求, "'", "''")
    '帐户年度信息
    Call Get帐户信息(TYPE_徐州, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_徐州 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 + mcur统筹支付 + mcur公务员 + mcur个帐支付 & _
        "," & cur统筹报销累计 + mcur统筹支付 + mcur公务员 + mcur个帐支付 & "," & int住院次数累计 & ",null,null,null,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_徐州 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur余额 & "," & cur帐户支出累计 & "," & _
        cur进入统筹累计 + mcur统筹支付 + mcur公务员 + mcur个帐支付 & "," & cur统筹报销累计 + mcur统筹支付 + mcur公务员 + mcur个帐支付 & _
        "," & int住院次数累计 & ",NULL,NULL,NULL," & mcur总额 & "," & mcur全自理 & "," & mcur部分自理 & _
        ",NULL," & mcur统筹支付 + mcur公务员 + mcur个帐支付 & ",NULL,NULL,NULL,NULL,NULL,NULL,'" & str出院请求 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "徐州医保")
    '改写IC卡状态
    If WriteCard(0) = False Then
        Err.Raise 9000, gstrSysName, "写卡状态时失败，但病人出院的操作不受影响"
    End If
     
    住院结算_徐州 = True
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    WriteInfo "发生错误:" & Err.Description
End Function

Public Function 住院结算冲销_徐州(lng结帐ID As Long) As Boolean
'功能:将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数:lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
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
    
    '退费
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "结算冲销", lng冲销ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        住院结算冲销_徐州 = False
        Exit Function
    End If
    
    cur个人帐户 = Nvl(rsTemp!个人帐户支付, 0)
    strTemp = Nvl(rsTemp!备注, "")
    
'    gstrSQL = "Select * From 保险帐户 Where 病人id=" & rsTemp!病人ID & " And 险类=" & TYPE_徐州
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If Nvl(rsTemp!当前状态, 0) = 0 Then
'        MsgBox "住院结算冲销前请先撤消病人出院"
'        Exit Function
'    End If
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州, lng结帐ID)
    lng病人ID = rsTemp!病人ID
    
    gstrSQL = "Select max(主页id) from 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = Nvl(rsTemp(0), 1)
    If 撤消医保出院_徐州(lng病人ID, lng主页ID) = False Then
        Exit Function
    End If
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_徐州, lng结帐ID)
    
    '帐户年度信息
    Call Get帐户信息(TYPE_徐州, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & TYPE_徐州 & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_徐州 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 - 1 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '改写IC卡状态
    If WriteCard(1) = False Then
        MsgBox "写卡状态时失败，但撤消出院的操作不受影响", vbInformation, "写卡"
    End If

    住院结算冲销_徐州 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "发生错误:" & Err.Description
    住院结算冲销_徐州 = False
End Function

